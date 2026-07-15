import os
import re
import subprocess
import time
from pathlib import Path
from typing import Dict, List, Optional, Tuple

from netmiko import ConnectHandler


PRECHECK_COMMANDS = [
    'show running-config',
    'show interfaces status',
    'show cdp neighbors',
    'show vlan',
    'show etherchannel summary',
]


def open_ssh_connection(ip_address: str):
    username = os.getenv('DNA_USERNAME')
    password = os.getenv('DNA_PASSWORD')
    if not username or not password:
        raise ValueError('Missing DNA_USERNAME or DNA_PASSWORD')

    device_type = os.getenv('SSH_DEVICE_TYPE', 'cisco_ios')
    timeout_raw = os.getenv('SSH_TIMEOUT_SECONDS', '12')
    try:
        timeout = max(5, int(timeout_raw))
    except ValueError:
        timeout = 12

    return ConnectHandler(
        device_type=device_type,
        host=ip_address,
        username=username,
        password=password,
        timeout=timeout,
        conn_timeout=timeout,
        banner_timeout=timeout,
        auth_timeout=timeout,
        fast_cli=False,
    )


def _parse_show_version_image(show_version_output: str) -> Optional[str]:
    patterns = [
        r'Cisco IOS XE Software, Version\s+([^,\s]+)',
        r'Cisco IOS Software.*Version\s+([^,\s]+)',
        r'System image file is\s+"([^"]+)"',
    ]
    for pat in patterns:
        m = re.search(pat, show_version_output, flags=re.IGNORECASE)
        if m:
            return m.group(1).strip()
    return None


def parse_show_version_details(show_version_output: str) -> Dict[str, str]:
    platform_patterns = [
        r'(?im)^\s*cisco\s+([A-Za-z0-9-]+)\s+\(',
        r'(?im)^\s*Model number\s*:\s*([A-Za-z0-9-]+)',
        r'(?im)^\s*Platform\s*:\s*([A-Za-z0-9-]+)',
        r'(?im)^\s*Processor board ID.*?\n\s*cisco\s+([A-Za-z0-9-]+)\s+\(',
    ]
    version_patterns = [
        r'Cisco IOS XE Software, Version\s+([^,\s]+)',
        r'Cisco IOS Software.*?Version\s+([^,\s]+)',
        r'(?im)^\s*Version\s+([^,\s]+)',
    ]
    image_patterns = [
        r'System image file is\s+"([^"]+)"',
        r'System image file is\s+([^\s]+)',
    ]

    platform_id = 'Unknown'
    for pattern in platform_patterns:
        match = re.search(pattern, show_version_output, flags=re.IGNORECASE | re.MULTILINE | re.DOTALL)
        if match:
            platform_id = match.group(1).strip()
            break

    software_version = 'Unknown'
    for pattern in version_patterns:
        match = re.search(pattern, show_version_output, flags=re.IGNORECASE | re.MULTILINE | re.DOTALL)
        if match:
            software_version = match.group(1).strip().rstrip(',')
            break

    system_image = 'Unknown'
    for pattern in image_patterns:
        match = re.search(pattern, show_version_output, flags=re.IGNORECASE | re.MULTILINE | re.DOTALL)
        if match:
            system_image = match.group(1).strip().rstrip(',')
            break

    return {
        'platform_id': platform_id,
        'software_version': software_version,
        'system_image': system_image,
        'raw_output': show_version_output,
    }


def parse_show_switch_members(show_switch_output: str) -> List[int]:
    members: List[int] = []
    for line in show_switch_output.splitlines():
        # Common formats:
        #  *1       Master ... Ready
        #   2       Member ... Ready
        # and some variants where state appears immediately after switch id.
        match = re.match(
            r'^\s*\*?(\d+)\s+(?:Master|Member|Active|Standby|Ready|Provisioned|Removed|Disabled|Init|Down|Unknown)\b',
            line,
            flags=re.IGNORECASE,
        )
        if match:
            member = int(match.group(1))
            if member not in members:
                members.append(member)
    return members or [1]


def run_timed_command(conn, command: str, prompt_responses: Optional[List[str]] = None) -> str:
    output = conn.send_command_timing(command, strip_prompt=False, strip_command=False)
    for response in list(prompt_responses or []):
        lowered = output.lower()
        if any(token in lowered for token in ('destination filename', 'confirm', '[confirm]', '[y/n]')):
            output += '\n' + conn.send_command_timing(response, strip_prompt=False, strip_command=False)
        else:
            break
    return output


def ssh_show_version(ip_address: str) -> Tuple[Optional[str], Optional[str]]:
    """Return (parsed_version_or_image, error_message).

    Uses DNAC credentials by default:
    - DNA_USERNAME
    - DNA_PASSWORD

    Optional env settings:
    - SSH_DEVICE_TYPE (default: cisco_ios)
    - SSH_TIMEOUT_SECONDS (default: 12)
    """
    username = os.getenv('DNA_USERNAME')
    password = os.getenv('DNA_PASSWORD')
    if not username or not password:
        return None, 'Missing DNA_USERNAME or DNA_PASSWORD'

    device_type = os.getenv('SSH_DEVICE_TYPE', 'cisco_ios')
    timeout_raw = os.getenv('SSH_TIMEOUT_SECONDS', '12')
    try:
        timeout = max(5, int(timeout_raw))
    except ValueError:
        timeout = 12

    conn = None
    try:
        conn = ConnectHandler(
            device_type=device_type,
            host=ip_address,
            username=username,
            password=password,
            timeout=timeout,
            conn_timeout=timeout,
            banner_timeout=timeout,
            auth_timeout=timeout,
            fast_cli=False,
        )
        output = conn.send_command('show version', read_timeout=max(20, timeout))
        parsed = _parse_show_version_image(output)
        if parsed:
            return parsed, None
        return None, 'Unable to parse image/version from show version output'
    except Exception as exc:
        return None, str(exc)
    finally:
        if conn is not None:
            try:
                conn.disconnect()
            except Exception:
                pass


def _build_connection(ip_address: str):
    username = os.getenv('DNA_USERNAME')
    password = os.getenv('DNA_PASSWORD')
    if not username or not password:
        raise ValueError('Missing DNA_USERNAME or DNA_PASSWORD')

    device_type = os.getenv('SSH_DEVICE_TYPE', 'cisco_ios')
    timeout_raw = os.getenv('SSH_TIMEOUT_SECONDS', '12')
    try:
        timeout = max(5, int(timeout_raw))
    except ValueError:
        timeout = 12

    return ConnectHandler(
        device_type=device_type,
        host=ip_address,
        username=username,
        password=password,
        timeout=timeout,
        conn_timeout=timeout,
        banner_timeout=timeout,
        auth_timeout=timeout,
        fast_cli=False,
    )


def _sanitize_filename(value: str) -> str:
    text = (value or 'unknown').strip()
    text = re.sub(r'[^A-Za-z0-9._-]+', '_', text)
    return text[:180] or 'unknown'


def collect_device_precheck_outputs(
    *,
    ip_address: str,
    device_name: str,
    site_code: str,
    base_directory: str = r'C:\Device_Files_Patching',
    file_suffix: str = '',
    commands: Optional[List[str]] = None,
) -> Dict[str, str]:
    command_list = commands or PRECHECK_COMMANDS
    target_dir = Path(base_directory) / _sanitize_filename(site_code)
    target_dir.mkdir(parents=True, exist_ok=True)
    suffix = _sanitize_filename(file_suffix).strip('_') if file_suffix else ''
    base_name = _sanitize_filename(device_name or ip_address)
    if suffix:
        file_name = f"{base_name}_{suffix}.txt"
    else:
        file_name = f"{base_name}.txt"
    file_path = target_dir / file_name

    conn = None
    try:
        conn = _build_connection(ip_address)
        sections: List[str] = []
        for command in command_list:
            try:
                output = conn.send_command(command, read_timeout=60)
            except Exception as exc:
                output = f'ERROR: {exc}'
            sections.append(f'===== {command} =====\n{output}'.rstrip())

        file_path.write_text('\n\n'.join(sections) + '\n', encoding='utf-8')
        return {
            'device_name': device_name,
            'ip_address': ip_address,
            'site_code': site_code,
            'file_path': str(file_path),
            'status': 'Success',
        }
    except Exception as exc:
        return {
            'device_name': device_name,
            'ip_address': ip_address,
            'site_code': site_code,
            'file_path': str(file_path),
            'status': f'Failed: {exc}',
        }
    finally:
        if conn is not None:
            try:
                conn.disconnect()
            except Exception:
                pass


def icmp_ping(
    ip_address: str,
    timeout_ms: int = 1500,
    attempts: int = 4,
    retry_delay_seconds: float = 0.5,
) -> Tuple[bool, str]:
    if not ip_address:
        return False, 'Missing IP address'

    max_attempts = max(1, int(attempts))
    delay = max(0.0, float(retry_delay_seconds))
    outputs: List[str] = []

    for attempt_index in range(max_attempts):
        command = ['ping', '-n', '1', '-w', str(max(250, timeout_ms)), ip_address]
        try:
            result = subprocess.run(command, capture_output=True, text=True, timeout=10, check=False)
            detail = (result.stdout or result.stderr or '').strip()
            outputs.append(f'Attempt {attempt_index + 1}/{max_attempts}: {detail}')
            if result.returncode == 0:
                return True, '\n\n'.join(outputs)
        except Exception as exc:
            outputs.append(f'Attempt {attempt_index + 1}/{max_attempts}: {exc}')

        if attempt_index < max_attempts - 1 and delay > 0:
            time.sleep(delay)

    return False, '\n\n'.join(outputs)
