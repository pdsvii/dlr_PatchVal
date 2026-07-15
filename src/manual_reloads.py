from __future__ import annotations

import io
import os
import re
from datetime import datetime
from pathlib import Path
from typing import Any, Dict, List

import pandas as pd
import streamlit as st
from dotenv import load_dotenv
from zoneinfo import ZoneInfo

try:
    from src.ssh_device_client import (
        icmp_ping,
        open_ssh_connection,
        parse_show_switch_members,
        parse_show_version_details,
        run_timed_command,
    )
except ModuleNotFoundError:
    from ssh_device_client import (
        icmp_ping,
        open_ssh_connection,
        parse_show_switch_members,
        parse_show_version_details,
        run_timed_command,
    )


load_dotenv(override=True)


TARGET_PLATFORM_PREFIX = os.getenv('MANUAL_RELOAD_PLATFORM_PREFIX', '2960').strip() or '2960'
TARGET_IMAGE_VERSION = os.getenv('MANUAL_RELOAD_TARGET_VERSION', '15.2(7)E14').strip() or '15.2(7)E14'
TARGET_IMAGE_NAME = os.getenv('MANUAL_RELOAD_IMAGE_NAME', 'c2960x-universalk9-mz.152-7.E14.bin').strip()
TARGET_IMAGE_FOLDER = os.getenv('MANUAL_RELOAD_IMAGE_FOLDER', 'c2960x-universalk9-mz.152-7.E14').strip()
TARGET_IMAGE_BASE_URL = os.getenv('MANUAL_RELOAD_IMAGE_BASE_URL', 'http://10.3.5.40:8500').strip().rstrip('/')
TARGET_IMAGE_URL = f'{TARGET_IMAGE_BASE_URL}/{TARGET_IMAGE_NAME}'


CANONICAL_FIELDS = [
    'Device Name',
    'Device IP',
    'DC',
    'Date',
    'DNAC EST Time',
    'Platform',
    'Business / Non Business Hours',
    'Assigned to',
    'Clean Up Status',
    'Patching Status',
    'Patching Status 2',
]


def _normalize_header(value: str) -> str:
    return re.sub(r'[^a-z0-9]+', '', (value or '').strip().lower())


def _current_est_text() -> str:
    return datetime.now(ZoneInfo('America/New_York')).strftime('%Y-%m-%d %H:%M:%S %Z')


def _read_imported_dataframe(uploaded_file) -> pd.DataFrame:
    file_name = getattr(uploaded_file, 'name', '') or ''
    suffix = Path(file_name).suffix.lower()
    file_bytes = uploaded_file.getvalue()
    data = io.BytesIO(file_bytes)
    if suffix == '.csv':
        return pd.read_csv(data, dtype=str)
    if suffix in {'.xls', '.xlsx'}:
        engine = 'xlrd' if suffix == '.xls' else 'openpyxl'
        return pd.read_excel(data, dtype=str, engine=engine)
    raise ValueError('Unsupported file type. Please upload a .csv or .xls file.')


def _canonicalize_manual_row(raw_row: Dict[str, Any]) -> Dict[str, str]:
    row: Dict[str, str] = {field: '' for field in CANONICAL_FIELDS}
    patching_slot = 0
    alias_map = {
        'devicename': 'Device Name',
        'device': 'Device Name',
        'hostname': 'Device Name',
        'fqdn': 'Device Name',
        'deviceip': 'Device IP',
        'ip': 'Device IP',
        'ipaddress': 'Device IP',
        'managementip': 'Device IP',
        'managementipaddress': 'Device IP',
        'dc': 'DC',
        'datacenter': 'DC',
        'site': 'DC',
        'date': 'Date',
        'dnacesttime': 'DNAC EST Time',
        'gettheesttimefromthednac': 'DNAC EST Time',
        'esttime': 'DNAC EST Time',
        'platform': 'Platform',
        'businessnonbusinesshours': 'Business / Non Business Hours',
        'businesshours': 'Business / Non Business Hours',
        'assignedto': 'Assigned to',
        'cleanupstatus': 'Clean Up Status',
        'patchingstatus': 'Patching Status',
    }

    for key, value in raw_row.items():
        header = _normalize_header(str(key))
        text = '' if value is None else str(value).strip()
        if not header:
            continue
        if header.startswith('patchingstatus'):
            if patching_slot == 0:
                row['Patching Status'] = text
                patching_slot = 1
            else:
                row['Patching Status 2'] = text
            continue
        field = alias_map.get(header)
        if field:
            row[field] = text

    if not row['DNAC EST Time']:
        row['DNAC EST Time'] = _current_est_text()

    return row


def import_manual_reload_rows(uploaded_file) -> List[Dict[str, str]]:
    frame = _read_imported_dataframe(uploaded_file)
    if frame.empty:
        return []
    frame = frame.fillna('')
    return [_canonicalize_manual_row(record) for record in frame.to_dict(orient='records')]


def _profile() -> Dict[str, str]:
    return {
        'platform_prefix': TARGET_PLATFORM_PREFIX,
        'target_version': TARGET_IMAGE_VERSION,
        'image_name': TARGET_IMAGE_NAME,
        'image_folder': TARGET_IMAGE_FOLDER,
        'image_url': TARGET_IMAGE_URL,
        'image_base_url': TARGET_IMAGE_BASE_URL,
    }


def _platform_matches(imported_platform: str, observed_platform: str, platform_prefix: str) -> bool:
    imported = (imported_platform or '').strip().lower()
    observed = (observed_platform or '').strip().lower()
    prefix = (platform_prefix or '').strip().lower()
    if prefix and prefix in observed:
        return True
    if imported and imported in observed:
        return True
    if observed and observed in imported:
        return True
    return False


def _version_matches(current_version: str, target_version: str) -> bool:
    current = (current_version or '').strip().lower()
    target = (target_version or '').strip().lower()
    return bool(current and target and current == target)


def _flash_targets(member_ids: List[int]) -> List[str]:
    members = sorted({int(member) for member in member_ids if int(member) > 0})
    if len(members) <= 1:
        return ['flash:/']
    return [f'flash{member}:/' for member in members]


def _boot_target_from_dir_output(flash_root: str, directory_output: str, profile: Dict[str, str]) -> str:
    image_name = profile['image_name']
    image_folder = profile['image_folder']
    if image_folder and image_name and f'{image_folder}/{image_name}' in directory_output:
        return f'{flash_root}{image_folder}/{image_name}'
    if image_name and image_name in directory_output:
        return f'{flash_root}{image_name}'
    if image_folder and image_folder in directory_output:
        return f'{flash_root}{image_folder}/{image_name}'
    return f'{flash_root}{image_name}'


def _has_image(directory_output: str, profile: Dict[str, str]) -> bool:
    return profile['image_name'] in directory_output or profile['image_folder'] in directory_output


def _cleanup_candidates(directory_output: str, profile: Dict[str, str]) -> List[tuple[str, bool]]:
    target_names = {profile['image_name'], profile['image_folder']}
    image_prefix = (profile['image_name'].split('-')[0] if profile['image_name'] else '').strip().lower()
    candidates: List[tuple[str, bool]] = []

    for line in directory_output.splitlines():
        tokens = line.strip().split()
        if not tokens:
            continue
        name = tokens[-1].strip()
        lowered = name.lower()
        if not image_prefix or not lowered.startswith(image_prefix):
            continue
        if name in target_names:
            continue
        is_directory = 'dir' in line.lower() or 'drwx' in line.lower()
        is_image_file = lowered.endswith('.bin') or is_directory
        if not is_image_file:
            continue
        candidate = (name, is_directory)
        if candidate not in candidates:
            candidates.append(candidate)

    return candidates


def _cleanup_old_images(conn, flash_targets: List[str], profile: Dict[str, str], logs: List[str]) -> str:
    deleted_any = False
    for flash_root in flash_targets:
        dir_output = conn.send_command(f'dir {flash_root}', read_timeout=60)
        for name, is_directory in _cleanup_candidates(dir_output, profile):
            delete_command = f'delete /force /recursive {flash_root}{name}' if is_directory else f'delete /force {flash_root}{name}'
            delete_output = run_timed_command(conn, delete_command, prompt_responses=[''])
            logs.append(f'cleanup {flash_root}{name}: {delete_output[-1200:]}')
            deleted_any = True
    return 'Completed' if deleted_any else 'No old images found'


def _validate_boot_statement(conn, boot_command: str, logs: List[str]) -> str:
    try:
        running_boot = conn.send_command('show running-config | include ^boot system', read_timeout=60)
        logs.append(f'boot validation: {running_boot[-1200:]}')
        expected = boot_command.strip().lower()
        lines = [line.strip().lower() for line in running_boot.splitlines() if line.strip()]
        if expected in lines:
            return 'Matched running-config'
        return 'Missing from running-config'
    except Exception as exc:
        logs.append(f'boot validation error: {exc}')
        return f'Validation failed: {exc}'


def _postcheck_commands(platform_text: str, flash_targets: List[str]) -> List[str]:
    platform = (platform_text or '').upper()
    commands = ['show running-config | include ^boot system']

    if any(token in platform for token in ('2960', '9200', '9300', '9500', '9410', 'IE-3')):
        commands.extend(['show version', 'show boot', 'show switch'])
    elif any(token in platform for token in ('ISR43', 'ISR44', 'C8300')):
        commands.extend(['show version', 'show bootvar'])
    else:
        commands.append('show version')

    for flash_root in flash_targets:
        commands.append(f'dir {flash_root}')

    deduped: List[str] = []
    for command in commands:
        if command not in deduped:
            deduped.append(command)
    return deduped


def _run_postchecks(conn, platform_text: str, flash_targets: List[str], logs: List[str]) -> Dict[str, str]:
    commands = _postcheck_commands(platform_text, flash_targets)
    outputs: List[str] = []
    for command in commands:
        try:
            output = conn.send_command(command, read_timeout=60)
            outputs.append(f'===== {command} =====\n{output}'.rstrip())
        except Exception as exc:
            outputs.append(f'===== {command} =====\nERROR: {exc}')
    logs.append('post-checks completed')
    return {
        'Post Check Status': 'Completed',
        'Post Check Commands': '; '.join(commands),
        'Post Check Output': '\n\n'.join(outputs),
    }


def _classify_ssh_login_error(exc: Exception) -> tuple[str, str]:
    message = str(exc).strip()
    lowered = message.lower()
    if any(token in lowered for token in ('authentication failed', 'auth failed', 'invalid password', 'permission denied')):
        return 'Authentication failed', message or 'SSH authentication failed'
    if any(token in lowered for token in ('timed out', 'timeout', 'no existing session')):
        return 'Connection timeout', message or 'SSH connection timed out'
    if any(token in lowered for token in ('refused', 'unable to connect', 'unreachable', 'no route to host')):
        return 'Connection refused/unreachable', message or 'SSH connection could not be established'
    if any(token in lowered for token in ('error reading ssh protocol banner', 'ssh protocol banner')):
        return 'SSH banner/protocol error', message or 'SSH protocol/banner negotiation failed'
    return 'Unknown SSH login error', message or 'Unknown SSH login error'


def _run_single_row(
    row: Dict[str, str],
    profile: Dict[str, str],
    execute_actions: bool,
    cleanup_old_images: bool = False,
    execute_reload: bool = True,
) -> Dict[str, str]:
    result = dict(row)
    logs: List[str] = []
    ip = (row.get('Device IP') or '').strip()
    imported_platform = (row.get('Platform') or '').strip()
    result['SSH Login Status'] = 'Not attempted'
    result['SSH Error Category'] = ''

    if not ip:
        result['Ping Status'] = 'Missing IP'
        result['ICMP Ping Status'] = 'Missing IP'
        result['SSH Login Status'] = 'Skipped'
        result['Workflow Status'] = 'Skipped'
        result['Workflow Notes'] = 'No device IP provided'
        return result

    is_online, ping_output = icmp_ping(ip)
    result['Ping Status'] = 'Online' if is_online else 'Offline'
    result['ICMP Ping Status'] = result['Ping Status']
    result['Ping Output'] = ping_output
    if not is_online:
        result['SSH Login Status'] = 'Skipped'
        result['Workflow Status'] = 'Skipped'
        result['Workflow Notes'] = 'Device did not respond to ping'
        return result

    conn = None
    try:
        try:
            conn = open_ssh_connection(ip)
            result['SSH Login Status'] = 'Success'
        except Exception as exc:
            category, detail = _classify_ssh_login_error(exc)
            result['SSH Login Status'] = 'Failed'
            result['SSH Error Category'] = category
            result['Workflow Status'] = 'SSH login failed'
            result['Workflow Notes'] = detail
            result['Workflow Log'] = '\n'.join(logs)
            return result

        show_version_output = conn.send_command('show version', read_timeout=60)
        details = parse_show_version_details(show_version_output)
        result['Observed Platform'] = details['platform_id']
        result['Observed Version'] = details['software_version']
        result['System Image'] = details['system_image']
        result['Post Check Status'] = ''
        result['Post Check Commands'] = ''
        result['Post Check Output'] = ''
        result['Platform Match'] = 'Yes' if _platform_matches(imported_platform, details['platform_id'], profile['platform_prefix']) else 'No'
        result['Version Match'] = 'Yes' if _version_matches(details['software_version'], profile['target_version']) else 'No'

        switch_output = conn.send_command('show switch', read_timeout=60)
        member_ids = parse_show_switch_members(switch_output)
        result['Switch Count'] = str(len(member_ids))

        flash_targets = _flash_targets(member_ids)
        boot_target = ''
        image_present = False
        for flash_root in flash_targets:
            dir_output = conn.send_command(f'dir {flash_root}', read_timeout=60)
            logs.append(f'{flash_root}: {"image found" if _has_image(dir_output, profile) else "image missing"}')
            if not boot_target and _has_image(dir_output, profile):
                boot_target = _boot_target_from_dir_output(flash_root, dir_output, profile)
                image_present = True

        if not boot_target:
            boot_target = _boot_target_from_dir_output(flash_targets[0], '', profile)

        result['Boot Target'] = boot_target
        result['Image Present'] = 'Yes' if image_present else 'No'

        if not result['Platform Match'] == 'Yes':
            result['Needs Upgrade'] = 'No'
            result['Workflow Status'] = 'Unsupported platform'
            result['Workflow Notes'] = f'Imported platform {imported_platform or "Unknown"} does not match target prefix {profile["platform_prefix"]}'
            return result

        needs_upgrade = not _version_matches(details['software_version'], profile['target_version'])
        result['Needs Upgrade'] = 'Yes' if needs_upgrade else 'No'

        if not needs_upgrade:
            result['Workflow Status'] = 'No action required'
            result['Workflow Notes'] = 'Target version already installed'
            return result

        if execute_actions and not image_present:
            for flash_root in flash_targets:
                copy_output = run_timed_command(conn, f'copy {profile["image_url"]} {flash_root}', prompt_responses=['', ''])
                logs.append(f'copy {flash_root}: {copy_output[-2500:]}')
            boot_target = _boot_target_from_dir_output(flash_targets[0], f'{profile["image_name"]}\n{profile["image_folder"]}', profile)
            result['Boot Target'] = boot_target
            result['Image Present'] = 'Yes'

        boot_command = f'boot system switch all {result["Boot Target"]}'
        result['Boot Command'] = boot_command

        if execute_actions:
            conn.send_config_set([boot_command])
            if cleanup_old_images:
                result['Clean Up Status'] = _cleanup_old_images(conn, flash_targets, profile, logs)
            else:
                result['Clean Up Status'] = 'Skipped'
            conn.send_command_timing('end', strip_prompt=False, strip_command=False)
            try:
                write_output = run_timed_command(conn, 'wr', prompt_responses=[''])
                logs.append(f'wr: {write_output[-1200:]}')
            except Exception as exc:
                logs.append(f'wr error: {exc}')
            result['Boot Validation Status'] = _validate_boot_statement(conn, boot_command, logs)
            result.update(_run_postchecks(conn, details['platform_id'] or imported_platform, flash_targets, logs))
            if execute_reload:
                try:
                    reload_output = run_timed_command(conn, 'reload', prompt_responses=[''])
                    logs.append(f'reload: {reload_output[-1200:]}')
                except Exception as exc:
                    logs.append(f'reload error: {exc}')
                result['Workflow Status'] = 'Reload sent'
                result['Workflow Notes'] = 'Reload command issued; verify that the device returns online'
                post_reload_ping, post_reload_output = icmp_ping(ip, attempts=3, timeout_ms=1000, retry_delay_seconds=2)
                result['Post Reload Ping'] = 'Online' if post_reload_ping else 'Offline'
                result['Post Reload Ping Output'] = post_reload_output
                if post_reload_ping:
                    try:
                        verify_conn = open_ssh_connection(ip)
                        try:
                            verify_output = verify_conn.send_command('show version', read_timeout=60)
                            verify_details = parse_show_version_details(verify_output)
                            result['Post Reload Version'] = verify_details['software_version']
                            result['Post Reload Platform'] = verify_details['platform_id']
                        finally:
                            try:
                                verify_conn.disconnect()
                            except Exception:
                                pass
                    except Exception as exc:
                        result['Workflow Notes'] = f'Post reload verification failed: {exc}'
            else:
                result['Workflow Status'] = 'Staged without reload'
                result['Workflow Notes'] = 'Image checked/downloaded, boot statement saved and validated, reload intentionally skipped for this phase'
        else:
            result['Workflow Status'] = 'Ready to execute'
            result['Workflow Notes'] = 'Review target image and boot statement, then enable execution'

        result['Workflow Log'] = '\n'.join(logs)
        return result

    except Exception as exc:
        if result.get('SSH Login Status') == 'Not attempted':
            category, _ = _classify_ssh_login_error(exc)
            result['SSH Login Status'] = 'Failed'
            result['SSH Error Category'] = category
        result['Workflow Status'] = 'Error'
        result['Workflow Notes'] = str(exc)
        result['Workflow Log'] = '\n'.join(logs)
        return result
    finally:
        if conn is not None:
            try:
                conn.disconnect()
            except Exception:
                pass


def _display_rows(rows: List[Dict[str, str]]) -> List[Dict[str, str]]:
    display: List[Dict[str, str]] = []
    for row in rows:
        display.append(
            {
                'Device Name': row.get('Device Name', ''),
                'Device IP': row.get('Device IP', ''),
                'DC': row.get('DC', ''),
                'Date': row.get('Date', ''),
                'DNAC EST Time': row.get('DNAC EST Time', ''),
                'Platform': row.get('Platform', ''),
                'Observed Platform': row.get('Observed Platform', ''),
                'Observed Version': row.get('Observed Version', ''),
                'Business / Non Business Hours': row.get('Business / Non Business Hours', ''),
                'Assigned to': row.get('Assigned to', ''),
                'Clean Up Status': row.get('Clean Up Status', ''),
                'Patching Status': row.get('Patching Status', ''),
                'Patching Status 2': row.get('Patching Status 2', ''),
                'Ping Status': row.get('Ping Status', ''),
                'SSH Login Status': row.get('SSH Login Status', ''),
                'SSH Error Category': row.get('SSH Error Category', ''),
                'Switch Count': row.get('Switch Count', ''),
                'Image Present': row.get('Image Present', ''),
                'Needs Upgrade': row.get('Needs Upgrade', ''),
                'Boot Target': row.get('Boot Target', ''),
                'Boot Command': row.get('Boot Command', ''),
                'Boot Validation Status': row.get('Boot Validation Status', ''),
                'Post Check Status': row.get('Post Check Status', ''),
                'Post Check Commands': row.get('Post Check Commands', ''),
                'Workflow Status': row.get('Workflow Status', ''),
                'Workflow Notes': row.get('Workflow Notes', ''),
            }
        )
    return display


def _template_rows() -> List[Dict[str, str]]:
    return [{field: '' for field in CANONICAL_FIELDS}]


def manual_reload_profile() -> Dict[str, str]:
    return _profile()


def build_manual_reload_profile(
    *,
    image_name: str | None = None,
    image_base_url: str | None = None,
    target_version: str | None = None,
    platform_prefix: str | None = None,
) -> Dict[str, str]:
    profile = _profile()
    if image_base_url:
        profile['image_base_url'] = image_base_url.strip().rstrip('/')
    if image_name:
        profile['image_name'] = image_name.strip()
        profile['image_folder'] = profile['image_name'].removesuffix('.bin')
    if target_version:
        profile['target_version'] = target_version.strip()
    if platform_prefix:
        profile['platform_prefix'] = platform_prefix.strip()
    profile['image_url'] = f"{profile['image_base_url']}/{profile['image_name']}"
    return profile


def run_manual_reload_workflow(
    rows: List[Dict[str, str]],
    *,
    execute_actions: bool = False,
    cleanup_old_images: bool = False,
    execute_reload: bool = True,
    profile: Dict[str, str] | None = None,
) -> List[Dict[str, str]]:
    active_profile = profile or _profile()
    return [_run_single_row(row, active_profile, execute_actions, cleanup_old_images, execute_reload) for row in rows]


def display_manual_reload_rows(rows: List[Dict[str, str]]) -> List[Dict[str, str]]:
    return _display_rows(rows)


def render_manual_reloads_panel() -> None:
    profile = _profile()
    st.divider()
    manual_tab, = st.tabs(['Manual Reloads'])

    with manual_tab:
        st.subheader('Manual Reloads')
        st.caption('Import a CSV or XLS file, analyze device reachability and image state, then stage actions before an optional reload.')
        st.caption(f'Target: platform {profile["platform_prefix"]}, image {profile["image_name"]}, version {profile["target_version"]}')

        template_bytes = pd.DataFrame(_template_rows()).to_csv(index=False).encode('utf-8')
        st.download_button('Download CSV Template', data=template_bytes, file_name='manual_reload_template.csv', mime='text/csv', use_container_width=True)

        upload = st.file_uploader('Manual reload importer', type=['csv', 'xls', 'xlsx'], key='manual_reload_upload')
        analyze_clicked = st.button('Analyze Import', use_container_width=True)
        execute_actions = st.checkbox('Execute copy, boot statement, and write memory steps', value=False, key='manual_reload_execute')
        cleanup_old_images = st.checkbox('Clean up old images before reload', value=False, key='manual_reload_cleanup')
        execute_reload = st.checkbox('Send reload after staging', value=False, key='manual_reload_execute_reload')
        execute_clicked = st.button('Run Reload Workflow', use_container_width=True)

        if 'manual_reload_rows' not in st.session_state:
            st.session_state['manual_reload_rows'] = []
        if 'manual_reload_error' not in st.session_state:
            st.session_state['manual_reload_error'] = ''
        if 'manual_reload_file_name' not in st.session_state:
            st.session_state['manual_reload_file_name'] = ''

        should_analyze = analyze_clicked or execute_clicked
        if should_analyze:
            if not upload:
                st.warning('Upload a CSV or XLS file first.')
            else:
                try:
                    rows = import_manual_reload_rows(upload)
                    if not rows:
                        st.error('No rows were found in the import file.')
                    else:
                        analyzed_rows: List[Dict[str, str]] = []
                        for row in rows:
                            analyzed_rows.append(
                                _run_single_row(
                                    row,
                                    profile,
                                    execute_actions and execute_clicked,
                                    cleanup_old_images and execute_clicked,
                                    execute_reload and execute_clicked,
                                )
                            )
                        st.session_state['manual_reload_rows'] = analyzed_rows
                        st.session_state['manual_reload_error'] = ''
                        st.session_state['manual_reload_file_name'] = upload.name
                except Exception as exc:
                    st.session_state['manual_reload_error'] = str(exc)

        if st.session_state.get('manual_reload_error'):
            st.error(st.session_state['manual_reload_error'])

        rows = st.session_state.get('manual_reload_rows', [])
        if rows:
            display_rows = _display_rows(rows)
            total = len(display_rows)
            online = sum(1 for row in display_rows if row.get('Ping Status') == 'Online')
            needs_upgrade = sum(1 for row in display_rows if row.get('Needs Upgrade') == 'Yes')
            ready = sum(1 for row in display_rows if row.get('Workflow Status') == 'Ready to execute')
            c1, c2, c3, c4 = st.columns(4)
            c1.metric('Rows', total)
            c2.metric('Online', online)
            c3.metric('Need Upgrade', needs_upgrade)
            c4.metric('Ready', ready)
            if st.session_state.get('manual_reload_file_name'):
                st.caption(f'Last import: {st.session_state["manual_reload_file_name"]}')
            st.dataframe(pd.DataFrame(display_rows), width='stretch')
