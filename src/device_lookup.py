import os
import time
import re
from dataclasses import dataclass
from datetime import datetime
from typing import Any, Dict, List, Optional, Tuple

from src.dna_client import DNAClient
from src.ssh_device_client import ssh_show_version


# Region-level inventory cache to speed up repeated lookups.
_DEVICE_CACHE: Dict[str, Tuple[float, List[Dict[str, Any]]]] = {}


@dataclass
class DeviceLookupResult:
    hostname: str
    ip_address: str
    region: str
    image_version: str
    online_status: str
    last_update_outcome: str
    last_update_time: str
    ssh_validated: bool
    ssh_image_or_version: str
    ssh_error: str
    raw_update_record: Dict[str, Any]


def _cache_ttl_seconds() -> int:
    raw = os.getenv('LOOKUP_CACHE_TTL_SECONDS', '300').strip()
    try:
        return max(0, int(raw))
    except ValueError:
        return 300


def _to_device_list(payload: Any) -> List[Dict[str, Any]]:
    if isinstance(payload, dict):
        if isinstance(payload.get('response'), list):
            return payload.get('response') or []
        if isinstance(payload.get('response'), dict):
            return [payload.get('response')]
    if isinstance(payload, list):
        return payload
    return []


def _normalize_online_status(device: Dict[str, Any]) -> str:
    candidates = [
        str(device.get('reachabilityStatus') or ''),
        str(device.get('collectionStatus') or ''),
        str(device.get('upTime') or ''),
    ]
    joined = ' '.join(candidates).lower()
    if any(k in joined for k in ('reachable', 'online', 'managed')):
        return 'Online'
    if any(k in joined for k in ('unreachable', 'offline', 'down')):
        return 'Offline'
    return 'Unknown'


def _parse_dt(value: Any) -> Optional[datetime]:
    if value is None:
        return None
    if isinstance(value, (int, float)):
        try:
            # DNAC often returns milliseconds epoch
            if value > 10_000_000_000:
                value = value / 1000
            return datetime.utcfromtimestamp(value)
        except Exception:
            return None
    if isinstance(value, str):
        v = value.strip()
        if not v:
            return None
        for fmt in (
            '%Y-%m-%dT%H:%M:%S.%fZ',
            '%Y-%m-%dT%H:%M:%SZ',
            '%Y-%m-%d %H:%M:%S',
            '%Y-%m-%d',
        ):
            try:
                return datetime.strptime(v, fmt)
            except ValueError:
                continue
    return None


def _best_time(record: Dict[str, Any]) -> Optional[datetime]:
    for key in ('endTime', 'completedTime', 'lastUpdateTime', 'updateTime', 'startTime', 'scheduleTime', 'scheduledAt'):
        ts = _parse_dt(record.get(key))
        if ts:
            return ts
    return None


def _extract_status_text(record: Dict[str, Any]) -> str:
    parts = [
        str(record.get('status') or ''),
        str(record.get('state') or ''),
        str(record.get('imageUpdateStatus') or ''),
        str(record.get('installStatus') or ''),
        str(record.get('taskStatus') or ''),
        str(record.get('description') or ''),
    ]
    return ' '.join(parts).strip().lower()


def _outcome_from_status_text(status_text: str) -> str:
    if any(k in status_text for k in ('success', 'completed', 'complete', 'installed')):
        return 'Success'
    if any(k in status_text for k in ('fail', 'failed', 'error', 'aborted', 'rollback')):
        return 'Failed'
    if any(k in status_text for k in ('in progress', 'running', 'pending', 'scheduled')):
        return 'In Progress'
    return 'Unknown'


def _version_tokens(value: str) -> List[str]:
    if not value:
        return []
    return re.findall(r'[A-Za-z]+|\d+', value)


def _versions_match(dnac_version: str, ssh_version: str) -> bool:
    a = _version_tokens(dnac_version.lower())
    b = _version_tokens(ssh_version.lower())
    if not a or not b:
        return False

    def normalize(tokens: List[str]) -> List[str]:
        out: List[str] = []
        for t in tokens:
            if t.isdigit():
                out.append(str(int(t)))
            else:
                out.append(t)
        return out

    return normalize(a) == normalize(b)


def _pick_latest_record(records: List[Dict[str, Any]]) -> Optional[Dict[str, Any]]:
    if not records:
        return None
    with_time: List[Tuple[datetime, Dict[str, Any]]] = []
    no_time: List[Dict[str, Any]] = []
    for r in records:
        ts = _best_time(r)
        if ts:
            with_time.append((ts, r))
        else:
            no_time.append(r)
    if with_time:
        with_time.sort(key=lambda item: item[0], reverse=True)
        return with_time[0][1]
    return no_time[0] if no_time else None


def _format_time(ts: Optional[datetime]) -> str:
    if not ts:
        return 'Unknown'
    return ts.strftime('%Y-%m-%d %H:%M:%S UTC')


def get_devices_for_region(region: str, client: DNAClient, refresh: bool = False) -> List[Dict[str, Any]]:
    ttl = _cache_ttl_seconds()
    now = time.time()

    if not refresh and ttl > 0 and region in _DEVICE_CACHE:
        cached_at, rows = _DEVICE_CACHE[region]
        if (now - cached_at) <= ttl:
            return rows

    payload = client.list_devices()
    devices = _to_device_list(payload)
    if ttl > 0:
        _DEVICE_CACHE[region] = (now, devices)
    return devices


def get_last_update_record(client: DNAClient, device_id: str) -> Dict[str, Any]:
    candidates: List[Dict[str, Any]] = []

    # Different DNAC versions expose SWIM/device history through slightly different shapes.
    endpoint_attempts = [
        (f'/dna/intent/api/v1/image/device/{device_id}', None),
        ('/dna/intent/api/v1/image/device', {'deviceUuid': device_id}),
        ('/dna/intent/api/v1/image/device', {'networkDeviceId': device_id}),
        ('/dna/intent/api/v1/image/device', {'deviceId': device_id}),
        ('/dna/intent/api/v1/image/device', {'id': device_id}),
    ]

    for ep, params in endpoint_attempts:
        try:
            payload = client.get(ep, params=params)
            records = _to_device_list(payload)
            for rec in records:
                rec_device_id = str(rec.get('deviceUuid') or rec.get('networkDeviceId') or rec.get('id') or '')
                if rec_device_id == device_id or ep.endswith(device_id):
                    candidates.append(rec)
                elif ep.endswith(device_id):
                    candidates.append(rec)
        except Exception:
            continue

    latest = _pick_latest_record(candidates)
    if latest:
        return latest

    return {}


def find_device_by_name(
    clients_by_region: List[Tuple[str, DNAClient]],
    device_name: str,
    use_cache: bool = True,
    use_ssh_validation: bool = False,
) -> Optional[DeviceLookupResult]:
    target = (device_name or '').strip().lower()
    if not target:
        return None

    for region, client in clients_by_region:
        devices = get_devices_for_region(region, client, refresh=not use_cache)
        for d in devices:
            hostname = (d.get('hostname') or '').strip().lower()
            fqdn = (d.get('fullyQualifiedDomainName') or '').strip().lower()
            if target in (hostname, fqdn):
                device_id = str(d.get('id') or d.get('instanceUuid') or '')
                update_record = get_last_update_record(client, device_id) if device_id else {}
                status_text = _extract_status_text(update_record)
                outcome = _outcome_from_status_text(status_text)
                update_time = _format_time(_best_time(update_record))
                ssh_image_or_version = 'Not Run'
                ssh_error = ''
                ssh_validated = False
                ip_address = d.get('managementIpAddress') or 'Unknown'
                if use_ssh_validation and ip_address != 'Unknown':
                    ssh_validated = True
                    parsed, err = ssh_show_version(ip_address)
                    if parsed:
                        ssh_image_or_version = parsed
                    else:
                        ssh_image_or_version = 'Unknown'
                    if err:
                        ssh_error = err

                    # Deterministic update outcome via direct device validation.
                    if ssh_error:
                        outcome = 'Failed (SSH Error)'
                    elif _versions_match(d.get('softwareVersion') or '', ssh_image_or_version):
                        outcome = 'Success'
                    else:
                        outcome = 'Failed (Version Mismatch)'

                # Safety guard: once SSH validation is requested, never leave outcome as Unknown.
                if use_ssh_validation and outcome == 'Unknown':
                    if ip_address == 'Unknown':
                        outcome = 'Failed (No Management IP)'
                    elif ssh_error:
                        outcome = 'Failed (SSH Error)'
                    elif ssh_image_or_version in ('Unknown', 'Not Run', ''):
                        outcome = 'Failed (SSH Parse Error)'
                    else:
                        outcome = 'Failed (Version Mismatch)'
                return DeviceLookupResult(
                    hostname=d.get('hostname') or d.get('fullyQualifiedDomainName') or device_name,
                    ip_address=ip_address,
                    region=region,
                    image_version=d.get('softwareVersion') or 'Unknown',
                    online_status=_normalize_online_status(d),
                    last_update_outcome=outcome,
                    last_update_time=update_time,
                    ssh_validated=ssh_validated,
                    ssh_image_or_version=ssh_image_or_version,
                    ssh_error=ssh_error,
                    raw_update_record=update_record,
                )
    return None


def find_many_devices(
    clients_by_region: List[Tuple[str, DNAClient]],
    device_names: List[str],
    use_cache: bool = True,
    use_ssh_validation: bool = False,
) -> List[Dict[str, Any]]:
    rows: List[Dict[str, Any]] = []
    for name in device_names:
        n = (name or '').strip()
        if not n:
            continue
        hit = find_device_by_name(
            clients_by_region,
            n,
            use_cache=use_cache,
            use_ssh_validation=use_ssh_validation,
        )
        if not hit:
            rows.append(
                {
                    'device': n,
                    'found': False,
                    'ip_address': 'Unknown',
                    'region': 'Unknown',
                    'image_version': 'Unknown',
                    'online_status': 'Unknown',
                    'last_update_outcome': 'Unknown',
                    'last_update_time': 'Unknown',
                    'ssh_validated': use_ssh_validation,
                    'ssh_image_or_version': 'Unknown',
                    'ssh_error': 'Device not found',
                }
            )
            continue
        rows.append(
            {
                'device': hit.hostname,
                'found': True,
                'ip_address': hit.ip_address,
                'region': hit.region,
                'image_version': hit.image_version,
                'online_status': hit.online_status,
                'last_update_outcome': hit.last_update_outcome,
                'last_update_time': hit.last_update_time,
                'ssh_validated': hit.ssh_validated,
                'ssh_image_or_version': hit.ssh_image_or_version,
                'ssh_error': hit.ssh_error,
            }
        )
    return rows


def check_update_endpoint_capabilities(clients_by_region: List[Tuple[str, DNAClient]]) -> Dict[str, List[Dict[str, Any]]]:
    checks = [
        ('/dna/intent/api/v1/image/device', {'offset': 1, 'limit': 1}),
        ('/dna/intent/api/v1/image/import', {'offset': 1, 'limit': 1}),
        ('/dna/intent/api/v1/task', {'offset': 1, 'limit': 1}),
    ]
    result: Dict[str, List[Dict[str, Any]]] = {}

    for region, client in clients_by_region:
        rows: List[Dict[str, Any]] = []
        for path, params in checks:
            try:
                client.get(path, params=params)
                rows.append({'endpoint': path, 'available': True, 'status': 200})
            except Exception as exc:
                status = None
                response = getattr(exc, 'response', None)
                if response is not None:
                    status = response.status_code
                rows.append({'endpoint': path, 'available': False, 'status': status or 'error'})
        result[region] = rows

    return result
