import os
import csv
import io
import re
import json
import threading
from uuid import uuid4
from typing import Dict, List, Tuple, Any
from datetime import datetime, time, timedelta
from pathlib import Path
from zoneinfo import ZoneInfo
from typing import Dict, List, Tuple, Any

import pandas as pd
import streamlit as st
from dotenv import load_dotenv

from src.dna_client import DNAClient
from src.device_lookup import find_device_by_name, find_many_devices
from src.manual_reloads import render_manual_reloads_panel
from src.outlook_messages import load_outlook_upgrade_messages
from src.ssh_device_client import collect_device_precheck_outputs, icmp_ping, ssh_show_version
from src.task_status import get_upgrade_tasks_for_date, get_upcoming_tasks_for_date


load_dotenv(override=True)


_UPCOMING_BG_LOCK = threading.Lock()
_UPCOMING_BG_THREADS: Dict[str, threading.Thread] = {}
_UPCOMING_BG_RESULTS: Dict[str, Dict[str, Any]] = {}
_OUTLOOK_BG_LOCK = threading.Lock()
_OUTLOOK_BG_THREADS: Dict[str, threading.Thread] = {}
_OUTLOOK_BG_RESULTS: Dict[str, Dict[str, Any]] = {}


def build_clients():
    verify_ssl = os.getenv('DNAC_VERIFY_SSL', 'true').strip().lower() in ('1', 'true', 'yes', 'on')
    regions = [
        ('EMEA', os.getenv('EMEA_DNAC_BASE_URL'), 'EMEA_DNAC_TOKEN'),
        ('US', os.getenv('US_DNAC_BASE_URL'), 'US_DNAC_TOKEN'),
        ('APAC', os.getenv('APAC_DNAC_BASE_URL'), 'APAC_DNAC_TOKEN'),
    ]
    clients = []
    for region, base_url, token_env in regions:
        if not base_url:
            continue
        clients.append((region, DNAClient(base_url=base_url, token_env=token_env, verify_ssl=verify_ssl)))
    return clients


def _format_runtime_error(exc: Exception, panel_name: str) -> str:
    message = str(exc).strip() or exc.__class__.__name__
    lower_message = message.lower()
    if '401' in lower_message or 'unauthorized' in lower_message or 'could not obtain dnac token' in lower_message:
        return (
            f'{panel_name} could not load because DNAC authentication failed. '
            'Check DNA_USERNAME/DNA_PASSWORD or the region token values in .env, then retry.'
        )
    if (
        'name resolutionerror' in lower_message
        or 'failed to resolve' in lower_message
        or 'getaddrinfo failed' in lower_message
    ):
        return (
            f'{panel_name} could not load because a DNAC hostname could not be resolved. '
            'Check EMEA_DNAC_BASE_URL/US_DNAC_BASE_URL/APAC_DNAC_BASE_URL hostnames, DNS, and VPN connectivity, then retry.'
        )
    return message


def _extract_device_names_from_csv(file_bytes: bytes) -> list:
    text = file_bytes.decode('utf-8', errors='ignore')
    reader = csv.DictReader(io.StringIO(text))
    if not reader.fieldnames:
        return []
    lower = {name.lower(): name for name in reader.fieldnames}
    preferred = None
    for key in ('device', 'device_name', 'hostname', 'fqdn'):
        if key in lower:
            preferred = lower[key]
            break
    if not preferred:
        preferred = reader.fieldnames[0]
    names = []
    for row in reader:
        value = (row.get(preferred) or '').strip()
        if value:
            names.append(value)
    return names


def _rows_to_csv_bytes(rows: List[Dict]) -> bytes:
    if not rows:
        return b''
    output = io.StringIO()
    writer = csv.DictWriter(output, fieldnames=list(rows[0].keys()))
    writer.writeheader()
    writer.writerows(rows)
    return output.getvalue().encode('utf-8')


def _rows_to_xls_bytes(rows: List[Dict]) -> bytes:
    import xlwt

    wb = xlwt.Workbook()
    ws = wb.add_sheet('Tasks')
    if not rows:
        stream = io.BytesIO()
        wb.save(stream)
        return stream.getvalue()

    columns = list(rows[0].keys())
    for col_idx, col_name in enumerate(columns):
        ws.write(0, col_idx, col_name)

    for row_idx, row in enumerate(rows, start=1):
        for col_idx, col_name in enumerate(columns):
            ws.write(row_idx, col_idx, str(row.get(col_name, '')))

    stream = io.BytesIO()
    wb.save(stream)
    return stream.getvalue()


def _rows_to_pdf_bytes(rows: List[Dict], title: str = 'Upgrade Task Status') -> bytes:
    from reportlab.lib import colors
    from reportlab.lib.pagesizes import landscape, letter
    from reportlab.platypus import SimpleDocTemplate, Table, TableStyle, Paragraph, Spacer
    from reportlab.lib.styles import getSampleStyleSheet

    stream = io.BytesIO()
    doc = SimpleDocTemplate(stream, pagesize=landscape(letter))
    styles = getSampleStyleSheet()
    story = [Paragraph(title, styles['Title']), Spacer(1, 8)]

    if not rows:
        story.append(Paragraph('No rows to export.', styles['Normal']))
        doc.build(story)
        return stream.getvalue()

    columns = list(rows[0].keys())
    table_data = [columns]
    for row in rows:
        table_data.append([str(row.get(col, '')) for col in columns])

    table = Table(table_data, repeatRows=1)
    table.setStyle(
        TableStyle(
            [
                ('BACKGROUND', (0, 0), (-1, 0), colors.lightgrey),
                ('TEXTCOLOR', (0, 0), (-1, 0), colors.black),
                ('GRID', (0, 0), (-1, -1), 0.25, colors.grey),
                ('FONTNAME', (0, 0), (-1, 0), 'Helvetica-Bold'),
                ('FONTSIZE', (0, 0), (-1, -1), 8),
            ]
        )
    )
    story.append(table)
    doc.build(story)
    return stream.getvalue()


def _parse_task_date(task_date_text: str):
    return datetime.strptime(task_date_text.strip(), '%Y-%m-%d').date()


def _current_task_config(task_date_text: str, force_fixed_est: bool, status_filter: List[str]) -> Dict:
    return {
        'task_date_text': task_date_text.strip(),
        'force_fixed_est': force_fixed_est,
        'status_filter': tuple(status_filter),
    }


def _extract_site_code(description: str) -> str:
    match = re.search(r'([A-Z]{2}-[A-Z0-9]+)\s*$', description or '')
    if match:
        return match.group(1)
    return 'UNKNOWN-SITE'


def _parse_sort_start_time(value: str):
    try:
        return datetime.fromisoformat(str(value or '').strip())
    except Exception:
        return None


def _dnac_row_upgrade_status(row: Dict) -> str:
    status = str(row.get('dnac_task_status') or '').strip().lower()
    if status == 'success':
        return 'Success'
    if status == 'failure':
        return 'Fail'
    return ''


def _build_completed_backfill_rows(
    completed_rows: List[Dict],
    task_date,
    tz_name: str,
    min_hour_local: int = 2,
) -> List[Dict]:
    tz = ZoneInfo(tz_name)
    cutoff = datetime.combine(task_date, time(hour=min_hour_local, minute=0), tzinfo=tz)
    rows: List[Dict] = []

    for item in completed_rows:
        status = str(item.get('task_status') or '').strip()
        if status not in {'Success', 'Failure'}:
            continue

        sort_end = _parse_sort_start_time(str(item.get('sort_end_time') or ''))
        if sort_end is None:
            continue
        if sort_end.tzinfo is None:
            sort_end = sort_end.replace(tzinfo=tz)

        if sort_end.date() != task_date or sort_end < cutoff:
            continue

        upgrade_status = 'Success' if status == 'Success' else 'Fail'
        task_id = str(item.get('task_id') or '').strip()
        device_ip = str(item.get('device_ip') or 'Unknown').strip() or 'Unknown'
        rows.append(
            {
                'description': f'DNAC completed upgrade task ({task_id or "no-task-id"})',
                'start_date_time_est': item.get('date_time_est', ''),
                'sort_start_time': sort_end.isoformat(),
                'device_name': item.get('device_name', 'Unknown'),
                'ip_address': device_ip,
                'current_image_version': item.get('current_image_version', 'Unknown'),
                'baseline_image_version': item.get('current_image_version', 'Unknown'),
                'devices_scheduled': 1,
                'region': item.get('region', ''),
                'task_id': task_id,
                'is_error': status == 'Failure',
                'source': 'task_history',
                'ip_reachability': 'Unknown',
                'upgrade_status': upgrade_status,
                'dnac_task_status': status,
                'ssh_validated_image': '',
                'ssh_validation_error': '',
                'last_upgrade_check': '',
            }
        )

    return rows


def _env_int(name: str, default: int, minimum: int = 0) -> int:
    raw = str(os.getenv(name, str(default))).strip()
    try:
        return max(minimum, int(raw))
    except ValueError:
        return default


def _history_cache_ttl_seconds(task_date, tz_name: str) -> int:
    tz = ZoneInfo(tz_name)
    today_local = datetime.now(tz).date()
    if task_date == today_local:
        return _env_int('DNAC_HISTORY_CACHE_SECONDS_TODAY', 600, minimum=30)
    return _env_int('DNAC_HISTORY_CACHE_SECONDS_PAST', 3600, minimum=60)


def _format_elapsed_duration(seconds_value: float) -> str:
    seconds_int = max(0, int(seconds_value))
    hours, remainder = divmod(seconds_int, 3600)
    minutes, seconds = divmod(remainder, 60)
    if hours > 0:
        return f'{hours}h {minutes}m {seconds}s'
    if minutes > 0:
        return f'{minutes}m {seconds}s'
    return f'{seconds}s'


def _set_upcoming_date_today() -> None:
    st.session_state['upcoming_query_date'] = datetime.now(ZoneInfo('America/New_York')).date()


def _get_cached_upgrade_history_rows(
    clients,
    task_date,
    tz_name: str,
    tz_label: str,
    include_alt_sources: bool | None = None,
) -> Tuple[List[Dict], bool]:
    cache = st.session_state.setdefault('upgrade_history_cache', {})
    if include_alt_sources is None:
        include_alt_sources = str(os.getenv('UPCOMING_INCLUDE_ALT_HISTORY', 'false')).strip().lower() in ('1', 'true', 'yes', 'on')
    cache_key = f"{task_date.isoformat()}|{tz_name}|{tz_label}|alt={int(bool(include_alt_sources))}"
    now_epoch = datetime.now().timestamp()
    ttl_seconds = _history_cache_ttl_seconds(task_date, tz_name)
    entry = cache.get(cache_key)
    if isinstance(entry, dict):
        fetched_at = float(entry.get('fetched_at', 0.0) or 0.0)
        if fetched_at and (now_epoch - fetched_at) < ttl_seconds:
            rows = entry.get('rows', [])
            if isinstance(rows, list):
                return rows, True

    history_page_size = _env_int('UPCOMING_HISTORY_PAGE_SIZE', 200, minimum=50)
    history_max_pages = _env_int('UPCOMING_HISTORY_MAX_PAGES', 1, minimum=1)
    rows = get_upgrade_tasks_for_date(
        clients,
        task_date,
        tz_name=tz_name,
        tz_label_override=tz_label,
        page_size=history_page_size,
        max_pages=history_max_pages,
        include_alt_sources=include_alt_sources,
    )
    cache[cache_key] = {
        'fetched_at': now_epoch,
        'rows': rows,
    }
    st.session_state['upgrade_history_cache'] = cache
    return rows, False


def _upcoming_snapshot_path() -> Path:
    root = Path(__file__).resolve().parents[1]
    return root / 'data' / 'upcoming_snapshot.json'


def _save_upcoming_snapshot(rows: List[Dict], meta: str, loaded_config: Dict) -> None:
    path = _upcoming_snapshot_path()
    path.parent.mkdir(parents=True, exist_ok=True)
    payload = {
        'saved_at': datetime.now().strftime('%Y-%m-%d %H:%M:%S'),
        'rows': rows,
        'meta': meta,
        'loaded_config': loaded_config,
    }
    path.write_text(json.dumps(payload, ensure_ascii=True), encoding='utf-8')


def _load_upcoming_snapshot() -> Dict[str, Any]:
    path = _upcoming_snapshot_path()
    if not path.exists():
        return {}
    try:
        payload = json.loads(path.read_text(encoding='utf-8'))
    except Exception:
        return {}
    if not isinstance(payload, dict):
        return {}
    rows = payload.get('rows', [])
    if not isinstance(rows, list) or not rows:
        return {}
    return payload


def _build_ran_rows_from_2am_to_now(completed_rows: List[Dict], task_date, tz_name: str) -> List[Dict]:
    tz = ZoneInfo(tz_name)
    cutoff = datetime.combine(task_date, time(hour=2, minute=0), tzinfo=tz)
    now_local = datetime.now(tz)
    rows: List[Dict] = []

    for item in completed_rows:
        status = str(item.get('task_status') or '').strip()
        if status not in {'Success', 'Failure', 'In Progress'}:
            continue

        sort_end = _parse_sort_start_time(str(item.get('sort_end_time') or ''))
        if sort_end is None:
            continue
        if sort_end.tzinfo is None:
            sort_end = sort_end.replace(tzinfo=tz)
        if sort_end.date() != task_date:
            continue
        if sort_end < cutoff or sort_end > now_local:
            continue

        if status == 'Success':
            upgrade_status = 'Success'
        elif status == 'Failure':
            upgrade_status = 'Fail'
        else:
            upgrade_status = 'Pending'

        task_id = str(item.get('task_id') or '').strip()
        device_ip = str(item.get('device_ip') or 'Unknown').strip() or 'Unknown'
        rows.append(
            {
                'description': f'DNAC ran task ({status}) ({task_id or "no-task-id"})',
                'start_date_time_est': item.get('date_time_est', ''),
                'sort_start_time': sort_end.isoformat(),
                'device_name': item.get('device_name', 'Unknown'),
                'ip_address': device_ip,
                'current_image_version': item.get('current_image_version', 'Unknown'),
                'baseline_image_version': item.get('current_image_version', 'Unknown'),
                'devices_scheduled': 1,
                'region': item.get('region', ''),
                'task_id': task_id,
                'is_error': status == 'Failure',
                'source': 'task_history_ran',
                'ip_reachability': 'Unknown',
                'upgrade_status': upgrade_status,
                'dnac_task_status': status,
                'ssh_validated_image': '',
                'ssh_validation_error': '',
                'last_upgrade_check': '',
                'last_upgrade_check_epoch': 0.0,
            }
        )

    return rows


def _fetch_upcoming_rows_payload(
    task_date_text: str,
    force_fixed_est: bool,
    status_filter: List[str],
    existing_rows: List[Dict],
) -> Dict[str, Any]:
    task_date = _parse_task_date(task_date_text)
    tz_name = 'Etc/GMT+5' if force_fixed_est else 'America/New_York'
    tz_label = 'EST' if force_fixed_est else ''
    clients = build_clients()

    rows = get_upcoming_tasks_for_date(clients, task_date, tz_name=tz_name, tz_label_override=tz_label)
    include_ran_tasks = str(os.getenv('UPCOMING_INCLUDE_RAN_TASKS', 'true')).strip().lower() in (
        '1',
        'true',
        'yes',
        'on',
    )
    ran_rows: List[Dict] = []
    if include_ran_tasks:
        include_alt_sources = str(os.getenv('UPCOMING_INCLUDE_ALT_HISTORY', 'true')).strip().lower() in (
            '1',
            'true',
            'yes',
            'on',
        )
        completed_rows, _ = _get_cached_upgrade_history_rows(
            clients,
            task_date,
            tz_name=tz_name,
            tz_label=tz_label,
            include_alt_sources=include_alt_sources,
        )
        ran_rows = _build_ran_rows_from_2am_to_now(completed_rows, task_date, tz_name=tz_name)
        rows.extend(ran_rows)

    # Return a fresh deduped snapshot; persistence policy is applied by _load_upcoming_rows.
    deduped_rows: List[Dict] = []
    seen = set()
    for row in sorted(rows, key=lambda item: str(item.get('sort_start_time') or '')):
        key = _monitor_key(row)
        if key in seen:
            continue
        seen.add(key)
        deduped_rows.append(_initial_upcoming_state(row))

    zone_note = 'fixed EST (UTC-5)' if force_fixed_est else 'Eastern local time (EST/EDT)'
    if include_ran_tasks:
        meta = (
            f'Upcoming list loaded for {task_date} using {zone_note}; '
            f'source=schedule_v4 + ran_tasks(>=2:00 to now {"EST" if force_fixed_est else "local"}), '
            f'ran_added={len(ran_rows)}'
        )
    else:
        meta = f'Upcoming list loaded for {task_date} using {zone_note}; source=schedule_v4'

    return {
        'upcoming_rows': deduped_rows,
        'upcoming_meta': meta,
        'upcoming_loaded_config': _current_task_config(task_date_text, force_fixed_est, status_filter),
        'upcoming_error': '',
    }


def _start_upcoming_background_refresh(
    session_id: str,
    task_date_text: str,
    force_fixed_est: bool,
    status_filter: List[str],
    existing_rows: List[Dict],
) -> bool:
    with _UPCOMING_BG_LOCK:
        current_thread = _UPCOMING_BG_THREADS.get(session_id)
        if current_thread and current_thread.is_alive():
            return False
        _UPCOMING_BG_RESULTS.pop(session_id, None)

    existing_snapshot = [dict(row) for row in existing_rows]
    filter_snapshot = list(status_filter)

    def _worker() -> None:
        try:
            payload = _fetch_upcoming_rows_payload(
                task_date_text=task_date_text,
                force_fixed_est=force_fixed_est,
                status_filter=filter_snapshot,
                existing_rows=existing_snapshot,
            )
            result = {'payload': payload, 'error': ''}
        except Exception as exc:
            result = {'payload': None, 'error': str(exc)}

        with _UPCOMING_BG_LOCK:
            _UPCOMING_BG_RESULTS[session_id] = result

    thread = threading.Thread(target=_worker, name=f'upcoming-refresh-{session_id}', daemon=True)
    with _UPCOMING_BG_LOCK:
        _UPCOMING_BG_THREADS[session_id] = thread
    thread.start()
    return True


def _collect_upcoming_background_result(session_id: str) -> Dict[str, Any] | None:
    with _UPCOMING_BG_LOCK:
        return _UPCOMING_BG_RESULTS.pop(session_id, None)


def _version_key(version_text: str):
    text = str(version_text or '').strip().lower()
    if not text or text == 'unknown':
        return None
    tokens = re.findall(r'\d+|[a-z]+', text)
    if not tokens:
        return None

    key = []
    for token in tokens:
        if token.isdigit():
            key.append((0, int(token)))
        else:
            key.append((1, token))
    return tuple(key)


def _is_higher_version(new_version: str, old_version: str):
    new_key = _version_key(new_version)
    old_key = _version_key(old_version)
    if new_key is None or old_key is None:
        return None
    return new_key > old_key


def _initial_upcoming_state(row: Dict) -> Dict:
    initialized = dict(row)
    baseline = str(row.get('current_image_version') or 'Unknown').strip() or 'Unknown'
    initialized['baseline_image_version'] = baseline
    initialized['ip_reachability'] = 'Unknown'
    initialized['upgrade_status'] = 'Pending'
    initialized['ssh_validated_image'] = ''
    initialized['ssh_validation_error'] = ''
    initialized['last_upgrade_check'] = ''
    initialized['last_upgrade_check_epoch'] = 0.0
    initialized['offline_seen_at'] = ''
    initialized['online_restored_at'] = ''
    return initialized


def _merge_persistent_upcoming_rows(existing_rows: List[Dict], fresh_rows: List[Dict], task_date) -> List[Dict]:
    existing_by_key = {_monitor_key(row): dict(row) for row in existing_rows}
    merged: List[Dict] = []
    fresh_keys = set()

    for fresh in fresh_rows:
        key = _monitor_key(fresh)
        fresh_keys.add(key)
        if key in existing_by_key:
            existing = existing_by_key[key]
            for field in (
                'description',
                'start_date_time_est',
                'sort_start_time',
                'device_name',
                'ip_address',
                'devices_scheduled',
                'region',
                'task_id',
                'is_error',
                'source',
            ):
                existing[field] = fresh.get(field, existing.get(field, ''))
            merged.append(existing)
        else:
            merged.append(_initial_upcoming_state(fresh))

    for key, row in existing_by_key.items():
        if key in fresh_keys:
            continue
        scheduled = _parse_sort_start_time(str(row.get('sort_start_time') or ''))
        if scheduled is None or scheduled.date() == task_date:
            # Keep rows visible even after their scheduled time so post-upgrade validation remains possible.
            merged.append(row)

    merged.sort(key=lambda row: str(row.get('sort_start_time') or ''))
    return merged


def _normalize_device_token(value: str) -> str:
    return (value or '').strip().lower()


def _build_known_device_map() -> Dict[str, str]:
    known: Dict[str, str] = {}
    for row in st.session_state.get('upcoming_rows', []):
        device_name = str(row.get('device_name') or '').strip()
        ip_address = str(row.get('ip_address') or '').strip()
        label = device_name or ip_address
        if device_name:
            known[_normalize_device_token(device_name)] = label
        if ip_address:
            known[_normalize_device_token(ip_address)] = label

    for section in ('ping_monitor_active', 'ping_monitor_completed'):
        for row in st.session_state.get(section, []):
            device_name = str(row.get('device_name') or '').strip()
            ip_address = str(row.get('ip_address') or '').strip()
            label = device_name or ip_address
            if device_name:
                known[_normalize_device_token(device_name)] = label
            if ip_address:
                known[_normalize_device_token(ip_address)] = label

    return known


def _build_outlook_rows_payload(
    lookback_hours: int,
    max_messages: int,
    mailbox_name: str,
    folder_path: str,
    class_filter: List[str],
    subject_contains: str,
    sender_contains: str,
    body_contains: str,
    to_contains: str,
    task_date_text: str,
    known_devices: Dict[str, str],
) -> Dict[str, Any]:
    target_date = _parse_task_date(task_date_text)
    rows, meta = load_outlook_upgrade_messages(
        lookback_hours=lookback_hours,
        max_messages=max_messages,
        mailbox_name=mailbox_name,
        folder_path=folder_path,
        target_date=target_date,
    )

    if class_filter:
        selected = set(class_filter)
        rows = [row for row in rows if row.get('classification') in selected]

    subject_query = (subject_contains or '').strip().lower()
    sender_query = (sender_contains or '').strip().lower()
    body_query = (body_contains or '').strip().lower()
    to_query = (to_contains or '').strip().lower()
    if subject_query:
        rows = [row for row in rows if subject_query in str(row.get('subject') or '').lower()]
    if sender_query:
        rows = [
            row
            for row in rows
            if sender_query in str(row.get('sender') or '').lower()
            or sender_query in str(row.get('sender_email') or '').lower()
        ]
    if body_query:
        rows = [row for row in rows if body_query in str(row.get('body') or '').lower()]
    if to_query:
        rows = [
            row
            for row in rows
            if to_query in str(row.get('to_recipients') or '').lower()
            or to_query in str(row.get('recipient_emails') or '').lower()
        ]

    enriched_rows: List[Dict] = []
    for row in rows:
        mapped = dict(row)
        match_label = ''
        ip_token = _normalize_device_token(str(row.get('ip_address') or ''))
        device_token = _normalize_device_token(str(row.get('device_name') or ''))
        if ip_token and ip_token in known_devices:
            match_label = known_devices[ip_token]
        elif device_token and device_token in known_devices:
            match_label = known_devices[device_token]

        mapped['correlated_device'] = match_label
        mapped['correlated'] = 'Yes' if match_label else 'No'
        enriched_rows.append(mapped)

    loaded_at = datetime.now().strftime('%Y-%m-%d %H:%M:%S')
    return {
        'rows': enriched_rows,
        'meta': meta,
        'last_run': loaded_at,
        'fetched_epoch': datetime.now().timestamp(),
    }


def _load_outlook_rows(
    lookback_hours: int,
    max_messages: int,
    mailbox_name: str,
    folder_path: str,
    class_filter: List[str],
    subject_contains: str,
    sender_contains: str,
    body_contains: str,
    to_contains: str,
    task_date_text: str,
) -> None:
    payload = _build_outlook_rows_payload(
        lookback_hours=lookback_hours,
        max_messages=max_messages,
        mailbox_name=mailbox_name,
        folder_path=folder_path,
        class_filter=class_filter,
        subject_contains=subject_contains,
        sender_contains=sender_contains,
        body_contains=body_contains,
        to_contains=to_contains,
        task_date_text=task_date_text,
        known_devices=_build_known_device_map(),
    )

    st.session_state['outlook_rows'] = payload.get('rows', [])
    st.session_state['outlook_meta'] = payload.get('meta', '')
    st.session_state['outlook_error'] = ''
    st.session_state['outlook_last_run'] = payload.get('last_run', datetime.now().strftime('%Y-%m-%d %H:%M:%S'))


def _start_outlook_background_refresh(session_id: str, outlook_config: Dict[str, Any], known_devices: Dict[str, str]) -> bool:
    with _OUTLOOK_BG_LOCK:
        current_thread = _OUTLOOK_BG_THREADS.get(session_id)
        if current_thread and current_thread.is_alive():
            return False
        _OUTLOOK_BG_RESULTS.pop(session_id, None)

    config_snapshot = dict(outlook_config)
    known_snapshot = dict(known_devices)

    def _worker() -> None:
        pythoncom_module = None
        try:
            try:
                import pythoncom as _pythoncom

                _pythoncom.CoInitialize()
                pythoncom_module = _pythoncom
            except Exception:
                pythoncom_module = None

            payload = _build_outlook_rows_payload(
                lookback_hours=int(config_snapshot.get('lookback_hours', 24)),
                max_messages=int(config_snapshot.get('max_messages', 200)),
                mailbox_name=str(config_snapshot.get('mailbox_name', '')),
                folder_path=str(config_snapshot.get('folder_path', 'Inbox')),
                class_filter=list(config_snapshot.get('class_filter', [])),
                subject_contains=str(config_snapshot.get('subject_contains', '')),
                sender_contains=str(config_snapshot.get('sender_contains', '')),
                body_contains=str(config_snapshot.get('body_contains', '')),
                to_contains=str(config_snapshot.get('to_contains', '')),
                task_date_text=str(config_snapshot.get('task_date_text', '')),
                known_devices=known_snapshot,
            )
            result = {'payload': payload, 'error': '', 'config': config_snapshot}
        except Exception as exc:
            result = {'payload': None, 'error': str(exc), 'config': config_snapshot}
        finally:
            if pythoncom_module is not None:
                try:
                    pythoncom_module.CoUninitialize()
                except Exception:
                    pass

        with _OUTLOOK_BG_LOCK:
            _OUTLOOK_BG_RESULTS[session_id] = result

    thread = threading.Thread(target=_worker, name=f'outlook-refresh-{session_id}', daemon=True)
    with _OUTLOOK_BG_LOCK:
        _OUTLOOK_BG_THREADS[session_id] = thread
    thread.start()
    return True


def _collect_outlook_background_result(session_id: str) -> Dict[str, Any] | None:
    with _OUTLOOK_BG_LOCK:
        return _OUTLOOK_BG_RESULTS.pop(session_id, None)


def _load_task_rows(task_date_text: str, force_fixed_est: bool, status_filter: List[str]) -> None:
    task_date = _parse_task_date(task_date_text)
    tz_name = 'Etc/GMT+5' if force_fixed_est else 'America/New_York'
    tz_label = 'EST' if force_fixed_est else ''
    clients = build_clients()
    rows = get_upgrade_tasks_for_date(clients, task_date, tz_name=tz_name, tz_label_override=tz_label)
    if status_filter:
        rows = [r for r in rows if r.get('task_status') in set(status_filter)]

    counts = {
        'In Progress': 0,
        'Success': 0,
        'Failure': 0,
        'Upcoming Tasks': 0,
    }
    for row in rows:
        if row['task_status'] in counts:
            counts[row['task_status']] += 1

    zone_note = 'fixed EST (UTC-5)' if force_fixed_est else 'Eastern local time (EST/EDT)'
    st.session_state['task_rows'] = rows
    st.session_state['task_counts'] = counts
    st.session_state['task_meta'] = f"Loaded for {task_date} using {zone_note}; filters={', '.join(status_filter) if status_filter else 'ALL'}"
    st.session_state['task_error'] = ''
    st.session_state['task_loaded_config'] = _current_task_config(task_date_text, force_fixed_est, status_filter)


def _load_upcoming_rows(task_date_text: str, force_fixed_est: bool, status_filter: List[str]) -> None:
    current_config = _current_task_config(task_date_text, force_fixed_est, status_filter)
    existing_rows = list(st.session_state.get('upcoming_rows', []))
    previous_config = st.session_state.get('upcoming_loaded_config')
    keep_persistent_snapshot = bool(existing_rows) and previous_config == current_config

    payload = _fetch_upcoming_rows_payload(
        task_date_text=task_date_text,
        force_fixed_est=force_fixed_est,
        status_filter=status_filter,
        existing_rows=existing_rows,
    )

    if keep_persistent_snapshot:
        st.session_state['upcoming_rows'] = existing_rows
        st.session_state['upcoming_meta'] = f"{payload.get('upcoming_meta', '')}; persistent snapshot retained"
    else:
        st.session_state['upcoming_rows'] = payload.get('upcoming_rows', [])
        st.session_state['upcoming_meta'] = payload.get('upcoming_meta', '')

    st.session_state['upcoming_error'] = ''
    st.session_state['upcoming_retry_after_epoch'] = 0.0
    loaded_now = datetime.now()
    st.session_state['upcoming_loaded_at_epoch'] = loaded_now.timestamp()
    st.session_state['upcoming_loaded_at_text'] = loaded_now.strftime('%Y-%m-%d %H:%M:%S')
    st.session_state['upcoming_loaded_config'] = payload.get(
        'upcoming_loaded_config',
        current_config,
    )
    _save_upcoming_snapshot(
        st.session_state.get('upcoming_rows', []),
        st.session_state.get('upcoming_meta', ''),
        st.session_state.get('upcoming_loaded_config', _current_task_config(task_date_text, force_fixed_est, status_filter)),
    )


def _refresh_upcoming_ping_status(
    force_fixed_est: bool,
    include_ssh_checks: bool = True,
    force_ssh_recheck: bool = False,
) -> None:
    rows = list(st.session_state.get('upcoming_rows', []))
    if not rows:
        return

    ping_cache: Dict[str, str] = {}
    ping_output_cache: Dict[str, str] = {}
    ssh_recheck_seconds = _env_int('UPCOMING_SSH_RECHECK_SECONDS', 900, minimum=30)
    online_hold_seconds = _env_int('UPCOMING_ONLINE_HOLD_SECONDS', 300, minimum=60)
    now_epoch = datetime.now().timestamp()
    now = _monitor_now(force_fixed_est)
    updated_rows: List[Dict] = []
    for row in rows:
        updated = dict(row)
        fallback_status = _dnac_row_upgrade_status(updated)
        updated.setdefault('offline_seen_at', '')
        updated.setdefault('online_restored_at', '')
        ip_address = str(updated.get('ip_address') or '').strip()
        if not ip_address or ip_address == 'Unknown':
            updated['ip_reachability'] = 'Unknown'
            updated['upgrade_status'] = fallback_status or 'Pending'
            updated_rows.append(updated)
            continue

        if ip_address not in ping_cache:
            try:
                is_online, ping_output = icmp_ping(ip_address)
                ping_cache[ip_address] = 'Online' if is_online else 'Offline'
                ping_output_cache[ip_address] = ping_output
            except Exception:
                ping_cache[ip_address] = 'Offline'
                ping_output_cache[ip_address] = ''

        updated['ip_reachability'] = ping_cache[ip_address]
        updated['last_ping_output'] = ping_output_cache.get(ip_address, '')

        scheduled_dt = _parse_sort_start_time(str(updated.get('sort_start_time') or ''))
        if scheduled_dt is None or now < scheduled_dt:
            updated['upgrade_status'] = fallback_status or 'Pending'
            updated['current_image_version'] = _best_image_version(updated)
            updated['offline_seen_at'] = ''
            updated['online_restored_at'] = ''
            updated_rows.append(updated)
            continue

        if updated['ip_reachability'] != 'Online':
            updated['upgrade_status'] = fallback_status or 'Pending'
            updated['current_image_version'] = _best_image_version(updated)
            if not str(updated.get('offline_seen_at') or '').strip():
                updated['offline_seen_at'] = now.isoformat()
            updated['online_restored_at'] = ''
            updated_rows.append(updated)
            continue

        online_restored_at = str(updated.get('online_restored_at') or '').strip()
        if not online_restored_at:
            updated['online_restored_at'] = now.isoformat()
            updated['current_image_version'] = _best_image_version(updated)
            updated['upgrade_status'] = fallback_status or 'Pending'
            updated_rows.append(updated)
            continue

        try:
            restored_dt = datetime.fromisoformat(online_restored_at)
        except ValueError:
            restored_dt = now
            updated['online_restored_at'] = now.isoformat()

        if not force_ssh_recheck and now < (restored_dt + timedelta(seconds=online_hold_seconds)):
            updated['current_image_version'] = _best_image_version(updated)
            updated['upgrade_status'] = fallback_status or 'Pending'
            updated_rows.append(updated)
            continue

        if not include_ssh_checks:
            # Fast pass: keep list responsive by filling reachability first, then defer SSH/image validation.
            updated['current_image_version'] = _best_image_version(updated)
            updated['upgrade_status'] = _best_upgrade_status(updated, force_fixed_est)
            updated_rows.append(updated)
            continue

        last_upgrade_check_epoch = float(updated.get('last_upgrade_check_epoch') or 0.0)
        if force_ssh_recheck or not (last_upgrade_check_epoch and (now_epoch - last_upgrade_check_epoch) < ssh_recheck_seconds):
            actual_image, ssh_error = ssh_show_version(ip_address)
            updated['ssh_validated_image'] = actual_image or ''
            updated['ssh_validation_error'] = ssh_error or ''
            updated['last_upgrade_check'] = now.strftime('%Y-%m-%d %H:%M:%S')
            updated['last_upgrade_check_epoch'] = now_epoch

            if actual_image:
                updated['current_image_version'] = actual_image

        updated['upgrade_status'] = _best_upgrade_status(updated, force_fixed_est)

        updated_rows.append(updated)

    st.session_state['upcoming_rows'] = updated_rows
    st.session_state['upcoming_ping_checked_at'] = datetime.now().strftime('%Y-%m-%d %H:%M:%S')


def _upcoming_status_style(value: object) -> str:
    text = str(value or '')
    if text == 'Online':
        return 'background-color: #DCFCE7; color: #166534; font-weight: 600;'
    if text == 'Offline':
        return 'background-color: #FEE2E2; color: #991B1B; font-weight: 600;'
    return 'background-color: #F3F4F6; color: #374151;'


def _best_image_version(row: Dict) -> str:
    ssh_value = str(row.get('ssh_validated_image') or '').strip()
    if ssh_value:
        return ssh_value
    current_value = str(row.get('current_image_version') or '').strip()
    if current_value:
        return current_value
    baseline_value = str(row.get('baseline_image_version') or '').strip()
    if baseline_value:
        return baseline_value
    return 'Unknown'


def _best_upgrade_status(row: Dict, force_fixed_est: bool) -> str:
    fallback_status = _dnac_row_upgrade_status(row)
    source = str(row.get('source') or '').strip().lower()
    if source.startswith('task_history') and fallback_status:
        return fallback_status

    now = _monitor_now(force_fixed_est)
    scheduled_dt = _parse_sort_start_time(str(row.get('sort_start_time') or ''))
    if scheduled_dt is None or now < scheduled_dt:
        return fallback_status or 'Pending'

    current = _best_image_version(row)
    ssh_current = str(row.get('ssh_validated_image') or '').strip()
    baseline = str(row.get('baseline_image_version') or 'Unknown').strip() or 'Unknown'
    last_upgrade_check_epoch = float(row.get('last_upgrade_check_epoch') or 0.0)

    if last_upgrade_check_epoch > 0 or ssh_current:
        if ssh_current and current != 'Unknown' and ssh_current == current:
            return 'Success'

        version_is_higher = _is_higher_version(current, baseline)
        if version_is_higher is True:
            return 'Success'
        if version_is_higher is False:
            return 'Fail'

        if baseline != 'Unknown' and current != 'Unknown' and baseline != current:
            return 'Success'
        if baseline != 'Unknown' and current == baseline:
            return 'Fail'

    if str(row.get('ip_reachability') or '').strip() != 'Online':
        return fallback_status or 'Pending'

    online_restored_at = str(row.get('online_restored_at') or '').strip()
    if not online_restored_at:
        return fallback_status or 'Pending'

    try:
        restored_dt = datetime.fromisoformat(online_restored_at)
    except ValueError:
        return fallback_status or 'Pending'

    if now < (restored_dt + timedelta(minutes=5)):
        return fallback_status or 'Pending'

    if last_upgrade_check_epoch <= 0 and not ssh_current:
        return fallback_status or 'Pending'

    if ssh_current and current != 'Unknown' and ssh_current == current:
        return 'Success'

    version_is_higher = _is_higher_version(current, baseline)
    if version_is_higher is True:
        return 'Success'
    if version_is_higher is False:
        return 'Fail'

    if baseline != 'Unknown' and current != 'Unknown' and baseline != current:
        return 'Success'
    if baseline != 'Unknown' and current == baseline:
        return 'Fail'

    return fallback_status or 'Pending'


def _has_past_due_upcoming_rows(rows: List[Dict], force_fixed_est: bool) -> bool:
    now = _monitor_now(force_fixed_est)
    for row in rows:
        scheduled_dt = _parse_sort_start_time(str(row.get('sort_start_time') or ''))
        if scheduled_dt is not None and now >= scheduled_dt:
            return True
    return False


def _upgrade_status_style(value: object) -> str:
    text = str(value or '')
    if text == 'Success':
        return 'background-color: #DCFCE7; color: #166534; font-weight: 600;'
    if text == 'Fail':
        return 'background-color: #FEE2E2; color: #991B1B; font-weight: 600;'
    return 'background-color: #DBEAFE; color: #1E3A8A; font-weight: 600;'


@st.fragment(run_every='60s')
def _render_upcoming_panel(task_date_text: str, force_fixed_est: bool, current_config: Dict[str, Any]) -> None:
    # Refresh only the Upcoming section so the loaded table remains visible without full-page reruns.
    auto_live_refresh_enabled = str(os.getenv('UPCOMING_AUTO_LIVE_REFRESH', 'true')).strip().lower() in ('1', 'true', 'yes', 'on')
    auto_ssh_refresh_seconds = _env_int('UPCOMING_AUTO_SSH_REFRESH_SECONDS', 180, minimum=60)
    if auto_live_refresh_enabled and st.session_state.get('upcoming_rows'):
        monitor_now = _monitor_now(force_fixed_est)
        task_date = _parse_task_date(task_date_text)
        if monitor_now.date() == task_date:
            try:
                _refresh_upcoming_ping_status(force_fixed_est, include_ssh_checks=False)
                st.session_state['upcoming_auto_live_refresh_epoch'] = datetime.now().timestamp()

                now_epoch = datetime.now().timestamp()
                last_auto_ssh_refresh_epoch = float(st.session_state.get('upcoming_auto_ssh_refresh_epoch') or 0.0)
                if (now_epoch - last_auto_ssh_refresh_epoch) >= auto_ssh_refresh_seconds:
                    _refresh_upcoming_ping_status(force_fixed_est, include_ssh_checks=True)
                    st.session_state['upcoming_auto_ssh_refresh_epoch'] = now_epoch

                _save_upcoming_snapshot(
                    st.session_state.get('upcoming_rows', []),
                    st.session_state.get('upcoming_meta', ''),
                    st.session_state.get('upcoming_loaded_config', current_config),
                )
            except Exception as exc:
                st.session_state['upcoming_error'] = _format_runtime_error(exc, 'Upcoming Tasks')
                st.session_state['upcoming_loaded_config'] = current_config

    st.divider()
    st.subheader('Upcoming Tasks (DNAC Schedule List)')
    st.caption('First successful DNAC pull is kept as a persistent table snapshot for this date/config. Ongoing updates only change IP Reachability (Online/Offline) and Upgrade Status (Success/Pending/Fail).')

    if st.session_state.get('upcoming_error'):
        st.error(st.session_state['upcoming_error'])

    if st.session_state.get('upcoming_meta'):
        st.caption(st.session_state['upcoming_meta'])
    upcoming_loaded_at_text = str(st.session_state.get('upcoming_loaded_at_text') or '').strip()
    upcoming_loaded_at_epoch = float(st.session_state.get('upcoming_loaded_at_epoch') or 0.0)
    if upcoming_loaded_at_text and upcoming_loaded_at_epoch > 0:
        elapsed_text = _format_elapsed_duration(datetime.now().timestamp() - upcoming_loaded_at_epoch)
        st.caption(f'Upcoming snapshot loaded at: {upcoming_loaded_at_text} ({elapsed_text} ago)')
    if st.session_state.get('upcoming_ping_checked_at'):
        st.caption(f"Ping status last checked: {st.session_state['upcoming_ping_checked_at']}")

    upcoming_rows = st.session_state.get('upcoming_rows', [])
    if not upcoming_rows:
        st.info('No upcoming tasks available yet for the selected date.')
        return

    available_sites = sorted({
        _extract_site_code(str(row.get('description') or ''))
        for row in upcoming_rows
        if str(row.get('description') or '').strip()
    })
    selected_sites = st.multiselect(
        'Upcoming Site Filter',
        options=available_sites,
        default=available_sites,
        key='upcoming_site_filter',
    )
    show_description_column = st.checkbox('Show Description column', value=False, key='upcoming_show_description')

    filtered_upcoming_rows = [
        row
        for row in upcoming_rows
        if not selected_sites or _extract_site_code(str(row.get('description') or '')) in set(selected_sites)
    ]

    online_count = sum(1 for row in filtered_upcoming_rows if row.get('ip_reachability') == 'Online')
    offline_count = sum(1 for row in filtered_upcoming_rows if row.get('ip_reachability') == 'Offline')

    u1, u2, u3 = st.columns(3)
    u1.metric('Filtered Upcoming Rows', len(filtered_upcoming_rows))
    u2.metric('Online', online_count)
    u3.metric(f'Offline ({offline_count})', offline_count)

    upcoming_display_rows = []
    for row in filtered_upcoming_rows:
        image_value = _best_image_version(row)
        status_value = _best_upgrade_status(row, force_fixed_est)
        display_row = {
            'Start Date/ Time': row.get('start_date_time_est', ''),
            'Device Name': row.get('device_name', ''),
            'IP Address': row.get('ip_address', ''),
            'IP Reachability': row.get('ip_reachability', 'Unknown'),
            'Current Image Version': image_value,
            'Upgrade Status': status_value,
        }
        if show_description_column:
            display_row = {'Description': row.get('description', ''), **display_row}
        upcoming_display_rows.append(display_row)

    upcoming_df = pd.DataFrame(upcoming_display_rows)
    upcoming_styled = upcoming_df.style.map(_upcoming_status_style, subset=['IP Reachability'])
    upcoming_styled = upcoming_styled.map(_upgrade_status_style, subset=['Upgrade Status'])
    st.dataframe(upcoming_styled, width='stretch')


@st.fragment(run_every='30s')
def _render_outlook_panel(task_date_text: str) -> None:
    st.divider()
    st.subheader('Outlook Upgrade Messages')

    outlook_enabled = os.getenv('OUTLOOK_ENABLED', 'true').strip().lower() in ('1', 'true', 'yes', 'on')
    if not outlook_enabled:
        st.info('Outlook integration is disabled. Set OUTLOOK_ENABLED=true in .env to use this panel.')
        st.dataframe(pd.DataFrame(columns=['Device', 'Upgrade Resukts', 'Missing Interfaces']), width='stretch')
        return

    outlook_lookback_hours = _env_int('OUTLOOK_LOOKBACK_HOURS', 24, minimum=1)
    outlook_max_messages = _env_int('OUTLOOK_MAX_MESSAGES', 200, minimum=10)
    outlook_mailbox = str(os.getenv('OUTLOOK_MAILBOX', '')).strip()
    outlook_folder = str(os.getenv('OUTLOOK_FOLDER', 'Inbox')).strip() or 'Inbox'
    outlook_class_filter = ('Failure', 'Success', 'Missing Ports', 'Upgrade Report')
    outlook_subject_contains = str(os.getenv('OUTLOOK_SUBJECT_CONTAINS', 'Cisco Upgrade Report')).strip()
    outlook_sender_contains = str(os.getenv('OUTLOOK_SENDER_CONTAINS', 'network-upgrade-reports@digitalrealty.com')).strip()
    outlook_body_contains = str(os.getenv('OUTLOOK_BODY_CONTAINS', '')).strip()
    outlook_to_contains = str(os.getenv('OUTLOOK_TO_CONTAINS', '')).strip()
    outlook_correlated_only = str(os.getenv('OUTLOOK_CORRELATED_ONLY', 'false')).strip().lower() in ('1', 'true', 'yes', 'on')
    outlook_auto_refresh = str(os.getenv('OUTLOOK_AUTO_REFRESH', 'false')).strip().lower() in ('1', 'true', 'yes', 'on')
    outlook_auto_refresh_minutes = _env_int('OUTLOOK_AUTO_REFRESH_MINUTES', 5, minimum=1)

    outlook_config = {
        'lookback_hours': int(outlook_lookback_hours),
        'max_messages': int(outlook_max_messages),
        'mailbox_name': str(outlook_mailbox or '').strip(),
        'folder_path': str(outlook_folder or '').strip(),
        'class_filter': tuple(outlook_class_filter),
        'subject_contains': str(outlook_subject_contains or '').strip(),
        'sender_contains': str(outlook_sender_contains or '').strip(),
        'body_contains': str(outlook_body_contains or '').strip(),
        'to_contains': str(outlook_to_contains or '').strip(),
        'task_date_text': task_date_text,
    }
    outlook_session_id = str(st.session_state.get('session_run_id') or 'outlook-default')
    completed_result = _collect_outlook_background_result(outlook_session_id)
    if completed_result is not None:
        st.session_state['outlook_bg_inflight'] = False
        st.session_state['outlook_bg_started_epoch'] = 0.0

        payload = completed_result.get('payload')
        if isinstance(payload, dict):
            st.session_state['outlook_rows'] = payload.get('rows', [])
            st.session_state['outlook_meta'] = payload.get('meta', '')
            st.session_state['outlook_last_run'] = payload.get('last_run', datetime.now().strftime('%Y-%m-%d %H:%M:%S'))
            st.session_state['outlook_error'] = ''
            st.session_state['outlook_loaded_config'] = completed_result.get('config', outlook_config)
            st.session_state['outlook_last_auto_refresh'] = datetime.now().strftime('%Y-%m-%d %H:%M:%S')
            st.session_state['outlook_last_poll_epoch'] = float(payload.get('fetched_epoch') or datetime.now().timestamp())
        else:
            st.session_state['outlook_error'] = str(completed_result.get('error') or 'Outlook background refresh failed.')

    monitor_interval_seconds = max(30, int(outlook_auto_refresh_minutes) * 60)
    now_epoch = datetime.now().timestamp()
    should_start_background_poll = False

    if not st.session_state.get('outlook_bg_inflight'):
        config_changed = st.session_state.get('outlook_loaded_config') != outlook_config
        has_rows = bool(st.session_state.get('outlook_rows'))
        last_poll_epoch = float(st.session_state.get('outlook_last_poll_epoch') or 0.0)

        if config_changed or not has_rows:
            should_start_background_poll = True
        elif outlook_auto_refresh and (now_epoch - last_poll_epoch) >= monitor_interval_seconds:
            should_start_background_poll = True

    if should_start_background_poll:
        started = _start_outlook_background_refresh(
            outlook_session_id,
            outlook_config,
            known_devices=_build_known_device_map(),
        )
        if started:
            st.session_state['outlook_bg_inflight'] = True
            st.session_state['outlook_bg_started_epoch'] = now_epoch

    displayed_outlook_rows = st.session_state.get('outlook_rows', [])
    if outlook_correlated_only:
        displayed_outlook_rows = [row for row in displayed_outlook_rows if row.get('correlated') == 'Yes']

    if displayed_outlook_rows:
        outlook_display_rows = [
            {
                'Device': row.get('device_name', ''),
                'Upgrade Resukts': row.get('upgrade_result', '') or row.get('classification', ''),
                'Missing Interfaces': row.get('missing_interfaces', ''),
            }
            for row in displayed_outlook_rows
        ]
        st.dataframe(outlook_display_rows, width='stretch')
    elif not st.session_state.get('outlook_error'):
        st.dataframe(pd.DataFrame(columns=['Device', 'Upgrade Resukts', 'Missing Interfaces']), width='stretch')
        st.info('No Outlook messages matched the current filters.')
    else:
        st.dataframe(pd.DataFrame(columns=['Device', 'Upgrade Resukts', 'Missing Interfaces']), width='stretch')

    if st.session_state.get('outlook_error'):
        st.error(st.session_state['outlook_error'])
    if st.session_state.get('outlook_meta'):
        st.caption(st.session_state['outlook_meta'])
    if st.session_state.get('outlook_last_run'):
        st.caption(f"Last Outlook load: {st.session_state['outlook_last_run']}")
    if st.session_state.get('outlook_bg_inflight'):
        started_epoch = float(st.session_state.get('outlook_bg_started_epoch') or 0.0)
        if started_epoch > 0:
            elapsed = _format_elapsed_duration(datetime.now().timestamp() - started_epoch)
            st.caption(f'Outlook background monitor: polling for new emails ({elapsed})')
        else:
            st.caption('Outlook background monitor: polling for new emails')


def _collect_precheck_files(upcoming_rows: List[Dict]) -> List[Dict]:
    results: List[Dict] = []
    seen = set()
    for row in upcoming_rows:
        ip_address = str(row.get('ip_address') or '').strip()
        device_name = str(row.get('device_name') or '').strip()
        description = str(row.get('description') or '')
        if not ip_address or ip_address == 'Unknown':
            continue
        key = (ip_address, description)
        if key in seen:
            continue
        seen.add(key)
        results.append(
            collect_device_precheck_outputs(
                ip_address=ip_address,
                device_name=device_name or ip_address,
                site_code=_extract_site_code(description),
            )
        )
    return results


def _monitor_key(row: Dict) -> str:
    return f"{row.get('task_id', '')}|{row.get('ip_address', '')}|{row.get('description', '')}"


def _build_ping_monitor_rows(upcoming_rows: List[Dict]) -> List[Dict]:
    monitor_rows: List[Dict] = []
    seen = set()
    for row in upcoming_rows:
        ip_address = str(row.get('ip_address') or '').strip()
        if not ip_address or ip_address == 'Unknown':
            continue
        key = _monitor_key(row)
        if key in seen:
            continue
        seen.add(key)
        monitor_rows.append(
            {
                'monitor_key': key,
                'description': row.get('description', ''),
                'start_date_time_est': row.get('start_date_time_est', ''),
                'sort_start_time': row.get('sort_start_time', ''),
                'device_name': row.get('device_name', ''),
                'ip_address': ip_address,
                'baseline_image_version': row.get('current_image_version', 'Unknown'),
                'current_image_version': row.get('current_image_version', 'Unknown'),
                'monitor_status': 'Waiting for scheduled time',
                'last_ping_status': 'Unknown',
                'last_ping_output': '',
                'last_checked_at': '',
                'offline_seen_at': '',
                'online_restored_at': '',
                'post_image_version': '',
                'upgrade_result': '',
                'ssh_validation_error': '',
                'postcheck_file_path': '',
                'postcheck_status': '',
            }
        )
    return monitor_rows


def _monitor_now(force_fixed_est: bool) -> datetime:
    tz_name = 'Etc/GMT+5' if force_fixed_est else 'America/New_York'
    return datetime.now(ZoneInfo(tz_name))


def _run_ping_monitor_cycle(task_date_text: str, force_fixed_est: bool, status_filter: List[str]) -> None:
    task_date = _parse_task_date(task_date_text)
    tz_name = 'Etc/GMT+5' if force_fixed_est else 'America/New_York'
    tz_label = 'EST' if force_fixed_est else ''
    history_cached = True
    completed_backfill_rows: List[Dict] = []

    # Keep the first loaded upcoming snapshot stable and only update runtime status fields.
    if not st.session_state.get('upcoming_rows'):
        clients = build_clients()
        current_upcoming_rows = get_upcoming_tasks_for_date(clients, task_date, tz_name=tz_name, tz_label_override=tz_label)
        completed_rows, history_cached = _get_cached_upgrade_history_rows(clients, task_date, tz_name=tz_name, tz_label=tz_label)
        completed_backfill_rows = _build_completed_backfill_rows(completed_rows, task_date, tz_name=tz_name, min_hour_local=2)
        current_upcoming_rows.extend(completed_backfill_rows)
        st.session_state['upcoming_rows'] = [_initial_upcoming_state(row) for row in current_upcoming_rows]

    _refresh_upcoming_ping_status(force_fixed_est)
    zone_note = 'fixed EST (UTC-5)' if force_fixed_est else 'Eastern local time (EST/EDT)'
    st.session_state['upcoming_meta'] = (
        f'Upcoming list loaded for {task_date} using {zone_note}; '
        f'first snapshot retained, runtime updates=ip_reachability+upgrade_status, '
        f'backfilled={len(completed_backfill_rows)}, history_cache={"hit" if history_cached else "refresh"}'
    )
    current_upcoming_keys = {_monitor_key(row) for row in st.session_state.get('upcoming_rows', [])}

    now = _monitor_now(force_fixed_est)
    active_rows: List[Dict] = []
    completed_rows = list(st.session_state.get('ping_monitor_completed', []))

    for row in st.session_state.get('ping_monitor_active', []):
        updated = dict(row)
        updated['last_checked_at'] = now.strftime('%Y-%m-%d %H:%M:%S')
        try:
            sort_start_time = str(updated.get('sort_start_time') or '')
            try:
                scheduled_dt = datetime.fromisoformat(sort_start_time) if sort_start_time else None
            except ValueError:
                scheduled_dt = None

            if scheduled_dt is None or now < scheduled_dt:
                updated['monitor_status'] = 'Waiting for scheduled time'
                active_rows.append(updated)
                continue

            if updated['monitor_key'] in current_upcoming_keys:
                updated['monitor_status'] = 'Past schedule time; still in upcoming task list'
                active_rows.append(updated)
                continue

            is_online, ping_output = icmp_ping(updated['ip_address'])
            updated['last_ping_status'] = 'Online' if is_online else 'Offline'
            updated['last_ping_output'] = ping_output

            if not updated.get('offline_seen_at'):
                if is_online:
                    updated['monitor_status'] = 'Waiting for device to go offline'
                    active_rows.append(updated)
                    continue
                updated['offline_seen_at'] = now.isoformat()
                updated['monitor_status'] = 'Offline detected'
                active_rows.append(updated)
                continue

            if not updated.get('online_restored_at'):
                if not is_online:
                    updated['monitor_status'] = 'Offline'
                    active_rows.append(updated)
                    continue
                updated['online_restored_at'] = now.isoformat()
                updated['monitor_status'] = 'Online restored; holding for 4 minutes'
                active_rows.append(updated)
                continue

            try:
                restored_dt = datetime.fromisoformat(str(updated['online_restored_at']))
            except ValueError:
                restored_dt = now
                updated['online_restored_at'] = now.isoformat()
            hold_until = restored_dt + timedelta(minutes=4)
            if now < hold_until:
                remaining = int((hold_until - now).total_seconds())
                updated['monitor_status'] = f'Online restored; validating in {remaining}s'
                active_rows.append(updated)
                continue

            post_image_version, ssh_error = ssh_show_version(updated['ip_address'])
            updated['post_image_version'] = post_image_version or 'Unknown'
            updated['ssh_validation_error'] = ssh_error or ''
            before = str(updated.get('baseline_image_version') or 'Unknown').strip()
            after = str(updated.get('post_image_version') or 'Unknown').strip()
            updated['upgrade_result'] = 'Success' if before != 'Unknown' and after != 'Unknown' and before != after else 'Failure'
            postcheck_result = collect_device_precheck_outputs(
                ip_address=updated['ip_address'],
                device_name=updated.get('device_name') or updated['ip_address'],
                site_code=_extract_site_code(updated.get('description', '')),
                file_suffix='postcheck',
            )
            updated['postcheck_file_path'] = postcheck_result.get('file_path', '')
            updated['postcheck_status'] = postcheck_result.get('status', '')
            updated['monitor_status'] = 'Completed'
            completed_rows.append(updated)
        except Exception as exc:
            updated['monitor_status'] = f'Monitor error: {exc}'
            active_rows.append(updated)

    st.session_state['ping_monitor_active'] = active_rows
    st.session_state['ping_monitor_completed'] = completed_rows
    st.session_state['ping_monitor_last_run'] = now.strftime('%Y-%m-%d %H:%M:%S')
    st.session_state['ping_monitor_error'] = ''


st.set_page_config(page_title='Patching Validator', layout='wide')

# Keep the page readable during Streamlit reruns by disabling stale/running fade.
disable_rerun_fade = str(os.getenv('DISABLE_RERUN_FADE', 'true')).strip().lower() in ('1', 'true', 'yes', 'on')
if disable_rerun_fade:
    st.markdown(
        """
        <style>
        .stApp[data-teststate="stale"],
        .stApp[data-teststate="running"],
        .stApp[data-test-script-state="running"] {
            opacity: 1 !important;
            filter: none !important;
        }

        [data-testid="stAppViewContainer"],
        [data-testid="stAppViewContainer"] * {
            transition: none !important;
            animation-duration: 0s !important;
        }
        </style>
        """,
        unsafe_allow_html=True,
    )

st.title('Patching Validator')
st.caption('Lookup DNAC device details and latest image update outcome')

if 'task_rows' not in st.session_state:
    st.session_state['task_rows'] = []
if 'task_counts' not in st.session_state:
    st.session_state['task_counts'] = {'In Progress': 0, 'Success': 0, 'Failure': 0, 'Upcoming Tasks': 0}
if 'task_meta' not in st.session_state:
    st.session_state['task_meta'] = ''
if 'task_error' not in st.session_state:
    st.session_state['task_error'] = ''
if 'upcoming_rows' not in st.session_state:
    st.session_state['upcoming_rows'] = []
if 'upcoming_meta' not in st.session_state:
    st.session_state['upcoming_meta'] = ''
if 'upcoming_error' not in st.session_state:
    st.session_state['upcoming_error'] = ''
if 'task_loaded_config' not in st.session_state:
    st.session_state['task_loaded_config'] = None
if 'upcoming_loaded_config' not in st.session_state:
    st.session_state['upcoming_loaded_config'] = None
if 'upcoming_last_refresh_epoch' not in st.session_state:
    st.session_state['upcoming_last_refresh_epoch'] = 0.0
if 'upcoming_retry_after_epoch' not in st.session_state:
    st.session_state['upcoming_retry_after_epoch'] = 0.0
if 'upcoming_ping_checked_at' not in st.session_state:
    st.session_state['upcoming_ping_checked_at'] = ''
if 'precheck_rows' not in st.session_state:
    st.session_state['precheck_rows'] = []
if 'precheck_error' not in st.session_state:
    st.session_state['precheck_error'] = ''
if 'ping_monitor_active' not in st.session_state:
    st.session_state['ping_monitor_active'] = []
if 'ping_monitor_completed' not in st.session_state:
    st.session_state['ping_monitor_completed'] = []
if 'ping_monitor_error' not in st.session_state:
    st.session_state['ping_monitor_error'] = ''
if 'ping_monitor_last_run' not in st.session_state:
    st.session_state['ping_monitor_last_run'] = ''
if 'outlook_rows' not in st.session_state:
    st.session_state['outlook_rows'] = []
if 'outlook_meta' not in st.session_state:
    st.session_state['outlook_meta'] = ''
if 'outlook_error' not in st.session_state:
    st.session_state['outlook_error'] = ''
if 'outlook_last_run' not in st.session_state:
    st.session_state['outlook_last_run'] = ''
if 'outlook_last_auto_refresh' not in st.session_state:
    st.session_state['outlook_last_auto_refresh'] = ''
if 'outlook_last_poll_epoch' not in st.session_state:
    st.session_state['outlook_last_poll_epoch'] = 0.0
if 'outlook_loaded_config' not in st.session_state:
    st.session_state['outlook_loaded_config'] = None
if 'outlook_folder_options' not in st.session_state:
    st.session_state['outlook_folder_options'] = []
if 'outlook_folder_error' not in st.session_state:
    st.session_state['outlook_folder_error'] = ''
if 'outlook_bg_inflight' not in st.session_state:
    st.session_state['outlook_bg_inflight'] = False
if 'outlook_bg_started_epoch' not in st.session_state:
    st.session_state['outlook_bg_started_epoch'] = 0.0
if 'upgrade_history_cache' not in st.session_state:
    st.session_state['upgrade_history_cache'] = {}
if 'session_run_id' not in st.session_state:
    st.session_state['session_run_id'] = uuid4().hex
if 'upcoming_bg_inflight' not in st.session_state:
    st.session_state['upcoming_bg_inflight'] = False
if 'upcoming_bg_last_started_at' not in st.session_state:
    st.session_state['upcoming_bg_last_started_at'] = ''
if 'upcoming_bg_started_epoch' not in st.session_state:
    st.session_state['upcoming_bg_started_epoch'] = 0.0
if 'upcoming_snapshot_saved_at' not in st.session_state:
    st.session_state['upcoming_snapshot_saved_at'] = ''
if 'upcoming_auto_live_refresh_epoch' not in st.session_state:
    st.session_state['upcoming_auto_live_refresh_epoch'] = 0.0
if 'upcoming_auto_ssh_refresh_epoch' not in st.session_state:
    st.session_state['upcoming_auto_ssh_refresh_epoch'] = 0.0
if 'upcoming_loaded_at_epoch' not in st.session_state:
    st.session_state['upcoming_loaded_at_epoch'] = 0.0
if 'upcoming_loaded_at_text' not in st.session_state:
    st.session_state['upcoming_loaded_at_text'] = ''

# Keep the upcoming table anchored after browser/app refresh by restoring persisted snapshot
# only when the current session has not already loaded rows.
snapshot = _load_upcoming_snapshot()
if snapshot:
    snapshot_saved_at = str(snapshot.get('saved_at') or '').strip()
    should_sync_snapshot = not st.session_state.get('upcoming_rows')
    if should_sync_snapshot:
        st.session_state['upcoming_rows'] = snapshot.get('rows', [])
        st.session_state['upcoming_meta'] = snapshot.get('meta', st.session_state.get('upcoming_meta', ''))
        st.session_state['upcoming_loaded_config'] = snapshot.get(
            'loaded_config',
            st.session_state.get('upcoming_loaded_config'),
        )
        st.session_state['upcoming_snapshot_saved_at'] = snapshot_saved_at
        if snapshot_saved_at:
            try:
                loaded_dt = datetime.strptime(snapshot_saved_at, '%Y-%m-%d %H:%M:%S')
                st.session_state['upcoming_loaded_at_epoch'] = loaded_dt.timestamp()
                st.session_state['upcoming_loaded_at_text'] = snapshot_saved_at
            except ValueError:
                pass

with st.sidebar:
    st.subheader('DNAC Regions')
    st.write('Configured regions are loaded from `.env`.')
    st.write('- EMEA (DNAC AMES)')
    st.write('- US (DNAC WEST)')
    st.write('- APAC (DNAC EAST)')

col1, col2 = st.columns([3, 1])
with col1:
    device_name = st.text_input('Device Hostname / FQDN', placeholder='cs01-los10-9500.digitalrealtytrust.com')
with col2:
    search = False

use_cache = st.checkbox('Use lookup cache (faster repeated searches)', value=True)
use_ssh_validation = st.checkbox('Validate image via SSH (show version)', value=True)

if search:
    name = (device_name or '').strip()
    if not name:
        st.warning('Enter a device hostname or FQDN.')
    else:
        with st.spinner('Querying DNAC regions...'):
            try:
                clients = build_clients()
                if not clients:
                    st.error('No DNAC regions configured. Update `.env` with region base URLs.')
                else:
                    result = find_device_by_name(
                        clients,
                        name,
                        use_cache=use_cache,
                        use_ssh_validation=use_ssh_validation,
                    )
                    if not result:
                        st.error(f'Device not found: {name}')
                    else:
                        m1, m2, m3, m4 = st.columns(4)
                        m1.metric('IP Address', result.ip_address)
                        m2.metric('Region', result.region)
                        m3.metric('Image Version', result.image_version)
                        m4.metric('Online Status', result.online_status)

                        st.subheader('Update Status')
                        st.write(f'Last image update outcome: **{result.last_update_outcome}**')
                        st.write(f'Last update time: **{result.last_update_time}**')
                        if use_ssh_validation:
                            st.caption('Outcome is confirmed by comparing DNAC image version to direct SSH `show version`.')

                        st.subheader('Direct SSH Validation')
                        st.write(f'SSH validated: **{result.ssh_validated}**')
                        st.write(f'SSH image/version (`show version`): **{result.ssh_image_or_version}**')
                        if result.ssh_error:
                            st.warning(f'SSH detail: {result.ssh_error}')

                        with st.expander('Raw update record (for troubleshooting)'):
                            st.json(result.raw_update_record)
            except Exception as exc:
                st.exception(exc)

st.divider()
st.subheader('Batch Lookup (CSV)')
st.caption('Upload a CSV with a `device`, `device_name`, `hostname`, or `fqdn` column.')

upload = st.file_uploader('Device CSV', type=['csv'])
run_batch = False

if run_batch:
    if not upload:
        st.warning('Upload a CSV file first.')
    else:
        try:
            names = _extract_device_names_from_csv(upload.getvalue())
            if not names:
                st.error('No device names found in the CSV.')
            else:
                with st.spinner(f'Looking up {len(names)} devices...'):
                    clients = build_clients()
                    rows = find_many_devices(
                        clients,
                        names,
                        use_cache=use_cache,
                        use_ssh_validation=use_ssh_validation,
                    )
                    st.dataframe(rows, width='stretch')
        except Exception as exc:
            st.exception(exc)

st.divider()
st.subheader('Upgrade Task Schedule & Status')
st.caption('Task times shown in Eastern Time (EST/EDT based on date).')

eastern_now = datetime.now(ZoneInfo('America/New_York'))
force_fixed_est = st.checkbox('Force fixed EST label year-round (UTC-5)', value=False)
today_task_date = eastern_now.date()
if 'upcoming_query_date' not in st.session_state:
    st.session_state['upcoming_query_date'] = today_task_date

date_col1, date_col2 = st.columns([3, 1])
with date_col1:
    selected_task_date = st.date_input(
        'Upcoming Task Date',
        value=st.session_state.get('upcoming_query_date', today_task_date),
        key='upcoming_query_date',
    )
with date_col2:
    st.write('')
    st.write('')
    st.button('Use Today', use_container_width=True, on_click=_set_upcoming_date_today)

task_date_text = selected_task_date.isoformat()
status_filter = st.multiselect(
    'Task Status Filter',
    options=['Upcoming Tasks', 'Failure', 'Success', 'In Progress'],
    default=['Upcoming Tasks', 'Failure', 'Success'],
)
auto_refresh_loaded_data = st.checkbox('Auto refresh loaded task panels when configuration changes', value=True)
task_action_col1, task_action_col2, task_action_col3, task_action_col4 = st.columns(4)
with task_action_col1:
    refresh_upcoming_live = st.button('Refresh Upcoming From DNAC', use_container_width=True)
with task_action_col2:
    refresh_upcoming_ping = st.button('Refresh IP Reachability', use_container_width=True)
with task_action_col3:
    refresh_upcoming_image = st.button('Refresh Image Version via SSH', use_container_width=True)
with task_action_col4:
    reset_upcoming_snapshot = st.button('Reset Snapshot', use_container_width=True)
load_tasks = False
capture_prechecks = False
start_ping_monitor = False
run_ping_monitor = False
clear_ping_monitor = False
auto_refresh_ping_monitor = st.checkbox('Run ping monitor cycle on every rerun', value=False)

current_config = _current_task_config(task_date_text, force_fixed_est, status_filter)

if load_tasks:
    try:
        try:
            _parse_task_date(task_date_text)
        except ValueError:
            st.error('Invalid date format. Use YYYY-MM-DD (example: 2026-07-01).')
            st.stop()
        with st.spinner('Loading upgrade tasks...'):
            _load_task_rows(task_date_text, force_fixed_est, status_filter)
    except Exception as exc:
        st.session_state['task_error'] = _format_runtime_error(exc, 'Upgrade Task Schedule & Status')

try:
    try:
        _parse_task_date(task_date_text)
    except ValueError:
        st.error('Invalid date format. Use YYYY-MM-DD (example: 2026-07-01).')
        st.stop()

    # Upcoming is snapshot-first and refreshes only on explicit manual action.
    st.session_state['upcoming_bg_inflight'] = False
    st.session_state['upcoming_bg_started_epoch'] = 0.0

    if refresh_upcoming_live:
        with st.spinner('Refreshing upcoming tasks from DNAC...'):
            _load_upcoming_rows(task_date_text, force_fixed_est, status_filter)
            _refresh_upcoming_ping_status(force_fixed_est, include_ssh_checks=False)
            st.session_state['upcoming_last_refresh_epoch'] = datetime.now().timestamp()
            st.session_state['upcoming_retry_after_epoch'] = 0.0

    if reset_upcoming_snapshot:
        with st.spinner('Resetting persistent upcoming snapshot...'):
            st.session_state['upcoming_rows'] = []
            st.session_state['upcoming_meta'] = ''
            st.session_state['upcoming_error'] = ''
            st.session_state['upcoming_loaded_config'] = None
            st.session_state['upcoming_retry_after_epoch'] = 0.0
            st.session_state['upcoming_snapshot_saved_at'] = ''
            _load_upcoming_rows(task_date_text, force_fixed_est, status_filter)
            _refresh_upcoming_ping_status(force_fixed_est, include_ssh_checks=False)
            st.session_state['upcoming_last_refresh_epoch'] = datetime.now().timestamp()

    if refresh_upcoming_ping:
        with st.spinner('Refreshing IP reachability for upcoming tasks...'):
            _refresh_upcoming_ping_status(force_fixed_est, include_ssh_checks=False)
            _save_upcoming_snapshot(
                st.session_state.get('upcoming_rows', []),
                st.session_state.get('upcoming_meta', ''),
                st.session_state.get('upcoming_loaded_config', current_config),
            )
            st.session_state['upcoming_error'] = ''

    if refresh_upcoming_image:
        with st.spinner('Refreshing image versions via SSH for reachable upcoming devices...'):
            _refresh_upcoming_ping_status(
                force_fixed_est,
                include_ssh_checks=True,
                force_ssh_recheck=True,
            )
            _save_upcoming_snapshot(
                st.session_state.get('upcoming_rows', []),
                st.session_state.get('upcoming_meta', ''),
                st.session_state.get('upcoming_loaded_config', current_config),
            )
            st.session_state['upcoming_error'] = ''
except Exception as exc:
    st.session_state['upcoming_error'] = _format_runtime_error(exc, 'Upcoming Tasks')
    st.session_state['upcoming_loaded_config'] = current_config
    st.session_state['upcoming_retry_after_epoch'] = datetime.now().timestamp() + 120.0

if auto_refresh_loaded_data and st.session_state.get('task_loaded_config') and st.session_state.get('task_loaded_config') != current_config:
    try:
        with st.spinner('Auto refreshing task status...'):
            _load_task_rows(task_date_text, force_fixed_est, status_filter)
    except Exception as exc:
        st.session_state['task_error'] = _format_runtime_error(exc, 'Upgrade Task Schedule & Status')

loaded_upcoming_config = st.session_state.get('upcoming_loaded_config')
loaded_upcoming_task_date_text = ''
if isinstance(loaded_upcoming_config, dict):
    loaded_upcoming_task_date_text = str(loaded_upcoming_config.get('task_date_text') or '').strip()

if (
    auto_refresh_loaded_data
    and (
        not st.session_state.get('upcoming_rows')
        or (loaded_upcoming_task_date_text and loaded_upcoming_task_date_text != task_date_text)
    )
    and not refresh_upcoming_live
    and datetime.now().timestamp() >= float(st.session_state.get('upcoming_retry_after_epoch') or 0.0)
):
    try:
        with st.spinner('Auto refreshing upcoming tasks...'):
            _load_upcoming_rows(task_date_text, force_fixed_est, status_filter)
            _refresh_upcoming_ping_status(force_fixed_est, include_ssh_checks=False)
            st.session_state['upcoming_last_refresh_epoch'] = datetime.now().timestamp()
    except Exception as exc:
        st.session_state['upcoming_error'] = _format_runtime_error(exc, 'Upcoming Tasks')
        st.session_state['upcoming_loaded_config'] = current_config
        st.session_state['upcoming_retry_after_epoch'] = datetime.now().timestamp() + 120.0

if capture_prechecks:
    try:
        with st.spinner('Collecting pre-task show commands...'):
            st.session_state['precheck_rows'] = _collect_precheck_files(st.session_state.get('upcoming_rows', []))
            st.session_state['precheck_error'] = ''
    except Exception as exc:
        st.session_state['precheck_error'] = str(exc)

if start_ping_monitor:
    st.session_state['ping_monitor_active'] = _build_ping_monitor_rows(st.session_state.get('upcoming_rows', []))
    st.session_state['ping_monitor_completed'] = []
    st.session_state['ping_monitor_error'] = ''
    st.session_state['ping_monitor_last_run'] = ''

if run_ping_monitor:
    try:
        with st.spinner('Running ping monitor cycle...'):
            _run_ping_monitor_cycle(task_date_text, force_fixed_est, status_filter)
    except Exception as exc:
        st.session_state['ping_monitor_error'] = str(exc)

if clear_ping_monitor:
    st.session_state['ping_monitor_active'] = []
    st.session_state['ping_monitor_completed'] = []
    st.session_state['ping_monitor_error'] = ''
    st.session_state['ping_monitor_last_run'] = ''

_render_outlook_panel(task_date_text)

if auto_refresh_ping_monitor and st.session_state.get('ping_monitor_active'):
    try:
        _run_ping_monitor_cycle(task_date_text, force_fixed_est, status_filter)
    except Exception as exc:
        st.session_state['ping_monitor_error'] = str(exc)

if st.session_state.get('task_error'):
    st.error(st.session_state['task_error'])

if st.session_state.get('task_meta'):
    st.caption(st.session_state['task_meta'])

rows = st.session_state.get('task_rows', [])
counts = st.session_state.get('task_counts', {'In Progress': 0, 'Success': 0, 'Failure': 0, 'Upcoming Tasks': 0})

c1, c2, c3, c4 = st.columns(4)
c1.metric('In Progress', counts.get('In Progress', 0))
c2.metric('Success', counts.get('Success', 0))
c3.metric('Failure', counts.get('Failure', 0))
c4.metric('Upcoming Tasks', counts.get('Upcoming Tasks', 0))

if not rows:
    pass
else:
    st.dataframe(rows, width='stretch')

if st.session_state.get('precheck_error'):
    st.error(st.session_state['precheck_error'])

if st.session_state.get('ping_monitor_error'):
    st.error(st.session_state['ping_monitor_error'])

_render_upcoming_panel(task_date_text, force_fixed_est, current_config)

precheck_rows = st.session_state.get('precheck_rows', [])
if precheck_rows:
    st.subheader('Pre-Task Command Capture')
    st.dataframe(precheck_rows, width='stretch')

active_ping_rows = st.session_state.get('ping_monitor_active', [])
completed_ping_rows = st.session_state.get('ping_monitor_completed', [])

show_icmp_section = bool(active_ping_rows or completed_ping_rows or st.session_state.get('ping_monitor_last_run'))
if show_icmp_section:
    st.divider()
    st.subheader('ICMP Upgrade Monitor')
    if st.session_state.get('ping_monitor_last_run'):
        st.caption(f"Last monitor cycle: {st.session_state['ping_monitor_last_run']}")

    if active_ping_rows:
        active_display_rows = [
            {
                'Description': row.get('description', ''),
                'Start Date/ Time': row.get('start_date_time_est', ''),
                'Device Name': row.get('device_name', ''),
                'IP Address': row.get('ip_address', ''),
                'Image Before': row.get('baseline_image_version', ''),
                'Ping Status': row.get('last_ping_status', ''),
                'Ping Output': row.get('last_ping_output', ''),
                'Monitor Status': row.get('monitor_status', ''),
                'Last Checked': row.get('last_checked_at', ''),
            }
            for row in active_ping_rows
        ]
        st.dataframe(active_display_rows, width='stretch')

    if completed_ping_rows:
        st.subheader('Completed Upgrade Validation')
        completed_display_rows = [
            {
                'Description': row.get('description', ''),
                'Device Name': row.get('device_name', ''),
                'IP Address': row.get('ip_address', ''),
                'Image Before': row.get('baseline_image_version', ''),
                'Image After': row.get('post_image_version', ''),
                'Upgrade Result': row.get('upgrade_result', ''),
                'Last Ping Output': row.get('last_ping_output', ''),
                'Postcheck Status': row.get('postcheck_status', ''),
                'Postcheck File': row.get('postcheck_file_path', ''),
                'SSH Error': row.get('ssh_validation_error', ''),
            }
            for row in completed_ping_rows
        ]
        st.dataframe(completed_display_rows, width='stretch')

render_manual_reloads_panel()



