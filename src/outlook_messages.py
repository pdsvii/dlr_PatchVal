import re
import sys
from datetime import date, datetime, timedelta
from typing import Dict, List, Optional, Tuple

try:
    import pythoncom  # type: ignore
except Exception:
    pythoncom = None


IP_REGEX = re.compile(r"\b(?:\d{1,3}\.){3}\d{1,3}\b")
HOST_REGEX = re.compile(r"\b[a-zA-Z0-9][a-zA-Z0-9._-]{2,}\b")
UPGRADE_TABLE_ROW_REGEX = re.compile(
    r"^\s*(?P<site>\S+)\s+"
    r"(?P<device>\S+)\s+"
    r"(?P<start_time>\d{2}:\d{2}:\d{2})\s+"
    r"(?P<completion_time>\d{2}:\d{2}:\d{2})\s+"
    r"(?P<new_version>\S+)\s+"
    r"(?P<upgrade_result>\S+)"
    r"(?:\s+(?P<missing_interfaces>.+))?\s*$",
    re.IGNORECASE,
)

SUCCESS_KEYWORDS = (
    'upgrade success',
    'upgrade successful',
    'successfully upgraded',
    'upgrade complete',
    'completed successfully',
    'task success',
    'activation success',
)

FAILURE_KEYWORDS = (
    'upgrade failed',
    'failure',
    'failed',
    'error',
    'aborted',
    'rollback',
    'timed out',
)

MISSING_PORT_KEYWORDS = (
    'missing port',
    'missing ports',
    'port missing',
    'ports missing',
    'no port',
    'no ports',
)


def _clean_text(value: object) -> str:
    if value is None:
        return ''
    return str(value)


def _classify_message(subject: str, body: str) -> str:
    text = f"{subject}\n{body}".lower()

    if any(term in text for term in MISSING_PORT_KEYWORDS):
        return 'Missing Ports'
    if any(term in text for term in FAILURE_KEYWORDS):
        return 'Failure'
    if any(term in text for term in SUCCESS_KEYWORDS):
        return 'Success'
    if 'cisco ios upgrade report results' in text or 'cisco upgrade report' in text:
        return 'Upgrade Report'
    return 'Other'


def _extract_identifiers(subject: str, body: str) -> Tuple[str, str]:
    text = f"{subject}\n{body}"
    ips = list(dict.fromkeys(IP_REGEX.findall(text)))

    host_candidates: List[str] = []
    for token in HOST_REGEX.findall(text):
        token_lower = token.lower()
        if token_lower in {'upgrade', 'success', 'failed', 'failure', 'error', 'task'}:
            continue
        if any(ch.isalpha() for ch in token) and ('-' in token or '.' in token):
            host_candidates.append(token)

    hostnames = list(dict.fromkeys(host_candidates))
    device_name = hostnames[0] if hostnames else ''
    ip_address = ips[0] if ips else ''
    return device_name, ip_address


def _classify_upgrade_result_row(upgrade_result: str, missing_interfaces: str) -> str:
    missing_text = (missing_interfaces or '').strip()
    result = (upgrade_result or '').strip().lower()

    if missing_text:
        return 'Missing Ports'
    if any(term in result for term in ('fail', 'error', 'abort', 'rollback', 'timeout')):
        return 'Failure'
    if any(term in result for term in ('success', 'complete', 'ok', 'passed')):
        return 'Success'
    return 'Upgrade Report'


def _parse_upgrade_report_table_rows(body: str) -> List[Dict[str, str]]:
    lines = [line.rstrip() for line in (body or '').splitlines()]
    header_index = -1
    for index, line in enumerate(lines):
        upper = line.upper()
        if (
            'SITE' in upper
            and 'DEVICE' in upper
            and 'START' in upper
            and 'COMPLETION' in upper
            and 'NEW VERSION' in upper
            and 'UPGRADE RESULTS' in upper
        ):
            header_index = index
            break

    if header_index < 0:
        return []

    parsed_rows: List[Dict[str, str]] = []
    for line in lines[header_index + 1:]:
        if not line.strip():
            continue
        match = UPGRADE_TABLE_ROW_REGEX.match(line)
        if not match:
            continue

        missing_interfaces = (match.group('missing_interfaces') or '').strip()
        parsed_rows.append(
            {
                'site': match.group('site') or '',
                'device_name': match.group('device') or '',
                'start_time': match.group('start_time') or '',
                'completion_time': match.group('completion_time') or '',
                'new_version': match.group('new_version') or '',
                'upgrade_result': match.group('upgrade_result') or '',
                'missing_interfaces': missing_interfaces,
                'classification': _classify_upgrade_result_row(match.group('upgrade_result') or '', missing_interfaces),
            }
        )

    return parsed_rows


def _to_python_datetime(value: object) -> Optional[datetime]:
    if isinstance(value, datetime):
        if value.tzinfo is not None:
            return value.astimezone().replace(tzinfo=None)
        return value
    if value is None:
        return None
    try:
        # Outlook COM datetime values usually support this cast.
        return datetime.fromtimestamp(value.timestamp())
    except Exception:
        pass
    try:
        return datetime.strptime(str(value), '%m/%d/%Y %I:%M:%S %p')
    except Exception:
        return None


def _resolve_folder(namespace, mailbox_name: str, folder_path: str):
    folder = None
    mailbox_name = mailbox_name.strip()
    folder_path = folder_path.strip()
    used_default_inbox = False

    if mailbox_name:
        folder = namespace.Folders.Item(mailbox_name)
    else:
        folder = namespace.GetDefaultFolder(6)  # 6 = Inbox
        used_default_inbox = True

    if folder_path:
        parts = [part for part in re.split(r'[\\/]+', folder_path) if part]
        if used_default_inbox and parts and parts[0].lower() == 'inbox':
            parts = parts[1:]
        for part in parts:
            folder = folder.Folders.Item(part)

    return folder


def list_outlook_folder_paths(mailbox_name: str = '') -> List[str]:
    """Return Outlook folder paths for the selected mailbox/profile."""

    com_initialized = False
    if pythoncom is not None:
        try:
            pythoncom.CoInitialize()
            com_initialized = True
        except Exception:
            com_initialized = False

    try:
        import win32com.client  # type: ignore
    except Exception as exc:
        raise RuntimeError(
            'Outlook integration requires pywin32 on Windows. '
            f'Interpreter={sys.executable}. '
            f'Detail={type(exc).__name__}: {exc}. '
            'Install with: python -m pip install pywin32'
        ) from exc

    try:
        try:
            outlook = win32com.client.Dispatch('Outlook.Application').GetNamespace('MAPI')
            mailbox_name = (mailbox_name or '').strip()
            if mailbox_name:
                root_folder = outlook.Folders.Item(mailbox_name)
            else:
                root_folder = outlook.GetDefaultFolder(6).Parent  # Mailbox root for default profile.
        except Exception as exc:
            raise RuntimeError(f'Unable to access Outlook mailbox folders: {exc}') from exc

        paths: List[str] = []

        def walk(folder, prefix: str) -> None:
            current = f'{prefix}/{folder.Name}' if prefix else str(folder.Name)
            paths.append(current)
            try:
                subfolders = folder.Folders
                count = int(getattr(subfolders, 'Count', 0))
                for index in range(1, count + 1):
                    walk(subfolders.Item(index), current)
            except Exception:
                return

        walk(root_folder, '')
        return sorted(set(paths))
    finally:
        if com_initialized and pythoncom is not None:
            try:
                pythoncom.CoUninitialize()
            except Exception:
                pass


def _extract_sender_fields(item) -> Tuple[str, str]:
    sender_name = _clean_text(getattr(item, 'SenderName', ''))
    sender_email = _clean_text(getattr(item, 'SenderEmailAddress', ''))

    try:
        sender_obj = getattr(item, 'Sender', None)
        exchange_user = sender_obj.GetExchangeUser() if sender_obj else None
        if exchange_user and getattr(exchange_user, 'PrimarySmtpAddress', None):
            sender_email = _clean_text(exchange_user.PrimarySmtpAddress)
    except Exception:
        pass

    sender_value = sender_email or sender_name
    return sender_value, sender_email


def _extract_recipient_fields(item) -> Tuple[str, str]:
    to_text = _clean_text(getattr(item, 'To', ''))
    recipient_emails: List[str] = []

    try:
        recipients = getattr(item, 'Recipients', None)
        count = int(getattr(recipients, 'Count', 0)) if recipients else 0
        for index in range(1, count + 1):
            recipient = recipients.Item(index)
            entry = _clean_text(getattr(recipient, 'Address', ''))
            if not entry:
                entry = _clean_text(getattr(recipient, 'Name', ''))
            if entry:
                recipient_emails.append(entry)
    except Exception:
        pass

    return to_text, '; '.join(recipient_emails)


def load_outlook_upgrade_messages(
    lookback_hours: int = 24,
    max_messages: int = 200,
    mailbox_name: str = '',
    folder_path: str = 'Inbox',
    target_date: Optional[date] = None,
) -> Tuple[List[Dict], str]:
    """Load and classify upgrade-related Outlook messages.

    Returns (rows, meta_message). Raises RuntimeError for environment access issues.
    """

    com_initialized = False
    if pythoncom is not None:
        try:
            pythoncom.CoInitialize()
            com_initialized = True
        except Exception:
            com_initialized = False

    try:
        try:
            import win32com.client  # type: ignore
        except Exception as exc:
            raise RuntimeError(
                'Outlook integration requires pywin32 on Windows. '
                f'Interpreter={sys.executable}. '
                f'Detail={type(exc).__name__}: {exc}. '
                'Install with: python -m pip install pywin32'
            ) from exc

        if lookback_hours <= 0:
            lookback_hours = 24
        if max_messages <= 0:
            max_messages = 200

        cutoff = None if target_date else (datetime.now() - timedelta(hours=lookback_hours))

        try:
            outlook = win32com.client.Dispatch('Outlook.Application').GetNamespace('MAPI')
            folder = _resolve_folder(outlook, mailbox_name=mailbox_name, folder_path=folder_path)
            items = folder.Items
            items.Sort('[ReceivedTime]', True)
        except Exception as exc:
            raise RuntimeError(f'Unable to access Outlook folder: {exc}') from exc

        rows: List[Dict] = []
        scanned = 0

        for item in items:
            if scanned >= max_messages:
                break
            scanned += 1

            try:
                message_class = _clean_text(getattr(item, 'MessageClass', ''))
                if message_class and not message_class.startswith('IPM.Note'):
                    continue

                received_dt = _to_python_datetime(getattr(item, 'ReceivedTime', None))
                if target_date is not None:
                    if not received_dt:
                        continue
                    received_date = received_dt.date()
                    if received_date > target_date:
                        continue
                    if received_date < target_date:
                        # Items are sorted descending; once older than target day we can stop.
                        break
                elif cutoff is not None and received_dt and received_dt < cutoff:
                    # Items are sorted descending; once past cutoff we can stop.
                    break

                subject = _clean_text(getattr(item, 'Subject', ''))
                body = _clean_text(getattr(item, 'Body', ''))
                if not subject and not body:
                    continue

                classification = _classify_message(subject, body)
                if classification == 'Other':
                    continue

                sender_value, sender_email = _extract_sender_fields(item)
                to_recipients, recipient_emails = _extract_recipient_fields(item)
                if classification == 'Upgrade Report':
                    table_rows = _parse_upgrade_report_table_rows(body)
                    if table_rows:
                        for parsed in table_rows:
                            device_name = str(parsed.get('device_name') or '')
                            ip_address = device_name if IP_REGEX.fullmatch(device_name) else ''
                            rows.append(
                                {
                                    'received_time': received_dt.strftime('%Y-%m-%d %H:%M:%S') if received_dt else '',
                                    'classification': str(parsed.get('classification') or 'Upgrade Report'),
                                    'source_classification': 'Upgrade Report',
                                    'subject': subject,
                                    'body': body,
                                    'device_name': device_name,
                                    'ip_address': ip_address,
                                    'sender': sender_value,
                                    'sender_email': sender_email,
                                    'to_recipients': to_recipients,
                                    'recipient_emails': recipient_emails,
                                    'site': str(parsed.get('site') or ''),
                                    'start_time': str(parsed.get('start_time') or ''),
                                    'completion_time': str(parsed.get('completion_time') or ''),
                                    'new_version': str(parsed.get('new_version') or ''),
                                    'upgrade_result': str(parsed.get('upgrade_result') or ''),
                                    'missing_interfaces': str(parsed.get('missing_interfaces') or ''),
                                }
                            )
                        continue

                device_name, ip_address = _extract_identifiers(subject, body)
                rows.append(
                    {
                        'received_time': received_dt.strftime('%Y-%m-%d %H:%M:%S') if received_dt else '',
                        'classification': classification,
                        'subject': subject,
                        'body': body,
                        'device_name': device_name,
                        'ip_address': ip_address,
                        'sender': sender_value,
                        'sender_email': sender_email,
                        'to_recipients': to_recipients,
                        'recipient_emails': recipient_emails,
                    }
                )
            except Exception:
                # Skip malformed items and continue scanning.
                continue

        rows.sort(key=lambda row: row.get('received_time', ''), reverse=True)
        if target_date is not None:
            summary = f'Scanned {scanned} messages from Outlook; matched {len(rows)} upgrade messages for {target_date.isoformat()}.'
        else:
            summary = f'Scanned {scanned} messages from Outlook; matched {len(rows)} upgrade messages.'
        return rows, summary
    finally:
        if com_initialized and pythoncom is not None:
            try:
                pythoncom.CoUninitialize()
            except Exception:
                pass
