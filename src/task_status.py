import re
import os
from concurrent.futures import ThreadPoolExecutor
from datetime import date, datetime
from typing import Any, Dict, List, Tuple
from zoneinfo import ZoneInfo

from src.device_lookup import get_devices_for_region


def _to_task_list(payload: Any) -> List[Dict[str, Any]]:
    if isinstance(payload, dict):
        response = payload.get("response")
        if isinstance(response, list):
            return response
    if isinstance(payload, list):
        return payload
    return []


def _to_record_list(payload: Any) -> List[Dict[str, Any]]:
    if isinstance(payload, dict):
        response = payload.get("response")
        if isinstance(response, list):
            return response
        if isinstance(response, dict):
            # Some DNAC versions return nested arrays under response.
            for key in ("records", "items", "response"):
                candidate = response.get(key)
                if isinstance(candidate, list):
                    return candidate
            return [response]
        for key in ("records", "items"):
            candidate = payload.get(key)
            if isinstance(candidate, list):
                return candidate
    if isinstance(payload, list):
        return payload
    return []


def _to_scheduled_job_list(payload: Any) -> List[Dict[str, Any]]:
    if isinstance(payload, dict):
        response = payload.get("response")
        if isinstance(response, list):
            return response
    if isinstance(payload, list):
        return payload
    return []


def _fetch_schedule_history_rows(client) -> List[Dict[str, Any]]:
    rows: List[Dict[str, Any]] = []
    # Pull non-WAITING/ACTIVE states so completed/failed jobs can still be surfaced.
    query_variants = [
        {
            "source": "EXTERNAL",
            "type": "DEFAULT",
            "sortBy": "lastUpdateTime",
            "order": "DESC",
            "module": "OS_UPDATE",
            "scheduleState": "COMPLETED",
        },
        {
            "source": "EXTERNAL",
            "type": "DEFAULT",
            "sortBy": "lastUpdateTime",
            "order": "DESC",
            "module": "OS_UPDATE",
            "scheduleState": "FAILED",
        },
        {
            "source": "EXTERNAL",
            "type": "DEFAULT",
            "sortBy": "lastUpdateTime",
            "order": "DESC",
            "module": "OS_UPDATE",
            "triggerState": "TRIGGERED",
        },
    ]

    for params in query_variants:
        try:
            payload = client.get("/api/schedule/v4/scheduled-job", params=params)
        except Exception:
            continue
        for row in _to_scheduled_job_list(payload):
            rows.append(row)
    return rows


def _from_epoch_millis(value: Any) -> datetime | None:
    if value is None:
        return None
    try:
        number = float(value)
        if number > 1e12:
            number /= 1000.0
        return datetime.fromtimestamp(number, tz=ZoneInfo("UTC"))
    except Exception:
        return None


def _extract_device_ip(progress: str) -> str:
    match = re.search(r"((?:\d{1,3}\.){3}\d{1,3})", progress or "")
    return match.group(1) if match else ""


def _extract_activation_description(task: Dict[str, Any]) -> str:
    """Extract the SWIM activation job description text shown in DNAC Upcoming Tasks."""
    candidates = [
        str(task.get("progress") or ""),
        str(task.get("data") or ""),
        str(task.get("serviceType") or ""),
    ]
    blob = "\n".join(candidates)
    match = re.search(r"(activation job[^\n\r]*)", blob, flags=re.IGNORECASE)
    if match:
        return match.group(1).strip()
    return ""


def _extract_start_utc(task: Dict[str, Any]) -> datetime | None:
    # Different DNAC builds expose slightly different start timestamp fields.
    for key in ("startTime", "scheduledTime", "lastUpdate", "endTime"):
        value = task.get(key)
        dt_value = _from_epoch_millis(value)
        if dt_value:
            return dt_value
    return None


def _extract_end_utc(task: Dict[str, Any]) -> datetime | None:
    # Keep this broad because DNAC task records differ by version/build.
    for key in ("endTime", "completedTime", "lastUpdate", "lastUpdateTime", "updateTime", "startTime", "scheduledTime"):
        value = task.get(key)
        dt_value = _from_epoch_millis(value)
        if dt_value:
            return dt_value
    return None


def _history_page_size() -> int:
    raw = os.getenv("DNAC_HISTORY_PAGE_SIZE", "200").strip()
    try:
        return max(50, min(1000, int(raw)))
    except ValueError:
        return 200


def _history_max_pages() -> int:
    raw = os.getenv("DNAC_HISTORY_MAX_PAGES", "2").strip()
    try:
        return max(1, min(200, int(raw)))
    except ValueError:
        return 2


def _extract_devices_scheduled(description: str, progress: str) -> int:
    text = f"{description} {progress}".lower()
    # Handle explicit "multiple devices" pattern seen in the DNAC sidebar.
    if "multiple devices" in text:
        return 2
    # Fall back to pending/total hints if present.
    for pattern in (r"devices\s*=\s*(\d+)", r"total\s*=\s*(\d+)", r"pending\s*=\s*(\d+)"):
        match = re.search(pattern, text)
        if match:
            try:
                return max(1, int(match.group(1)))
            except ValueError:
                pass
    return 1


def _format_catalyst_datetime(dt_value: datetime) -> str:
    # Match Catalyst sidebar style, e.g. "Jul 1, 2026 8:45 AM".
    month = dt_value.strftime("%b")
    hour = dt_value.strftime("%I").lstrip("0") or "0"
    minute = dt_value.strftime("%M")
    am_pm = dt_value.strftime("%p")
    return f"{month} {dt_value.day}, {dt_value.year} {hour}:{minute} {am_pm}"


def _classify_task(progress: str, is_error: bool) -> str:
    text = (progress or "").lower()
    pending_match = re.search(r"pending\s*=\s*(\d+)", text)
    pending_count = int(pending_match.group(1)) if pending_match else 0

    if pending_count > 0 or any(k in text for k in ("scheduled", "upcoming", "queued", "waiting for device")):
        return "Upcoming Tasks"
    if any(k in text for k in ("running", "in progress")):
        return "In Progress"
    if is_error or any(k in text for k in ("failed", "failure", "error", "aborted", "rollback")):
        return "Failure"
    if any(k in text for k in ("success", "completed successfully")):
        return "Success"
    # Keep output constrained to requested categories.
    return "In Progress"


def _fetch_endpoint_rows(client, path: str, page_size: int, max_pages: int) -> List[Dict[str, Any]]:
    rows: List[Dict[str, Any]] = []
    offset = 1
    for _ in range(max_pages):
        try:
            payload = client.get(path, params={"offset": offset, "limit": page_size})
        except Exception:
            break
        page_rows = _to_record_list(payload)
        if not page_rows:
            break
        rows.extend(page_rows)
        if len(page_rows) < page_size:
            break
        offset += page_size
    return rows


def _fetch_tasks(
    client,
    page_size: int | None = None,
    max_pages: int | None = None,
    include_alt_task_endpoint: bool = True,
) -> List[Dict[str, Any]]:
    page_size = page_size or _history_page_size()
    max_pages = max_pages or _history_max_pages()
    paths = ["/dna/intent/api/v1/task"]
    if include_alt_task_endpoint:
        paths.append("/dna/intent/api/v1/tasks")

    rows: List[Dict[str, Any]] = []
    seen = set()
    for path in paths:
        fetched = _fetch_endpoint_rows(client, path, page_size, max_pages)
        for task in fetched:
            key = str(task.get("id") or "")
            if not key:
                key = f"{task.get('serviceType', '')}|{task.get('startTime', '')}|{task.get('endTime', '')}|{task.get('progress', '')}"
            if key in seen:
                continue
            seen.add(key)
            rows.append(task)
    return rows


def _fetch_image_device_history(client, page_size: int | None = None, max_pages: int | None = None) -> List[Dict[str, Any]]:
    page_size = page_size or _history_page_size()
    max_pages = max_pages or _history_max_pages()
    return _fetch_endpoint_rows(client, "/dna/intent/api/v1/image/device", page_size, max_pages)


def _extract_device_id(record: Dict[str, Any]) -> str:
    for key in ("networkDeviceId", "deviceUuid", "deviceId", "id"):
        value = str(record.get(key) or "").strip()
        if value:
            return value
    return ""


def _classify_history_record(status_text: str, is_error: bool) -> str:
    text = (status_text or "").lower()
    if is_error or any(k in text for k in ("failed", "failure", "error", "aborted", "rollback")):
        return "Failure"
    if any(k in text for k in ("success", "successful", "completed", "installed")):
        return "Success"
    if any(k in text for k in ("running", "in progress")):
        return "In Progress"
    if any(k in text for k in ("pending", "scheduled", "queued", "waiting")):
        return "Upcoming Tasks"
    return "In Progress"


def get_upgrade_tasks_for_date(
    clients_by_region: List[Tuple[str, Any]],
    target_date: date,
    tz_name: str = "America/New_York",
    tz_label_override: str = "",
    page_size: int | None = None,
    max_pages: int | None = None,
    include_alt_sources: bool = True,
) -> List[Dict[str, Any]]:
    tz = ZoneInfo(tz_name)
    output: List[Dict[str, Any]] = []

    for region, client in clients_by_region:
        try:
            devices = get_devices_for_region(region, client, refresh=False)
        except Exception:
            # Keep partial data loading alive when one region is temporarily unreachable.
            devices = []
        device_id_to_name: Dict[str, str] = {}
        device_id_to_ip: Dict[str, str] = {}
        device_id_to_version: Dict[str, str] = {}
        for d in devices:
            device_name = d.get("hostname") or d.get("fullyQualifiedDomainName") or "Unknown"
            device_ip = str(d.get("managementIpAddress") or "")
            device_version = d.get("softwareVersion") or "Unknown"
            for candidate in (d.get("id"), d.get("instanceUuid"), d.get("deviceId"), d.get("networkDeviceId")):
                device_id = str(candidate or "").strip()
                if not device_id:
                    continue
                device_id_to_name[device_id] = device_name
                device_id_to_ip[device_id] = device_ip
                device_id_to_version[device_id] = device_version

        ip_to_name = {
            str(d.get("managementIpAddress") or ""): (d.get("hostname") or d.get("fullyQualifiedDomainName") or "Unknown")
            for d in devices
            if d.get("managementIpAddress")
        }
        ip_to_version = {
            str(d.get("managementIpAddress") or ""): (d.get("softwareVersion") or "Unknown")
            for d in devices
            if d.get("managementIpAddress")
        }

        try:
            tasks = _fetch_tasks(
                client,
                page_size=page_size,
                max_pages=max_pages,
                include_alt_task_endpoint=include_alt_sources,
            )
        except Exception:
            tasks = []
        if include_alt_sources:
            try:
                history_rows = _fetch_image_device_history(client, page_size=page_size, max_pages=max_pages)
            except Exception:
                history_rows = []
            try:
                schedule_history_rows = _fetch_schedule_history_rows(client)
            except Exception:
                schedule_history_rows = []
        else:
            history_rows = []
            schedule_history_rows = []
        emitted_keys = set()

        for task in tasks:
            progress = str(task.get("progress") or "")
            service = str(task.get("serviceType") or "")
            data = str(task.get("data") or "")
            operation = str(task.get("operationName") or "")
            blob = (service + " " + operation + " " + progress + " " + data).lower()
            if not any(k in blob for k in ("swim", "image", "upgrade", "install")):
                continue

            end_utc = _extract_end_utc(task)
            if not end_utc:
                continue

            end_local = end_utc.astimezone(tz)
            if end_local.date() != target_date:
                continue

            device_ip = _extract_device_ip(progress)
            task_status = _classify_task(progress, bool(task.get("isError")))
            if tz_label_override:
                time_value = end_local.strftime("%Y-%m-%d %I:%M:%S %p") + f" {tz_label_override}"
            else:
                time_value = end_local.strftime("%Y-%m-%d %I:%M:%S %p %Z")

            output.append(
                {
                    "date_time_est": time_value,
                    "sort_end_time": end_local.isoformat(),
                    "region": region,
                    "task_status": task_status,
                    "device_ip": device_ip or "Unknown",
                    "device_name": ip_to_name.get(device_ip, "Unknown"),
                    "current_image_version": ip_to_version.get(device_ip, "Unknown"),
                    "task_id": task.get("id") or "",
                    "progress": progress,
                }
            )
            emitted_keys.add(f"task|{task.get('id') or ''}|{device_ip or 'Unknown'}|{end_local.isoformat()}")

        for history in history_rows:
            status_blob = " ".join(
                [
                    str(history.get("status") or ""),
                    str(history.get("state") or ""),
                    str(history.get("imageUpdateStatus") or ""),
                    str(history.get("installStatus") or ""),
                    str(history.get("taskStatus") or ""),
                    str(history.get("description") or ""),
                ]
            )
            if not status_blob.strip():
                continue

            end_utc = _extract_end_utc(history)
            if not end_utc:
                continue
            end_local = end_utc.astimezone(tz)
            if end_local.date() != target_date:
                continue

            device_id = _extract_device_id(history)
            device_ip = str(history.get("deviceIp") or history.get("managementIpAddress") or "").strip()
            if not device_ip and device_id:
                device_ip = device_id_to_ip.get(device_id, "")
            task_id = str(history.get("taskId") or history.get("id") or "").strip()
            dedupe_key = f"history|{task_id}|{device_ip or 'Unknown'}|{end_local.isoformat()}"
            if dedupe_key in emitted_keys:
                continue

            task_status = _classify_history_record(status_blob, bool(history.get("isError")))
            if task_status not in {"Success", "Failure", "In Progress", "Upcoming Tasks"}:
                continue

            if tz_label_override:
                time_value = end_local.strftime("%Y-%m-%d %I:%M:%S %p") + f" {tz_label_override}"
            else:
                time_value = end_local.strftime("%Y-%m-%d %I:%M:%S %p %Z")

            fallback_name = str(history.get("deviceName") or history.get("hostname") or "").strip()
            resolved_name = fallback_name or device_id_to_name.get(device_id, "Unknown")
            if not resolved_name or resolved_name == "Unknown":
                resolved_name = ip_to_name.get(device_ip, "Unknown") if device_ip else "Unknown"

            version = str(history.get("currentImageVersion") or history.get("softwareVersion") or "").strip()
            if not version and device_id:
                version = device_id_to_version.get(device_id, "Unknown")
            if not version and device_ip:
                version = ip_to_version.get(device_ip, "Unknown")

            output.append(
                {
                    "date_time_est": time_value,
                    "sort_end_time": end_local.isoformat(),
                    "region": region,
                    "task_status": task_status,
                    "device_ip": device_ip or "Unknown",
                    "device_name": resolved_name,
                    "current_image_version": version or "Unknown",
                    "task_id": task_id,
                    "progress": status_blob,
                }
            )
            emitted_keys.add(dedupe_key)

        for job in schedule_history_rows:
            description = str(job.get("description") or "").strip()
            blob = (
                description
                + " "
                + str(job.get("scheduleState") or "")
                + " "
                + str(job.get("triggerState") or "")
                + " "
                + str(job.get("progress") or "")
            ).lower()
            if not description or "activation job" not in blob:
                continue

            end_utc = _extract_end_utc(job)
            if not end_utc:
                continue
            end_local = end_utc.astimezone(tz)
            if end_local.date() != target_date:
                continue

            schedule_state = str(job.get("scheduleState") or "").lower()
            trigger_state = str(job.get("triggerState") or "").lower()
            jobs_failed = int(job.get("jobsFailed") or 0)

            if jobs_failed > 0 or "fail" in schedule_state or "error" in schedule_state:
                task_status = "Failure"
            elif any(k in schedule_state for k in ("completed", "success", "succeeded")) or any(
                k in trigger_state for k in ("triggered", "completed", "success")
            ):
                task_status = "Success"
            elif any(k in schedule_state for k in ("pending", "active", "queued")):
                task_status = "Upcoming Tasks"
            else:
                task_status = "In Progress"

            time_value = end_local.strftime("%Y-%m-%d %I:%M:%S %p")
            if tz_label_override:
                time_value = f"{time_value} {tz_label_override}"
            else:
                time_value = f"{time_value} {end_local.tzname() or ''}".strip()

            scheduled_devices = job.get("scheduledDevices")
            task_id = str(job.get("taskId") or job.get("id") or "").strip()
            if isinstance(scheduled_devices, list) and scheduled_devices:
                for scheduled in scheduled_devices:
                    device_id = str(scheduled.get("deviceId") or scheduled.get("id") or "").strip()
                    device_ip = device_id_to_ip.get(device_id, "")
                    dedupe_key = f"schedule|{task_id}|{device_ip or 'Unknown'}|{end_local.isoformat()}"
                    if dedupe_key in emitted_keys:
                        continue

                    output.append(
                        {
                            "date_time_est": time_value,
                            "sort_end_time": end_local.isoformat(),
                            "region": region,
                            "task_status": task_status,
                            "device_ip": device_ip or "Unknown",
                            "device_name": device_id_to_name.get(device_id, "Unknown"),
                            "current_image_version": device_id_to_version.get(device_id, "Unknown"),
                            "task_id": task_id,
                            "progress": description,
                        }
                    )
                    emitted_keys.add(dedupe_key)
            else:
                dedupe_key = f"schedule|{task_id}|Unknown|{end_local.isoformat()}"
                if dedupe_key in emitted_keys:
                    continue
                output.append(
                    {
                        "date_time_est": time_value,
                        "sort_end_time": end_local.isoformat(),
                        "region": region,
                        "task_status": task_status,
                        "device_ip": "Unknown",
                        "device_name": "Unknown",
                        "current_image_version": "Unknown",
                        "task_id": task_id,
                        "progress": description,
                    }
                )
                emitted_keys.add(dedupe_key)

    output.sort(key=lambda r: str(r.get("sort_end_time") or ""), reverse=True)
    return output


def get_upcoming_tasks_for_date(
    clients_by_region: List[Tuple[str, Any]],
    target_date: date,
    tz_name: str = "America/New_York",
    tz_label_override: str = "",
) -> List[Dict[str, Any]]:
    """Return upcoming SWIM activation jobs for a specific date.

    This mirrors the DNAC "Upcoming Tasks" sidebar shape with description,
    start date/time, and devices scheduled.
    """
    tz = ZoneInfo(tz_name)
    output: List[Dict[str, Any]] = []

    # Primary path: use the same scheduling endpoint as Catalyst Upcoming Tasks.
    stop_after_first_nonempty = os.getenv("UPCOMING_STOP_AFTER_FIRST_NONEMPTY", "true").strip().lower() in (
        "1",
        "true",
        "yes",
        "on",
    )
    for region, client in clients_by_region:
        before_count = len(output)
        device_by_id: Dict[str, Dict[str, Any]] = {}

        use_region_inventory = os.getenv("UPCOMING_USE_REGION_INVENTORY", "true").strip().lower() in (
            "1",
            "true",
            "yes",
            "on",
        )
        if use_region_inventory:
            # Optional enrichment path: index region inventory once.
            try:
                region_devices = get_devices_for_region(region, client, refresh=False)
            except Exception:
                region_devices = []
            for dev in region_devices:
                for key_name in ("id", "instanceUuid", "deviceId", "networkDeviceId"):
                    device_id = str(dev.get(key_name) or "").strip()
                    if device_id:
                        device_by_id[device_id] = dev

        def _load_device(device_id: str) -> tuple[str, Dict[str, Any]]:
            if not device_id:
                return "", {}
            try:
                payload = client.get("/dna/intent/api/v1/network-device", params={"id": device_id})
                response = payload.get("response") if isinstance(payload, dict) else None
                if isinstance(response, list) and response:
                    return device_id, response[0]
            except Exception:
                pass
            try:
                payload = client.get_device_by_id(device_id)
                if isinstance(payload, dict) and isinstance(payload.get("response"), dict):
                    return device_id, payload.get("response") or {}
                if isinstance(payload, dict):
                    return device_id, payload
            except Exception:
                pass
            return device_id, {}

        query_variants = [
            {
                "source": "EXTERNAL",
                "type": "DEFAULT",
                "sortBy": "lastUpdateTime",
                "order": "DESC",
                "scheduleState": "ACTIVE",
                "module": "OS_UPDATE",
                "triggerState": "WAITING",
            },
            {
                "source": "EXTERNAL",
                "type": "DEFAULT",
                "sortBy": "lastUpdateTime",
                "order": "DESC",
                "scheduleState": "PENDING",
                "module": "OS_UPDATE",
            },
            {
                "source": "EXTERNAL",
                "type": "DEFAULT",
                "sortBy": "lastUpdateTime",
                "order": "DESC",
                "module": "OS_UPDATE",
            },
            {
                "sortBy": "lastUpdateTime",
                "order": "DESC",
                "module": "OS_UPDATE",
                "triggerState": "WAITING",
            },
        ]

        jobs: List[Dict[str, Any]] = []
        seen_job_keys = set()
        for idx, params in enumerate(query_variants):
            try:
                payload = client.get("/api/schedule/v4/scheduled-job", params=params)
            except Exception:
                # DNAC API query support differs by version; keep trying fallback shapes.
                continue

            for job in _to_scheduled_job_list(payload):
                key = str(job.get("taskId") or job.get("id") or "").strip()
                if not key:
                    key = (
                        f"{job.get('startTime','')}|{job.get('nextTriggerTime','')}|"
                        f"{job.get('description','')}|{job.get('scheduleState','')}|{job.get('triggerState','')}"
                    )
                if key in seen_job_keys:
                    continue
                seen_job_keys.add(key)
                jobs.append(job)

            # First successful payload is usually enough; avoid extra calls that slow page rendering.
            if jobs:
                break

        filtered_jobs: List[Dict[str, Any]] = []
        needed_device_ids: set[str] = set()

        for job in jobs:
            description = str(job.get("description") or "").strip()
            if not description:
                continue
            desc_lower = description.lower()
            if not any(token in desc_lower for token in ("activation job", "os update", "swim", "image")):
                continue

            start_utc = _extract_start_utc(job) or _from_epoch_millis(job.get("startTime") or job.get("nextTriggerTime"))
            if not start_utc:
                continue

            start_local = start_utc.astimezone(tz)
            if start_local.date() != target_date:
                continue

            filtered_jobs.append(job)
            scheduled_devices = job.get("scheduledDevices")
            if isinstance(scheduled_devices, list):
                for scheduled in scheduled_devices:
                    device_id = str(scheduled.get("deviceId") or scheduled.get("id") or "")
                    if device_id:
                        needed_device_ids.add(device_id)

        if needed_device_ids:
            enable_device_api_enrichment = os.getenv("UPCOMING_ENRICH_DEVICE_API", "false").strip().lower() in (
                "1",
                "true",
                "yes",
                "on",
            )
            if enable_device_api_enrichment:
                missing_device_ids = sorted([device_id for device_id in needed_device_ids if device_id not in device_by_id])
                if missing_device_ids:
                    max_workers = min(8, len(missing_device_ids))
                    with ThreadPoolExecutor(max_workers=max_workers) as executor:
                        for device_id, device in executor.map(_load_device, missing_device_ids):
                            if device_id:
                                device_by_id[device_id] = device

        for job in filtered_jobs:
            description = str(job.get("description") or "").strip()
            start_utc = _from_epoch_millis(job.get("startTime") or job.get("nextTriggerTime"))
            start_local = start_utc.astimezone(tz)

            time_value = _format_catalyst_datetime(start_local)
            if tz_label_override:
                time_value = f"{time_value} {tz_label_override}"

            scheduled_devices = job.get("scheduledDevices")
            devices_scheduled = len(scheduled_devices) if isinstance(scheduled_devices, list) and scheduled_devices else 1
            task_id_value = job.get("taskId") or job.get("id") or ""
            is_error_value = int(job.get("jobsFailed") or 0) > 0

            if isinstance(scheduled_devices, list) and scheduled_devices:
                for scheduled in scheduled_devices:
                    device_id = str(scheduled.get("deviceId") or scheduled.get("id") or "")
                    device = device_by_id.get(device_id, {})
                    scheduled_name = (
                        scheduled.get("hostname")
                        or scheduled.get("deviceName")
                        or scheduled.get("name")
                        or ""
                    )
                    scheduled_ip = (
                        scheduled.get("managementIpAddress")
                        or scheduled.get("ipAddress")
                        or scheduled.get("ip")
                        or ""
                    )
                    scheduled_version = (
                        scheduled.get("softwareVersion")
                        or scheduled.get("currentImageVersion")
                        or scheduled.get("version")
                        or ""
                    )
                    output.append(
                        {
                            "description": description,
                            "start_date_time_est": time_value,
                            "sort_start_time": start_local.isoformat(),
                            "device_name": scheduled_name
                            or device.get("hostname")
                            or device.get("fullyQualifiedDomainName")
                            or "Unknown",
                            "ip_address": scheduled_ip or device.get("managementIpAddress") or "Unknown",
                            "current_image_version": scheduled_version or device.get("softwareVersion") or "Unknown",
                            "devices_scheduled": devices_scheduled,
                            "region": region,
                            "task_id": task_id_value,
                            "is_error": is_error_value,
                            "source": "schedule_v4",
                        }
                    )
            else:
                output.append(
                    {
                        "description": description,
                        "start_date_time_est": time_value,
                        "sort_start_time": start_local.isoformat(),
                        "device_name": "Unknown",
                        "ip_address": "Unknown",
                        "current_image_version": "Unknown",
                        "devices_scheduled": devices_scheduled,
                        "region": region,
                        "task_id": task_id_value,
                        "is_error": is_error_value,
                        "source": "schedule_v4",
                    }
                )

        if stop_after_first_nonempty and len(output) > before_count:
            break

    output.sort(key=lambda r: r.get("sort_start_time") or "")
    return output
