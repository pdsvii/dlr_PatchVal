import json
import os
import sqlite3
from dataclasses import dataclass
from datetime import datetime
from typing import Any, Dict, List, Optional


@dataclass
class ValidationRuleSet:
    min_software_version: Optional[str] = None
    required_platform_prefix: Optional[str] = None


def _normalize_version(v: Optional[str]) -> List[int]:
    if not v:
        return []
    parts = []
    token = ''
    for ch in v:
        if ch.isdigit() or ch == '.':
            token += ch
        elif token:
            break
    for p in token.split('.'):
        if p.isdigit():
            parts.append(int(p))
    return parts


def version_at_least(current: Optional[str], minimum: Optional[str]) -> bool:
    if not minimum:
        return True
    c = _normalize_version(current)
    m = _normalize_version(minimum)
    if not c or not m:
        return False
    max_len = max(len(c), len(m))
    c += [0] * (max_len - len(c))
    m += [0] * (max_len - len(m))
    return c >= m


def evaluate_device(device: Dict[str, Any], rules: ValidationRuleSet) -> Dict[str, Any]:
    hostname = device.get('hostname') or device.get('managementIpAddress') or 'unknown'
    platform = device.get('platformId') or ''
    software = device.get('softwareVersion') or ''

    checks = {
        'min_software_version': version_at_least(software, rules.min_software_version),
        'platform_prefix': (
            True if not rules.required_platform_prefix else platform.startswith(rules.required_platform_prefix)
        ),
    }
    compliant = all(checks.values())

    return {
        'device_id': device.get('id') or device.get('instanceUuid') or '',
        'hostname': hostname,
        'platform': platform,
        'software_version': software,
        'compliant': compliant,
        'checks': checks,
    }


def init_db(db_path: str) -> None:
    os.makedirs(os.path.dirname(db_path), exist_ok=True)
    with sqlite3.connect(db_path) as conn:
        conn.execute(
            '''
            CREATE TABLE IF NOT EXISTS validation_results (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                run_ts TEXT NOT NULL,
                region TEXT NOT NULL,
                device_id TEXT,
                hostname TEXT,
                platform TEXT,
                software_version TEXT,
                compliant INTEGER NOT NULL,
                checks_json TEXT NOT NULL
            )
            '''
        )
        conn.commit()


def save_results(db_path: str, region: str, rows: List[Dict[str, Any]]) -> None:
    run_ts = datetime.utcnow().isoformat(timespec='seconds') + 'Z'
    with sqlite3.connect(db_path) as conn:
        conn.executemany(
            '''
            INSERT INTO validation_results (
                run_ts, region, device_id, hostname, platform,
                software_version, compliant, checks_json
            ) VALUES (?, ?, ?, ?, ?, ?, ?, ?)
            ''',
            [
                (
                    run_ts,
                    region,
                    r.get('device_id', ''),
                    r.get('hostname', ''),
                    r.get('platform', ''),
                    r.get('software_version', ''),
                    1 if r.get('compliant') else 0,
                    json.dumps(r.get('checks', {}), sort_keys=True),
                )
                for r in rows
            ],
        )
        conn.commit()


def build_report(results_by_region: Dict[str, List[Dict[str, Any]]]) -> Dict[str, Any]:
    summary = {}
    for region, rows in results_by_region.items():
        total = len(rows)
        compliant = sum(1 for r in rows if r.get('compliant'))
        non_compliant = total - compliant
        summary[region] = {
            'total': total,
            'compliant': compliant,
            'non_compliant': non_compliant,
            'compliance_pct': round((compliant / total) * 100, 2) if total else 0.0,
        }

    return {
        'generated_at': datetime.utcnow().isoformat(timespec='seconds') + 'Z',
        'summary': summary,
        'regions': results_by_region,
    }


def write_json_report(report: Dict[str, Any], output_path: str) -> None:
    os.makedirs(os.path.dirname(output_path), exist_ok=True)
    with open(output_path, 'w', encoding='utf-8') as f:
        json.dump(report, f, indent=2)
