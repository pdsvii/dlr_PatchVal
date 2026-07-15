#!/usr/bin/env python3
"""Quick runner for the Patch Validator scaffold.

Examples:
  python src/main.py --list-devices
"""
import argparse
import os
from datetime import datetime
from dotenv import load_dotenv

load_dotenv()

from src.dna_client import DNAClient
from src.device_lookup import check_update_endpoint_capabilities, find_device_by_name
from src.validator import (
    ValidationRuleSet,
    build_report,
    evaluate_device,
    init_db,
    save_results,
    write_json_report,
)


def main():
    parser = argparse.ArgumentParser(description='Patch Validator runner')
    parser.add_argument('--list-devices', action='store_true', help='List devices from DNA Center')
    parser.add_argument('--test-auth', action='store_true', help='Test authentication for all configured regions')
    parser.add_argument('--validate', action='store_true', help='Validate patch compliance and write SQLite + JSON report')
    parser.add_argument('--lookup-device', default=None, help='Lookup one device and print IP/region/image/online/update status')
    parser.add_argument('--check-endpoints', action='store_true', help='Check DNAC endpoint availability for update-status fidelity')
    parser.add_argument('--no-cache', action='store_true', help='Disable lookup cache for this run')
    parser.add_argument('--ssh-validate', action='store_true', help='Run direct SSH show version validation on the matched device')
    parser.add_argument('--min-version', default=os.getenv('MIN_SOFTWARE_VERSION'), help='Minimum allowed software version')
    parser.add_argument('--platform-prefix', default=os.getenv('REQUIRED_PLATFORM_PREFIX'), help='Required platform prefix')
    parser.add_argument('--report-path', default=None, help='Optional custom JSON report output path')
    args = parser.parse_args()

    # Regions configured via environment vars (see .env.example)
    regions = [
        ("EMEA", os.getenv('EMEA_DNAC_BASE_URL'), 'EMEA_DNAC_TOKEN'),
        ("US", os.getenv('US_DNAC_BASE_URL'), 'US_DNAC_TOKEN'),
        ("APAC", os.getenv('APAC_DNAC_BASE_URL'), 'APAC_DNAC_TOKEN'),
    ]

    verify_ssl = os.getenv('DNAC_VERIFY_SSL', 'true').strip().lower() in ('1', 'true', 'yes', 'on')

    def make_client(base_url, token_env):
        if not base_url:
            return None
        return DNAClient(base_url=base_url, token_env=token_env, verify_ssl=verify_ssl)

    if args.test_auth:
        for name, base, token_env in regions:
            if not base:
                print(f'{name}: no base URL configured (skipping)')
                continue
            client = make_client(base, token_env)
            try:
                token = client.get_token()
                if token:
                    print(f'{name}: auth ok, token present=True')
                else:
                    print(f'{name}: auth failed: token request was rejected or returned no token')
            except Exception as e:
                print(f'{name}: auth failed: {e}')

    if args.list_devices:
        for name, base, token_env in regions:
            if not base:
                print(f'{name}: no base URL configured (skipping)')
                continue
            client = make_client(base, token_env)
            try:
                devices = client.list_devices()
                count = len(devices.get('response', devices) if isinstance(devices, dict) else (devices or []))
                print(f'{name}: Found {count} devices')
                sample = devices.get('response', devices) if isinstance(devices, dict) else devices
                for d in (sample or [])[:10]:
                    print('-', d.get('hostname') or d.get('platformId') or d.get('managementIpAddress'))
            except Exception as e:
                print(f'{name}: Error fetching devices: {e}')
                return 2

    if args.lookup_device:
        clients = []
        for name, base, token_env in regions:
            if not base:
                continue
            clients.append((name, make_client(base, token_env)))

        if not clients:
            print('No regions configured in environment.')
            return 1

        result = find_device_by_name(
            clients,
            args.lookup_device,
            use_cache=not args.no_cache,
            use_ssh_validation=args.ssh_validate,
        )
        if not result:
            print(f'Device not found: {args.lookup_device}')
            return 1

        print(f'Hostname: {result.hostname}')
        print(f'IP Address: {result.ip_address}')
        print(f'Region: {result.region}')
        print(f'Image Version: {result.image_version}')
        print(f'Online Status: {result.online_status}')
        print(f'Last Update Outcome: {result.last_update_outcome}')
        print(f'Last Update Time: {result.last_update_time}')
        print(f'SSH Validated: {result.ssh_validated}')
        print(f'SSH Image/Version: {result.ssh_image_or_version}')
        if result.ssh_error:
            print(f'SSH Error: {result.ssh_error}')

    if args.check_endpoints:
        clients = []
        for name, base, token_env in regions:
            if not base:
                continue
            clients.append((name, make_client(base, token_env)))

        if not clients:
            print('No regions configured in environment.')
            return 1

        capabilities = check_update_endpoint_capabilities(clients)
        for region, checks in capabilities.items():
            print(f'[{region}]')
            for row in checks:
                print(f"- {row['endpoint']}: available={row['available']} status={row['status']}")

    if args.validate:
        db_path = os.getenv('DB_PATH', './data/patches.db')
        init_db(db_path)

        rules = ValidationRuleSet(
            min_software_version=args.min_version,
            required_platform_prefix=args.platform_prefix,
        )

        results_by_region = {}
        had_error = False

        for name, base, token_env in regions:
            if not base:
                print(f'{name}: no base URL configured (skipping)')
                continue
            client = make_client(base, token_env)
            try:
                devices_payload = client.list_devices()
                devices = devices_payload.get('response', devices_payload) if isinstance(devices_payload, dict) else (devices_payload or [])
                rows = [evaluate_device(d, rules) for d in devices]
                save_results(db_path, name, rows)
                results_by_region[name] = rows
                compliant = sum(1 for r in rows if r.get('compliant'))
                print(f'{name}: validated {len(rows)} devices, compliant={compliant}, non_compliant={len(rows) - compliant}')
            except Exception as e:
                had_error = True
                results_by_region[name] = []
                print(f'{name}: validation failed: {e}')

        if not results_by_region:
            print('No regions were configured, nothing validated.')
            return 1

        report = build_report(results_by_region)
        if args.report_path:
            report_path = args.report_path
        else:
            stamp = datetime.utcnow().strftime('%Y%m%d_%H%M%S')
            report_path = f'./reports/validation_report_{stamp}.json'
        write_json_report(report, report_path)
        print(f'Report written: {report_path}')
        print(f'SQLite updated: {db_path}')

        if had_error:
            return 2

    return 0


if __name__ == '__main__':
    raise SystemExit(main())
