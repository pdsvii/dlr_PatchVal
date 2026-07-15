# Changelog

All notable changes to this project will be documented in this file.

The format is based on "Keep a Changelog" and follows [Semantic Versioning](https://semver.org/).

## [Unreleased]

### Added
- 2026-07-01: Initialize changelog for Patching Validator project.
- 2026-07-01: Added `src/validator.py` for compliance rule evaluation, SQLite persistence, and JSON report generation.
- 2026-07-01: Added `--validate`, `--min-version`, `--platform-prefix`, and `--report-path` options to `src/main.py`.
- 2026-07-01: Updated README with region-based DNAC auth/test/list/validate commands.
- 2026-07-01: Added `src/device_lookup.py` for cross-region device lookup and SWIM update outcome parsing.
- 2026-07-01: Added Streamlit UI in `src/app.py` for hostname lookup (IP, region, image, online status, last update outcome).
- 2026-07-01: Added CLI option `--lookup-device` in `src/main.py`.
- 2026-07-01: Added batch CSV lookup in Streamlit UI (`src/app.py`) for multi-device results table.
- 2026-07-01: Added CLI endpoint capability check via `--check-endpoints` to assess per-region DNAC update endpoint availability.
- 2026-07-01: Added direct SSH validation module `src/ssh_device_client.py` to run `show version` and parse image/version.
- 2026-07-01: Added `--ssh-validate` to CLI device lookup and SSH validation fields in Streamlit UI.
- 2026-07-01: Added upgrade task schedule panel in app with status buckets (In Progress, Success, Failure, Upcoming Tasks).
- 2026-07-01: Added device-name resolution for task rows using DNAC inventory mapping by management IP.
- 2026-07-01: Added Eastern Time display and date filtering for task timestamps in the app.
- 2026-07-01: Normalized upgrade task status output to four categories only: In Progress, Success, Failure, Upcoming Tasks.
- 2026-07-01: Added app toggle to force fixed EST label/year-round UTC-5 for task timestamps.
- 2026-07-01: Added manual date input (`YYYY-MM-DD`) for task filtering in the app.
- 2026-07-01: Updated task classifier so "Waiting for device" entries are counted under Upcoming Tasks.
- 2026-07-01: Added task status filter control in app for Upcoming Tasks, Failure, Success, and In Progress.
- 2026-07-01: Added task list export options in app: PDF, CSV, and XLS.

### Changed
- 2026-07-01: DNAC client now normalizes base URLs like `/dna/home` to host root before calling API paths.
- 2026-07-01: Added `DNAC_VERIFY_SSL` support in runner to handle environments with self-signed certificates.
- 2026-07-01: Updated `requirements.txt` to include `streamlit` and updated README with UI run steps.
- 2026-07-01: Added inventory cache support in `src/device_lookup.py` with configurable TTL (`LOOKUP_CACHE_TTL_SECONDS`).
- 2026-07-01: Added `.env` configuration options for SSH device type and timeout (`SSH_DEVICE_TYPE`, `SSH_TIMEOUT_SECONDS`).
- 2026-07-01: When SSH validation is enabled, update outcome is now derived from DNAC-vs-SSH version comparison (no longer `Unknown`).

### Fixed
- 2026-07-01: Fixed token retry logic to ignore placeholder env tokens and refresh using username/password after 401.
- 2026-07-01: Reduced noisy traceback logging for expected 404 responses from optional DNAC image-history endpoints.
- 2026-07-01: Fixed SSH-validation path so `Last image update outcome` cannot remain `Unknown` when SSH validation is enabled.

---

## Guidelines
- Add entries under the `Unreleased` heading as you make changes.
- Use one-line items prefixed with the date in YYYY-MM-DD format.
- Move `Unreleased` items into release sections and add a version when cutting a release.
- Keep entries short and factual; link to issues/PRs when available.

