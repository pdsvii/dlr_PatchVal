# Patch Validator

## MCP Filesystem Server

This workspace includes a local MCP filesystem server for Copilot at [.vscode/mcp.json](.vscode/mcp.json).
It launches [scripts/mcp_filesystem_server.py](scripts/mcp_filesystem_server.py) with the workspace root locked to this repo.

Scaffold for validating patching status of network devices using Cisco DNA Center.

Quick start:

1. Copy `.env.example` to `.env` and set region endpoints and credentials.
2. Install dependencies:

```bash
python -m pip install -r requirements.txt
```

3. Test auth across configured regions:

```bash
python -m src.main --test-auth
```

4. List devices (example):

```bash
python -m src.main --list-devices
```

5. Run validation and generate a report:

```bash
python -m src.main --validate --min-version 17.9.4
```

Outputs:
- SQLite history at `DB_PATH` (default `./data/patches.db`)
- JSON report at `./reports/validation_report_<timestamp>.json`

Optional flags:
- `--platform-prefix C9` to enforce platform ID prefix
- `--report-path ./reports/latest.json` for a fixed report path

6. Lookup one device from CLI:

```bash
python -m src.main --lookup-device cs01-los10-9500.digitalrealtytrust.com
```

Optional lookup flags:
- `--no-cache` disables in-memory inventory cache for one run.
- `--ssh-validate` runs direct device SSH (`show version`) using `DNA_USERNAME`/`DNA_PASSWORD`.

Example with SSH validation:

```bash
python -m src.main --lookup-device cs01-los10-9500.digitalrealtytrust.com --ssh-validate
```

7. Check DNAC endpoint capability for update-status fidelity:

```bash
python -m src.main --check-endpoints
```

8. Run the web app UI:

```bash
streamlit run src/app.py
```

9. Run the distribution-failure staging app:

```bash
python -m streamlit run src/dist_failure_app.py
```

The UI returns:
- IP Address
- Region
- Image Version
- Online/Offline
- Last image update outcome (Success/Failed/In Progress/Unknown)

Environment variables:
- `DNA_USERNAME`, `DNA_PASSWORD`
- `DNAC_VERIFY_SSL` (`true` by default; set to `false` for self-signed lab certificates)
- `EMEA_DNAC_BASE_URL`, `EMEA_DNAC_TOKEN`
- `US_DNAC_BASE_URL`, `US_DNAC_TOKEN`
- `APAC_DNAC_BASE_URL`, `APAC_DNAC_TOKEN`
- `DB_PATH`, `MIN_SOFTWARE_VERSION`, `REQUIRED_PLATFORM_PREFIX`
- `LOOKUP_CACHE_TTL_SECONDS` (default `300`)
- `SSH_DEVICE_TYPE` (default `cisco_ios`)
- `SSH_TIMEOUT_SECONDS` (default `12`)

UI note:
- Enable "Validate image via SSH (show version)" to add direct CLI validation from the device.
- When SSH validation is enabled, update outcome is derived from DNAC vs SSH version comparison:
	- `Success` if versions match
	- `Failed (Version Mismatch)` if versions differ
	- `Failed (SSH Error)` if direct SSH validation cannot run

Base URL note:
- Use the DNAC host base URL (for example, `https://dnac.example.com`).
- If you paste a UI URL like `https://dnac.example.com/dna/home`, the client normalizes it automatically.

## Streamlit Deploy Prep

The distribution-failure app entrypoint is `src/dist_failure_app.py`.

Before deploying to Streamlit Community Cloud:

1. Install Git for Windows on this machine.
2. Initialize this folder as a Git repository.
3. Push the code to a GitHub repository.
4. In Streamlit deploy, select that repository and set the app file to `src/dist_failure_app.py`.

Do not commit live secrets:

- `.env` is intentionally ignored.
- Put deployment secrets into Streamlit app secrets or environment settings instead of source control.

Distribution-failure staging outputs are written locally to:

- `data/dist_failure_remediation.db`
- the copied workbook folder under OneDrive for seed/report/post-check exports
