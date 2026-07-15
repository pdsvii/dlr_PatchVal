from pathlib import Path

from mcp.server.fastmcp import FastMCP

mcp = FastMCP("SystemsPTU")

ROOT = Path(r"C:\systems_ptu")
ROOT.mkdir(exist_ok=True)


# --------------------------
# Filesystem Tools
# --------------------------


@mcp.tool()
def list_files():
    """List all files under C:\\systems_ptu"""
    return [str(p) for p in ROOT.rglob("*")]


@mcp.tool()
def read_file(path: str):
    """Read a file"""
    target = ROOT / path
    return target.read_text(encoding="utf-8")


@mcp.tool()
def write_file(path: str, content: str):
    """Write a file"""
    target = ROOT / path
    target.parent.mkdir(parents=True, exist_ok=True)
    target.write_text(content, encoding="utf-8")
    return f"Saved: {target}"


# --------------------------
# Systems PTU Tools
# --------------------------


@mcp.tool()
def get_project_status():
    return {
        "project": "Systems_PTU",
        "root": str(ROOT),
        "inventory": str(ROOT / "inventory"),
        "reports": str(ROOT / "reports"),
        "logs": str(ROOT / "logs"),
        "status": "READY",
    }


@mcp.tool()
def save_device_audit(hostname: str, version: str, image: str, boot_path: str):
    report_dir = ROOT / "reports"
    report_dir.mkdir(exist_ok=True)

    report = report_dir / f"{hostname}.txt"

    report.write_text(
        f'''Hostname: {hostname}
Version: {version}
Image: {image}
Boot Path: {boot_path}
''',
        encoding="utf-8",
    )

    return f"Created {report}"


if __name__ == "__main__":
    mcp.run()