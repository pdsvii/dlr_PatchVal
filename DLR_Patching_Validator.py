import os
import subprocess
from dotenv import load_dotenv

# Load .env variables
load_dotenv()

PROJECT_NAME = "DLR Patching Validator"

PROJECT_PATH = os.getenv(
    "PROJECT_PATH",
    r"C:\Patch_Validator"
)

WEB_URL = os.getenv("WEB_URL")


def open_vscode():
    """
    Open VS Code in the project folder.
    VS Code must already be installed and
    the 'code' command available in PATH.
    """

    try:
        subprocess.Popen(
            [
                "code",
                PROJECT_PATH
            ]
        )

        print("[SUCCESS] VS Code opened.")

    except Exception as error:
        print(f"[ERROR] Unable to open VS Code: {error}")


def open_website():
    """
    Open website using PowerShell.
    """

    if not WEB_URL:
        print("[WARNING] WEB_URL not configured.")
        return

    try:
        subprocess.Popen(
            [
                "powershell.exe",
                "-Command",
                f"Start-Process '{WEB_URL}'"
            ]
        )

        print(f"[SUCCESS] Opened {WEB_URL}")

    except Exception as error:
        print(f"[ERROR] Unable to open website: {error}")


def main():

    print("")
    print("=" * 60)
    print(f"{PROJECT_NAME}")
    print("=" * 60)
    print("")

    # Startup Actions

    open_vscode()

    open_website()

    print("")
    print("DLR Patching Validator Initialized")
    print("")


if __name__ == "__main__":
    main()