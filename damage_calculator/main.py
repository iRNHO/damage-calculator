import json
import shutil
import subprocess
import urllib.request
from pathlib import Path

from platformdirs import user_data_dir


APP_NAME = "Damage Calculator"
AUTHOR = "iRNHO"
REPO = "iRNHO/damage-calculator-data"


def safe_request(url):
    try:
        with urllib.request.urlopen(url, timeout=5) as r:
            return r.read()
    except Exception:
        return None


def read_version(path):
    if not path.exists():
        return None
    return path.read_text().strip()


def run_local(root):
    app_path = root / "main.py"

    if not app_path.exists():
        print("No local install found.")
        return

    print("Running app...\n")
    subprocess.run(["python", str(app_path)])


def main():
    print("Launcher starting...")

    root = Path(user_data_dir(APP_NAME, AUTHOR))
    root.mkdir(parents=True, exist_ok=True)

    print(f"Install location: {root}")

    version_file = root / "version.txt"
    local_version = read_version(version_file)

    print("Checking GitHub for latest version...")

    data = safe_request(f"https://api.github.com/repos/{REPO}/releases/latest")

    if not data:
        print("Offline mode.")
        run_local(root)
        return

    latest_version = json.loads(data)["tag_name"]

    print(f"Latest version: {latest_version}")
    print(f"Installed version: {local_version}")

    if latest_version != local_version:
        print("Updating...\n")

        # Clear install
        for item in root.iterdir():
            if item.is_dir():
                shutil.rmtree(item)
            else:
                item.unlink()

        # Download required files
        for file in ["main.py", "version.txt"]:
            url = f"https://github.com/{REPO}/releases/latest/download/{file}"
            data = safe_request(url)

            if not data:
                print(f"Failed to download {file}")
                run_local(root)
                return

            (root / file).write_bytes(data)

    else:
        print("Already up to date.\n")

    run_local(root)


if __name__ == "__main__":
    main()