import json
import shutil
import subprocess
import urllib.request
import zipfile
import io
import time

from pathlib import Path
from platformdirs import user_data_dir


APP_NAME = "Damage Calculator"
AUTHOR = "iRNHO"
REPO = "iRNHO/damage-calculator-data"


# -------------------- Utilities -------------------- #

def safe_request(url, timeout=10):
    try:
        with urllib.request.urlopen(url, timeout=timeout) as r:
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


def download_and_extract_zip(url, extract_to):
    data = safe_request(url)

    if not data:
        return False

    with zipfile.ZipFile(io.BytesIO(data)) as z:
        z.extractall(extract_to)

    return True


# -------------------- Main -------------------- #

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

        # Small delay to avoid GitHub cache race condition
        time.sleep(2)

        zip_url = f"https://github.com/{REPO}/archive/refs/tags/{latest_version}.zip"

        print("Downloading full release...")

        success = download_and_extract_zip(zip_url, root)

        if not success:
            print("Failed to download release")
            run_local(root)
            return

        # Find extracted folder (GitHub naming pattern)
        extracted_folders = list(root.glob(f"*{latest_version.lstrip('v')}*"))

        if not extracted_folders:
            print("Failed to locate extracted folder")
            run_local(root)
            return

        extracted_folder = extracted_folders[0]

        # Clear existing install (except extracted folder)
        for item in root.iterdir():
            if item == extracted_folder:
                continue
            if item.is_dir():
                shutil.rmtree(item)
            else:
                item.unlink()

        # Move new files into root
        for item in extracted_folder.iterdir():
            target = root / item.name
            if target.exists():
                if target.is_dir():
                    shutil.rmtree(target)
                else:
                    target.unlink()
            item.rename(target)

        shutil.rmtree(extracted_folder)

        print("Update complete.\n")

    else:
        print("Already up to date.\n")

    run_local(root)


if __name__ == "__main__":
    main()