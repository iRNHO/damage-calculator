import subprocess
import sys
import argparse
import json
import tempfile
import zipfile
import io
import shutil
import time

from urllib.request import urlopen
from pathlib import Path
from platformdirs import user_data_dir

LAUNCHER_VERSION = "0.1.5"

def get_latest_launcher_version():
    url = "https://pypi.org/pypi/irnho-damage-calculator/json"
    try:
        with urlopen(url, timeout=5) as response:
            data = json.loads(response.read())
        return data["info"]["version"]
    except Exception:
        return None

def safe_request(url):
    """
    Safely fetches data from a URL; returning 'None' if offline or unreachable.

    Parameters:
        • url (str): The URL to fetch data from.
    
    """
    try:
        with urlopen(url, timeout=5) as response:
            return response.read()

    except Exception:
        return None

def main():
    release_data = safe_request("https://pypi.org/pypi/irnho-damage-calculator/json")

    if release_data:
        latest_launcher_version = json.loads(release_data)["info"]["version"]

        if LAUNCHER_VERSION != latest_launcher_version:
            print(f"This launcher is outdated (v{LAUNCHER_VERSION} vs v{latest_launcher_version}). Please reinstall the laucher using:\n\nuv tool install --reinstall irnho-damage-calculator\n")
            return

    parser = argparse.ArgumentParser(
        usage="damage-calculator [-h] [-f]",
        description="iRNHO's Damage Calculator Launcher"
    )
    parser.add_argument(
        "-f", "--force",
        action="store_true",
        help="Force the launcher to install the latest version of the application."
    )
    parser.add_argument(
        "-h", "--help",
        action="help",
        help="Show this help message and exit."
    )
    args = parser.parse_args()

    print("Attempting to find a local installation of the application...")
    root_directory = Path(user_data_dir("iRNHO's Damage Calculator", "iRNHO"))
    root_directory.mkdir(parents=True, exist_ok=True)
    version_path = root_directory / "version.txt"

    if version_path.exists():
        local_version = version_path.read_text()
        print(f"Successfully found a local installation of the application at version '{local_version}'.\n")

    else:
        local_version = None
        print("Failed to find local installation of the application.\n")

    print("Attempting to fetch the latest release information from GitHub...")
    release_data = safe_request("https://api.github.com/repos/iRNHO/damage-calculator-data/releases/latest")

    if not release_data:
        if local_version:
            print("Failed to reach the GitHub API; attempting to launch a previous installation of the application...")
            subprocess.run([sys.executable, str(root_directory / "main.py")])
            return

        print("Failed to reach the GitHub API; please check your internet connection and try again.")
        return
    
    try:
        latest_version = json.loads(release_data)["tag_name"]

    except (json.JSONDecodeError, KeyError):
        if local_version:
            print("Failed to parse the latest release information from GitHub; attempting to launch a previous installation of the application...")
            subprocess.run([sys.executable, str(root_directory / "main.py")])
            return

        print("Failed to parse the latest release information from GitHub; please try again later.")
        return

    print(f"Successfully reached the GitHub API; latest version is '{latest_version}'.\n")

    if latest_version != local_version or args.force_install:
        print("Attempting to download the application data at the latest version...")

        for attempt in range(4):
            if attempt == 3:
                print("Failed to download the application data after multiple attempts; please check your internet connection and try again.")
                return

            data_bytes = safe_request(f"https://github.com/iRNHO/damage-calculator-data/releases/download/{latest_version}/data.zip")

            if data_bytes:
                with tempfile.TemporaryDirectory() as temp_directory:
                    temp_path = Path(temp_directory)
                    try:
                        with zipfile.ZipFile(io.BytesIO(data_bytes)) as zip_file:
                            zip_file.extractall(temp_path)
                    except zipfile.BadZipFile:
                        continue

                    for item in root_directory.iterdir():
                        if item.is_dir():
                            shutil.rmtree(item)
                        else:
                            item.unlink()

                    for item in temp_path.iterdir():
                        shutil.move(str(item), root_directory / item.name)

                version_path.write_text(latest_version)
                print("Successfully downloaded the application data and updated the local installation.\n")
                break

            time.sleep(2 ** attempt)

        print("Attempting to launch the application...")

    else:
        print("The local installation is already up to date; attempting to launch the application...")
        
    subprocess.run([sys.executable, str(root_directory / "main.py")])

if __name__ == "__main__":
    main()
