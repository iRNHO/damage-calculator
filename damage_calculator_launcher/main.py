import json
import argparse
import subprocess
import sys
import tempfile
import zipfile
import io
import shutil
import time

from . import __version__
from urllib.request import urlopen
from packaging.version import Version
from pathlib import Path
from platformdirs import user_data_dir

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
        if Version(__version__) < Version(latest_launcher_version):
            print(f"This launcher version is outdated (v{__version__} vs v{latest_launcher_version}); please reinstall the launcher using:\n\nuv tool install --reinstall irnho-damage-calculator\n")
            return

    parser = argparse.ArgumentParser(
        usage="dcl [-h] [-f]",
        description=f"iRNHO's Damage Calculator Launcher (v{__version__})",
        add_help=False        
    )
    parser.add_argument(
        "-h", "--help",
        action="help",
        help="Show this help message and exit."
    )
    parser.add_argument(
        "-f", "--factory-reset",
        action="store_true",
        help="Wipe all application data and perform a clean installation from scratch."
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
        print("Failed to find a local installation of the application.\n")

    print("Attempting to fetch the latest release information from GitHub...")
    release_data = safe_request("https://api.github.com/repos/iRNHO/damage-calculator-data/releases/latest")

    if not release_data:
        if local_version:
            print("Failed to fetch the latest release information; attempting to launch the local installation of the application...")
            subprocess.run([sys.executable, str(root_directory / "main.py")])
            return

        print("Failed to fetch the latest release information; please check your internet connection and try again.")
        return
    
    latest_version = json.loads(release_data)["tag_name"]
    print(f"Successfully fetched the latest release information; the latest version is '{latest_version}'.\n")

    if not local_version or Version(local_version) < Version(latest_version) or args.factory_reset:
        print("Attempting to download the application data at the latest version...")

        for attempt in range(4):
            if attempt == 3:
                print("Failed to download the application data after multiple attempts; please check your internet connection and try again.")
                return

            application_data = safe_request(f"https://github.com/iRNHO/damage-calculator-data/releases/download/{latest_version}/data.zip")

            if application_data:
                with tempfile.TemporaryDirectory() as temp_directory:
                    temp_path = Path(temp_directory)
                    try:
                        with zipfile.ZipFile(io.BytesIO(application_data)) as zip_file:
                            zip_file.extractall(temp_path)
                    except zipfile.BadZipFile:
                        continue

                    if args.factory_reset:
                        shutil.rmtree(root_directory)
                        root_directory.mkdir(parents=True, exist_ok=True)
                        for item in temp_path.iterdir():
                            shutil.move(str(item), root_directory / item.name)
                    else:
                        for item in temp_path.iterdir():
                            target = root_directory / item.name
                            if target.exists():
                                if target.is_dir():
                                    shutil.rmtree(target)
                                else:
                                    target.unlink()
                            shutil.move(str(item), target)

                version_path.write_text(latest_version)
                print(f"Successfully downloaded the application data and {"reset the" if args.factory_reset else "updated the" if local_version else "created a"} local installation.\n")
                break

            time.sleep(2 ** attempt)

        print("Attempting to launch the application...")

    else:
        print("The local installation is already up to date; attempting to launch the application...")
        
    subprocess.run([sys.executable, str(root_directory / "main.py")])

if __name__ == "__main__":
    main()
