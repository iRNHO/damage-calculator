"""
This script is designed to be the distributable component of the calculator, and performs the following operations:

    • Checks for the latest release version on GitHub.
    • Compares it to the locally stored version (if any).
    • If the local version is outdated or non-existent, downloads the latest release files from GitHub.
    • Ensures that all required Python packages (as specified in 'requirements.txt') are installed.
    • Finally, it runs 'main.py' to start the application.

"""

__author__ = "iRNHO"
__contact__ = "Message iRNHO on XBOX or irnho on discord regarding any questions, feedback, bug-reporting etc."
__discord__ = "https://discord.gg/--------"


#################### SECTION BREAK ####################

##### CONSTANTS #####

FORCE_FRESH_INSTALL = False


#################### SECTION BREAK ####################

##### IMPORT STATEMENTS #####

import sys
import json
import shutil
import subprocess
import urllib.request
import importlib.metadata

from pathlib import Path
from packaging import version


#################### SECTION BREAK ####################

##### UTILITY FUNCTIONS #####

def setup_package(package_string: str) -> None:
    """
    Ensures that a given package is installed and at an acceptable version.
        
    Parameters:
        • package_string (str): The package name, necessarily followed by either '>=version' or '==version'.
        
    """

    if ">=" in package_string:
        package_name, minimum_version = package_string.split(">=")
        print(f"Ensuring package '{package_name}' is installed at version '{minimum_version}' or higher...")

        try:
            installed_version = importlib.metadata.version(package_name)

            if version.parse(installed_version) >= version.parse(minimum_version): 
                print(f"Package '{package_name}' is already installed at version '{installed_version}'; no action required.\n")
                return
            
            print(f"Package '{package_name}' is already installed at version '{installed_version}'; lower than the required minimum version '{minimum_version}'.")

        except importlib.metadata.PackageNotFoundError: 
            print(f"Package '{package_name}' is not currently installed.")
            pass
        
        try:
            print(f"Attempting to install the latest version of package '{package_name}'...")
            subprocess.run([sys.executable, "-m", "pip", "install", "--quiet", package_name], check=True)
            installed_version = importlib.metadata.version(package_name)

            if version.parse(installed_version) >= version.parse(minimum_version):
                print(f"Successfully installed package '{package_name}' at version '{installed_version}'.\n")
                return

        except subprocess.CalledProcessError:
            pass
        
        raise ValueError(f"Failed to install package '{package_name}' at version '{minimum_version}' or higher; please contact iRNHO for assistance.")

    elif "==" in package_string:
        package_name, exact_version = package_string.split("==")
        print(f"Ensuring package '{package_name}' is installed at version '{exact_version}' exactly...")

        try:
            installed_version = importlib.metadata.version(package_name)

            if installed_version == exact_version: 
                print(f"Package '{package_name}' is already installed at version '{installed_version}'; no action required.\n")
                return
            
            print(f"Package '{package_name}' is already installed at version '{installed_version}'; does not match the required version '{exact_version}' exactly.")

        except importlib.metadata.PackageNotFoundError: 
            print(f"Package '{package_name}' is not currently installed.")
            pass
        
        try:
            print(f"Attempting to install package '{package_name}' at version '{exact_version}' exactly...")
            subprocess.run([sys.executable, "-m", "pip", "install", "--quiet", f"{package_name}=={exact_version}"], check=True)
            installed_version = importlib.metadata.version(package_name)

            if installed_version == exact_version:
                print(f"Successfully installed package '{package_name}' at version '{installed_version}'.\n")
                return

        except subprocess.CalledProcessError:
            pass
        
        raise Exception(f"Failed to install package '{package_name}' at version '{exact_version}' exactly; please contact iRNHO for assistance.")

    else:
        raise ValueError(f"Package string '{package_string}' is not in a recognised format; please contact iRNHO for assistance.")


def safe_request(url, timeout=10):
    """
    Safely fetches bytes from a URL; returning 'None' if offline or unreachable.

    Parameters:
        • url (str): The URL to request.
        • timeout (int): The timeout duration in seconds.    
    
    """

    try:
        with urllib.request.urlopen(url, timeout=timeout) as response: return response.read()

    except Exception: return None


#################### SECTION BREAK ####################

##### MAIN #####

if __name__ == "__main__":

    setup_package("platformdirs>=4.3.6")

    from platformdirs import user_data_dir

    print("Attempting to find a previous installation of the application...")
    root_directory = Path(user_data_dir("Damage Calculator", "iRNHO"))
    root_directory.mkdir(parents=True, exist_ok=True)
    version_file = root_directory / "version.json"

    if version_file.exists():
        try: 
            local_version = json.loads(version_file.read_text())["version"]

        except Exception: local_version = None

    if local_version:
        print(f"Successfully found evidence of a previous installation of the application at version '{local_version}'.\n")

    else:
        print("Failed to find evidence of a previous installation of the application.\n")

    print("Attempting to fetch the latest release information from GitHub...")

    data_bytes = safe_request("https://api.github.com/repos/iRNHO/damage-calculator/releases/latest")

    if not data_bytes:
        if local_version:
            print("Failed to reach the GitHub API; attempting to launch a previous installation of the application...")
            subprocess.run([sys.executable, str(root_directory / "main.py")], check=True)
            exit(0)

        else:
            raise Exception("Failed to reach the GitHub API; please check your internet connection and try again.")
        
    else:
        latest_version = json.loads(data_bytes)["tag_name"]
        print(f"Successfully reached the GitHub API; latest version is '{latest_version}'.\n")

        if latest_version != local_version or FORCE_FRESH_INSTALL:
            print("Attempting to install the application at the latest version..." if not local_version else "Attempting to reinstall the application at the latest version..." if FORCE_FRESH_INSTALL else "Attempting to update the application to the latest version...")

            for file in ["main.py", "requirements.txt"]:
                data_bytes = safe_request(f"https://github.com/iRNHO/damage-calculator/releases/latest/download/{file}")

                if data_bytes:
                    (root_directory / file).write_bytes(data_bytes)

                else:
                    raise Exception(f"Failed to fetch file '{file}' from GitHub.")

            for folder in ["data", "assets"]:
                folder_path = root_directory / folder

                if folder_path.exists():
                    shutil.rmtree(folder_path)

                folder_path.mkdir()
                listing = safe_request(f"https://api.github.com/repos/iRNHO/damage-calculator/contents/{folder}?ref=main")

                if not listing:
                    raise Exception(f"Failed to fetch the file listing for folder '{folder}' from GitHub.")

                for item in json.loads(listing.decode()):
                    if item["type"] == "file":
                        file_data = safe_request(item["download_url"])

                        if file_data:
                            (folder_path / item["name"]).write_bytes(file_data)

                        else:
                            raise Exception(f"Failed to fetch file '{folder}/{item["name"]}' from GitHub.")
            
            print("Successfully downloaded core application files.\n")
            requirements_file = root_directory / "requirements.txt"

            if requirements_file.exists():
                for line in requirements_file.read_text().splitlines():
                    if line.strip() and not line.startswith("#"):
                        setup_package(line.strip())

            print("Successfully ensured all required packages are installed at the required versions.\n")
            version_file.write_text(json.dumps({"version": latest_version}))
            print("Successfully installed the latest version; attempting to launch the application..." if not local_version else "Successfully updated the previous installation to the latest version; attempting to launch the application...")
            subprocess.run([sys.executable, str(root_directory / "main.py")], check=True)
            exit(0)

        else:
            print("The previous installation is already up to date; attempting to launch the application...")
            subprocess.run([sys.executable, str(root_directory / "main.py")], check=True)
            exit(0)
