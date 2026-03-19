import shutil
import subprocess
from pathlib import Path

from platformdirs import user_data_dir


APP_NAME = "Damage Calculator"
AUTHOR = "iRNHO"


def read_version(path):
    if not path.exists():
        return None
    return path.read_text().strip()


def main():
    print("Launcher starting...")

    root = Path(user_data_dir(APP_NAME, AUTHOR))
    root.mkdir(parents=True, exist_ok=True)

    print(f"Install location: {root}")

    source = Path.cwd() / "app_source"
    install = root

    source_version_file = source / "version.txt"
    install_version_file = install / "version.txt"

    source_version = read_version(source_version_file)
    install_version = read_version(install_version_file)

    print(f"Source version: {source_version}")
    print(f"Installed version: {install_version}")

    if source_version != install_version:
        print("Updating local install...\n")

        # Clear old install
        for item in install.iterdir():
            if item.is_dir():
                shutil.rmtree(item)
            else:
                item.unlink()

        # Copy new files
        shutil.copytree(source, install, dirs_exist_ok=True)

    else:
        print("Already up to date.\n")

    app_path = install / "main.py"

    if not app_path.exists():
        print("App not found after install.")
        return

    print("Running app...\n")
    subprocess.run(["python", str(app_path)])


if __name__ == "__main__":
    main()