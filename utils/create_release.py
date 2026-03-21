import json
import subprocess

from urllib.request import urlopen

def change_line(file_path, starts_with, new_line):
    with open(file_path, "r", encoding="utf-8") as f:
        lines = f.readlines()
    with open(file_path, "w", encoding="utf-8") as f:
        for line in lines:
            if line.startswith(starts_with):
                f.write(new_line + "\n")
            else:
                f.write(line)

try:
    release_data = urlopen("https://pypi.org/pypi/irnho-damage-calculator/json").read()
    print(f"Latest release: v{json.loads(release_data)["info"]["version"]}")

except Exception:
    print("Failed to fetch latest release information.")

new_version = input("Enter new release version (e.g. 'v1.2.3'): ").strip()
commit_message = input("Enter commit message: ").strip()

change_line("pyproject.toml", "version = \"", f"version = \"{new_version[1:]}\"")
change_line("damage_calculator_launcher/__init__.py", "__version__ = \"", f"__version__ = \"{new_version[1:]}\"")

for args in [
    ["add", "-A"],
    ["commit", "-m", commit_message],
    ["push", "origin", "main"],
    ["tag", "-f", new_version],
    ["push", "-f", "origin", new_version]
]:
    subprocess.run(["git"] + args)
