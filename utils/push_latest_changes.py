import json
import subprocess
import urllib.request

try:
    data = urllib.request.urlopen(f"https://api.github.com/repos/iRNHO/damage-calculator/releases/latest").read()
    latest = json.loads(data)["tag_name"]
    print(f"Latest release: {latest}")

except Exception:
    print("Error occured while attempting to fetch latest release info from GitHub; assuming no latest release.")
    latest = None

new_tag = input("Enter new release version (e.g. '1.2.0' or 'v1.2.0'): ").strip()
commit_message = input("Enter commit message: ").strip()

if not new_tag.startswith("v"):
    new_tag = "v" + new_tag

commands = [
    ["git", "add", "-A"],
    ["git", "commit", "-m", commit_message],
    ["git", "push", "origin", "main"],
    ["git", "tag", "-d", new_tag],
    ["git", "push", "origin", f":refs/tags/{new_tag}"],
    ["git", "tag", "-a", new_tag, "-m", f"Release {new_tag}"],
    ["git", "push", "origin", new_tag]
]

for command in commands:
    subprocess.run(command, check=True)
