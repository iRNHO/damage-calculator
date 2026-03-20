import subprocess

PYPROJECT_PATH = "pyproject.toml"
MAIN_PATH = "damage_calculator/main.py"

# Helper to update version in pyproject.toml
def update_pyproject_version(new_version):
	with open(PYPROJECT_PATH, "r", encoding="utf-8") as f:
		lines = f.readlines()
	with open(PYPROJECT_PATH, "w", encoding="utf-8") as f:
		for line in lines:
			if line.startswith("version = "):
				f.write(f"version = \"{new_version}\"\n")
			else:
				f.write(line)

# Helper to update LAUNCHER_VERSION in main.py
def update_launcher_version(new_version):
	with open(MAIN_PATH, "r", encoding="utf-8") as f:
		lines = f.readlines()
	with open(MAIN_PATH, "w", encoding="utf-8") as f:
		for line in lines:
			if line.startswith("LAUNCHER_VERSION = "):
				f.write(f"LAUNCHER_VERSION = \"{new_version}\"\n")
			else:
				f.write(line)

def main():
	# Fetch current version from PyPI
	import json
	from urllib.request import urlopen
	try:
		with urlopen("https://pypi.org/pypi/irnho-damage-calculator/json", timeout=5) as response:
			data = json.loads(response.read())
		current_version = data["info"]["version"]
	except Exception:
		current_version = None
	print(f"Current version on PyPI: {current_version}")
	new_version = input("Enter new version (e.g. '0.1.8'): ").strip()
	if not new_version:
		print("No version entered. Aborting.")
		return
	update_pyproject_version(new_version)
	update_launcher_version(new_version)
	subprocess.run(["git", "add", "-A"])
	commit_message = input("Enter commit message: ").strip()
	subprocess.run(["git", "commit", "-m", commit_message])
	subprocess.run(["git", "push", "origin", "main"])
	subprocess.run(["git", "tag", "-f", f"v{new_version}"])
	subprocess.run(["git", "push", "-f", "origin", f"v{new_version}"])

if __name__ == "__main__":
	main()
