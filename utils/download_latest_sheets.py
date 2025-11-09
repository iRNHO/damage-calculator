import shutil
import urllib.request

from pathlib import Path

file_ids = {
    "Armor": "1aXimE69Om-RQBeZSF1WNi1A-bxvBI9-EntkpohaVN7E",
    "Skill": "1CCh1e-uOkx23xhLrmSHZxgKsCvOXZm-judyUiFxcyMQ",
    "Talent": "1OW3mZmidotVKX3b_kYtkdT7AcnQxUK3qfVRbg49iX44",
    "Weapon": "1Ce2l-5BwG0NjQ9kartKT-m48rp1eEWvGiH-e0BmNx70",
}

data_directory = Path("data")

if data_directory.exists():
    shutil.rmtree(data_directory)

data_directory.mkdir()

for file_name, file_id in file_ids.items():
    url = f"https://docs.google.com/spreadsheets/d/{file_id}/export?format=xlsx"

    with urllib.request.urlopen(urllib.request.Request(url)) as response:
        (data_directory / f"{file_name}.xlsx").write_bytes(response.read())
