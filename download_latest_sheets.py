import os
import shutil
import urllib.request

file_ids = {
    "armour": "1aXimE69Om-RQBeZSF1WNi1A-bxvBI9-EntkpohaVN7E",
    "skill": "1CCh1e-uOkx23xhLrmSHZxgKsCvOXZm-judyUiFxcyMQ",
    "talent": "1OW3mZmidotVKX3b_kYtkdT7AcnQxUK3qfVRbg49iX44",
    "weapon": "1Ce2l-5BwG0NjQ9kartKT-m48rp1eEWvGiH-e0BmNx70"
}

if os.path.exists("data"):
    shutil.rmtree("data")

os.makedirs("data")

for file_name, file_id in file_ids.items():
    url = f"https://docs.google.com/spreadsheets/d/{file_id}/export?format=xlsx"
    output_path = os.path.join("data", f"{file_name}.xlsx")
    
    with urllib.request.urlopen(urllib.request.Request(url)) as response:
        with open(output_path, "wb") as file:
            file.write(response.read())
    
    print(f"Successfully downloaded '{file_name}' sheet to '{output_path}'.")
