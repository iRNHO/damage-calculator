import pandas as pd
from openpyxl import load_workbook
from pathlib import Path

from platformdirs import user_data_dir

root_directoy = Path(user_data_dir("Damage Calculator", "iRNHO"))

data = {name: pd.read_excel(root_directoy / "data" / f"{name}.xlsx", sheet_name=None) for name in ["Armor", "Skill", "Talent", "Weapon"]}

print(data.keys())
print(data["Armor"]["Exotic"].head())
