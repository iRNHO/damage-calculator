import pandas as pd
from openpyxl import load_workbook
from pathlib import Path

data = {name: pd.read_excel(Path("data") / f"{name}.xlsx") for name in ["Armor", "Skill", "Talent", "Weapon"]}

print(data.keys())
print(data["Armor"]["Exotic"].head())