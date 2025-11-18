import urllib.request

from pathlib import Path

with urllib.request.urlopen(urllib.request.Request("https://docs.google.com/spreadsheets/d/1V8TNWd33-PuOggIC3dM186NVpuf4b0cIPDVTpwXlohQ/export?format=xlsx")) as response:
    Path("iRNHO's Spreadsheet.xlsx").write_bytes(response.read())
