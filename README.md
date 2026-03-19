# iRNHO's Damage Calculator

A launcher that automatically installs and keeps the Damage Calculator up to date.

## Installation

You do **not** need Python installed. You only need [uv](https://docs.astral.sh/uv/getting-started/installation/).

**Install uv** (run once):

- **Windows:** `winget install --id=astral-sh.uv -e` or download from [astral.sh/uv](https://docs.astral.sh/uv/getting-started/installation/)
- **macOS/Linux:** `curl -LsSf https://astral.sh/uv/install.sh | sh`

**Install the calculator** (run once):

```
uv tool install irnho-damage-calculator
```

## Running

```
damage-calculator
```

On first launch it will download the latest version automatically. Subsequent launches check for updates and start immediately if already up to date.

## Updating

Updates are applied automatically on launch. No manual steps required.
