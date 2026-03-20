# iRNHO's Damage Calculator Launcher

The launcher will attempt to find a local installation of the application, check for the latest version on GitHub, and update the local installation if necessary before launching the application.

## Prerequisites

You must have the [uv](https://docs.astral.sh/uv/getting-started/installation/) package manager in order to install the launcher. Use the link and follow the instructions for your platform, which should amount to a single one-time command. 

## Installation

One-time command:
```
uv tool install irnho-damage-calculator
```


## Usage

To run the launcher, use the command:
```
dcl
```

## Parameters (Optional)

- `-h`, `--help`: Show usage and available options.

	```
	dcl -h
	```
- `-f`, `--force`: Force a fresh installation of the application regardless of local version.

	```
	dcl -f
	```
