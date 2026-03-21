# iRNHO's Damage Calculator Launcher

This launcher was created to provide easy installation and versioning for iRNHO's Damage Calculator application.

The launcher checks for the latest release of the application on GitHub and compares it to the locally installed version if present. It will then download and install/update the core application files as necessary, ensuring to preserve build data.

In cases of application file corruption, a factory reset option is available, which wipes all application data and performs a clean installation from scratch. Please note that this will not preserve build data.

The launcher also verifies its own version against PyPI to ensure users are running the latest launcher code, prompting the user to update when necessary. All operations are performed in the user's platform-specific data directory for isolation and reliability.

## Prerequisites

You must have the [uv](https://docs.astral.sh/uv/getting-started/installation/) package manager in order to install the launcher. Please follow the instructions at the provided link for your platform; installation typically requires running a single command in your terminal.

## Installation

One-time installation command:
```
uv tool install irnho-damage-calculator
```


## Usage

Once installed, use the following command to run the launcher:
```
dcl
```

## Parameters (Optional)

- `-h`, `--help`: Show usage and available options.

	```
	dcl -h
	```
- `-f`, `--factory-reset`: During setup, wipe all application data and perform a clean installation from scratch.

	```
	dcl -f
	```

## Contact & Support

For help, troubleshooting, or more information about iRNHO's Damage Calculator and its launcher, please join the **[OMD Community Discord](https://discord.gg/xxxxxx)**. All support, updates, and project information are provided through the Discord server.

## License

This project is licensed under the Creative Commons Attribution-NonCommercial 4.0 International License (CC BY-NC 4.0). You are free to use, modify, and share the code for non-commercial purposes, provided you give appropriate credit to the author (iRNHO). Commercial use is not permitted.

See the LICENSE file or visit [CC BY-NC 4.0](https://creativecommons.org/licenses/by-nc/4.0/) for full details.