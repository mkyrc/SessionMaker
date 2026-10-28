# SessionMaker

<!-- ```
 _____               _            ___  ___      _
/  ___|             (_)           |  \/  |     | |
\ `--.  ___  ___ ___ _  ___  _ __ | .  . | __ _| | _____ _ __
 `--. \/ _ \/ __/ __| |/ _ \| '_ \| |\/| |/ _` | |/ / _ \ '__|
/\__/ /  __/\__ \__ \ | (_) | | | | |  | | (_| |   <  __/ |
\____/ \___||___/___/_|\___/|_| |_\_|  |_/\__,_|_|\_\___|_|

```                                                                 -->

- [SessionMaker](#sessionmaker)
  - [Description](#description)
    - [Changelog](#changelog)
    - [Supported features](#supported-features)
  - [Session Maker](#session-maker)
    - [Usage](#usage)
    - [Example](#example)
  - [Session Reader](#session-reader)
    - [Usage](#usage-1)
    - [Example](#example-1)
  - [Excel workbook structure](#excel-workbook-structure)
    - ['sessions' worksheet columns](#sessions-worksheet-columns)
    - ['rdm-credentials' worksheet columns](#rdm-credentials-worksheet-columns)
    - ['rdm-host' worksheet columns](#rdm-host-worksheet-columns)
    - ['scrt-credentials' worksheet columns](#scrt-credentials-worksheet-columns)
    - ['scrt-firewalls' worksheet columns](#scrt-firewalls-worksheet-columns)
  - [Custom configuration](#custom-configuration)
  - [Author](#author)
  - [License](#license)

## Description

Excel workbook to SecureCRT and/or Devolutions RDM sessions (and vice-versa) converter. The goal of the project is to make life easier for network engineers ;-).

> **Important:**<br>
> Session Reader or Session Maker never works with secrets/password! There is possible define only usernames (including credentials).

There are two parts:

- [**Session Maker**](#session-maker) - generates sessions from Excel book to:
  - Generate SecureCRT sessions (`.xml`)
  - Devolutions RDM sessions (`.json`)
  - Generic RDP files (`.rdp`)
  - Remmina profile files (`.remmina`)
- [**Session Reader**](#session-reader) - Generate Excel book from SecureCRT sessions file

### Changelog

- Starting version 0.3.0, are worksheets or/and column names for unused settings not required.
- Starting version 0.4.0, are supported new RDM object types: "host object" and "PowerShell script before open"
- Starting version 0.5.0, can user define session defaults in the file

For more details see [CHANGELOG.md](CHANGELOG.md) file.

### Supported features

SecureCRT, Devolution RDM support:

| Feature         | SecureCRT | Devolution RDM | Description                                                 |
| --------------- | :-------: | :------------: | ----------------------------------------------------------- |
| Folders         |    yes    |      yes       | Path to session (or any other object, e.g. credentials,...) |
| Username        |    yes    |      yes       | Username (not used when credentials is defined)             |
| Credentials     |    yes    |      yes       | Credentials definitions (higher preference than username)   |
| SSH Session     |    yes    |      yes       | SSH session support                                         |
| RDP Session     |    no     |      yes       | RDP session support                                         |
| WEB Session     |     -     |      yes       | WEB session support                                         |
| Colorscheme     |    yes    |       no       | Terminal colorscheme                                        |
| Keywords        |    yes    |       no       | Keyword highlighting (per session)                          |
| Firewall groups |    yes    |       -        | Firewall groups definition (SecureCRT feature)              |

Remmina and RDP support:

| Feature         | Remmina |  RDP  | Description                                                 |
| --------------- | :-----: | :---: | ----------------------------------------------------------- |
| Folders         |   yes   |   -   | Path to session (or any other object, e.g. credentials,...) |
| Username        |   yes   |  yes  | Username (not used when credentials is defined)             |
| SSH Session     |   yes   |   -   | SSH session support                                         |
| RDP Session     |   no    |  yes  | RDP session support                                         |
| Alternate shell |   yes   |  yes  | Alternate shell (Exec shell)                                |

## Session Maker

Reads Excel workbook and generate SecureCRT sessions content (XML file) or Devolutions RDM connections (JSON file).

```mermaid
graph LR;
    Excel["Excel (xlsx)"]-->SM("Session Maker")
    SM-->scrt["SecureCRT sessions file (.xml)"]
    SM-->rdm["DevolutionsRDM sessions file (.json)"];
    SM-->rdp["RDP files (.rdp)"];
    SM-->remmina["Remmina files (.remmina)"];
```

### Usage

It is simple - read help :).

Run the scripts from the `src` folder (`cd src`); all paths below are relative to it.

```
$ python3 session_maker.py -h
usage: session_maker.py [-h] [--config CONFIG] [--type {scrt,rdm,rdp,remmina}] [--write DESTINATION | -p] [-q | -v] source

Read Excel file (source) and generate sessions output for [SecureCRT|Devolutions|RDP|Remmina].

positional arguments:
  source                Source (XLS) file

options:
  -h, --help            show this help message and exit
  --config CONFIG       Configuration settings file (default=config.yaml)
  --type {scrt,rdm,rdp,remmina}
                        Destination type: scrt=SecureCRT (default), rdm=DevolutionsRDM, rdp=RDP files (.rdp), remmina=Remmina profiles (.remmina)
  --write DESTINATION, -w DESTINATION
                        Write to file. If not specified, write to 'export' subfolder as the source.
  -p, --print           Print to screen only (don't write it to the file).
  -q, --quiet           Quiet output.
  -v, --verbose         Verbose output. (use: -v, -vv)
  --version             show program's version number and exit
```

Export can be written as:

- **file/folder**: Option `--write`. If not defined, output is stored in `export` subfolder
- **stdout**: Option `--print`.

> **Note:** For `--type rdp`, export is generated as multiple files in a destination folder (one file per session) and `--print` is not supported. Only RDP sessions are exported (other types are ignored) as `.rdp` files built from the template `rdp.template.rdp` in `config.yaml` (default `templates/rdp_default.rdp`) with `full address`, `username` and `alternate shell` (Excel column `rdp alternate shell`) taken from Excel.

> **Note:** For `--type remmina`, export is generated as `.remmina` files in one destination folder (`--print` is not supported). Only RDP sessions are exported. File name is `<group>_<type>_<name>_<ip>.remmina`, where `group` is the Excel `folder` value with `/` replaced by `-`. Defaults are read from the template `remmina.template.rdp` in `config.yaml` (default `templates/remmina_rdp_default.remmina`) and `name`, `group`, `server`, `username` and `exec` (Excel column `rdp alternate shell`) are taken from Excel.

### Example

<details>
  <summary>Show Excel -> SecureCRT example</summary>

**Source file**

Excel (source) file:

```
$ ls data/EXAMPLE/
devices-v0.3.x.xlsx
```

**Build process**

Build XML content for SecureCRT from Excel source:

```
$ python3 session_maker.py data/EXAMPLE/devices-v0.3.x.xlsx
Reading arguments...
Done.
Reading Excel book...
Done: 4 sessions (ssh: 4), 2 credential group(s), 0 firewall group(s) from Excel.
Building sessions...
Done.
Writing to 'data/EXAMPLE/export/devices-v0.3.x.xml'...
Done.
```

**Destination file**

XML file is exported to `export` subfolder (because option `--write` or `--print` is not defined):

```
$ ls data/EXAMPLE/export/
devices-v0.3.x.xml
```

</details>

<details>
  <summary>Show Excel -> Devolution RDM example</summary>

**Source file**

Excel (source) file:

```
$ ls data/EXAMPLE/
devices-v0.3.x.xlsx
```

**Build process**

Build XML content for SecureCRT from Excel source:

```
$ python3 session_maker.py data/EXAMPLE/devices-v0.3.x.xlsx --type rdm

Reading arguments...
Done.
Reading Excel book...
Done. 12 session(s) (ssh: 4, rdp: 4, web: 4), 2 credential(s) from Excel.
Building sessions...
Done.
Writing to 'data/EXAMPLE/export/devices-v0.3.x.json'...
Done.
```

**Destination file**

JSON file is exported to `export` subfolder (because option `--write` or `--print` is not defined):

```
$ ls data/EXAMPLE/export/
devices-v0.3.x.json
```

</details>

## Session Reader

Reads SecureCRT sessions XML file (SecureCRT menu: `Tools -> Export settings...`) and export it to Excel workbook.

```mermaid
graph LR;
    scrt["SecureCRT (XML)"]-->SR("Session Reader")
    SR-->Excel["Excel (xlsx)"]

```

### Usage

```
$ python session_reader.py -h
usage: session_reader.py [-h] [--config CONFIG] [--type {scrt}] [-w DESTINATION] [-q | -v] source [source ...]

Read SecureCRT XML and export them to Excel file (write to destination).

positional arguments:
  source                Source file(s): SecureCRT XML.

options:
  -h, --help            show this help message and exit
  --config CONFIG       Configuration settings file (default=config.yaml)
  --type {scrt}         Source type: scrt=SecureCRT XML (default).
  -w DESTINATION, --write DESTINATION
                        Write to destination Excel (xlsx) file. If not defined, write to the 'export' subfolder.
  -q, --quiet           Quiet output.
  -v, --verbose         Verbose output (use: -v, -vv).
```

If `--write` option is not defined, destination file is exported to `export` subfolder.

### Example

<details>
  <summary>SecureCRT (XML) -> Excel example</summary>

**Source file**

SecureCRT (source) file (it is previously generated Excel file):

```
$ ls data/EXAMPLE/export/
devices-v0.3.x.xml
```

**Build process**

Generate Excel file from SecureCRT XML source file:

```
$ python3 session_reader.py data/EXAMPLE/export/devices-v0.3.x.xml
Reading arguments...
Done.
Reading SecureCRT sessions XML file...
Done. 4 session(s), 2 credential group(s), 0 firewall group(s) from XML file.
Writing Excel file...
Done.
```

**Destination file**

Excel workbook is exported to `export` subfolder (because option `--write` is not defined):

```
$ ls data/EXAMPLE/export/export/
devices.xlsx
```

</details>

## Excel workbook structure

Excel workbook contains 4 worksheets:

- **[sessions](#sessions-worksheet-columns)**: list of device sessions
- **[rdm-credentials](#rdm-credentials-worksheet-columns)**: Devolutions RDM credentials list
- **[scrt-credentials](#scrt-credentials-worksheet-columns)**: SecureCRT credential groups list
- **[scrt-firewalls](#scrt-firewalls-worksheet-columns)**: SecureCRT firewall groups list

> **Note:** <br>
> All column names described below can be changed in [config.yaml](src/config.yaml) file.

### 'sessions' worksheet columns

| column name (Excel)     | required | default    | description                                                    |
| ----------------------- | :------: | ---------- | -------------------------------------------------------------- |
| folder                  |          |            | Path/hierarchy to session                                      |
| session                 |   yes    |            | Session name                                                   |
| hostname                |          |            | Device hostname (DNS name or IP address)                       |
| type                    |   yes    |            | Session type: [ssh \| rdp \| web]                              |
| port                    |          | 22 or 3389 | TCP port (default: 22 for SSH, 3389 for RDP, not used for WEB) |
| username                |          |            | Username                                                       |
| rdp alternate shell     |          |            | RDP alternative shell                                          |
| **Devolution RDM only** |
| rdm credential          |          |            | Credential object name                                         |
| rdm host                |          |            | Host object name                                               |
| rdm script before open  |          |            | PowerShell script before open path                             |
| web login form          |          |            | WEB session: Login form name (HTML id)                         |
| web login field         |          |            | WEB session: Login field name (HTML id)                        |
| web passwd field        |          |            | WEB session: Password field name (HTML id)                     |
| **SecureCRT only**      |
| scrt credentials        |          |            | Credential group name                                          |
| scrt colorscheme        |          |            | Color Scheme name                                              |
| scrt keywords           |          |            | Keyword Highlighting List name                                 |
| scrt firewall group     |          |            | Firewall group or Session name (path/to/session)               |

> **Note:**
> For details see `['excel']['col_names_sessions']` in [config.yaml](src/config.yaml) file.

### 'rdm-credentials' worksheet columns

| column name     | required | default | description               |
| --------------- | :------: | ------- | ------------------------- |
| folder          |          |         | Path to credential object |
| credential name |   yes    |         | Credential name           |
| username        |          |         | Username                  |

> **Note:**
> For details see `['excel']['col_names_rdm_credentials']` in [config.yaml](config.yaml) file.

### 'rdm-host' worksheet columns

| column name    | required | default | description                |
| -------------- | :------: | ------- | -------------------------- |
| folder         |          |         | Path to RDM host object    |
| name           |   yes    |         | RDM Host object name       |
| host           |   yes    |         | IP od hostname             |
| rdm credential |          |         | RDM Credential object path |

> **Note:**
> For details see `['excel']['col_names_rdm_credentials']` in [config.yaml](src/config.yaml) file.

### 'scrt-credentials' worksheet columns

| column name      | required | default | description           |
| ---------------- | :------: | ------- | --------------------- |
| credential group |   yes    |         | Credential group name |
| username         |          |         | Username              |

> **Note:**
> For details see `['excel']['col_names_scrt_credentials']` in [config.yaml](src/config.yaml) file.

### 'scrt-firewalls' worksheet columns

| column name    | required | default | description            |
| -------------- | :------: | ------- | ---------------------- |
| firewall group |   yes    |         | Firewall group name    |
| address        |          |         | IP address or DNS name |
| port           |          |         | TCP port               |
| username       |          |         | Username               |

> **Note:**
> For details see `['excel']['col_names_scrt_firewalls']` in [config.yaml](src/config.yaml) file.

## Custom configuration

Excel worksheet and/or column names and SecureCRT custom settings (not configurable by Session Maker) is possible to customize. 

For more detials see [config.yaml](src/config.yaml) file.

## Author

- **Martin Kyrc** - 2022-2026 - [@mkyrc](https://github.com/mkyrc) on Github

## License

This project is licensed under the MIT License - see the [LICENSE.txt](LICENSE.txt) file for details.
