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
  - [Session Maker](#session-maker)
    - [Usage](#usage)
    - [Example](#example)
  - [Session Reader](#session-reader)
    - [Usage](#usage-1)
    - [Example](#example-1)
  - [Excel workbook structure](#excel-workbook-structure)
    - [Columns of sessions worksheet](#columns-of-sessions-worksheet)
    - [Columns of credentials worksheet](#columns-of-credentials-worksheet)
    - [Columns of firewalls worksheet](#columns-of-firewalls-worksheet)

## Description

Excel workbook to SecureCRT sessions (and vice-versa) converter. There are two parts:

- [**Session Maker**](#session-maker) - Generate SecureCRT XML file from Excel book source (Excel -> XML)
- [**Session Reader**](#session-reader) - Generate Excel book from SecureCRT XML sessions export file (XML -> Excel)

## Session Maker

Reads Excel workbook and generate XML session content for SecureCRT.

### Usage

It is simple - read help :).

```
$ python3 session_maker.py -h
usage: session_maker.py [-h] [--config CONFIG] [--type {scrt,rdm}] [--write DESTINATION | -p] [-q | -v] source

Read Excel file (source) and generate sessions XML file for SecureCRT.

positional arguments:
  source                Source (XLS) file

options:
  -h, --help            show this help message and exit
  --config CONFIG       Configuration settings file (default=config.yaml)
  --type {scrt,rdm}     Destination type: scrt=SecureCRT (default), rdm=DevolutionsRDM
  --write DESTINATION, -w DESTINATION
                        Write to file. If not specified, write to 'export' subfolder as the source.
  -p, --print           Print to screen only (don't write it to the file).
  -q, --quiet           Quiet output.
  -v, --verbose         Verbose output. (use: -v, -vv)

```

XML content can be exported to:

- **file**: option `--write`. if not defined, the file is stored in `export` subfolder
- **stdout**: option `--print`

### Example

**Source file**

Excel (source) file:

```
$ ls data/EXAMPLE/
devices.xlsx
```

**Build process**

Build XML content for SecureCRT from Excel source:

```
$ python3 session_maker.py  data/EXAMPLE/devices.xlsx
Reading arguments...
Done.
Exporting sessions for SecureCRT...
Exported: 5 sessions, 2 credential groups, 2 firewall groups.
Done.
```

**Destination file**

XML file is exported to `export` subfolder (because option `--write` or `--print` is not defined):

```
$ ls data/EXAMPLE/export/
devices.xml
```

## Session Reader

Reads SecureCRT sessions file (SecureCRT menu: `Tools -> Export settings...`) and export it to Excel workbook.

### Usage

```
$ python session_reader.py -h
usage: session_reader.py [-h] [--config CONFIG] [-w DESTINATION] [-q | -v] source

Read SecureCRT sessions XML file (source) and export it to Excel file (write to destination).

positional arguments:
  source                SecureCRT sessions XML file (export from SecureCRT).

options:
  -h, --help            show this help message and exit
  --config CONFIG       Configuration settings file (default=config.yaml)
  -w DESTINATION, --write DESTINATION
                        Write to destination Excel (xlsx) file. If not defined, write to the 'export' subfolder.
  -q, --quiet           Quiet output.
  -v, --verbose         Verbose output (use: -v, -vv).

```

If `--write` option is not defined, destination file is exported to `export` subfolder.

### Example

**Source file**

SecureCRT (source) file (it is previously generated Excel file):

```
$ ls data/EXAMPLE/export/
devices.xml
```

**Build process**

Generate Excel file from SecureCRT XML source file:

```
$ python session_reader.py data/EXAMPLE/export/devices.xml
Reading arguments...
Done.
Read SecureCRT sessions XML file...
Done. Imported: 5 sessions, 2 credential groups, 2 firewall groups.
Writing Excel file...
Done.
```

**Destination file**

Excel workbook is exported to `export` subfolder (because option `--write` is not defined):

```
$ ls data/EXAMPLE/export/export/
devices.xlsx
```

## Excel workbook structure

Excel file workbook contains 3 worksheets:

- **sessions**: list of device sessions
- **credentials**: list of credential groups
- **firewalls**: list of firewall groups

### Columns of sessions worksheet

| column name      | required | default | description                                           |
| ---------------- | :------: | ------- | ----------------------------------------------------- |
| folder           |          |         | Path/hierarchy to session                             |
| session          |   yes    |         | Session name                                          |
| hostname         |          |         | Device hostname (DNS name or IP address)              |
| port             |          | 22      | TCP port                                              |
| username         |          |         | Username                                              |
| credential group |          |         | Credential group name                                 |
| colorscheme      |          |         | Color Scheme name                                     |
| keywords         |          |         | Keyword Highlighting List name                        |
| firewall group   |          |         | Firewall group or Session name (use: path/to/session) |

**SecureCRT specific fields:**

- credential group
- colorscheme
- keywords
- firewall group

### Columns of credentials worksheet

| column name      | required | default | description           |
| ---------------- | :------: | ------- | --------------------- |
| credential group |   yes    |         | Credential group name |
| username         |          |         | Username              |

### Columns of firewalls worksheet

| column name    | required | default | description            |
| -------------- | :------: | ------- | ---------------------- |
| firewall group |   yes    |         | Firewall group name    |
| address        |          |         | IP address or DNS name |
| port           |          |         | TCP port               |
| username       |          |         | Username               |
