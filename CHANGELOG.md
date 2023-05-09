
# Changelog
<!-- https://keepachangelog.com/en/1.0.0/ -->

## 0.3.0-rc1 (2023-05-08)

### Added
- RDP support for Devolutions RDM (Excel->RDM). SecureCRT is not supported because it has poor support of the RDP protocol.
- WEB support for Devolutions RDM (Excel->RDM). SecureCRT has no support for web-based sessions.
  
### Changed
- Column names in Excel are changed (for details see [README.md](README.md) file or [config.yaml](config.yaml)).
- Worksheets and/or column names for unused settings is not required.

### Fixed
- Missing worksheet in Excel file is skipped with warning message.
- Missing optional column in Excel file is skipped with warning message.
- Missing required column in Excel file stop processing with error message.
  
## 0.2.4 (2023-04-28)

### Fixed
- Devolutions RDM (Credentials): Credentials (ConnectionType 26) 'ID' added. Interconnect between SSH Session (Type 77) and Credentials (Type 26) corrected.
  
## 0.2.3 (2023-04-27)

### Fixed
- SecureCRT sessions export (Excel -> SecureCRT): Keyword list from Excel is now inserted to SecureCRT XML file correctly
- Devolutions RDM (Excel->RDM): SSH session port number is now inserted into JSON RDM file as correct attribute
## 0.2.2 (2022-01-18)

### Fixed
- SecureCRT sessions export (Excel -> SecureCRT): SSH Authentications changed from 'keyboard-interactive' only to 'keyboard-interactive,password'

## 0.2.1 (2022-12-31)

### Other

- README update
- CHANGELOG and LICENSE added

## 0.2.0 (2022-12-29)

### New

- Support for Devolutions RDM (Excel -> Devolutions RDM (json))

### Changed 

- Excel worksheets column names added (Devolutions RDM) or changed (SecureCRT)
- Script help (RDM support)

## 0.1.0 (2022-11-25)

### New

- Initial version of Session Maker (supporting Excel -> SecureCRT (xml) and vice-versa).
