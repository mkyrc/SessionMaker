"""Devolutions Remote Desktop Manager (RDM) session generator"""

import logging
import xml.etree.ElementTree as ET
import uuid

from lib.sm import SessionMaker


# ========================================
# Class SMDevolutionsRDM
# ========================================
class SMDevolutionsRdm(SessionMaker):
    """SessionMaker - Devolutions RDM sessions generator class"""

    def __init__(
        self,
        settings: dict | None = None,
        excel_file: str | None = None,
        json_file: str = "",
        read_excel_file=False,
        credentials: dict | None = None,
        hosts: dict | None = None,
        session_defaults_rdm: dict = {},
        # **kwargs,
    ):
        """Initial method

        Args:
            parent:
            settings (dict): Configuartion settings (config.yaml content)
            excel_fle (str): Excel file path (source or destination)

            self:
            scrt_file (str): SecureCRT file path (destination or source)
            credentials (dict): Ordered dict of credentials
        """

        # parent class attribiutes:
        # - self._settings
        # - self._excel_file
        # - self._excel_obj
        # - self._xml_file
        # - self._xml_obj
        # - self._sessions_dict
        # - self._credentials_dict
        # - self._xml_session_file
        super().__init__(
            settings,
            excel_file,
            read_excel_file=True,
            session_defaults=session_defaults_rdm,
        )

        # rdm credential dict
        self.set_credentials_dict(credentials)

        # rdm hosts
        self._rdm_hosts_dict = {}
        self.set_hosts_dict(hosts)

        # session defaults
        # self.session_defaults = session_defaults_scrt

        # JSON file
        self.json_file = ""
        self._json_sessions = {}
        self._json_hosts = {}
        self._rdm_connection_list = []
        self.set_json_file(json_file, read_json_file=False)

    # ========================================
    # Private methods
    # ========================================

    # ========================================
    # Public methods
    # ========================================

    def excel_read_sheet_credentials(self, sheet_name: str) -> dict | list | bool:
        """Read excel sheet 'rdm_credentials' and return content as dict/array.

        Args:
            sheet_name (str): Sheet's name

        Returns:
            ordered dict: Column/Row-based dictionary (when get=['column', 'row']
            False: In case of error
        """
        credentials_dict_ret = self.excel_read_sheet(sheet_name, "column")

        if credentials_dict_ret is False:
            credentials_dict = None
        else:
            credentials_dict = self.col_name_normalize(
                credentials_dict_ret,
                self._settings["excel"]["col_names_rdm_credentials"],
            )

        if self.set_credentials_dict(credentials_dict) is False:
            return False

        return self._credentials_dict

    def excel_read_sheet_rdm_hosts(self, sheet_name: str) -> dict | list | bool:
        """Read excel sheet 'rdm_hosts' and return content as dict/array.

        Args:
            sheet_name (str): Sheet's name

        Returns:
            ordered dict: Column/Row-based dictionary (when get=['column', 'row']
            False: In case of error
        """
        sheet_dict_ret = self.excel_read_sheet(sheet_name, "column")
        if sheet_dict_ret is False:
            content_dict = None
        else:
            content_dict = self.col_name_normalize(
                sheet_dict_ret,
                self._settings["excel"]["col_names_rdm_hosts"],
            )
        if self.set_hosts_dict(content_dict):
            return self._rdm_hosts_dict

        return False

    def set_credentials_dict(self, credentials=None):
        """Set credentials dictionary. If not set, create empty.

        Args:
            credentials (dict): Ordered dict of RDM credentials.
        """
        excel_col_name = self._settings["excel"]["col_names_rdm_credentials"]
        keys = ["folder", "credential", "username"]
        required_keys = ["credential"]

        if credentials is None:
            for key in excel_col_name:
                self._credentials_dict[key] = []
        else:
            for key in excel_col_name:
                if key in keys:
                    try:
                        self._credentials_dict[key] = list(map(str, credentials[key]))
                    except KeyError:
                        logging.warning(
                            "Missing column name '%s' (key: '%s').",
                            excel_col_name[key],
                            key,
                        )
                        if key in required_keys:
                            logging.error(
                                "Missing required column '%s'.", excel_col_name[key]
                            )
                            return False
                        else:
                            logging.warning(
                                "Creating empty column name '%s'.", excel_col_name[key]
                            )
                            self._credentials_dict[key] = [""] * len(
                                credentials["credential"]
                            )

    def set_hosts_dict(self, hosts: dict | None = None):
        """Set hosts dictionary. If not set, create empty.

        Args:
            hosts (dict): Ordered dict of RDM credentials.
        """
        excel_col_name = self._settings["excel"]["col_names_rdm_hosts"]
        keys = ["folder", "name", "host", "rdm_credential"]
        required_keys = ["name"]

        if hosts is None:
            for key in excel_col_name:
                self._rdm_hosts_dict[key] = []
        else:
            for key in excel_col_name:
                if key in keys:
                    try:
                        self._rdm_hosts_dict[key] = list(map(str, hosts[key]))
                    except KeyError:
                        logging.warning(
                            "Missing column name '%s' (key: '%s').",
                            excel_col_name[key],
                            key,
                        )
                        if key in required_keys:
                            logging.error(
                                "Missing required column '%s'.", excel_col_name[key]
                            )
                            return False

                        logging.warning(
                            "Creating empty column name '%s'.", excel_col_name[key]
                        )
                        self._rdm_hosts_dict[key] = [""] * len(hosts["name"])
        return True

    def set_json_file(self, json_file: str | None = None, read_json_file=False):
        """Set JSON file attribute. If read_json_file is True, read content.

        Args:
            json_file (str): JSON file (source or destination)
            read_json_file (Bool): if True - read JSON file content
        """
        self.json_file = json_file

        # if self.json_file != "" and read_json_file:
        #     # self._xml_obj = SMXml(xml_file=self.xml_file, read_xml_file=True)
        #     self.parse_xml_file()

    def set_sessions_dict(self, sessions=None) -> bool:
        """Set Devolutions RDM specific fields session dictionary. If not set, initiate it.

        Args:
            sessions (dict): sessions dictionary

        Return:
            False in case of error (missing required column)
        """

        # set general fields for session dict
        if super().set_sessions_dict(sessions) is False:
            return False

        # set rdm specific fields
        excel_col_name = self._settings["excel"]["col_names_sessions"]
        keys = [
            "rdm_credential",
            "rdm_host",
            "rdm_script_before_open",
            "rdm_web_form",
            "rdm_web_login",
            "rdm_web_passwd",
        ]
        required_keys = []

        if sessions is None or len(sessions) == 0:
            for key in excel_col_name:
                if key in keys:
                    self._sessions_dict[key] = []
        else:
            for key in excel_col_name:
                if key in keys:
                    try:
                        self._sessions_dict[key] = list(map(str, sessions[key]))
                    except KeyError:
                        logging.warning(
                            "Missing column name '%s' (key: '%s').",
                            excel_col_name[key],
                            key,
                        )
                        if key in required_keys:
                            logging.error(
                                "Missing required column '%s'.", excel_col_name[key]
                            )
                            return False

                        logging.warning(
                            "Creating empty column name '%s'.", excel_col_name[key]
                        )
                        self._sessions_dict[key] = [""] * len(sessions["session"])

        return True

    # ====================
    # Prepare XML to ordered dict (From XML to Excel)
    # ====================

    ### private methods

    def __set_credentials_dict_from_xml(self, root: ET.Element):
        """Set self._credentials_dict from XML content"""

        # walk through all "key tags" and read folders and sessions
        idx = 0
        for child in root.iterfind("key"):
            # set session parameters from XML content

            self._credentials_dict["credential"].insert(idx, child.attrib["name"])

            for sub_et in child.findall("./*/[@name='Username']"):
                text = "" if sub_et.text is None else sub_et.text
                self._credentials_dict["username"].insert(idx, text)

            logging.debug(
                " {0:>3} | {1:<20} | {2:<33}".format(
                    idx + 1,
                    self._credentials_dict["credential"][idx],
                    self._credentials_dict["username"][idx],
                )
            )
            idx += 1

    def __set_sessions_dict_from_xml(self, root: ET.Element, folder):
        """Set self._sessions_dict from XML content"""

        # walk through all "key tags" and read folders and sessions
        for child in root.iterfind("key"):
            # get correct folder path
            while len(folder) > 0 and root.attrib["name"] != folder[len(folder) - 1]:
                if len(folder) > 1:
                    del folder[len(folder) - 1]
                else:
                    folder = []
                    break

            if len(child.findall("key")) > 0:
                # build folder path
                folder.append(child.attrib["name"])
            else:
                # build session
                idx = len(self._sessions_dict["folder"])
                logging.debug(
                    " {0:>3} | {1:<30} | {2:<30}".format(
                        idx + 1, "/".join(folder), child.attrib["name"]
                    )
                )

                # set session parameters from XML content
                self._sessions_dict["folder"].insert(idx, "/".join(folder))

                self._sessions_dict["session"].insert(idx, child.attrib["name"])

                for sub_et in child.findall("./*/[@name='Hostname']"):
                    text = "" if sub_et.text is None else sub_et.text
                    self._sessions_dict["hostname"].insert(idx, text)

                for sub_et in child.findall("./*/[@name='[SSH2] Port']"):
                    text = "" if sub_et.text is None else sub_et.text
                    self._sessions_dict["port"].insert(idx, text)

                for sub_et in child.findall("./*/[@name='Username']"):
                    text = "" if sub_et.text is None else sub_et.text
                    self._sessions_dict["username"].insert(idx, text)

                for sub_et in child.findall("./*/[@name='Credential Title']"):
                    text = "" if sub_et.text is None else sub_et.text
                    self._sessions_dict["credential"].insert(idx, text)

                for sub_et in child.findall("./*/[@name='Keyword Set']"):
                    text = "" if sub_et.text is None else sub_et.text
                    self._sessions_dict["keywords"].insert(idx, text)

                for sub_et in child.findall("./*/[@name='Color Scheme']"):
                    text = "" if sub_et.text is None else sub_et.text
                    self._sessions_dict["colorscheme"].insert(idx, text)

                for sub_et in child.findall("./*/[@name='Firewall Name']"):
                    text = "" if sub_et.text is None else sub_et.text
                    self._sessions_dict["firewall"].insert(idx, text)

            self.__set_sessions_dict_from_xml(child, folder)

    ### public methods

    def get_rdm_hosts_dict_count(self):
        """Return credentials dictionary size (int)."""
        return len(self._rdm_hosts_dict["name"])

    def build_dict_from_xml(self):
        """Read SecureCRT XML session file and set all dictionaries.

        Read XML file, set sessions_xml attribute and set:
            - self._sessions_dict
            - self._credentials_dict
            - self._firewalls_dict
        """

        # get XML element/content from XML file
        # xml_element = self.parse_xml_file(self.xml_file)
        # xml_element = self._xml_obj.parse_xml_file(self.xml_file)
        if self._xml_sessions == None:
            return False

        # self.set_xml_sessions(xml_element)

        self.set_sessions_dict_from_xml()
        self.set_credentials_dict_from_xml()

    def set_credentials_dict_from_xml(self) -> None | dict:
        """Read SecureCRT export (self._sessions_xml) and set self._credentials_dict.

        Returns:
            None | dict: Ordered dict (self._credentials_dict) or None.
        """

        if self._xml_sessions is None:
            return None

        base_root = self._xml_sessions
        credentials_root = base_root.find("./key[@name='Credentials']")

        if credentials_root is not None:
            folder = []
            logging.info("Importing credentials from XML file...")
            logging.debug(
                " {0:>3} | {1:<20} | {2:<20}".format(
                    "#", "credential group", "username"
                )
            )
            logging.debug(" {0:->3}-+-{1:-<20}-+-{2:-<33}".format("", "", ""))
            self.set_credentials_dict()
            self.__set_credentials_dict_from_xml(credentials_root)
            logging.debug(" {0:->3}-+-{1:-<20}-+-{2:-<33}".format("", "", ""))
            logging.info("Imported %d record(s).", self.get_credentials_dict_count())

        return self._credentials_dict

    def set_sessions_dict_from_xml(self) -> None | dict:
        """Read SecureCRT export (self._sessions_xml) and set self._sessions_dict.

        Returns:
            None | dict: Ordered dict (self._sessions_dict) or None.
        """

        if self._xml_sessions is None:
            return None

        base_root = self._xml_sessions
        sessions_root = base_root.find("./key[@name='Sessions']")

        if sessions_root is not None:
            folder = []
            logging.info("Importing sessions from XML file...")
            logging.debug(
                " {0:>3} | {1:<30} | {2:<30}".format("#", "folder path", "session name")
            )
            logging.debug(" {0:->3}-+-{1:-<30}-+-{2:-<30}".format("", "", ""))
            self.set_sessions_dict()
            self.__set_sessions_dict_from_xml(sessions_root, folder)
            logging.debug(" {0:->3}-+-{1:-<30}-+-{2:-<30}".format("", "", ""))
            logging.info("Imported %d record(s).", self.get_sessions_dict_count())

        return self._sessions_dict

    def write_excel(self, **kwargs):
        excel_file = str(kwargs.get("excel_file", self.excel_file))
        credentials_dict = kwargs.get("credentials_dict", self._credentials_dict)
        sessions_dict = kwargs.get("sessions_dict", self._sessions_dict)
        firewalls_dict = kwargs.get("firewalls_dict", {})

        self._excel_obj.write_excel_book(
            excel_file=excel_file,
            sessions_dict=sessions_dict,
            credentials_dict=credentials_dict,
            firewalls_dict=firewalls_dict,
        )

    # ====================
    # Prepare RDM structure from ordered dicts (from Excel to JSON)
    # ====================

    def normalize_string_path(self, data: str) -> str:
        """
        Normalize the path separators in the given data.
        This method replaces all forward slashes ("/") with backslashes ("\\")
        in the provided data. If the data is a dictionary, it will recursively
        normalize the paths for any keys specified in the valid_keys list.
        Args:
            data (str | dict): The data to normalize. Can be a string or a dictionary.

        Returns:
            str | dict: The normalized data with path separators replaced.
        """

        return data.replace("/", "\\")

    def normalize_dict_path(self, data: dict, valid_keys: list[str] = []) -> dict:
        """
        Normalize the path separators in the given data.
        This method replaces all forward slashes ("/") with backslashes ("\\")
        in the provided data. If the data is a dictionary, it will recursively
        normalize the paths for any keys specified in the valid_keys list.
        Args:
            data (str | dict): The data to normalize. Can be a string or a dictionary.
            valid_keys (list[str] | None): A list of keys whose values should be
                                           normalized if the data is a dictionary.
                                           Defaults to None.
        Returns:
            str | dict: The normalized data with path separators replaced.
        """

        if isinstance(data, dict):
            for key, value in data.items():
                if isinstance(value, dict):
                    data[key] = self.normalize_dict_path(value, valid_keys)
                elif key in valid_keys:
                    data[key] = self.normalize_string_path(value)

        return data

    def _build_rdm_connection_folder(self, **kwargs):
        """Set RDM connection folder (type 25).

        Check if parent exists, if not, create it (recursively).

        Args:
            folder (str): Folder path as a string
        """
        # arguments
        folder = kwargs.get("folder", "")

        # prepare folder path and add parent recursively
        # if "/" in folder:
        #     folder = folder.replace("/", "\\")
        folder = self.normalize_string_path(folder)
        folder_list = folder.split("\\")
        folder_name = folder_list[-1]
        if len(folder_list) > 1:
            self._build_rdm_connection_folder(folder="\\".join(folder_list[0:-1]))

        # build folder object
        conn_obj = {}
        conn_obj["ConnectionType"] = 25
        conn_obj["Group"] = folder
        conn_obj["Name"] = folder_name

        # check if folder dict exists
        if conn_obj in self._rdm_connection_list:
            return

        #  append folder dict to self.__rdm_connection_list
        self._rdm_connection_list.append(conn_obj)

    def _build_rdm_connection_credential(self, folder="", credential="", username=""):
        """Set RDM Credential (type 26)

        Check if credential not exists in self.__rdm_connection_list and add it.

        Args:
            folder (str, optional, default=""): folder path
            credential (str): credential name
            username (str, optional, default: ""): username
        """
        # arguments
        self._build_rdm_connection_folder(folder=folder)
        # credential = kwargs.get("credential", "")
        # username = kwargs.get("username", "")

        if credential is None or credential == "":
            logging.warning("Credential without name. Skipping.")
            return

        # object: credential
        conn_obj = {}
        conn_obj["ConnectionType"] = 26
        conn_obj["Group"] = folder
        conn_obj["Name"] = credential
        conn_obj["CredentialConnectionID"] = str(uuid.uuid4())
        conn_obj["ID"] = str(uuid.uuid4())
        conn_obj["Credentials"] = {}
        if username != "":
            conn_obj["Credentials"]["UserName"] = username

        # check duplicity
        if conn_obj in self._rdm_connection_list:
            return

        self._rdm_connection_list.append(conn_obj)

    def _build_rdm_connection_host(
        self, folder="", name="", host="", rdm_credential=""
    ):
        """
        Builds a Connection object type Host (type 53).

        Args:
            folder (str): Session path (optional)
            name (str): Session name (required)
            host (str): Host/IP (optional)
            rdm_credential (str): Path to credential (optional)

        Returns:
            None
        """

        # build Host object (type 53)

        # arguments
        self._build_rdm_connection_folder(folder=folder)

        # rdm_credential
        # rdm_vault=self._rdm_hosts_dict["rdm_vault"][idx]
        rdm_credential = rdm_credential.replace("/", "\\")

        if name == "":
            logging.warning("Host object without name. Skipping.")
            return

        # object: Host
        conn_obj = {}
        conn_obj["ConnectionType"] = 53
        conn_obj["Group"] = folder
        conn_obj["Name"] = name
        # generate unique UUID (when using in other connection types)
        conn_obj["ID"] = str(uuid.uuid4())
        # host/ip
        conn_obj["HostDetails"] = {}
        if host != "":
            conn_obj["HostDetails"]["Host"] = host
        # credential (if defined)
        if rdm_credential != "":
            conn_obj["CredentialConnectionSavedPath"] = rdm_credential
            credential_uuid = self._get_rdm_connection_uuid(rdm_credential)
            conn_obj["CredentialConnectionID"] = credential_uuid

        # check duplicity
        if conn_obj in self._rdm_connection_list:
            return

        self._rdm_connection_list.append(conn_obj)

    def _build_rdm_connection_rdp_session(
        self,
        folder="",
        session="",
        hostname="",
        port="",
        username="",
        rdm_credential="",
        alternate_shell="",
        rdm_host="",
        rdm_script_before_open="",
        **kwargs,
    ):
        """Set RDM RDP session (type 1)

        Check if RDP session not exists in self.__rdm_connection_list and add it.

        Args:
            folder (str, optional, default=""): folder path
            session (str): RDP session name
            hostname (str, optional, default: ""): hostname/IP
            port (str, optional, default: "3389"): port
            alternate_shell (str, optional, default: ""): alternate shell (command executed on connection)
        """

        # arguments
        # folder = kwargs.get("folder", "")
        self._build_rdm_connection_folder(folder=folder)
        # session = kwargs.get("session", "")
        # hostname = kwargs.get("hostname", "")
        # port = kwargs.get("port", "3389")
        # username = kwargs.get("username", "")
        # credential = kwargs.get("credential", "")
        rdm_credential = rdm_credential.replace("/", "\\")
        rdm_host = rdm_host.replace("/", "\\")
        # credential_uuid = self.get_credential()
        # alternate_shell = kwargs.get("alternate_shell", "")

        if session == "":
            logging.warning("Session without name. Skipping.")
            return

        # sesion default data
        sdd_excel = self.session_defaults.get("rdp", {}).get("excel", {})
        keys_to_normalize = ["folder", "rdm_credential", "rdm_host"]
        sdd_excel = self.normalize_dict_path(sdd_excel, keys_to_normalize)

        sdd_raw = self.session_defaults.get("rdp", {}).get("raw", {})

        # session data
        sd_excel = {
            "port": port,
            "rdm_credential": rdm_credential,
            "rdm_host": rdm_host,
            "rdm_script_before_open": rdm_script_before_open,
        }

        sd_excel = self.merge_session_data(sd_excel, sdd_excel)
        sd_excel["folder"] = folder
        sd_excel["session"] = session
        sd_excel["alternate_shell"] = alternate_shell

        conn_obj = {}
        conn_obj["ConnectionType"] = 1
        conn_obj["Group"] = folder
        conn_obj["Name"] = session

        # rdp specific
        # conn_obj["Terminal"] = {}

        # alternate shell
        if alternate_shell != "":
            conn_obj["AlternateShell"] = alternate_shell

        # rdp
        conn_obj["RDP"] = {}
        conn_obj["RDP"]["NetworkLevelAuthentication"] = "false"

        # any other
        conn_obj["AuthentificationLevel"] = 2
        conn_obj["OpenEmbedded"] = True

        # script before open
        if sd_excel["rdm_script_before_open"] != "":
            # conn_obj["AllowPasswordVariable"] = True
            conn_obj["Events"] = {}
            conn_obj["Events"]["BeforeConnectionEmbeddedPowerShellScript"] = sd_excel[
                "rdm_script_before_open"
            ]
            conn_obj["Events"]["BeforeConnectionEvent"] = 5
            conn_obj["Events"]["BeforeConnectionWaitForExit"] = True
            conn_obj["Events"]["ConnectionPause"] = 10
            conn_obj["Events"]["ConnectionUseDefaultWorkingDirectory"] = False

        # username, linked credential, private vault
        if username != "" and sd_excel["rdm_credential"] == "":
            conn_obj["RDP"]["Username"] = username
            conn_obj["PromptCredentials"] = "true"

        elif sd_excel["rdm_credential"] != "":

            if sd_excel["rdm_credential"].startswith("private:"):
                # private vault
                cr = sd_excel["rdm_credential"].lstrip("private:").strip()
                conn_obj["CredentialPrivateVaultSearchString"] = cr
            else:
                # linked credential
                conn_obj["CredentialConnectionSavedPath"] = sd_excel["rdm_credential"]
                credential_uuid = self._get_rdm_connection_uuid(
                    sd_excel["rdm_credential"]
                )
                conn_obj["CredentialConnectionID"] = credential_uuid

        # host+port, linked host
        if hostname != "" and sd_excel["rdm_host"] == "":
            # hostname, port
            conn_obj["Url"] = hostname
            if sd_excel["port"] != "":
                conn_obj["Port"] = sd_excel["port"]

        elif sd_excel["rdm_host"] != "":
            # linked host
            conn_obj["HostSourceMode"] = 1
            conn_obj["HostConnectionSavedPath"] = sd_excel["rdm_host"]
            host_uuid = self._get_rdm_connection_uuid(sd_excel["rdm_host"])
            conn_obj["HostConnectionID"] = host_uuid

        # merge raw data
        conn_obj = self.append_session_data(conn_obj, sdd_raw)

        # check duplicity
        if conn_obj in self._rdm_connection_list:
            return

        self._rdm_connection_list.append(conn_obj)

    def _build_rdm_connection_ssh_session(
        self,
        folder="",
        session="",
        hostname="",
        port="",
        username="",
        rdm_credential="",
        rdm_host="",
        rdm_script_before_open="",
        # **kwargs,
    ):
        """Set RDM SSH session (type 77)

        Check if SSH session not exists in self.__rdm_connection_list and add it.

        Args:
            folder (str, optional, default=""): folder path
            session (str): SSH session name
            hostname (str, optional, default: ""): hostname/IP
            port (str, optional, default: "22"): port
        """
        # arguments
        # folder = kwargs.get("folder", "")
        self._build_rdm_connection_folder(folder=folder)
        # session = kwargs.get("session", "")
        # hostname = kwargs.get("hostname", "")
        # port = kwargs.get("port", "")
        # username = kwargs.get("username", "")
        # credential = kwargs.get("credential", "")
        rdm_credential = rdm_credential.replace("/", "\\")
        rdm_host = rdm_host.replace("/", "\\")
        # rdm_script_before_open = rdm_script_before_open.replace("\\", "\\\\")

        if session == "":
            logging.debug("Session without name. Skipping.")
            return

        # sesion default data
        sdd_excel = self.session_defaults.get("ssh", {}).get("excel", {})
        sdd_raw = self.session_defaults.get("ssh", {}).get("raw", {})

        # session data
        sd_excel = {
            "port": port,
            "rdm_credential": rdm_credential,
            "rdm_host": rdm_host,
            "rdm_script_before_open": rdm_script_before_open,
        }

        sd_excel = self.merge_session_data(sd_excel, sdd_excel)

        port = sd_excel["port"]

        conn_obj = {}
        conn_obj["ConnectionType"] = 77
        conn_obj["Group"] = folder
        conn_obj["Name"] = session
        
        conn_obj["Terminal"] = {}
        

        # ssh specific

        # script before open
        if sd_excel["rdm_script_before_open"] != "":
            # conn_obj["AllowPasswordVariable"] = True
            conn_obj["Events"] = {}
            conn_obj["Events"]["BeforeConnectionEmbeddedPowerShellScript"] = sd_excel[
                "rdm_script_before_open"
            ]
            conn_obj["Events"]["BeforeConnectionEvent"] = 5
            conn_obj["Events"]["BeforeConnectionWaitForExit"] = True
            conn_obj["Events"]["ConnectionPause"] = 10
            conn_obj["Events"]["ConnectionUseDefaultWorkingDirectory"] = False

        # username, linked credential, private vault
        if username != "" and sd_excel["rdm_credential"] == "":
            conn_obj["Terminal"]["Username"] = username

        elif rdm_credential != "":

            if sd_excel["rdm_credential"].startswith("private:"):
                # private vault
                cr = sd_excel["rdm_credential"].lstrip("private:").strip()
                conn_obj["CredentialPrivateVaultSearchString"] = cr
            else:
                # linked credential
                conn_obj["CredentialConnectionSavedPath"] = sd_excel["rdm_credential"]
                credential_uuid = self._get_rdm_connection_uuid(
                    sd_excel["rdm_credential"]
                )
                conn_obj["CredentialConnectionID"] = credential_uuid

        # host+port, linked host
        if hostname != "" and sd_excel["rdm_host"] == "":
            conn_obj["Terminal"]["Host"] = hostname
            if sd_excel["port"] != "":
                conn_obj["Terminal"]["HostPort"] = sd_excel["port"]

        elif sd_excel["rdm_host"] != "":
            # linked host
            conn_obj["HostSourceMode"] = 1
            conn_obj["HostConnectionSavedPath"] = sd_excel["rdm_host"]
            host_uuid = self._get_rdm_connection_uuid(sd_excel["rdm_host"])
            conn_obj["HostConnectionID"] = host_uuid

        # merge raw data
        conn_obj = self.append_session_data(conn_obj, sdd_raw)

        # check duplicity
        if conn_obj in self._rdm_connection_list:
            return

        self._rdm_connection_list.append(conn_obj)

    def _build_rdm_connection_web_session(self, **kwargs):
        """Set RDM Web based session (type 32)

        Check if Web session not exists in self.__rdm_connection_list and add it.

        Args:
            folder (str, optional, default=""): folder path
            session (str): session name
            hostname (str, optional, default: ""): hostname/IP
            credential (str, optional): credential name/path
            username (str, optional): username
            web_form (str, optional): web login form id
            web_login (str, optional): web login field id
            web_passwd (str, optional): web password field id
        """
        # arguments
        folder = kwargs.get("folder", "")
        self._build_rdm_connection_folder(folder=folder)
        session = kwargs.get("session", "")
        hostname = kwargs.get("hostname", "")
        username = kwargs.get("username", "")
        credential = kwargs.get("credential", "")
        credential = credential.replace("/", "\\")
        web_form = kwargs.get("web_form", "")
        web_login = kwargs.get("web_login", "")
        web_passwd = kwargs.get("web_passwd", "")

        if session == "":
            logging.warning("Session without name. Skipping.")
            return

        conn_obj = dict()
        conn_obj["ConnectionType"] = 32
        conn_obj["Group"] = folder
        conn_obj["Name"] = session
        conn_obj["OpenEmbedded"] = True
        conn_obj["DataEntry"] = {}
        conn_obj["DataEntry"]["Url"] = hostname
        # conn_obj["DataEntry"]["BrowserExtensionLinkerCompareTypeWeb"] = 7
        conn_obj["DataEntry"]["ConnectionTypeInfos"] = [{"DataEntryConnectionType": 11}]
        conn_obj["DataEntry"]["WebBrowserApplication"] = 3  # chrome=3
        conn_obj["DataEntry"]["WebBrowserExtensionMode"] = 1

        # webform autofill
        conn_obj["DataEntry"]["WebFormIdHtmlElementName"] = web_form
        conn_obj["DataEntry"]["WebUsernameHtmlElementName"] = web_login
        conn_obj["DataEntry"]["WebPasswordHtmlElementName"] = web_passwd
        conn_obj["DataEntry"]["WebSubmitHtmlElementName"] = "[ENTER]"

        # username
        if username != "" and credential == "":
            conn_obj["DataEntry"]["WebUserName"] = username

        # credential
        if credential != "":
            # conn_obj["DataEntry"]["CredentialConnectionSavedPath"] = credential
            credential_uuid = self._get_rdm_connection_uuid(credential)
            conn_obj["DataEntry"]["CredentialConnectionID"] = credential_uuid

        # check duplicity
        if conn_obj in self._rdm_connection_list:
            return

        self._rdm_connection_list.append(conn_obj)

    def _get_rdm_connection_uuid(self, connection_path):
        """Return UUID of the connection based on full path.

        Args:
            connection_path (str):Connection name including folder path.

        Return:
            UUID of the connection record
        """
        connection_path = connection_path.replace("/", "\\")
        conn_path_list = connection_path.split("\\")
        for conn_obj in self._rdm_connection_list:
            if (
                conn_obj["Group"].rstrip("\\") == "\\".join(conn_path_list[0:-1])
                and conn_obj["Name"] == conn_path_list[-1]
            ):
                if conn_obj["ConnectionType"] == 26:
                    # credential obj
                    return conn_obj["ID"]

                if conn_obj["ConnectionType"] == 53:
                    # host obj
                    return conn_obj["ID"]

    def __sessions_dict_to_json_connections(self):
        """Set __rdm_connection_list from _sessions_dict"""

        # get folder path and session in a loop
        for idx, session_row in enumerate(self._sessions_dict["session"]):
            # get folders structure
            folder_path = self._sessions_dict["folder"][idx]
            folder_path = folder_path.replace("/", "\\")
            session_type = self._sessions_dict["type"][idx]

            # ssh session (#77)
            if session_type == "ssh":
                self._build_rdm_connection_ssh_session(
                    folder=folder_path,
                    session=self._sessions_dict["session"][idx],
                    hostname=self._sessions_dict["hostname"][idx],
                    port=self._sessions_dict["port"][idx],
                    username=self._sessions_dict["username"][idx],
                    rdm_credential=self._sessions_dict["rdm_credential"][idx],
                    rdm_host=self._sessions_dict["rdm_host"][idx],
                    rdm_script_before_open=self._sessions_dict[
                        "rdm_script_before_open"
                    ][idx],
                )

            # rdp session (#1)
            if session_type == "rdp":
                self._build_rdm_connection_rdp_session(
                    folder=folder_path,
                    session=self._sessions_dict["session"][idx],
                    hostname=self._sessions_dict["hostname"][idx],
                    port=self._sessions_dict["port"][idx],
                    username=self._sessions_dict["username"][idx],
                    rdm_credential=self._sessions_dict["rdm_credential"][idx],
                    alternate_shell=self._sessions_dict["rdp_alternate"][idx],
                    rdm_script_before_open=self._sessions_dict[
                        "rdm_script_before_open"
                    ][idx],
                )

            # web session (#32)
            if session_type == "web":
                self._build_rdm_connection_web_session(
                    folder=folder_path,
                    session=self._sessions_dict["session"][idx],
                    hostname=self._sessions_dict["hostname"][idx],
                    port=self._sessions_dict["port"][idx],
                    username=self._sessions_dict["username"][idx],
                    credential=self._sessions_dict["rdm_credential"][idx],
                    web_form=self._sessions_dict["rdm_web_form"][idx],
                    web_login=self._sessions_dict["rdm_web_login"][idx],
                    web_passwd=self._sessions_dict["rdm_web_passwd"][idx],
                )

    def _credentials_dict_to_json_connections(self):
        """Set __rdm_connection_list from _credentials_dict"""

        # get credentials/credentials in a loop
        for idx, vault_row in enumerate(self._credentials_dict["credential"]):
            # get folders structure
            folder_path = self._credentials_dict["folder"][idx]
            folder_path = folder_path.replace("/", "\\")

            # credential (#26)
            self._build_rdm_connection_credential(
                folder=folder_path,
                credential=self._credentials_dict["credential"][idx],
                username=self._credentials_dict["username"][idx],
            )

        return self._rdm_connection_list

    def _rdm_hosts_dict_to_json_connections(self):
        """Set __rdm_connection_list from _rdm_hosts_dict"""

        # get credentials/credentials in a loop
        for idx, host_row in enumerate(self._rdm_hosts_dict["name"]):
            # get folders structure
            folder_path = self._rdm_hosts_dict["folder"][idx]
            folder_path = folder_path.replace("/", "\\")

            # host (#53)
            self._build_rdm_connection_host(
                folder=folder_path,
                name=self._rdm_hosts_dict["name"][idx],
                host=self._rdm_hosts_dict["host"][idx],
                rdm_credential=self._rdm_hosts_dict["rdm_credential"][idx],
            )

        return self._rdm_connection_list

    ### public methods

    def build_json_from_dict(self):
        """Build DevolutionsRDM JSON content.
        Method set's attribute self._json_sessions.

        Returns:
            (dict()): JSON content of sessions for importing to Devolutions RDM.
        """

        self._credentials_dict_to_json_connections()
        self._rdm_hosts_dict_to_json_connections()
        self.__sessions_dict_to_json_connections()
        self._json_sessions = dict()
        self._json_sessions["Connections"] = self._rdm_connection_list

        return self._json_sessions
