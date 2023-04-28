"""Devolutions Remote Desktop Manager (RDM) session generator"""

import logging
import xml.etree.ElementTree as ET
import uuid
from .sm_class import SessionMaker
from .sm_xml import SMXml


# ========================================
# Class SMDevolutionsRDM
# ========================================
class SMDevolutionsRdm(SessionMaker):
    """SessionMaker - Devolutions RDM sessions generator class"""

    def __init__(self, **kwargs):
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
        super().__init__(**kwargs)

        # credential groups dict
        # self._credentials_dict = dict()
        self.set_credentials_dict(kwargs.get("credentails", None))

        # JSON file
        self.json_file = ""
        self._json_sessions = dict()
        self.__rdm_connection_list = []
        self.set_json_file(
            kwargs.get("json_file", ""), kwargs.get("read_json_file", False)
        )

    # ========================================
    # Private methods
    # ========================================

    # ========================================
    # Public methods
    # ========================================

    def excel_read_sheet_credentials(self, sheet_name: str) -> dict | list | bool:
        """Read excel sheet 'rdm-credentials' and return content as dict/array.

        Args:
            sheet_name (str): Sheet's name

        Returns:
            ordered dict: Column/Row-based dictionary (when get=['column', 'row']
            False: In case of error
        """
        credentials_dict = self.excel_read_sheet(sheet_name, "column")
        credentials_dict = self.col_name_normalize(
            credentials_dict, self._settings["excel"]["col_names_rdm_credentials"]
        )
        self.set_credentials_dict(credentials_dict)

        return self._credentials_dict

    def set_credentials_dict(self, credentials=None):
        """Set credentials dictionary. If not set, create empty.

        Args:
            credentials (dict): Ordered dict of RDM credentials.
        """
        col_name = self._settings["excel"]["col_names_rdm_credentials"]
        # keys = ["folder", "credential", "username"]

        if credentials is None:
            for key in col_name:
                self._credentials_dict[key] = []
        else:
            for key in col_name:
                # self._credentials_dict[key] = list(map(str, credentials[col_name[key]]))
                self._credentials_dict[key] = list(map(str, credentials[key]))

    def set_json_file(self, json_file: str, read_json_file=False):
        """Set JSON file attribute. If read_json_file is True, read content.

        Args:
            json_file (str): JSON file (source or destination)
            read_json_file (Bool): if True - read JSON file content
        """
        self.json_file = json_file

        # if self.json_file != "" and read_json_file:
        #     # self._xml_obj = SMXml(xml_file=self.xml_file, read_xml_file=True)
        #     self.parse_xml_file()

    def set_sessions_dict(self, sessions=None):
        """Set (SecureCRT specific fields) session dictionary. If not set, initiate it.

        Args:
            sessions (dict): sessions dictionary
        """
        super().set_sessions_dict(sessions)

        excel_col_name = self._settings["excel"]["col_names_sessions"]
        # keys = ["rdm-credential", "credential", "colorscheme", "keywords", "firewall"]
        keys = ["rdm-credential"]

        if sessions is None or len(sessions) == 0:
            for key in excel_col_name:
                if key in keys:
                    self._sessions_dict[key] = []
        else:
            for key in excel_col_name:
                if key in keys:
                    # self._sessions_dict[key] = list(map(str, sessions[excel_col_name[key]]))
                    self._sessions_dict[key] = list(map(str, sessions[key]))

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
    # Prepare XML from ordered dicts (from Excel to XML)
    # ====================

    ### private methods

    def __build_rdm_connection_folder(self, **kwargs):
        """Set RDM connection folder (type 25).

        Check if parent exists, if not, create it (recursively).

        Args:
            folder (str): Folder path as a string
        """
        # arguments
        folder = kwargs.get("folder", "")

        # prepare folder path and add parent recursively
        if "/" in folder:
            folder = folder.replace("/", "\\")
        folder_list = folder.split("\\")
        folder_name = folder_list[-1]
        if len(folder_list) > 1:
            self.__build_rdm_connection_folder(folder="\\".join(folder_list[0:-1]))

        # build folder object
        conn_obj = dict()
        conn_obj["ConnectionType"] = 25
        conn_obj["Group"] = folder
        conn_obj["Name"] = folder_name

        # check if folder dict exists
        if conn_obj in self.__rdm_connection_list:
            return

        #  append folder dict to self.__rdm_connection_list
        self.__rdm_connection_list.append(conn_obj)

    def __build_rdm_connection_ssh_session(self, **kwargs):
        """Set RDM SSH session (type 77)

        Check if SSH session not exists in self.__rdm_connection_list and add it.

        Args:
            folder (str, optional, default=""): folder path
            session (str): SSH session name
            hostname (str, optional, default: ""): hostname/IP
            port (str, optional, default: "22"): port
        """
        # arguments
        folder = kwargs.get("folder", "")
        self.__build_rdm_connection_folder(folder=folder)
        session = kwargs.get("session", "")
        hostname = kwargs.get("hostname", "")
        port = kwargs.get("port", "")
        username = kwargs.get("username", "")
        credential = kwargs.get("credential", "")
        credential = credential.replace("/", "\\")
        # credential_uuid = self.get_credential()

        if session == "":
            logging.warning("Session without name. Skipping.")
            return

        conn_obj = dict()
        conn_obj["ConnectionType"] = 77
        conn_obj["Group"] = folder
        conn_obj["Name"] = session
        conn_obj["Terminal"] = dict()
        if port == "":
            conn_obj["Terminal"]["Host"] = hostname
        else:
            conn_obj["Terminal"]["Host"] = hostname
            conn_obj["Terminal"]["HostPort"] = port

        # username
        if username != "":
            conn_obj["Terminal"]["Username"] = username

        # credential
        if username == "" and credential != "":
            conn_obj["CredentialConnectionSavedPath"] = credential
            credential_uuid = self.__get_rdm_connection_uuid(credential)
            conn_obj["CredentialConnectionID"] = credential_uuid

        # check duplicity
        if conn_obj in self.__rdm_connection_list:
            return

        self.__rdm_connection_list.append(conn_obj)

    def __get_rdm_connection_uuid(self, connection_path):
        """Return UUID of the connection based on full path.

        Args:
            connection_path (str):Connection name including folder path.

        Return:
            UUID of the connection record
        """
        conn_path_list = connection_path.split("\\")
        for conn_obj in self.__rdm_connection_list:
            if (
                conn_obj["Group"] == "\\".join(conn_path_list[0:-1])
                and conn_obj["Name"] == conn_path_list[-1]
            ):
                if conn_obj["ConnectionType"] == 26:
                    # return conn_obj["CredentialConnectionID"]
                    return conn_obj["ID"]

    def __build_rdm_connection_credential(self, **kwargs):
        """Set RDM Credential (type 26)

        Check if credential not exists in self.__rdm_connection_list and add it.

        Args:
            folder (str, optional, default=""): folder path
            credential (str): credential name
            username (str, optional, default: ""): username
        """
        # arguments
        folder = kwargs.get("folder", "")
        self.__build_rdm_connection_folder(folder=folder)
        credential = kwargs.get("credential", "")
        username = kwargs.get("username", "")
        # id = str(uuid.uuid4())

        if credential == "":
            logging.warning("Credential without name. Skipping.")
            return

        conn_obj = dict()
        conn_obj["ConnectionType"] = 26
        conn_obj["Group"] = folder
        conn_obj["Name"] = credential
        conn_obj["CredentialConnectionID"] = str(uuid.uuid4())
        conn_obj["ID"] = str(uuid.uuid4())
        conn_obj["Credentials"] = dict()
        if username != "":
            conn_obj["Credentials"]["UserName"] = username

        # check duplicity
        if conn_obj in self.__rdm_connection_list:
            return

        self.__rdm_connection_list.append(conn_obj)

    def __sessions_dict_to_json_connections(self):
        """Set __rdm_connection_list from _sessions_dict"""

        # get folder path and session in a loop
        for idx, session_row in enumerate(self._sessions_dict["session"]):
            # get folders structure
            folder_path = self._sessions_dict["folder"][idx]
            folder_path = folder_path.replace("/", "\\")

            # ssh session (#77)
            self.__build_rdm_connection_ssh_session(
                folder=folder_path,
                session=self._sessions_dict["session"][idx],
                hostname=self._sessions_dict["hostname"][idx],
                port=self._sessions_dict["port"][idx],
                username=self._sessions_dict["username"][idx],
                credential=self._sessions_dict["rdm-credential"][idx],
            )

    def __credentials_dict_to_json_connections(self):
        """Set __rdm_connection_list from _credentials_dict"""

        # get credentials/credentials in a loop
        for idx, vault_row in enumerate(self._credentials_dict["credential"]):
            # get folders structure
            folder_path = self._credentials_dict["folder"][idx]
            folder_path = folder_path.replace("/", "\\")

            # credential (#26)
            self.__build_rdm_connection_credential(
                folder=folder_path,
                credential=self._credentials_dict["credential"][idx],
                username=self._credentials_dict["username"][idx],
            )

        return self.__rdm_connection_list

    ### public methods

    def build_json_from_dict(self):
        """Build DevolutionsRDM JSON content.
        Method set's attribute self._json_sessions.

        Returns:
            (dict()): JSON content of sessions for importing to Devolutions RDM.
        """

        self.__credentials_dict_to_json_connections()
        self.__sessions_dict_to_json_connections()
        self._json_sessions = dict()
        self._json_sessions["Connections"] = self.__rdm_connection_list

        return self._json_sessions
