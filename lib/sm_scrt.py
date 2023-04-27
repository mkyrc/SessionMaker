"""SecureCRT session generator"""
import logging
import xml.etree.ElementTree as ET

from .sm_class import SessionMaker
from .sm_xml import SMXml

# ========================================
# Class SMSecureCrt
# ========================================
class SMSecureCrt(SessionMaker):
    """SessionMaker - SecureCRT sessions generator class"""

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
        self.set_credentials_dict(kwargs.get("credentials", None))

        # firewall groups dict
        self._firewalls_dict = dict()
        self.set_firewalls_dict(kwargs.get("firewalls", None))

    # ========================================
    # Private methods
    # ========================================

    # ========================================
    # Public methods
    # ========================================
    
    def excel_read_sheet_credentials(self, sheet_name: str) -> dict | list | bool:
        """Read excel sheet 'scrt-credentials' and return content as dict/array.

        Args:
            sheet_name (str): Sheet's name

        Returns:
            ordered dict: Column/Row-based dictionary (when get=['column', 'row']
            False: In case of error
        """
        credentials_dict = self.excel_read_sheet(sheet_name, "column")
        credentials_dict = self.col_name_normalize(
            credentials_dict, self._settings["excel"]["col_names_scrt_credentials"]
        )
        self.set_credentials_dict(credentials_dict)

        return self._credentials_dict    
    
    def excel_read_sheet_firewalls(self, sheet_name: str) -> dict | list | bool:
        """Read excel sheet 'scrt-firewalls' and return content as dict/array.

        Args:
            sheet_name (str): Sheet's name

        Returns:
            ordered dict: Column/Row-based dictionary (when get=['column', 'row']
            False: In case of error
        """
        firewalls_dict = self.excel_read_sheet(sheet_name, "column")
        firewalls_dict = self.col_name_normalize(
            firewalls_dict, self._settings["excel"]["col_names_scrt_firewalls"]
        )
        self.set_firewalls_dict(firewalls_dict)

        return self._credentials_dict      

    def get_firewalls_dict(self):
        """Return firewall groups dictionary (ordered dict)."""
        return self._firewalls_dict

    def get_firewalls_dict_count(self):
        """Return firewall groups dictionary size (int)."""
        return len(self._firewalls_dict["firewall"])

    def set_credentials_dict(self, credentials=None):
        """Set (SecureCRT specific fields) credentials dictionary. If not set, create empty.

        Args:
            credentials (dict): Ordered dict.
        """
        col_name = self._settings["excel"]["col_names_scrt_credentials"]
        # keys = ["scrt-credential", "scrt-username"]

        if credentials is None:
            for key in col_name:
                self._credentials_dict[key] = []
        else:
            for key in col_name:
                # self._credentials_dict[key] = list(map(str, credentials[col_name[key]]))
                self._credentials_dict[key] = list(map(str, credentials[key]))

    def set_firewalls_dict(self, firewalls=None):
        """Set (SecureCRT specific fields) credentials dictionary. If not set, initiate it.

        Args:
            firewalls (dict): Ordered dict.
        """
        col_name = self._settings["excel"]["col_names_scrt_firewalls"]
        # keys = ["firewall", "address", "port", "username"]

        if firewalls is None:
            for key in col_name:
                self._firewalls_dict[key] = []
        else:
            for key in col_name:
                # self._firewalls_dict[key] = list(map(str, firewalls[col_name[key]]))
                self._firewalls_dict[key] = list(map(str, firewalls[key]))

    def set_sessions_dict(self, sessions=None):
        """Set (SecureCRT specific fields) session dictionary. If not set, initiate it.

        Args:
            sessions (dict): sessions dictionary
        """
        super().set_sessions_dict(sessions)

        excel_col_name = self._settings["excel"]["col_names_sessions"]
        keys = ["scrt-credential", "scrt-colorscheme", "scrt-keywords", "scrt-firewall"]

        if sessions is None or len(sessions) == 0:
            for key in excel_col_name:
                if key in keys:
                    self._sessions_dict[key] = []
        else:
            for key in excel_col_name:
                if key in keys:
                    # self._sessions_dict[key] = list(map(str, sessions[col_name[key]]))
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

    def __set_firewalls_dict_from_xml(self, root: ET.Element):
        """Set self._firewalls_dict from XML content"""

        # walk through all "key tags" and read folders and sessions
        idx = 0
        for child in root.iterfind("key"):

            # set firewall parameters from XML content

            self._firewalls_dict["firewall"].insert(idx, child.attrib["name"])

            for sub_et in child.findall("./*/[@name='Firewall Address']"):
                text = "" if sub_et.text is None else sub_et.text
                self._firewalls_dict["address"].insert(idx, text)

            for sub_et in child.findall("./*/[@name='Firewall Port']"):
                text = "" if sub_et.text is None else sub_et.text
                self._firewalls_dict["port"].insert(idx, text)

            for sub_et in child.findall("./*/[@name='Firewall User']"):
                text = "" if sub_et.text is None else sub_et.text
                self._firewalls_dict["username"].insert(idx, text)

            logging.debug(
                " {0:>3} | {1:<20} | {2:<20} | {3:<10}".format(
                    idx + 1,
                    self._firewalls_dict["firewall"][idx],
                    self._firewalls_dict["address"][idx],
                    self._firewalls_dict["port"][idx],
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
                    self._sessions_dict["scrt-credential"].insert(idx, text)

                for sub_et in child.findall("./*/[@name='Keyword Set']"):
                    text = "" if sub_et.text is None else sub_et.text
                    self._sessions_dict["scrt-keywords"].insert(idx, text)

                for sub_et in child.findall("./*/[@name='Color Scheme']"):
                    text = "" if sub_et.text is None else sub_et.text
                    self._sessions_dict["scrt-colorscheme"].insert(idx, text)

                for sub_et in child.findall("./*/[@name='Firewall Name']"):
                    text = "" if sub_et.text is None else sub_et.text
                    self._sessions_dict["scrt-firewall"].insert(idx, text)

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
        self.set_firewalls_dict_from_xml()

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

    def set_firewalls_dict_from_xml(self) -> None | dict:
        """Read SecureCRT export (self._sessions_xml) and set self._firewalls_dict.

        Returns:
            None | dict: Ordered dict (self._firewalls_dict) or None.
        """

        if self._xml_sessions is None:
            return None

        base_root = self._xml_sessions
        firewalls_root = base_root.find("./key[@name='Firewalls']")

        if firewalls_root is not None:
            folder = []
            logging.info("Importing firewalls from XML file...")
            logging.debug(
                " {0:>3} | {1:<20} | {2:<20} | {3:<10}".format(
                    "#", "firewall group", "address", "port"
                )
            )
            logging.debug(
                " {0:->3}-+-{1:-<20}-+-{2:-<20}-+-{3:-<10}".format("", "", "", "")
            )
            self.set_firewalls_dict()
            self.__set_firewalls_dict_from_xml(firewalls_root)
            logging.debug(
                " {0:->3}-+-{1:-<20}-+-{2:-<20}-+-{3:-<10}".format("", "", "", "")
            )
            logging.info("Imported %d record(s).", self.get_firewalls_dict_count())

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
        """Write credentials, sessions and firewalls to Excel file

        Args:
            excel_file (str, default: self.excel_file): Excel file to write content
            credentials_dict (str, default: self._credentials_dict): Credentials dict to write content
            sessions_dict (str, default: self._sessions_dict): Sessions dict to write content
            firewalls_dict (str, default: self._firewalls_dict): Firewalls dict to write content
        """

        excel_file = str(kwargs.get("excel_file", self.excel_file))
        credentials_dict = kwargs.get("credentials_dict", self._credentials_dict)
        sessions_dict = kwargs.get("sessions_dict", self._sessions_dict)
        firewalls_dict = kwargs.get("firewalls_dict", self._firewalls_dict)

        self._excel_obj.write_excel_book(
            excel_file=excel_file,
            sessions_dict=sessions_dict,
            scrt_credentials_dict=credentials_dict,
            scrt_firewalls_dict=firewalls_dict,            
        )

    # ====================
    # Prepare XML from ordered dicts (from Excel to XML)
    # ====================

    ### private methods

    def __credentials_dict_to_xml(self) -> ET.Element:
        """Read self._credentials_dict and return credentials hierarchy as XML object.

        Returns:
            (ET.Element): XML object for credentials
        """
        ret_xml = ET.Element("CREDENTIALS")
        if not self._credentials_dict:
            return ret_xml

        for idx, credential_row in enumerate(self._credentials_dict["credential"]):
            # build credentials data in XML format
            credential_xml = self.__xml_build_credential(
                xml_tpl_credential=self.__xml_tpl_get_credential(),
                credential=self._credentials_dict["credential"][idx],
                username=self._credentials_dict["username"][idx],
            )

            # return session_xml only (no folder path defined)
            if isinstance(credential_xml, ET.Element):
                ret_xml.append(credential_xml)

        return ret_xml

    def __firewalls_dict_to_xml(self) -> ET.Element:
        """Read self._firewalls_dict and return firewalls hierarchy as XML object.

        Returns:
            (ET.Element): XML object for firewalls
        """
        ret_xml = ET.Element("FIREWALLS")
        if not self._firewalls_dict:
            return ret_xml

        for idx, firewall_row in enumerate(self._firewalls_dict["firewall"]):
            # build firewalls data in XML format
            firewall_xml = self.__xml_build_firewall(
                xml_tpl_firewall=self.__xml_tpl_get_firewall(),
                firewall=self._firewalls_dict["firewall"][idx],
                address=self._firewalls_dict["address"][idx],
                port=self._firewalls_dict["port"][idx],
                username=self._firewalls_dict["username"][idx],
            )

            # return session_xml only (no folder path defined)
            if isinstance(firewall_xml, ET.Element):
                ret_xml.append(firewall_xml)

        return ret_xml

    def __sessions_dict_to_xml(self) -> ET.Element:
        """Read self._sessions_dict and return firewalls hierarchy as XML object.

        Returns:
            (ET.Element): XML object for sessions
        """
        # root object for return
        ret_xml = ET.Element("SESSION")

        # get folder path and session in a loop
        for idx, session_row in enumerate(self._sessions_dict["session"]):
            # get folders structure
            folder_path = self._sessions_dict["folder"][idx]
            if folder_path == "":
                # no folder path
                folders_xml = None
            else:
                folders_xml = self.__xml_build_folder_path(folder_path.split("/"))

            # get session data in XML format
            session_xml = self.__xml_get_session(
                # template
                xml_tpl_session=self.__xml_tpl_get_session(),
                # values
                session=self._sessions_dict["session"][idx],
                hostname=self._sessions_dict["hostname"][idx],
                port=self._sessions_dict["port"][idx],
                username=self._sessions_dict["username"][idx],
                credential=self._sessions_dict["scrt-credential"][idx],
                colorscheme=self._sessions_dict["scrt-colorscheme"][idx],
                keywords=self._sessions_dict["scrt-keywords"][idx],
                firewall=self._sessions_dict["scrt-firewall"][idx],
            )

            # add session XML to folder path XML
            if isinstance(folders_xml, ET.Element):
                # return folders_xml (with session_xml)
                xml_filter = "." + "/key" * (len(folder_path.split("/")) - 1)
                last_folder = folders_xml.find(xml_filter)
                if isinstance(last_folder, ET.Element):
                    last_folder.append(session_xml)
                else:
                    folders_xml.append(session_xml)
                ret_xml.append(folders_xml)
            else:
                # return session_xml only (no folder prefix/path defined)
                if isinstance(session_xml, ET.Element):
                    ret_xml.append(session_xml)

        # normalize folder paths structure (merge duplicate folder paths)
        ret_xml = self.__xml_merge_sessions_folder_path(ret_xml)

        return ret_xml

    def __xml_build_credential(self, **kwargs) -> ET.Element:
        """Read XML credential template and set XML object based on arguments.

        Args:
            xml_tpl_credential (str): XML template file for credential
            credential (str): Credential group name
            username (str): Firewall user name

        Returns:
            (ET.Element): Credential group XML object
        """
        # set XML root Element
        # from template XML file if exists, else create new
        credential_root = kwargs.get("xml_tpl_credential", ET.Element("key"))

        # credential parameters
        par_username = kwargs.get("username", "")
        par_credential = kwargs.get("credential", "")

        ### create XML object ###
        # modify credential group name
        credential_root.set("name", par_credential)

        # set credential parameters (find and modify/replace)
        if par_username:
            for sub_et in credential_root.findall("./*/[@name='Username']"):
                sub_et.text = par_username

        return credential_root

    def __xml_build_folder_path(self, key_list: list) -> ET.Element:
        """Return SecureCRT folder structure (recursively) as ET.Element.

        Args:
            key_list (list): Folder path as a list

        Returns:
            (ET.Element): Folder root object
        """

        if len(key_list) > 1:
            folder_root = ET.Element("key", name=key_list[0])
            child = self.__xml_build_folder_path(key_list[1:])
            folder_root.append(child)
        else:
            folder_root = ET.Element("key", name=key_list[0])

        return folder_root

    def __xml_build_firewall(self, **kwargs) -> ET.Element:
        """Read XML firewall template and set XML object based on arguments.

        Args:
            xml_tpl_firewall (str): XML template file for firewall
            firewall (str): Firewall group name
            address (str): Firewall address
            port (str): Firewall port
            username (str): Firewall user name

        Returns:
            (ET.Element): Firewall group XML object
        """
        # set XML root Element
        # from template XML file if exists, else create new
        firewall_root = kwargs.get("xml_tpl_firewall", ET.Element("key"))

        # credential parameters
        par_firewall = kwargs.get("firewall", "")
        par_address = kwargs.get("address", "")
        par_port = kwargs.get("port", "")
        par_username = kwargs.get("username", "")

        ### create XML object ###
        # modify firewall group name
        firewall_root.set("name", par_firewall)

        # set address parameters (find and modify/replace)
        if par_address:
            for sub_et in firewall_root.findall("./*/[@name='Firewall Address']"):
                sub_et.text = par_address

        # set port parameters (find and modify/replace)
        if par_port:
            for sub_et in firewall_root.findall("./*/[@name='Firewall Port']"):
                sub_et.text = par_port

        # set username parameters (find and modify/replace)
        if par_username:
            for sub_et in firewall_root.findall("./*/[@name='Firewall User']"):
                sub_et.text = par_username

        return firewall_root

    def __xml_get_session(self, **kwargs) -> ET.Element:
        """Read XML session template and set XML object based on arguments.

        Args:
            xml_tpl_session (str): XML template file for session
            session (str): Session name
            address (str): Session address
            port (str, default: "22"): Session port
            username (str): Session user name
            credential (str): Credential group
            keyword (str): Keyword name
            colorscheme (str): Colorscheme name
            firewall (str): Firewall group name (or path to other session)

        Returns:
            (ET.Element): Session XML object
        """
        # set XML root Element
        # use XML file template if exists, else create new
        session_root = kwargs.get("xml_tpl_session", ET.Element("key"))

        # session parameters
        par_session = kwargs.get("session", "default-session")
        par_hostname = kwargs.get("hostname", "")
        par_port = kwargs.get("port", "22")
        par_username = kwargs.get("username", "")
        par_credential = kwargs.get("credential", "")
        par_keyword = kwargs.get("keywords", "")
        par_colorscheme = kwargs.get("colorscheme", "")
        par_firewall = kwargs.get("firewall", "")

        # when firewall contains path to session,
        # set firewall name to "Session:<session_path>"
        if "/" in par_firewall:
            if "Session:" not in par_firewall:
                par_firewall = "Session:" + par_firewall

        ### create XML object ###
        # modify session name
        session_root.set("name", par_session)

        # set session parameters (find and modify/replace)
        if par_hostname:
            for sub_et in session_root.findall("./*/[@name='Hostname']"):
                sub_et.text = par_hostname

        if par_port:
            for port in session_root.findall("./*/[@name='[SSH2] Port']"):
                port.text = str(par_port)

        if par_username:
            for sub_et in session_root.findall("./*/[@name='Username']"):
                sub_et.text = par_username

        if par_credential:
            for sub_et in session_root.findall("./*/[@name='Credential Title']"):
                sub_et.text = par_credential

        if par_keyword:
            for sub_et in session_root.findall("./*/[@name='Keyword Set']"):
                sub_et.text = par_keyword

        if par_colorscheme:
            for sub_et in session_root.findall("./*/[@name='Color Scheme']"):
                sub_et.text = par_colorscheme

        if par_firewall:
            for sub_et in session_root.findall("./*/[@name='Firewall Name']"):
                sub_et.text = par_firewall

        return session_root

    def __xml_merge_sessions_folder_path(self, parent_element) -> ET.Element:
        """Normalize folder path structure (remove duplicities).

        Read sessions folder/path in a loop and merge the same paths to one sub folder.

        Args:
            parent_element (ET.Element): Parent XML element

        Return:
            (ET.Element)
        """
        # get all childrens
        child_list = parent_element.findall("./key")
        new_child = []

        for child in child_list:
            if not child.attrib in new_child:
                new_child.append(child.attrib)
            else:
                # find child element instance to extend with current child
                find_string = "./key[@name='" + child.attrib["name"] + "']"
                child_element = parent_element.find(find_string)

                # if child is folder path (has 'key(s)')
                # add it to new path and remove it from old place
                if child.find("./key") is None:
                    # child has no 'key type' childrens
                    parent_element.remove(child)
                else:
                    # child has 'key' childrens
                    child_element.extend(child)
                    parent_element.remove(child)

                # iterate new child_element
                self.__xml_merge_sessions_folder_path(child_element)

        return parent_element

    def __xml_tpl_get_root(self):
        """Return root template Element object"""
        xml_obj = SMXml()
        return xml_obj.parse_xml_file(self._settings["scrt"]["template"]["root"])

    def __xml_tpl_get_credential(self):
        """Return credential template Element object"""
        xml_obj = SMXml()
        return xml_obj.parse_xml_file(self._settings["scrt"]["template"]["credential"])

    def __xml_tpl_get_firewall(self):
        """Return firewall template Element object"""
        xml_obj = SMXml()
        return xml_obj.parse_xml_file(self._settings["scrt"]["template"]["firewall"])

    def __xml_tpl_get_session(self):
        """Return session template Element object"""
        xml_obj = SMXml()
        return xml_obj.parse_xml_file(self._settings["scrt"]["template"]["session"])

    ### public methods

    def build_xml_from_dict(self):
        """Build SecureCRT XML content from template (root+sessions+credentials+firewalls).
        Method set's attribute self._sessions_xml.

        Returns:
            (ET.Element): XML content of sessions for importing to SecureCRT.
        """
        # read default base(root) XML file structure
        base_root = self.__xml_tpl_get_root()

        # read all sessions as XML structures
        sessions_root = self.__sessions_dict_to_xml()

        # read all credentials as XML structures
        credentials_root = self.__credentials_dict_to_xml()

        # read all credentials as XML structures
        firewalls_root = self.__firewalls_dict_to_xml()

        if base_root:
            # add sessions to base xml on correct place (key.name=Sessions)
            sub_sessions = base_root.find("./key[@name='Sessions']")
            if isinstance(sub_sessions, ET.Element):
                for session in sessions_root.findall("./"):
                    sub_sessions.append(session)

            # add credentials to base xml on correct place (key.name=Credentials)
            sub_credentials = base_root.find("./key[@name='Credentials']")
            if isinstance(sub_credentials, ET.Element):
                for credential in credentials_root.findall("./"):
                    sub_credentials.append(credential)

            # add firewalls to base xml on correct place (key.name=Firewalls)
            sub_firewalls = base_root.find("./key[@name='Firewalls']")
            if isinstance(sub_firewalls, ET.Element):
                for firewall in firewalls_root.findall("./"):
                    sub_firewalls.append(firewall)

        self._xml_sessions = base_root

        return self._xml_sessions
