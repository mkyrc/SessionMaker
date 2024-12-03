"""Session Maker master class"""

import logging

# import os.path
# from pathlib import Path
import xml.etree.ElementTree as ET
import xmltodict

from lib.io import SMExcel, SMXml, SMJson


# logging.basicConfig(format="%(levelname)s: %(message)s", level=logging.INFO)

# logger = logging.getLogger(__name__)


# ========================================
# Class SessionMaker
# ========================================
class SessionMaker:
    """SessionMaker SecureCRT sessions generator class

    Private variables:
        _settings
    """

    def __init__(
        self,
        settings: dict | None = None,
        excel_file: str | None = None,
        read_excel_file=False,
        xml_file="",
        read_xml_file=False,
        # json_file="",
        # read_json_file=False,
        **kwargs,
    ):
        """Initial class method.

        Args:
            settings (dict): Configuration settings content (default: config.yaml)
            excel_file (str): Excel file path (source or destination). Default: "".
            read_excel_file (bool):  Read excel file? Default: False.
            xml_file (str): XML file path (source or destination). Default: ""
            read_xml_file (bool):  Read xml file? Default: False.
            json_file (str): XML file path (source or destination). Default: ""
            read_json_file (bool):  Read xml file? Default: False.
            # sessions (dict): essions dictionary (default: create empty).
        """

        # settings (dict, config.yaml content)
        if settings is None:
            self._settings = {}
        else:
            self.set_settings(settings)

        # excel_file path (str). if set, initiate excel_obj
        self.excel_file = ""
        self._excel_obj = SMExcel(settings=self._settings, read_excel_file=False)
        self.set_excel_file(excel_file, read_excel_file)

        # sessions (ordered dict)
        self._sessions_dict = {}
        # self.set_sessions_dict(kwargs.get("sessions", None))

        # credential groups dict
        self._credentials_dict = {}

        # XML file
        self.xml_file = ""
        self._xml_obj = SMXml()
        self._xml_sessions = None
        self.set_xml_file(xml_file, read_xml_file)

        # XML file to export to XLS
        self._sessions_dict_file_xml = ""

        # JSON
        self._json_obj = SMJson()
        self._json_sessions = None

        self.json_file = ""
        self.set_json_file(
            kwargs.get("json_file", ""), kwargs.get("read_json_file", False)
        )

    # ====================
    # private methods
    # ====================

    # ====================
    # protected methods
    # ====================

    # ====================
    # public methods
    # ====================

    def get_credentials_dict(self):
        """Return credentials dictionary (ordered dict)."""
        return self._credentials_dict

    def get_credentials_dict_count(self):
        """Return credentials dictionary size (int)."""
        return len(self._credentials_dict["credential"])

    def get_sessions_dict(self) -> dict:
        """Return sessions dictionary.

        Returns:
            (dict): Sessions groups dict
        """
        return self._sessions_dict

    def get_sessions_dict_count(self, type=[""]) -> int:
        """Return sessions dictionary size.

        Args:
            type (str): Type of session. If empty, return all sessions
        Returns:
            (int): Sessions groups dict count
        """
        # return len(self._sessions_dict["session"])
        session_count = 0

        for idx, session_type in enumerate(self._sessions_dict["type"]):
            if self._sessions_dict["session"][idx] == "":
                # empty session name, skip it
                continue
            if "" in type:
                session_count += 1
            elif session_type in type:
                session_count += 1

        return session_count

    def get_xml_sessions(self) -> ET.Element | None:
        """Return sessions in XML format.

        Returns:
            (ET.Element): _sessions_xml attribute
        """
        return self._xml_sessions

    def parse_xml_file(self, xml_file="") -> ET.Element | None:
        """Read XML file and return ET.Element object."""
        if xml_file == "":
            xml_file = self.xml_file

        self._xml_sessions = SMXml(xml_file=xml_file).parse_xml_file()
        return self._xml_sessions

    def set_excel_file(self, excel_file: str | None = None, read_excel_file=False):
        """Set excel_file attribute.

        Args:
            excel_file (str): Excel file (source or destination)
        """
        self.excel_file = excel_file

        if self.excel_file is not None and read_excel_file:
            self.excel_read_book()

    def set_json_file(self, json_file: str, read_json_file=False):
        """Set XML file attribute. If xml_file is not empty, initialize self._xml_obj (read content).

        Args:
            xml_file (str): XML file (source or destination)
        """
        self.json_file = json_file

        # TODO
        # if self.json_file != "" and read_json_file:
        #     # self._xml_obj = SMXml(xml_file=self.xml_file, read_xml_file=True)
        #     self.parse_json_file()

    def set_xml_file(self, xml_file: str | None, read_xml_file=False):
        """Set XML file attribute. If xml_file is not empty, initialize self._xml_obj (read content).

        Args:
            xml_file (str): XML file (source or destination)
        """
        self.xml_file = xml_file

        if self.xml_file != "" and read_xml_file:
            # self._xml_obj = SMXml(xml_file=self.xml_file, read_xml_file=True)
            self.parse_xml_file()

    def set_sessions_dict(self, sessions=None) -> bool:
        """Set sessions dictionary. If not set initiate it.

        Args:
            sessions (dict): Sessions ordered dictionary

        Return:
            False in case of error (missing required column)
            True when success
        """
        # general fields (columns) in excel file
        excel_col_name = self._settings["excel"]["col_names_sessions"]
        keys = [
            "folder",
            "session",
            "type",
            "hostname",
            "port",
            "username",
            "rdp_alternate",
        ]
        required_keys = ["session", "type", "hostname"]

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

    def set_settings(self, settings: dict):
        """Set configuration settings dict (default: config.yaml).

        Args:
            settings (dict): config.yaml content
        """
        self._settings = settings

    def set_xml_sessions(self, xml_element: ET.Element):
        """Set XML sessions attribute."""
        self._xml_sessions = xml_element

    # ====================
    # Excel methods
    # ====================

    def excel_read_book(self) -> bool:
        """Read excel_file workbook.

        Returns:
            True: When success
            False: If not
        """

        if self.excel_file != "":
            self._excel_obj = SMExcel(
                excel_file=self.excel_file,
                settings=self._settings,
                read_excel_file=True,
            )
            return True

        return False

    def excel_read_sheet(self, sheet_name: str, type="column") -> dict | list | bool:
        """Read excel sheet and return content as dict/array.

        Args:
            sheet_name (str): Sheet's name
            get (str): One of ['column', 'row', 'array']. Defaults to "column".

        Returns:
            ordered dict: Column/Row-based dictionary (when get=['column', 'row']
            array: Array (when get=['array']
            False: In case of error
        """
        return self._excel_obj.read_excel_sheet(sheet_name, type)

    def excel_read_sheet_sessions(self, sheet_name: str) -> dict | list | bool:
        """Read excel sheet 'sessions' and return content as dict/array.

        Args:
            sheet_name (str): Sheet's name

        Returns:
            ordered dict: Column/Row-based dictionary (when get=['column', 'row']
            False: In case of error
        """
        sessions_dict_ret = self.excel_read_sheet(sheet_name, "column")
        if sessions_dict_ret:
            sessions_dict = self.col_name_normalize(
                sessions_dict_ret, self._settings["excel"]["col_names_sessions"]
            )
        else:
            sessions_dict = None

        if self.set_sessions_dict(sessions_dict):
            return self._sessions_dict
        return False

    def col_name_normalize(self, ordered_dict, col_names):
        """Change dictionary key from excel column name to system key defined in config.yaml

        Args:
            ordered_dict (ordered dict): Ordered dictionary (excel worksheet)
            col_names (dict): Dictionary "dict_key_name: excel_column_name"

        Returns:
            ordered dict: normalized ordered dict
        """

        ret_object = {}
        for dict_key in ordered_dict:
            for key, value in col_names.items():
                if dict_key == value:
                    ret_object[key] = ordered_dict[dict_key]

        return ret_object

    # ====================
    # XML methods
    # ====================

    def xml_write(self, **kwargs) -> None:
        """Write XML Element to file

        Args:
            xml_element (ET.Element, optional): XML object.
            xml_file (str, optional): Destination file. If not set, use self.xml_file.
        """

        xml_element = kwargs.get("xml_element", self._xml_sessions)

        # when xml_file is not defined, use object's self.xml_file attribute
        if "xml_file" in kwargs and kwargs.get("xml_file") != "":
            dst_file = kwargs["xml_file"]
        else:
            dst_file = self.xml_file

        self._xml_obj.write_xml_file(xml_element=xml_element, xml_file=dst_file)

    # ====================
    # JSON methods
    # ====================

    def print_json(self, **kwargs):
        """Print JSON to stdout as formated JSON.
        If not set, use self._json_sessions attribute.

        Args:
            json_content (optional): JSON to print.
        """

        json_content = kwargs.get("json_content", self._json_sessions)
        self._json_obj.print_json(json_content)

    def write_json(self, **kwargs) -> None:
        """Write JSON to file

        Args:
            json_content (, optional): JSON object.
            json_file (str, optional): Destination file. If not set, use self.xml_file.
        """

        json_content = kwargs.get("json_content", self._json_sessions)

        # when xml_file is not defined, use object's self.xml_file attribute
        if "json_file" in kwargs and kwargs.get("json_file") != "":
            dst_file = kwargs["json_file"]
        else:
            dst_file = self.json_file

        self._json_obj.write_json_file(json_content=json_content, json_file=dst_file)

    # ====================
    # general methods
    # ====================

    def xml_to_dict(self, xml_data):
        """Parse dict and return xml"""
        return xmltodict.parse(xml_data)

    def dict_to_xml(self, dict_data):
        """Parse dict and return xml"""
        return xmltodict.unparse(dict_data)

    def print_xml(self, **kwargs):
        """
        Print XML (ET.Element) object 'xml_element' to stdout as formated XML. 
        If not set, use self._session_xml attribute.

        Args:
            xml_element (ET.Element, optional): Element to print.
        """

        xml_element = kwargs.get("xml_element", self._xml_sessions)
        self._xml_obj.print_xml(xml_element=xml_element)
