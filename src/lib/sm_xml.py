"""SessionMaker XML module

Class - SMXml:
    Basic XML operations (read and write).

Author:
    Martin Kyrc

Version list:
    = 1.0 (20221205)
        - initial version

"""
import logging
import os.path
from pathlib import Path

import xml.etree.ElementTree as ET

# import pyexcel

# from jinja2 import Environment, FileSystemLoader
# from ruamel.yaml import YAML

# ========================================
# Class SMExcel
# ========================================
class SMXml:
    """SessionMaker XML manipulation class.

    Top-Level class for XML content manipulation.

    Attributes:
        Public:
            xml_file (str): Path to XML  file.
            excel_sheets (list): List of excel book sheets. Default: empty (all)
            settings (dict): COnfiguration file content

        Private:
        _excel_book (obj): Excel book content.
    """

    def __init__(self, **kwargs):

        ### public attributes
        self.xml_file = ""
        self._xml_element = ET.Element
        self.set_xml_file(
            kwargs.get("xml_file", ""), kwargs.get("read_xml_file", False)
        )


    # ========================================
    # Private methods
    # ========================================

    # ========================================
    # Protected methods
    # ========================================

    # ========================================
    # Public methods
    # ========================================

    def parse_xml_file(self, xml_file="") -> ET.Element | None:
        """Read XML file and return ET.Element root object."""

        if xml_file == "":
            xml_file = self.xml_file

        try:
            logging.info("Parsing XML file '%s'...", xml_file)
            root = ET.parse(xml_file)
            self._xml_element = root.getroot()
            logging.info("Success.")
        except ET.ParseError as err:
            logging.error("Unable to parse XML file '%s'", xml_file)
            logging.error("%s", err)
            self._xml_element = None
        except FileNotFoundError as err:
            logging.error("Unable to read XML file '%s'", xml_file)
            logging.error("%s", err)
            self._xml_element = None

        return self._xml_element

    def print_xml(self, **kwargs):
        """Print ElementTree object to stdout as formated XML"""

        xml_element = kwargs.get("xml_element", self._xml_element)

        if type(xml_element) is ET.Element:
            tree = ET.ElementTree(element=xml_element)
            ET.indent(tree, space="\t", level=0)
            # print(ET.tostring(et_data, encoding="utf8").decode("utf8"))
            print(
                ET.tostring(
                    xml_element, short_empty_elements=False, encoding="utf8"
                ).decode("utf8")
            )

    def set_xml_file(self, xml_file: str, read_xml_file=False):
        """Set XML file attribute"""
        self.xml_file = xml_file
        if xml_file != "" and read_xml_file:
            self.parse_xml_file()

    def write_xml_file(self, **kwargs) -> None:
        """Write XML object to file."""

        xml_file = str(kwargs.get("xml_file", self.xml_file))
        # xml_element = ET.Element(kwargs.get("xml_element", self._xml_element))
        xml_element = kwargs.get("xml_element", self._xml_element)

        dst = os.path.split(xml_file)
        if os.path.isdir(dst[0]) is False:
            # create parent folders if not exists
            logging.info("Creating subfolder '%s'.", dst[0])
            Path(dst[0]).mkdir(parents=True, exist_ok=True)

        if os.path.exists(xml_file):
            logging.warning("Destination file '%s' exists. Overwriting.", xml_file)

        logging.info("Writing XML file '%s'.", xml_file)
        if type(xml_element) is ET.Element:
            # ET.indent(xml_element, space="\t", level=0)
            tree = ET.ElementTree(element=xml_element)
            ET.indent(tree, space="\t", level=0)
            try:
                tree.write(xml_file, encoding="utf8")
            except FileNotFoundError as err:
                logging.error(
                    "Unable to write. Destination XML file not set.",
                )
                logging.error("%s", err)
                return
        else:
            logging.error("Wrong XML element type")
            return
