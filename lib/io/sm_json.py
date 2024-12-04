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

import json


# ========================================
# Class SMJson
# ========================================
class SMJson:
    """SessionMaker JSON manipulation class.

    Top-Level class for JSON content manipulation.

    Attributes:
        Public:
            xml_file (str): Path to XML  file.
            excel_sheets (list): List of excel book sheets. Default: empty (all)
            settings (dict): COnfiguration file content

        Private:
        _json (obj): JSON  content.
    """

    def __init__(self, **kwargs):

        ### public attributes
        self.json_file = ""
        self._json_content = None
        self.set_json_file(
            kwargs.get("json_file", ""), kwargs.get("read_json_file", False)
        )
        # self._json = dict()

    # ========================================
    # Private methods
    # ========================================

    # ========================================
    # Protected methods
    # ========================================

    # ========================================
    # Public methods
    # ========================================

    def print_json(self, json_content=None):
        """Print formated JSON to stdout.

        Args:
            json (optional): JSON object. Defaults to dict().
        """
        if json_content is None:
            return

        # obj = json.loads(json_content)
        json_formated = json.dumps(json_content, indent=4)
        print(json_formated)

    def set_json_file(self, json_file: str, read_json_file=False):
        """Set JSON file attribute"""
        self.json_file = json_file
        # if json_file != "" and read_json_file:
        #     self.parse_json_file()

    def write_json_file(self, json_file: str | None = None, json_content=None) -> None:
        """
        Writes JSON content to a specified file.
        
        Args:
            json_file (str | None, optional): The path to the JSON file. If None, defaults to self.json_file.
            json_content (any, optional): The content to be written to the JSON file. If None, defaults to self._json_content.
        
        Raises:
            FileNotFoundError: If the specified file path does not exist and cannot be created.
        
        Logs:
            Info: When creating subfolders that do not exist.
            Warning: If the destination file already exists and will be overwritten.
            Error: If unable to write to the JSON file due to a FileNotFoundError.
        """
        

        # json_file = str(kwargs.get("json_file", self.json_file))
        if json_file is None:
            json_file = self.json_file

        # json_content = kwargs.get("json_content", self._json_content)
        if json_content is None:
            json_content = self._json_content

        json_object = json.dumps(json_content, indent=4)

        dst = os.path.split(json_file)
        if os.path.isdir(dst[0]) is False:
            # create parent folders if not exists
            logging.info("Creating subfolder '%s'.", dst[0])
            Path(dst[0]).mkdir(parents=True, exist_ok=True)

        if os.path.exists(json_file):
            logging.warning("Destination file '%s' exists. Overwriting.", json_file)

        # write to file
        try:
            with open(json_file, "w", encoding="utf8") as outfile:
                outfile.write(json_object)
        except FileNotFoundError as err:
            logging.error(
                "Unable to write. JSON file destination not set.",
            )
            logging.error("%s", err)

    #     # xml_element = ET.Element(kwargs.get("xml_element", self._xml_element))

    #     dst = os.path.split(xml_file)
    #     if os.path.isdir(dst[0]) is False:
    #         # create parent folders if not exists
    #         logging.info("Creating subfolder '%s'.", dst[0])
    #         Path(dst[0]).mkdir(parents=True, exist_ok=True)

    #     if os.path.exists(xml_file):
    #         logging.warning("Destination file '%s' exists. Overwriting.", xml_file)

    #     logging.info("Writing XML file '%s'.", xml_file)
    #     if type(xml_element) is ET.Element:
    #         # ET.indent(xml_element, space="\t", level=0)
    #         tree = ET.ElementTree(element=xml_element)
    #         ET.indent(tree, space="\t", level=0)
    #         try:
    #             tree.write(xml_file, encoding="utf8")
    #         except FileNotFoundError as err:
    #             logging.error(
    #                 "Unable to write. Destination XML file not set.",
    #             )
    #             logging.error("%s", err)
    #             return
    #     else:
    #         logging.error("Wrong XML element type")
    #         return
