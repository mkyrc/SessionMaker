"""SessionMaker Excel module

Class - SMExcel:
    Basic Excel sheet operations (read book and sheet data).

Author:
    Martin Kyrc

Version list:
    = 1.0 (20221117)
        - initial version

"""

import logging

# from os import path
import os.path
from pathlib import Path
from datetime import datetime

import pyexcel
import xlsxwriter

# from jinja2 import Environment, FileSystemLoader
# from ruamel.yaml import YAML


# ========================================
# Class SMExcel
# ========================================
class SMExcel:
    """SessionMaker Excel manipulation class.

    Top-Level class for Excel (data) manipulation.

    Attributes:
        Public:
            excel_file (str): Path to excel book file.
            excel_sheets (list): List of excel book sheets. Default: empty (all)
            settings (dict): COnfiguration file content

        Private:
        _excel_book (obj): Excel book content.
    """

    def __init__(self, **kwargs):
        ### public attributes

        ### private attributes
        self._excel_book = {}  # excel book dictionary

        # settings
        self._settings = {}  # settings
        self.set_settings(kwargs.get("settings", {}))

        # excel sheets (if empty use all)
        self._excel_sheets = []  # list of sheets to read
        self.set_excel_sheets(kwargs.get("sheets", []))

        # excel file name
        self._excel_file = ""
        # self._excel_obj = ""
        if kwargs.get("excel_file", "") != "":
            self.set_excel_file(
                kwargs.get("excel_file", ""), kwargs.get("read_excel_file", True)
            )

    # ========================================
    # Private methods
    # =======================================

    # def _get_nested(self, array):
    #     """Return nested dictionary based on array.
    #     First row defines names for nested distionary.
    #     """

    #     iteration = 0
    #     for item in array:
    #         if iteration == 0:
    #             break

    #         iteration += 1

    def read_excel_book(self):
        """Read whole excel book into 'self._excel_book'.

        Returns:
            True: When read excel book is successfully.
            False: When excel book file reading is not successful.
        """

        excel_book = {}
        try:
            if not self._excel_sheets:
                # read all sheets
                excel_book = pyexcel.get_book_dict(file_name=self._excel_file)
            else:
                # read defined sheets only
                excel_book = pyexcel.get_book_dict(
                    file_name=self._excel_file, sheets=self._excel_sheets
                )

            pyexcel.free_resources()
            logging.info("Loading excel book '%s' complete.", self._excel_file)

        except OSError as err:
            logging.error("Unable to load file or sheet(s).")
            logging.error("%s", err)
            return False

        self._excel_book = excel_book
        return True

    def set_excel_file(self, excel_file: str, read_excel_file=True):
        """Set self._excel_file variable

        Args:
            excel_file (str): Excel file path

        Returns:
            True: When success
            False: When not success
        """

        if not os.path.isfile(excel_file):
            logging.error("File path '%s' to Excel file is not valid.", excel_file)
            self._excel_file = ""
            return False

        self._excel_file = excel_file
        if excel_file != "" and read_excel_file:
            self.read_excel_book()

        return True

    def set_excel_sheets(self, excel_sheets):
        """Set self._excel_sheets variable

        Args:
            excel_sheets (list): Excel sheets list

        Returns:
            True: When success
            False: When not success
        """

        if not isinstance(excel_sheets, list):
            self._excel_sheets = []
            logging.error("Sheets must be list.")
            return False

        self._excel_sheets = excel_sheets
        return True

    def set_settings(self, settings):
        """Set interested excel sheets to work with."""
        self._settings = settings

    # ==========
    # Excel file manipulation
    # ==========

    def read_excel_sheet(self, sheet_name: str, get="column"):
        """Read sheet and return data.

        Read sheet's data from previously readed excel book.
        Return array (default) or column/row based dict.

        Arg:
            sheet_name (str): Sheet name.
            get (enum):
                column: return column based dict (key = first row)
                row: return row based dict (key = first column)
                array: return array (default)

        Returns:
            False: if error occured.
            (dict): when success.
        """
        try:
            sheet_array = self._excel_book[sheet_name]
            logging.info("Loading sheet '%s' from the book complete.", sheet_name)
        except KeyError:
            logging.warning(
                "Unable to load sheet '%s' from Excel file '%s'.",
                sheet_name,
                self._excel_file,
            )
            return False

        try:
            if get == "row":
                # get row-base dict. key is first col.
                sheet_content = pyexcel.get_dict(
                    array=sheet_array, name_columns_by_row=-1, name_rows_by_column=0
                )
            elif get == "column":
                # get column-base dict. key is first row.
                sheet_content = pyexcel.get_dict(
                    array=sheet_array, name_columns_by_row=0
                )
            else:
                # get array (no key-val based dict)
                sheet_content = pyexcel.get_array(array=sheet_array)
            logging.info("Reading data from sheet '%s' complete.", sheet_name)
        except Exception:
            logging.error("Unable to read data from the sheet '%s'.", sheet_name)
            return False

        return sheet_content

    def write_excel_book(self, **kwargs):
        # parse kwargs
        excel_file = str(kwargs.get("excel_file", self._excel_file))
        sessions_dict = kwargs.get("sessions_dict", {})
        rdm_credentials_dict = kwargs.get("rdm_credentials_dict", {})
        scrt_credentials_dict = kwargs.get("scrt_credentials_dict", {})
        scrt_firewalls_dict = kwargs.get("scrt_firewalls_dict", {})

        # Create a workbook and add a worksheet.
        workbook = xlsxwriter.Workbook(excel_file)
        # sheet_sessions = workbook.add_worksheet(name="sessions")

        today = datetime.today()
        workbook.set_properties(
            {
                "title": "Device sessions list",
                "subject": "",
                "author": "SessionMaker",
                "company": "Soitron",
                # 'category': 'Example spreadsheets',
                # 'keywords': 'Sample, Example, Properties',
                "created": today,
                "comments": "Sessions workbook generated by SessionMaker on "
                + today.strftime("%d/%m/%Y, %H:%M:%S"),
            }
        )

        workbook = self._write_sheet(
            workbook,
            "sessions",
            self._settings["excel"]["col_names_sessions"],
            sessions_dict,
        )
        workbook = self._write_sheet(
            workbook,
            "rdm-credentials",
            self._settings["excel"]["col_names_rdm_credentials"],
            rdm_credentials_dict,
        )
        workbook = self._write_sheet(
            workbook,
            "scrt-credentials",
            self._settings["excel"]["col_names_scrt_credentials"],
            scrt_credentials_dict,
        )
        workbook = self._write_sheet(
            workbook,
            "scrt-firewalls",
            self._settings["excel"]["col_names_scrt_firewalls"],
            scrt_firewalls_dict,
        )

        # preparing destination file
        logging.info("Writing Excel file '%s'.", excel_file)
        dst = os.path.split(excel_file)
        if os.path.isdir(dst[0]) is False:
            # create parent folders if not exists
            logging.info("Creating subfolder '%s'.", dst[0])
            Path(dst[0]).mkdir(parents=True, exist_ok=True)

        if os.path.exists(excel_file):
            logging.warning("Destination file '%s' exists. Overwriting.", excel_file)

        workbook.close()

    def _write_sheet(
        self, workbook, sheet_name, col_names=[], data=[], title_bg_color=""
    ):
        sheet = workbook.add_worksheet(name=sheet_name)

        # title format
        title_general = workbook.add_format({"bold": 1})
        if title_bg_color != "":
            title_general.set_fg_color(title_bg_color)
        title_scrt = workbook.add_format({"bold": 1})
        title_scrt.set_fg_color("#60b1b5")
        title_rdm = workbook.add_format({"bold": 1})
        title_rdm.set_fg_color("#3f8df3")

        logging.info("Creating workbook sheet '%s'", sheet_name)
        col = 0
        for key in col_names:
            if key.startswith("scrt_") or sheet_name.startswith("scrt "):
                # title_format.set_fg_color('#60b1b5')
                sheet.write(0, col, col_names[key], title_scrt)
            elif key.startswith("rdm_") or sheet_name.startswith("rdm "):
                # title_format.set_fg_color('#3f8df3')
                sheet.write(0, col, col_names[key], title_rdm)
            else:
                # title_format.set_fg_color('#ffffff')
                sheet.write(0, col, col_names[key], title_general)

                # 3f8df3
            if not key in data.keys():
                data[key] = [""]

            # set column width
            if len(max(data[key], key=len)) > len(col_names[key]):
                # if len(max(len(data[key]), len(int(key)))) > len(col_names[key]):
                # if max(len(data[key]), len(key)) > len(col_names[key]):
                col_width = len(max(data[key], key=len))
            else:
                col_width = len(col_names[key])
            sheet.set_column(col, col, col_width + 1)
            sheet.write_column(1, col, data[key])

            col += 1

        return workbook
