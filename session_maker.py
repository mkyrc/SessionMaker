"""
Session Maker - SecureCRT and Devolutions RDM session generator

Arguments:


Author:
    Martin Kyrc,
    Soitron NetOps Team

Revision:
    1.0 (2022-11-08)
        - initial version
"""

# import logging
import os.path
from pathlib import Path

# import lib
from lib.parseargs import parse_maker_args
from lib.logging import init_logging
from lib.settings import set_config_file
from lib.settings import read_config_file

# from lib.sm_excel import SMExcel
from lib.sm_scrt import SMSecureCrt
from lib.sm_rdm import SMDevolutionsRdm

# ====================
# Main functions
# ====================


def main():
    """Main function of the script"""

    ## default settings
    config_file = "config.yaml"  # default settings file
    src_file = None  # Excel file
    dst_file = None  # SecureCRT XML or Devolutions RDM JSON file

    # arguments
    # ==========

    if not ARGS.quiet:
        print("Reading arguments...")

    # read config file
    # if undefined, use 'config.yaml'
    if ARGS.config:
        config_file = set_config_file(ARGS.config.strip(), config_file)

    config_data = read_config_file(config_file)
    if config_data is False:
        return

    # source file (excel)
    if ARGS.source:
        src_file = ARGS.source

    # destination file (xml or json)
    # if undefined, export to 'export' subfolder
    if not ARGS.print:
        if ARGS.write != None:
            dst_file = ARGS.write
        else:
            src_folder = os.path.split(ARGS.source)
            filename = Path(src_folder[1]).stem
            if ARGS.type == "scrt":
                dst_file = src_folder[0] + "/export/" + filename + ".xml"
            if ARGS.type == "rdm":
                dst_file = src_folder[0] + "/export/" + filename + ".json"

    if not ARGS.quiet:
        print("Done.")

    # ===========
    # Make a sessions
    # ===========

    if ARGS.type == "scrt":
        # SecureCRT sessions (XML content) maker

        scrt_maker(
            settings=config_data,
            src_file=src_file,
            dst_file=dst_file,
            quiet=ARGS.quiet,
            stdout=ARGS.print,
        )

    if ARGS.type == "rdm":
        # Devolutions RDM session (JSON content) maker
        rdm_maker(
            settings=config_data,
            src_file=src_file,
            dst_file=dst_file,
            quiet=ARGS.quiet,
            stdout=ARGS.print,
        )


# ====================
# Functions
# ====================


def scrt_maker(**kwargs):
    """Reading Excel and export sessions to SecureCRT."""

    # arguments
    settings = kwargs.get("settings", {})
    src_file = kwargs.get("src_file", "")  # src Excel file
    dst_file = kwargs.get("dst_file", "")  # dst XML file
    quiet = kwargs.get("quiet", False)
    stdout = kwargs.get("stdout", False)

    # Reading Excel
    # ==========

    if not quiet:
        print("Reading Excel book...")

    sm_scrt = SMSecureCrt(settings=settings, excel_file=src_file, read_excel_file=True)

    # get excel content (and set object's attribute(s))
    sessions_dict = sm_scrt.excel_read_sheet_sessions(settings["excel"]["tab_sessions"])
    credentials_dict = sm_scrt.excel_read_sheet_credentials(
        settings["excel"]["tab_scrt_credentials"]
    )
    firewalls_dict = sm_scrt.excel_read_sheet_firewalls(
        settings["excel"]["tab_scrt_firewalls"]
    )

    if sessions_dict == False or credentials_dict == False or firewalls_dict == False:
        if not ARGS.quiet:
            print("Exit.")
        return

    # summary
    if not quiet:
        print(
            "Done: %d sessions (ssh: %d), %d credential group(s), %d firewall group(s) from Excel."
            % (
                sm_scrt.get_sessions_dict_count(["ssh"]),
                sm_scrt.get_sessions_dict_count(["ssh"]),
                sm_scrt.get_credentials_dict_count(),
                sm_scrt.get_firewalls_dict_count(),
            )
        )

    # Building SecureCRT sessions
    # ==========

    if not quiet:
        print("Building sessions...")

    scrt_xml = sm_scrt.build_xml_from_dict()

    if scrt_xml == None:
        if not quiet:
            print("No sessions. Exit.")
        return
    if not quiet:
        print("Done.")

    # Exporting
    # ==========

    if stdout:
        # print to stdout
        if not quiet:
            print("XML content...")
        sm_scrt.print_xml()
    else:
        # write to file
        if not quiet:
            print(f"Writing to '{dst_file}'...")
        sm_scrt.set_xml_file(dst_file)
        sm_scrt.xml_write()
        # alebo takto:
        # sm_scrt.xml_write(xml_file=dst_file)

    if not quiet:
        print("Done.")


def rdm_maker(**kwargs):
    """Export session to DevolutionsRDM - by settings source file only."""

    # arguments
    settings = kwargs.get("settings", {})
    src_file = kwargs.get("src_file", "")  # src Excel file
    dst_file = kwargs.get("dst_file", "")  # dst JSON file
    quiet = kwargs.get("quiet", False)
    stdout = kwargs.get("stdout", False)

    # Reading Excel
    # ==========

    if not quiet:
        print("Reading Excel book...")

    sm_rdm = SMDevolutionsRdm(
        settings=settings, excel_file=src_file, read_excel_file=True
    )

    # get content (and set object's attribute(s))
    sessions_dict = sm_rdm.excel_read_sheet_sessions(settings["excel"]["tab_sessions"])
    credentials_dict = sm_rdm.excel_read_sheet_credentials(
        settings["excel"]["tab_rdm_credentials"]
    )

    if sessions_dict == False or credentials_dict == False:
        if not ARGS.quiet:
            print("Exit.")
        return

    # summary
    if not quiet:
        print(
            "Done. %d session(s) (ssh: %s, rdp: %s, web: %s), %d credential(s) from Excel."
            % (
                sm_rdm.get_sessions_dict_count(["ssh","rdp","web"]),
                sm_rdm.get_sessions_dict_count(["ssh"]),
                sm_rdm.get_sessions_dict_count(["rdp"]),                
                sm_rdm.get_sessions_dict_count(["web"]),                
                sm_rdm.get_credentials_dict_count(),
            )
        )

    # Building Devolutions RDM sessions
    # ==========

    if not quiet:
        print("Building sessions...")

    rdm_json = sm_rdm.build_json_from_dict()

    if rdm_json == None:
        if not quiet:
            print("No sessions. Exit.")
        return
    if not quiet:
        print("Done.")

    # Exporting
    # ==========

    if stdout:
        # print to stdout
        if not quiet:
            print("JSON content...")
        sm_rdm.print_json()
    else:
        # write to file
        if not quiet:
            print(f"Writing to '{dst_file}'...")
        sm_rdm.set_json_file(dst_file)
        sm_rdm.write_json()
        # sm_rdm.xml_write(xml_file=dst_file)

    if not quiet:
        print("Done.")


# ====================
# Initial function
# ====================

if __name__ == "__main__":
    ARGS = parse_maker_args()
    init_logging(ARGS.verbose)
    main()
