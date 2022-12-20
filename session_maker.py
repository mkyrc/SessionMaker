"""
Session Maker - SecureCRT session generator

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

# import pprint

# import lib
from lib.parseargs import parse_maker_args
from lib.logging import init_logging
from lib.settings import set_config_file
from lib.settings import read_config_file
from lib.sm_excel import SMExcel
from lib.sm_scrt import SMSecureCrt


# ====================
# Main functions
# ====================


def main():
    """Main function of the script"""

    ## default settings
    config_file = "config.yaml"  # default settings file
    dst_file = None
    src_file = None

    ## arguments
    if not ARGS.quiet:
        print("Reading arguments...")

    # set and read config file
    if ARGS.config:
        config_file = set_config_file(ARGS.config.strip(), config_file)

    config_data = read_config_file(config_file)
    if config_data is False:
        return

    # set source excel file
    if ARGS.source:
        src_file = ARGS.source

    # set destination file
    if not ARGS.print:
        if ARGS.write != None:
            dst_file = ARGS.write
        else:
            src_folder = os.path.split(ARGS.source)
            filename = Path(src_folder[1]).stem
            dst_file = src_folder[0] + "/export/" + filename + ".xml"

    if not ARGS.quiet:
        print("Done.")

    ## /arguments

    # export sessions to SecureCRT
    if ARGS.type == "scrt":

        example = 2

        ### how to #1 ###

        if example == 1:
            # read Excel file and sheets content
            if not ARGS.quiet:
                print(f"Reading Excel file '{ARGS.source}'...")

            excel_obj = SMExcel(
                settings=config_data, excel_file=ARGS.source, read_excel_file=True
            )
            # excel_obj.read_excel_book()
            cfg_excel = config_data["excel"]
            sessions_dict = excel_obj.read_excel_sheet(cfg_excel["tab_sessions"])
            credentials_dict = excel_obj.read_excel_sheet(cfg_excel["tab_credentials"])
            firewalls_dict = excel_obj.read_excel_sheet(cfg_excel["tab_firewalls"])

            if (
                sessions_dict == False
                or credentials_dict == False
                or firewalls_dict == False
            ):

                if not ARGS.quiet:
                    print("Exit.")
                return

            if not ARGS.quiet:
                print("Done.")

            # build output based on above dictionaries
            scrt_maker_1(
                settings=config_data,
                dst_file=dst_file,
                quiet=ARGS.quiet,
                stdout=ARGS.print,
                sessions=sessions_dict,  # <<<
                credentials=credentials_dict,  # <<<
                firewalls=firewalls_dict,  # <<<
            )

        ### how to #2 ###

        if example == 2:
            scrt_maker_2(
                settings=config_data,
                src_file=src_file,  # <<<
                dst_file=dst_file,
                quiet=ARGS.quiet,
                stdout=ARGS.print,
            )


# ====================
# Functions
# ====================


def scrt_maker_2(**kwargs):
    """Export session to SecureCRT - by settings source file only."""

    settings = kwargs.get("settings", {})
    src_file = kwargs.get("src_file", "")
    dst_file = kwargs.get("dst_file", "")
    quiet = kwargs.get("quiet", False)
    stdout = kwargs.get("stdout", False)

    if not quiet:
        print("Exporting sessions for SecureCRT...")

    # sm_scrt = SMSecureCrt(settings=settings, excel_file=src_file, xml_file=dst_file)
    sm_scrt = SMSecureCrt(settings=settings, excel_file=src_file, read_excel_file=True)
    # sm_scrt.excel_read_book()

    # get content (and set object's attribute(s))
    sessions_dict = sm_scrt.excel_read_sheet(settings["excel"]["tab_sessions"])
    credentials_dict = sm_scrt.excel_read_sheet(settings["excel"]["tab_credentials"])
    firewalls_dict = sm_scrt.excel_read_sheet(settings["excel"]["tab_firewalls"])

    if sessions_dict == False or credentials_dict == False or firewalls_dict == False:
        if not ARGS.quiet:
            print("Exit.")
        return

    sm_scrt.set_sessions_dict(sessions_dict)
    sm_scrt.set_credentials_dict(credentials_dict)
    sm_scrt.set_firewalls_dict(firewalls_dict)

    if not quiet:
        print(
            "Exported: %d sessions, %d credential groups, %d firewall groups."
            % (
                sm_scrt.get_sessions_dict_count(),
                sm_scrt.get_credentials_dict_count(),
                sm_scrt.get_firewalls_dict_count(),
            )
        )
    # get XML from above initiated sessions/credentials
    scrt_xml = sm_scrt.build_xml_from_dict()
    if scrt_xml == None:
        if not quiet:
            print("No sessions. Exit.")
        return

    if stdout:
        # print to stdout
        sm_scrt.print_xml()
    else:
        # write to file
        # bud takto (ak nie je dst_file nastaveny pri inicializacii objektu):
        sm_scrt.set_xml_file(dst_file)
        sm_scrt.xml_write()
        # alebo takto:
        # sm_scrt.xml_write(xml_file=dst_file)

    if not quiet:
        print("Done.")


def scrt_maker_1(**kwargs):
    """Export session to SecureCRT - #1.

    First get sessions, credentials and firewalls and then creat 'sm_scrt' object with these content.
    """

    settings = kwargs.get("settings", {})
    dst_file = kwargs.get("dst_file", "")
    quiet = kwargs.get("quiet", False)
    stdout = kwargs.get("stdout", False)

    sessions_dict = kwargs.get("sessions")
    credentials_dict = kwargs.get("credentials")
    firewalls_dict = kwargs.get("firewalls")

    if not quiet:
        print("Exporting sessions for SecureCRT...")

    sm_scrt = SMSecureCrt(
        settings=settings,
        xml_file=dst_file,
        sessions=sessions_dict,
        credentials=credentials_dict,
        firewalls=firewalls_dict,
    )

    # get XML from above initiated sessions/credentials
    scrt_xml = sm_scrt.build_xml_from_dict()
    if scrt_xml == None:
        if not quiet:
            print("No sessions.")
        return

    if stdout:
        # print to stdout
        sm_scrt.print_xml()
    else:
        # print to file
        sm_scrt.set_xml_file(dst_file)
        # sm_scrt.xml_write(xml_file=dst_file)
        sm_scrt.xml_write()

    if quiet is False:
        print("Done.")

    return True


# ====================
# Initial function
# ====================

if __name__ == "__main__":

    # parse args and settings
    ARGS = parse_maker_args()

    init_logging(ARGS.verbose)

    # run the main method
    main()
