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

import logging
import os.path
from pathlib import Path
from datetime import datetime

# import lib
from lib import parse_maker_args, init_logging
from lib import SMSecureCrt, SMDevolutionsRdm, Settings

# ====================
# Main functions
# ====================

# global ARGS


def main():
    """Main function of the script"""

    ARGS = parse_maker_args()
    init_logging(ARGS.verbose)

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
        config_file = ARGS.config.strip()

    stg = Settings(config_file)

    # config_data = read_config_file(config_file)
    if stg.app_config is None:
        logging.error("Unable to read configuration file '%s'.", config_file)
        return

    # source file (excel)
    if ARGS.source:
        src_file = ARGS.source

    # destination file (xml or json)
    # if undefined, export to 'export' subfolder
    if not ARGS.print:
        if ARGS.write:
            dst_file = ARGS.write
        else:
            src_folder = os.path.split(ARGS.source)
            filename = Path(src_folder[1]).stem
            current_date = datetime.now().strftime("%Y%m%d")

            if ARGS.type == "scrt":
                dst_file = f"{src_folder[0]}/export/{current_date}-{filename}-scrt.xml"
            if ARGS.type == "rdm":
                dst_file = f"{src_folder[0]}/export/{current_date}-{filename}-rdm.json"

    if not ARGS.quiet:
        print("Done.")

    # ===========
    # Make a sessions
    # ===========

    if ARGS.type == "scrt":
        # SecureCRT sessions (XML content) maker

        scrt_maker(
            stg=stg,
            src_file=src_file,
            dst_file=dst_file,
            quiet=ARGS.quiet,
            stdout=ARGS.print,
        )

    if ARGS.type == "rdm":
        # Devolutions RDM session (JSON content) maker
        rdm_maker(
            stg=stg,
            src_file=src_file,
            dst_file=dst_file,
            quiet=ARGS.quiet,
            stdout=ARGS.print,
        )


# ====================
# Functions
# ====================


def scrt_maker(
    stg: Settings,
    src_file: str | None = None,
    dst_file: str | None = None,
    quiet=False,
    stdout=False,
):
    """Reading Excel and export sessions to SecureCRT."""

    # arguments
    # settings = kwargs.get("settings", {})
    # src_file = kwargs.get("src_file", "")  # src Excel file
    # dst_file = kwargs.get("dst_file", "")  # dst XML file
    # quiet = kwargs.get("quiet", False)
    # stdout = kwargs.get("stdout", False)

    # Reading Excel
    # ==========

    if not quiet:
        print("Reading Excel book...")

    sm_scrt = SMSecureCrt(
        settings=stg.app_config,
        excel_file=src_file,
        read_excel_file=True,
    )

    # get excel content (and set object's attribute(s))
    sessions_dict = sm_scrt.excel_read_sheet_sessions(
        stg.app_config["excel"]["tab_sessions"]
    )
    credentials_dict = sm_scrt.excel_read_sheet_credentials(
        stg.app_config["excel"]["tab_scrt_credentials"]
    )
    firewalls_dict = sm_scrt.excel_read_sheet_firewalls(
        stg.app_config["excel"]["tab_scrt_firewalls"]
    )

    if sessions_dict is False or credentials_dict is False or firewalls_dict is False:
        if not quiet:
            print("Exit.")
        return

    # summary
    if not quiet:
        c_s = sm_scrt.get_sessions_dict_count(["ssh"])
        c_s_ssh = sm_scrt.get_sessions_dict_count(["ssh"])
        c_cred = sm_scrt.get_credentials_dict_count()
        c_fw = sm_scrt.get_firewalls_dict_count()
        p_sessions = f"{c_s} session(s) (ssh: {c_s_ssh})"
        p_credentials = f"{c_cred} credential(s)"
        p_firewalls = f"{c_fw} firewall(s)"

        print(f"Done. {p_sessions}, {p_credentials}, {p_firewalls} from Excel.")

    # Building SecureCRT sessions
    # ==========

    if not quiet:
        print("Building sessions...")

    scrt_xml = sm_scrt.build_xml_from_dict()

    if scrt_xml is None:
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


def rdm_maker(
    stg: Settings,
    src_file: str | None = None,
    dst_file: str | None = None,
    quiet=False,
    stdout=False,
):
    """
    Generates Devolutions RDM sessions from an Excel file and exports them to JSON.

    Args:
        src_file (str, optional): Path to the source Excel file. Defaults to None.
        dst_file (str, optional): Path to the destination JSON file. Defaults to None.
        stg (Settings): App settings. Defaults
        quiet (bool, optional): If True, suppresses output messages. Defaults to False.
        stdout (bool, optional): If True, prints the JSON content to stdout 
                                 instead of writing to a file. Defaults to False.

    Returns:
        None
    """

    # arguments
    # settings = kwargs.get("settings", {})
    # src_file = kwargs.get("src_file", "")  # src Excel file
    # dst_file = kwargs.get("dst_file", "")  # dst JSON file
    # quiet = kwargs.get("quiet", False)
    # stdout = kwargs.get("stdout", False)

    # Reading Excel
    # ==========

    if not quiet:
        print("Reading Excel book...")

    stg.read_session_defaults(client_type="rdm")
    sm_rdm = SMDevolutionsRdm(
        settings=stg.app_config,
        excel_file=src_file,
        read_excel_file=True,
        session_defaults_rdm=stg.rdm_session_defaults,
    )

    # get content (and set object's attribute(s))
    sessions_dict = sm_rdm.excel_read_sheet_sessions(
        stg.app_config["excel"]["tab_sessions"]
    )
    credentials_dict = sm_rdm.excel_read_sheet_credentials(
        stg.app_config["excel"]["tab_rdm_credentials"]
    )
    hosts_dict = sm_rdm.excel_read_sheet_rdm_hosts(
        stg.app_config["excel"]["tab_rdm_hosts"]
    )

    # if sessions_dict is False or credentials_dict is False or hosts_dict is False:
    if sessions_dict is False:
        if not quiet:
            print("Exit.")
        return

    # summary
    if not quiet:
        c_s = sm_rdm.get_sessions_dict_count(["ssh", "rdp", "web"])
        c_s_ssh = sm_rdm.get_sessions_dict_count(["ssh"])
        c_s_rdp = sm_rdm.get_sessions_dict_count(["rdp"])
        c_s_web = sm_rdm.get_sessions_dict_count(["web"])
        c_cred = sm_rdm.get_credentials_dict_count()
        c_host = sm_rdm.get_rdm_hosts_dict_count()

        p_sessions = (
            f"{c_s} session(s) (ssh: {c_s_ssh}, rdp: {c_s_rdp}, web: {c_s_web})"
        )
        p_credentials = f"{c_cred} credential(s)"
        p_hosts = f"{c_host} host(s)"

        print(f"Done. {p_sessions}, {p_credentials}, {p_hosts} from Excel.")

    # Building Devolutions RDM sessions
    # ==========

    if not quiet:
        print("Building sessions...")

    rdm_json = sm_rdm.build_json_from_dict()

    if rdm_json is None:
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

    main()
