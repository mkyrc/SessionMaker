"""
Session Reader - SecureCRT session reader

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
from lib.parseargs import parse_reader_args
from lib.logging import init_logging
from lib.settings import set_config_file
from lib.settings import read_config_file
from lib.sm_scrt import SMSecureCrt
from lib.sm_remmina import SMRemmina

# ====================
# Main function
# ====================

def main():
    """Main function of the script"""

    ## default settings
    config_file = "config.yaml"  # default settings file
    src_input = []  # source file(s)/folder(s)
    dst_file = None # Excel file

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

    # source file(s)
    if ARGS.source:
        src_input = ARGS.source

    # destination file (excel)
    # if undefined, export to 'export' subfolder
    if ARGS.write != None:
        dst_file = ARGS.write
    else:
        first_source = src_input[0]
        if os.path.isdir(first_source):
            source_dir = first_source
            source_name = Path(first_source).name
        else:
            src_folder = os.path.split(first_source)
            source_dir = src_folder[0]
            source_name = Path(src_folder[1]).stem

        if source_name == "":
            source_name = "sessions"
        dst_file = source_dir + "/export/" + source_name + ".xlsx"

    if not ARGS.quiet:
        print("Done.")

    if ARGS.type == "scrt":
        scrt_reader_2(
            settings=config_data,
            src_file=src_input[0],
            dst_file=dst_file,
            quiet=ARGS.quiet,
        )

    if ARGS.type == "remmina":
        remmina_reader(
            settings=config_data,
            src_list=src_input,
            dst_file=dst_file,
            quiet=ARGS.quiet,
        )


# ====================
# Functions
# ====================


def scrt_reader_1(**kwargs):
    """Read SecureCRT XML sessions file and export it to Excel/stdout.

    EXAMPLE #1: step-by-step
    """

    # parse kwargs
    settings = kwargs.get("settings", {})
    src_file = kwargs.get("src_file", "")
    dst_file = kwargs.get("dst_file", "")
    quiet = kwargs.get("quiet", False)

    if not quiet:
        print("Read SecureCRT sessions XML file...")

    sm_scrt = SMSecureCrt(settings=settings, xml_file=src_file, read_xml_file=True)

    if sm_scrt.get_xml_sessions() == None:
        if not quiet:
            print("Exit.")
        return

    if not quiet:
        print("Done.")

    # sessions
    if not quiet:
        print("Reading SecureCRT sessions...")
    sm_scrt.set_sessions_dict_from_xml()
    if not quiet:
        print("Done. Imported: %d" % sm_scrt.get_sessions_dict_count())

    # credentials
    if not quiet:
        print("Reading SecureCRT credential groups...")
    sm_scrt.set_credentials_dict_from_xml()
    if not quiet:
        print("Done. Imported: %d" % sm_scrt.get_credentials_dict_count())

    # firewalls
    if not quiet:
        print("Reading SecureCRT firewall groups...")
    sm_scrt.set_firewalls_dict_from_xml()
    if not quiet:
        print("Done. Imported: %d" % sm_scrt.get_firewalls_dict_count())


def scrt_reader_2(**kwargs):
    """Read SecureCRT XML sessions file and export it to Excel/stdout.

    EXAMPLE #2: all in one.
    """

    ## parse kwargs
    settings = kwargs.get("settings", {})
    src_file = kwargs.get("src_file", "")
    dst_file = kwargs.get("dst_file", "")
    quiet = kwargs.get("quiet", False)

    # parse XML and prepare dictionaries
    # ==========    

    if not quiet:
        print("Reading SecureCRT sessions XML file...")

    sm_scrt = SMSecureCrt(settings=settings, xml_file=src_file, read_xml_file=True)

    if sm_scrt.build_dict_from_xml() == False:
        if not quiet:
            print("No Sessions. Exit")

    if not quiet:
        print(
            "Done. %d session(s), %d credential group(s), %d firewall group(s) from XML file."
            % (
                sm_scrt.get_sessions_dict_count(),
                sm_scrt.get_credentials_dict_count(),
                sm_scrt.get_firewalls_dict_count(),
            )
        )

    # Write to Excel file
    # ==========    

    if not quiet:
        print("Writing Excel file...")
    sm_scrt.set_excel_file(dst_file, False)
    sm_scrt.write_excel()
    if not quiet:
        print("Done.")


def _find_remmina_profile_files(src_list: list[str]) -> list[str]:
    profile_files = []
    for src_item in src_list:
        if os.path.isdir(src_item):
            for root, _, files in os.walk(src_item):
                for file_name in files:
                    if file_name.lower().endswith(".remmina"):
                        profile_files.append(os.path.join(root, file_name))
        elif src_item.lower().endswith(".remmina") and os.path.isfile(src_item):
            profile_files.append(src_item)

    profile_files = sorted(set(profile_files))
    return profile_files


def remmina_reader(**kwargs):
    settings = kwargs.get("settings", {})
    src_list = kwargs.get("src_list", [])
    dst_file = kwargs.get("dst_file", "")
    quiet = kwargs.get("quiet", False)

    if not quiet:
        print("Reading Remmina profile(s)...")

    profile_files = _find_remmina_profile_files(src_list)
    if len(profile_files) == 0:
        if not quiet:
            print("No Remmina profiles found. Exit.")
        return

    sm_remmina = SMRemmina(settings=settings)
    imported_count, skipped_count = sm_remmina.load_sessions_from_remmina_files(profile_files)

    if imported_count == 0:
        if not quiet:
            print("No valid Remmina sessions found. Exit.")
        return

    if not quiet:
        c_total = sm_remmina.get_sessions_dict_count([""])
        c_ssh = sm_remmina.get_sessions_dict_count(["ssh"])
        c_rdp = sm_remmina.get_sessions_dict_count(["rdp"])
        print(
            "Done. %d session(s) (ssh: %d, rdp: %d), skipped %d profile(s), source files: %d."
            % (c_total, c_ssh, c_rdp, skipped_count, len(profile_files))
        )

    if not quiet:
        print("Writing Excel file...")
    sm_remmina.write_excel(dst_file)
    if not quiet:
        print("Done.")


# ====================
# Initial functions
# ====================


if __name__ == "__main__":
    ARGS = parse_reader_args()
    init_logging(ARGS.verbose)
    main()
