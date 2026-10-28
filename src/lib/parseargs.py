"""Parse arguments library"""

import argparse

# import os.path
# import logging

# from pathlib import Path
# from ruamel.yaml import YAML

# VERSION = __version__

def get_version():
    # Read the version from the VERSION file
    version_file = "VERSION"
    try:
        with open(version_file, "r", encoding="utf8") as file:
            version = file.read().strip()
    except FileNotFoundError:
        version = "unknown"
    return version

def parse_maker_args():
    """Parse arguments for export sessions

    Returns:
        object: Arguments
    """
    parser = argparse.ArgumentParser(
        description="Read Excel file (source) and generate sessions output for [SecureCRT|Devolutions|RDP|Remmina]."
    )
    group1 = parser.add_mutually_exclusive_group()
    group2 = parser.add_mutually_exclusive_group()

    parser.add_argument(
        "--config",
        type=str,
        metavar="CONFIG",
        help="Configuration settings file (default=config.yaml)",
        default="config.yaml",
    )
    parser.add_argument("source", type=str, help="Source (XLS) file")
    parser.add_argument(
        "--type",
        "-t", 
        choices=["scrt", "rdm", "rdp", "remmina"],
        default="scrt",
        help="Destination type: scrt=SecureCRT (default), rdm=Devolutions RDM, rdp=RDP files, remmina=Remmina profiles",
    )
    group1.add_argument(
        "--write",
        "-w",
        metavar="DESTINATION",
        # action="store_true",
        required=False,
        help="Write to destination file/folder. If not specified, write to 'export' subfolder as the source.",
    )
    group1.add_argument(
        "-p",
        "--print",
        action="store_true",
        required=False,
        help="Print to screen only (not supported for rdp and remmina export).",
    )
    group2.add_argument(
        "-q",
        "--quiet",
        action="store_true",
        required=False,
        help="Quiet output.",
    )
    group2.add_argument(
        "-v",
        "--verbose",
        dest="verbose",
        action="count",
        required=False,
        help="Verbose output. (use: -v, -vv)",
    )
    group2.add_argument(        
        "--version", action="version",
        version = f"{parser.prog} version  {get_version()}"
    )
    arg = parser.parse_args()

    return arg


def parse_reader_args():
    """Parse arguments for import sessions

    Returns:
        object: Arguments
    """
    parser = argparse.ArgumentParser(
        description="Read SecureCRT XML and export them to Excel file (write to destination)."
    )
    group1 = parser.add_mutually_exclusive_group()
    group2 = parser.add_mutually_exclusive_group()

    parser.add_argument(
        "--config",
        type=str,
        metavar="CONFIG",
        help="Configuration settings file (default=config.yaml)",
        default="config.yaml",
    )
    parser.add_argument(
        "source",
        nargs="+",
        type=str,
        help="Source file(s): SecureCRT XML.",
    )
    parser.add_argument(
        "--type",
        choices=["scrt"],
        default="scrt",
        help="Source type: scrt=SecureCRT XML (default).",
    )

    group1.add_argument(
        "-w",
        "--write",
        metavar="DESTINATION",
        dest="write",
        required=False,
        help="Write to destination Excel (xlsx) file. If not defined, write to the 'export' subfolder.",
    )
    group2.add_argument(
        "-q",
        "--quiet",
        action="store_true",
        required=False,
        help="Quiet output.",
    )
    group2.add_argument(
        "-v",
        "--verbose",
        dest="verbose",
        action="count",
        required=False,
        help="Verbose output (use: -v, -vv).",
    )
    group2.add_argument(        
        "--version", action="version",
        version = f"{parser.prog} version  {get_version()}"
    ) 
    arg = parser.parse_args()

    return arg
