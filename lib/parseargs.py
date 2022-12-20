"""Parse arguments library"""

import argparse

# import os.path
# import logging

# from pathlib import Path
# from ruamel.yaml import YAML


def parse_maker_args():
    """Parse arguments for export sessions

    Returns:
        object: Arguments
    """
    parser = argparse.ArgumentParser(
        description="Read Excel file (source) and generate sessions XML file for [SecureCRT|Devolutions]."
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
        choices=["scrt", "rdm"],
        default="scrt",
        help="Destination type: scrt=SecureCRT (default), rdm=DevolutionsRDM",
    )
    group1.add_argument(
        "--write",
        "-w",
        metavar="DESTINATION",
        # action="store_true",
        required=False,
        help="Write to file. If not specified, write to 'export' subfolder as the source.",
    )
    group1.add_argument(
        "-p",
        "--print",
        action="store_true",
        required=False,
        help="Print to screen only (don't write it to the file).",
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
    arg = parser.parse_args()

    return arg


def parse_reader_args():
    """Parse arguments for import sessions

    Returns:
        object: Arguments
    """
    parser = argparse.ArgumentParser(
        description="Read SecureCRT sessions XML file (source) and export it to Excel file (write to destination)."
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
    parser.add_argument("source", type=str, help="SecureCRT sessions XML file (export from SecureCRT).")

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
    arg = parser.parse_args()

    return arg
