"""Parse settings file library"""

import logging
import os.path


from pathlib import Path
from ruamel.yaml import YAML


def set_config_file(config_file, default_config_file):
    """Return correct 'config_file'. If not existst, return 'default_config_file'."""

    config_file = config_file.strip()
    if os.path.isfile(config_file):
        return config_file

    logging.info(
        "Config file path '%s' is not valid. Reading '%s'.",
        config_file,
        default_config_file,
    )
    return default_config_file


def read_config_file(config_file: str):
    """Read configuration file and return content as nested object.

    Attributes:
        setting_file (str): Path to setting file

    Returns:
        settings_content as nested object
    """
    yaml = YAML(typ="safe")
    path = Path(config_file)
    with open(path, "r", encoding="utf-8") as file:
        try:
            config_data = yaml.load(file)
            logging.info("Loading config file '%s'.", config_file)
            return config_data
        except OSError:
            logging.error("Unable to read configuration file '%s'.", config_file)

    return False
