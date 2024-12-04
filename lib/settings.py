"""Parse settings file library"""

import logging
import os.path

from pathlib import Path
from ruamel.yaml import YAML


class Settings:
    """Settings class to read configuration file and return settings."""

    def __init__(self, config_file: str):

        # settings content
        self.app_config: dict = {}

        # configuration file path
        self.config_file: str | None = None
        self.set_config_file(config_file)

        # session defaults
        self.rdm_session_defaults = {}

    def _read_config_file(self):
        """
        Reads the configuration file specified by `self.config_file` and loads
        its contents into `self.settings`.

        Returns:
            bool: True if the configuration file was successfully read and loaded, False otherwise.

        Logs:
            - Info: When the configuration file is successfully loaded.
            - Error: When there is an issue reading the configuration file.
        """

        if self.config_file is None:
            return False

        yaml = YAML(typ="safe")
        path = Path(self.config_file)
        with open(path, "r", encoding="utf-8") as file:
            try:
                self.app_config = yaml.load(file)
                logging.info("Loading config file '%s'.", self.config_file)
                return True

            except OSError:
                logging.error(
                    "Unable to read configuration file '%s'.", self.config_file
                )
                return False

    def read_session_defaults(self, type="scrt"):
        
        if type == "rdm":
            self._read_session_defaults_rdm()
            
    def _read_session_defaults_rdm(self):

        if self.app_config is None:
            raise ValueError("App config is not valid.")

        path = Path(self.app_config["rdm"]["session_defaults"])
        yaml = YAML(typ="safe")
        
        data = {}

        if not path.is_dir():
            raise ValueError(f"The path '{path}' is not a valid directory.")

        # Read and merge all YAML files in the folder
        for yaml_file in path.glob('*.yaml'):
            with open(yaml_file, "r", encoding="utf-8") as file:
                try:
                    data = yaml.load(file) or {}
                    self.rdm_session_defaults.update(data)

                except Exception as e:
                    print(f"Error reading {yaml_file}: {e}")

    def set_config_file(self, config_file):
        """Set configuration file path.

        Return:
            False - in case of non-existing file
            True - in case of existing file

        """

        if os.path.isfile(config_file.strip()):
            self.config_file = config_file.strip()
            return self._read_config_file()

        logging.info("Config file path '%s' is not valid.", config_file)
        return False
