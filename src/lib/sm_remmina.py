"""Remmina profile generator (Excel -> .remmina)"""

import logging
import os
from pathlib import Path

from .sm_class import SessionMaker


class SMRemmina(SessionMaker):
    """SessionMaker - Remmina profile generator class"""

    def __init__(
        self,
        settings: dict | None = None,
        excel_file: str | None = None,
        read_excel_file=False,
    ):
        super().__init__(
            settings=settings,
            excel_file=excel_file,
            read_excel_file=read_excel_file,
        )

    def _normalize_folder(self, folder: str) -> str:
        # folder = folder.strip().replace("\\", "/")
        parts = [part for part in folder.split("/") if part not in ("", ".", "..")]
        return "/".join(parts)

    def _safe_part(self, value: str) -> str:
        value = value.strip()
        for char in ["/", "\\", ":", "*", "?", '"', "<", ">", "|"]:
            value = value.replace(char, "_")
        return value

    def _build_filename(self, group: str, session_type: str, name: str, ip: str) -> str:
        """File name: <group>_<type>_<name>_<ip>.remmina (group: '/' replaced by '-')."""
        parts = [group.replace("/", "-"), session_type, name, ip]
        parts = [self._safe_part(part) for part in parts if part.strip() != ""]
        return "_".join(parts) + ".remmina"

    def _build_server(self, hostname: str, port: str) -> str:
        host = hostname.strip()
        if host == "":
            return ""

        clean_port = port.strip()
        if clean_port == "":
            return host

        return f"{host}:{clean_port}"

    def _read_remmina_template(self) -> list[list[str]]:
        """Read Remmina template file as list of [key, value] items."""
        template_file = self._settings.get("remmina", {}).get("template", {}).get("rdp", "")
        items: list[list[str]] = []
        if template_file == "":
            return items

        try:
            with open(template_file, "r", encoding="utf-8-sig") as remmina_file:
                for line in remmina_file:
                    line = line.rstrip("\r\n")
                    if line.strip() == "" or line.startswith("["):
                        continue
                    key, sep, value = line.partition("=")
                    if sep == "":
                        continue
                    items.append([key.strip(), value])
        except OSError:
            logging.warning("Unable to read Remmina template '%s'. Using empty template.", template_file)

        return items

    def _build_remmina_content(
        self,
        name: str,
        group: str,
        server: str,
        username: str,
        alternate_shell: str,
    ) -> str:
        """Build .remmina content from template, overridden by values from Excel."""
        values = {
            "name": name,
            "group": group,
            "server": server,
            "username": username,
            "exec": alternate_shell,
            "protocol": "RDP",
        }

        items = self._read_remmina_template()
        used_keys: set[str] = set()
        for item in items:
            if item[0] in values:
                used_keys.add(item[0])
                # empty Excel value keeps default from template (except for mandatory ones)
                if values[item[0]] != "" or item[0] in ("name", "server", "protocol"):
                    item[1] = values[item[0]]

        for key, value in values.items():
            if key not in used_keys and value != "":
                items.append([key, value])

        lines = ["[remmina]"] + [f"{key}={value}" for key, value in items]
        return "\n".join(lines) + "\n"

    def write_remmina_files(self, export_dir: str):
        profile_count = 0
        skipped_count = 0
        written_files: set[Path] = set()

        if not self._sessions_dict or "session" not in self._sessions_dict:
            return profile_count, skipped_count

        destination_dir = Path(export_dir)

        for idx, session_name in enumerate(self._sessions_dict["session"]):
            if session_name.strip() == "":
                # Skip sessions with empty names.
                continue

            session_type = self._sessions_dict["type"][idx].strip().lower()
            if session_type != "rdp":
                # only RDP sessions are exported
                continue

            hostname = self._sessions_dict["hostname"][idx].strip()
            server = self._build_server(hostname, self._sessions_dict["port"][idx])
            if server == "":
                skipped_count += 1
                logging.warning("Session '%s' has empty hostname. Skipping.", session_name)
                continue

            group = self._normalize_folder(self._sessions_dict["folder"][idx])
            profile_content = self._build_remmina_content(
                name=session_name.strip(),
                group=group,
                server=server,
                username=self._sessions_dict["username"][idx].strip(),
                alternate_shell=self._sessions_dict["rdp_alternate"][idx].strip(),
            )

            destination_dir.mkdir(parents=True, exist_ok=True)
            destination_file = destination_dir / self._build_filename(
                group, session_type, session_name, hostname
            )

            if destination_file in written_files:
                base_stem = destination_file.stem
                counter = 2
                while destination_file in written_files:
                    destination_file = destination_dir / f"{base_stem}~{counter}.remmina"
                    counter += 1
                logging.warning(
                    "Duplicate session '%s'. Writing as '%s'.",
                    session_name,
                    destination_file.name,
                )
            elif os.path.exists(destination_file):
                logging.warning("Destination file '%s' exists. Overwriting.", destination_file)
            written_files.add(destination_file)

            with open(destination_file, "w", encoding="utf-8") as remmina_file:
                remmina_file.write(profile_content)

            profile_count += 1

        return profile_count, skipped_count
