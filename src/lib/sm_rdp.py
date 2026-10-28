"""RDP file generator (Excel -> .rdp)"""

import logging
import os
from pathlib import Path

from .sm_class import SessionMaker


class SMRdp(SessionMaker):
    """SessionMaker - RDP file generator class"""

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

    def _normalize_folder(self, folder: str) -> tuple[str, list[str]]:
        folder = folder.strip().replace("\\", "/")
        folder = folder.strip("/")
        if folder == "":
            return "", []

        parts = [part for part in folder.split("/") if part not in ("", ".", "..")]
        return "/".join(parts), parts

    def _safe_filename(self, session_name: str, session_type: str) -> str:
        session_name = session_name.strip()

        for char in ['/', '\\', ':', '*', '?', '"', '<', '>', '|']:
            session_name = session_name.replace(char, "_")

        session_suffix = session_type.strip().lower()
        if session_suffix == "":
            return f"{session_name}.rdp"

        return f"{session_name}-{session_suffix}.rdp"

    def _build_server(self, hostname: str, port: str) -> str:
        host = hostname.strip()
        if host == "":
            return ""

        clean_port = port.strip()
        if clean_port == "":
            return host

        return f"{host}:{clean_port}"

    def _read_rdp_template(self) -> list[list[str]]:
        """Read RDP template file as list of [key, type, value] items."""
        template_file = self._settings.get("rdp", {}).get("template", {}).get("rdp", "")
        items: list[list[str]] = []
        if template_file == "":
            return items

        try:
            with open(template_file, "r", encoding="utf-8-sig") as rdp_file:
                for line in rdp_file:
                    line = line.rstrip("\r\n")
                    parts = line.split(":", 2)
                    if len(parts) == 3 and parts[0].strip() != "":
                        items.append([parts[0].strip(), parts[1].strip(), parts[2]])
        except OSError:
            logging.warning("Unable to read RDP template '%s'. Using empty template.", template_file)

        return items

    def _build_rdp_content(
        self,
        server: str,
        username: str,
        alternate_shell: str,
    ) -> str:
        """Build .rdp content from template, overridden by values from Excel."""
        values = {
            "full address": server,
            "username": username,
            "alternate shell": alternate_shell,
        }

        items = self._read_rdp_template()
        used_keys: set[str] = set()
        for item in items:
            key = item[0].lower()
            if key in values:
                used_keys.add(key)
                # empty Excel value keeps default from template (except for mandatory address)
                if values[key] != "" or key == "full address":
                    item[1] = "s"
                    item[2] = values[key]

        for key, value in values.items():
            if key not in used_keys and value != "":
                items.append([key, "s", value])

        return "\r\n".join(f"{key}:{kind}:{value}" for key, kind, value in items) + "\r\n"

    def write_rdp_files(self, export_dir: str):
        profile_count = 0
        skipped_count = 0
        written_files: set[Path] = set()

        if not self._sessions_dict or "session" not in self._sessions_dict:
            return profile_count, skipped_count

        for idx, session_name in enumerate(self._sessions_dict["session"]):
            if session_name.strip() == "":
                # Skip sessions with empty names.
                continue

            session_type = self._sessions_dict["type"][idx]
            if session_type.strip().lower() != "rdp":
                # only RDP sessions are exported
                continue

            folder = self._sessions_dict["folder"][idx]
            _, folder_parts = self._normalize_folder(folder)

            hostname = self._sessions_dict["hostname"][idx]
            port = self._sessions_dict["port"][idx]
            username = self._sessions_dict["username"][idx]
            rdp_alternate = self._sessions_dict["rdp_alternate"][idx]
            server = self._build_server(hostname, port)
            if server == "":
                skipped_count += 1
                logging.warning(
                    "Session '%s' has empty hostname. Skipping.",
                    session_name,
                )
                continue

            file_name = self._safe_filename(session_name, session_type)
            destination_dir = Path(export_dir, *folder_parts)
            destination_dir.mkdir(parents=True, exist_ok=True)

            destination_file = destination_dir / file_name
            profile_content = self._build_rdp_content(
                server=server,
                username=username.strip(),
                alternate_shell=rdp_alternate.strip(),
            )

            if destination_file in written_files:
                counter = 2
                while True:
                    destination_file = destination_dir / f"{destination_file.stem.rsplit('~', 1)[0]}~{counter}{destination_file.suffix}"
                    if destination_file not in written_files:
                        break
                    counter += 1
                logging.warning(
                    "Duplicate session '%s' in '%s'. Writing as '%s'.",
                    session_name,
                    destination_dir,
                    destination_file.name,
                )
            elif os.path.exists(destination_file):
                logging.warning(
                    "Destination file '%s' exists. Overwriting.",
                    destination_file,
                )
            written_files.add(destination_file)

            with open(destination_file, "w", encoding="utf-8", newline="") as rdp_file:
                rdp_file.write(profile_content)

            profile_count += 1

        return profile_count, skipped_count
