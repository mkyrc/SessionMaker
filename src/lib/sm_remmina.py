"""Remmina session generator"""

import configparser
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

    def _normalize_folder(self, folder: str) -> tuple[str, list[str]]:
        folder = folder.strip().replace("\\", "/")
        folder = folder.strip("/")
        if folder == "":
            return "", []

        parts = [part for part in folder.split("/") if part not in ("", ".", "..")]
        return "/".join(parts), parts

    def _safe_filename(self, session: str, session_type: str) -> str:
        session_name = session.strip()
        if session_name == "":
            session_name = "session"

        for char in ['/', '\\', ':', '*', '?', '"', '<', '>', '|']:
            session_name = session_name.replace(char, "_")

        session_suffix = session_type.strip().lower()
        if session_suffix == "":
            return f"{session_name}.remmina"

        return f"{session_name}-{session_suffix}.remmina"

    def _build_server(self, hostname: str, port: str) -> str:
        host = hostname.strip()
        if host == "":
            return ""

        clean_port = port.strip()
        if clean_port == "":
            return host

        return f"{host}:{clean_port}"

    def _session_type_to_protocol(self, session_type: str) -> str | None:
        mapping = {
            "ssh": "SSH",
            "rdp": "RDP",
        }
        return mapping.get(session_type.strip().lower())

    def _protocol_to_session_type(self, protocol: str) -> str | None:
        protocol_name = protocol.strip().lower()
        mapping = {
            "ssh": "ssh",
            "rdp": "rdp",
        }
        return mapping.get(protocol_name)

    def _split_server(self, server: str) -> tuple[str, str]:
        value = server.strip()
        if value == "":
            return "", ""

        if value.startswith("[") and "]" in value:
            end_bracket = value.find("]")
            host = value[1:end_bracket].strip()
            port = ""
            if len(value) > end_bracket + 1 and value[end_bracket + 1] == ":":
                port = value[end_bracket + 2 :].strip()
            return host, port

        if value.count(":") == 1:
            host, port = value.rsplit(":", 1)
            return host.strip(), port.strip()

        return value, ""

    def _read_profile_file(self, profile_file: str) -> dict | None:
        parser = configparser.RawConfigParser(interpolation=None)
        parser.optionxform = str

        read_files = parser.read(profile_file, encoding="utf-8")
        if not read_files or not parser.has_section("remmina"):
            return None

        get_value = lambda key: parser.get("remmina", key, fallback="").strip()
        protocol = get_value("protocol")
        session_type = self._protocol_to_session_type(protocol)
        if session_type is None:
            return None

        session_name = get_value("name")
        if session_name == "":
            session_name = Path(profile_file).stem

        group = get_value("group").replace("\\", "/").strip("/")
        server = get_value("server")
        hostname, port = self._split_server(server)

        return {
            "folder": group,
            "session": session_name,
            "type": session_type,
            "hostname": hostname,
            "port": port,
            "username": get_value("username"),
            "rdp_alternate": "",
        }

    def load_sessions_from_remmina_files(self, profile_files: list[str]) -> tuple[int, int]:
        self.set_sessions_dict()
        imported_count = 0
        skipped_count = 0

        for profile_file in profile_files:
            try:
                profile_data = self._read_profile_file(profile_file)
            except (configparser.Error, OSError):
                logging.warning("Unable to parse Remmina profile '%s'. Skipping.", profile_file)
                skipped_count += 1
                continue

            if profile_data is None:
                skipped_count += 1
                logging.warning(
                    "Remmina profile '%s' is missing required values or has unsupported protocol. Skipping.",
                    profile_file,
                )
                continue

            if profile_data["hostname"] == "":
                skipped_count += 1
                logging.warning("Remmina profile '%s' has empty server. Skipping.", profile_file)
                continue

            for key in self._sessions_dict:
                self._sessions_dict[key].append(profile_data.get(key, ""))

            imported_count += 1

        return imported_count, skipped_count

    def write_excel(self, excel_file: str):
        self._excel_obj.write_excel_book(
            excel_file=excel_file,
            sessions_dict=self._sessions_dict,
            rdm_credentials_dict={},
            scrt_credentials_dict={},
            scrt_firewalls_dict={},
        )

    def _build_profile_content(
        self,
        name: str,
        group: str,
        protocol: str,
        server: str,
        username: str,
    ) -> str:
        lines = ["[remmina]"]
        lines.append(f"name={name}")
        if group != "":
            lines.append(f"group={group}")
        lines.append(f"protocol={protocol}")
        lines.append(f"server={server}")
        if username != "":
            lines.append(f"username={username}")

        return "\n".join(lines) + "\n"

    def write_remmina_profiles(self, export_dir: str):
        profile_count = 0
        skipped_count = 0
        written_files: set[Path] = set()

        if not self._sessions_dict or "session" not in self._sessions_dict:
            return profile_count, skipped_count

        for idx, session_name in enumerate(self._sessions_dict["session"]):
            session_type = self._sessions_dict["type"][idx]
            protocol = self._session_type_to_protocol(session_type)
            if protocol is None:
                skipped_count += 1
                logging.warning(
                    "Session '%s' has unsupported type '%s' for Remmina. Skipping.",
                    session_name,
                    session_type,
                )
                continue

            folder = self._sessions_dict["folder"][idx]
            group, folder_parts = self._normalize_folder(folder)

            hostname = self._sessions_dict["hostname"][idx]
            port = self._sessions_dict["port"][idx]
            username = self._sessions_dict["username"][idx]
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
            profile_content = self._build_profile_content(
                name=session_name,
                group=group,
                protocol=protocol,
                server=server,
                username=username.strip(),
            )

            if destination_file in written_files:
                counter = 2
                while True:
                    destination_file = destination_dir / f"{destination_file.stem.rsplit('~', 1)[0]}~{counter}.remmina"
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

            with open(destination_file, "w", encoding="utf-8") as remmina_file:
                remmina_file.write(profile_content)

            profile_count += 1

        return profile_count, skipped_count
