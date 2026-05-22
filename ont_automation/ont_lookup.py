"""ONT 조회 — discovery, active scan, and port/ID helpers."""

import logging
import re
from typing import Dict, List, Optional


class ONTLookupMixin:
    """Mixin for querying ONT inventory on the OLT."""

    def get_unprovisioned_onts(self, discovery_cmd: str) -> List[Dict[str, str]]:
        if self.connection:
            output = self.send_command(discovery_cmd)
            matches = re.findall(r"(\d+(?:/\d+)+)\s+([A-Za-z]{4}[A-Fa-f0-9]{8})", output)
            if matches:
                onts = [{"chanpair": m[0], "serial": m[1]} for m in matches]
                unique_onts = {item["serial"]: item for item in onts}.values()
                return list(unique_onts)
            serials = re.findall(r"([A-Za-z]{4}[A-Fa-f0-9]{8})", output)
            return [{"chanpair": "unknown", "serial": s} for s in set(serials)]
        return []

    def scan_active_onts(self, channel_prefix: str = "1/1/1") -> List[Dict[str, str]]:
        """Scan channel-pairs 1..8 for ONTs in UP state."""
        items = []
        if not self.connection:
            return items
        for i in range(1, 9):
            chanpair = f"{channel_prefix}/{i}"
            out = self.send_command(f"show equipment ont status channel-pair {chanpair}", timeout=5)
            matches = re.findall(
                r"(\d+(?:/\d+)+)\s+([a-zA-Z0-9:-]+(?:/\d+)+)\s+([A-Za-z]{4}:?[A-Fa-f0-9]{8})\s+(\S+)\s+(\S+)",
                out,
            )
            for m in matches:
                if m[4].lower() == "up":
                    items.append({"chanpair": m[0], "port": m[1], "serial": m[2].replace(":", "")})
        return items

    def suggest_pon_from_chanpair(self, chanpair: str) -> Optional[str]:
        parts = chanpair.split("/")
        if len(parts) >= 4:
            return f"ng2:{parts[3]}/{parts[2]}"
        return None

    def find_available_ont_id(self, chanpair: str) -> Optional[str]:
        """Return the first unused ONT ID (1-128) on the given channel-pair."""
        if not self.connection:
            return None
        logging.info(f"Scanning channel {chanpair} for available ID...")
        out = self.send_command(f"show equipment ont status channel-pair {chanpair}", timeout=5)
        used_ids = set()
        matches = re.findall(r"([a-zA-Z0-9:-]+(?:/\d+)+)\s+([A-Za-z]{4}:?[A-Fa-f0-9]{8})", out)
        for port_str, _ in matches:
            last_digit = port_str.split("/")[-1]
            if last_digit.isdigit():
                used_ids.add(int(last_digit))
        for i in range(1, 129):
            if i not in used_ids:
                return str(i)
        return None
