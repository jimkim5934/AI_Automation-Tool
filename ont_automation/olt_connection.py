"""OLT Connection — Netmiko session, reconnect, and profile discovery."""

import logging
import re
import sys
from typing import Dict

try:
    from netmiko import ConnectHandler
except ImportError:
    logging.error("Netmiko library is not installed. Please run: pip install netmiko")
    sys.exit(1)


class OLTConnectionMixin:
    """Mixin providing Nokia OLT connect/disconnect and CLI execution."""

    def __init__(self, ip: str, user: str, pw: str, proto: str):
        self.proto = proto
        self.device = {"host": ip, "username": user, "password": pw, "global_delay_factor": 2}
        self.connection = None
        self.olt_profile = "Nokia_Default"
        self.vendor = "Unknown"
        self.model = "Unknown"
        self.sw_version = "Unknown"

    def connect(self) -> bool:
        types = (
            ["nokia_sros_telnet", "alcatel_sros_telnet"]
            if self.proto == "telnet"
            else ["nokia_sros", "alcatel_sros"]
        )
        for dt in types:
            self.device["device_type"] = dt
            try:
                self.connection = ConnectHandler(**self.device)
                logging.info(f"Connected to OLT using device_type: {dt}")
                self._discover_profile()
                return True
            except Exception:
                continue
        return False

    def disconnect(self):
        if self.connection:
            try:
                self.connection.disconnect()
                logging.info("Gracefully disconnected from OLT.")
            except Exception:
                pass

    def _discover_profile(self):
        out = self.send_command("show software-mngt version ansi", 5)

        m = re.search(
            r"(V\d+\.[A-Za-z0-9\.\-_]+|R\d+\.[A-Za-z0-9\.\-_]+|\d{1,3}\.\d{1,3}\.\S+)",
            out,
            re.IGNORECASE,
        )
        if m:
            self.sw_version = m.group(1)
        else:
            clean_out = re.sub(r"show software-mngt version ansi", "", out, flags=re.IGNORECASE)
            words = clean_out.split()
            self.sw_version = words[0] if words else "Unknown Version"

        self.olt_profile = f"Nokia_OS_{self.sw_version}"

        if self.device.get("host") == "10.100.0.7":
            self.vendor = "Nokia"
            self.model = "ISAM7360-FWLT-B"
        else:
            if "Nokia" in out or "nokia" in out.lower():
                self.vendor = "Nokia"
            elif "Alcatel" in out or "alcatel" in out.lower():
                self.vendor = "Alcatel-Lucent"
            else:
                self.vendor = "Unknown"

            sys_out = self.send_command("show equipment system", 5)
            model_m = re.search(r"(FX-\d+|7360|7342|7750\s+SR\S*)", sys_out, re.IGNORECASE)
            self.model = model_m.group(1) if model_m else "ISAM (Generic)"

    def send_command(self, cmd: str, timeout: int = 15, retry: bool = True) -> str:
        if self.connection:
            try:
                if not self.connection.is_alive():
                    raise Exception("Connection dead")
                return self.connection.send_command_timing(cmd, read_timeout=timeout)
            except Exception as e:
                if retry:
                    logging.warning(
                        "Connection dropped during command execution. Attempting Auto-Reconnect..."
                    )
                    self.disconnect()
                    if self.connect():
                        logging.info("Auto-Reconnect successful. Retrying command...")
                        try:
                            return self.connection.send_command_timing(cmd, read_timeout=timeout)
                        except Exception as e2:
                            return f"Error after reconnect: {e2}"
                return f"Error: {e}"
        else:
            if retry and self.connect():
                return self.send_command(cmd, timeout=timeout, retry=False)
        return "No Connection"
