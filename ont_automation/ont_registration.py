"""ONT Registration — register and verify ONT on the OLT."""

import logging
import time
from typing import Dict

from ont_automation.database import DatabaseManager
from ont_automation.gui_helpers import gui_ask


class ONTRegistrationMixin:
    """Mixin for ONT registration and verification CLI sequences."""

    def format_serial(self, serial_raw: str) -> str:
        vendor_id = serial_raw[:4]
        sn_rem = serial_raw[4:]
        return f"{vendor_id}:{sn_rem}"

    def register_ont(self, config: Dict, db: DatabaseManager) -> bool:
        port_full = f"{config['pon']}/{config['ont_id']}"
        serial_raw = config["ont_serial"]
        serial_formatted = self.format_serial(serial_raw)

        reg_key = f"Registration_{config['ont_type'].upper()}"
        learned_cmds = db.get_learned_command(self.olt_profile, reg_key)

        if learned_cmds and "admin-state up" not in learned_cmds.lower():
            logging.warning(
                "Learned registration command is missing critical 'admin-state up' step. Reverting to default sequence."
            )
            learned_cmds = None

        if learned_cmds:
            cmd_templates = [c.strip() for c in learned_cmds.split(";") if c.strip()]
        else:
            if config["ont_type"] == "sfu":
                cmd_templates = [
                    "configure equipment ont interface {port} sw-ver-pland disabled sernum {serial_formatted} fec-up enable sw-dnload-version disabled pref-channel-pair {chanpair}",
                    "configure equipment ont interface {port} admin-state up",
                    "configure equipment ont slot {port}/1 planned-card-type ethernet plndnumdataports {lan_ports} plndnumvoiceports 0",
                    "configure interface port uni:{port}/1/1 admin-up",
                ]
            else:
                cmd_templates = [
                    "configure equipment ont interface {port} sw-ver-pland disabled sernum {serial_formatted} fec-up enable sw-dnload-version disabled pref-channel-pair {chanpair}",
                    "configure equipment ont interface {port} admin-state up",
                    "configure equipment ont slot {port}/14 planned-card-type veip plndnumdataports {lan_ports} plndnumvoiceports 0",
                ]

        logging.info(f"--- [STEP 1] Initiating ONT Registration for {serial_raw} on {port_full} ---")
        error_kws = [
            "invalid command",
            "invalid token",
            "unknown command",
            "bad parameter",
            "incomplete command",
        ]

        while True:
            success = True
            for template in cmd_templates:
                cmd = template.format(
                    serial_formatted=serial_formatted,
                    port=port_full,
                    chanpair=config.get("chanpair", ""),
                    lan_ports=config.get("lan_ports", "1"),
                )

                out = self.send_command(cmd, timeout=10)
                out_lower = out.lower()
                if (any(kw in out_lower for kw in error_kws) or "^" in out) and "pattern not detected" not in out_lower:
                    print(f"\n[ REGISTRATION FAILED ] Command rejected: {cmd}")
                    self._cleanup_failed_provisioning(port_full)

                    new_cmds = gui_ask(
                        "Registration failed. Enter corrective CLI sequence (use ';' for multiple):"
                    )
                    if new_cmds.lower() == "exit" or not new_cmds:
                        return False

                    db.save_learned_command(self.olt_profile, reg_key, new_cmds)
                    cmd_templates = [c.strip() for c in new_cmds.split(";") if c.strip()]
                    success = False
                    break

            if success:
                return True

    def verify_registration(self, config: Dict, db: DatabaseManager) -> bool:
        port_full = f"{config['pon']}/{config['ont_id']}"
        chanpair = config.get("chanpair", "")
        serial_formatted = self.format_serial(config["ont_serial"])

        cmd_template = db.get_learned_command(self.olt_profile, "Reg_Verify_Cmd")
        default_verify_cmd = "show equipment ont status channel-pair {chanpair}"

        if not cmd_template or ("{chanpair}" not in cmd_template and "{port}" not in cmd_template):
            db.save_learned_command(self.olt_profile, "Reg_Verify_Cmd", default_verify_cmd)
            cmd_template = default_verify_cmd

        logging.info(f"--- [STEP 2] Verifying ONT Registration on {port_full} ---")

        while True:
            cmd = cmd_template.format(port=port_full, chanpair=chanpair)
            success = False

            for _attempt in range(1, 6):
                last_out = self.send_command(cmd, timeout=5)
                match_line = [
                    line
                    for line in last_out.split("\n")
                    if serial_formatted.lower() in line.lower() and port_full.lower() in line.lower()
                ]
                if match_line:
                    tokens = match_line[0].strip().lower().split()
                    if "up" in tokens:
                        success = True
                        break
                time.sleep(5)

            if success:
                return True

            print(
                f"\n[ VERIFICATION TIMEOUT ] Could not find oper status 'UP' for serial {serial_formatted} on port {port_full}."
            )
            choice = gui_ask(
                "Verification Failed. Options: [1] Retry [2] Change Cmd [3] Force Pass [4] Abort (Enter number)"
            )
            if choice == "1":
                continue
            elif choice == "2":
                new_cmd = gui_ask("Enter new verification command template:")
                if new_cmd and new_cmd.lower() != "exit":
                    cmd_template = new_cmd
                    db.save_learned_command(self.olt_profile, "Reg_Verify_Cmd", cmd_template)
            elif choice == "3":
                return True
            else:
                return False
