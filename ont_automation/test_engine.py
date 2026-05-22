import json
import logging
import re
import time
from datetime import datetime
from typing import Dict

from ont_automation.database import DatabaseManager
from ont_automation.gui_helpers import gui_ask
from ont_automation.nokia_olt import NokiaOLTConnector
from ont_automation.throughput_test import run_throughput_test


class TestAutomationEngine:
    def __init__(self, db: DatabaseManager, olt: NokiaOLTConnector, config: Dict):
        self.db = db
        self.olt = olt
        self.config = config
        self.serial = config["ont_serial"]
        self.pon = config["pon"]
        self.ont_id = config["ont_id"]
        self.chanpair = config.get("chanpair", "")
        self.port_full = f"{self.pon}/{self.ont_id}"

        vendor_id = self.serial[:4]
        sn_rem = self.serial[4:]
        self.serial_formatted = f"{vendor_id}:{sn_rem}"

    def execute_tests(self) -> bool:
        scenarios = self.db.load_test_cases(self.serial)
        retry_required = False

        for sn in scenarios:
            logging.info(f"Executing [{sn['order']}] {sn['type']}")
            if "wait" in sn["parameters"] and sn["parameters"]["wait"] > 0:
                logging.info(f"Waiting {sn['parameters']['wait']} seconds...")
                time.sleep(sn["parameters"]["wait"])

            res = ""

            if sn["type"] == "Speed_Test":
                try:
                    res, _, _ = run_throughput_test(self.config)
                except FileNotFoundError:
                    res = "[SPEED_TEST_FAILED] iperf3 executable not found in system PATH."
                    logging.error(res)
                except Exception as e:
                    res = f"[SPEED_TEST_FAILED] Unexpected error: {e}"
                    logging.error(res)

            elif sn["type"] == "Reboot_Test":
                logging.info("Sending Reboot Command to OLT...")
                res_initial = self.olt.send_command(sn["command"])

                if "^" in res_initial or "invalid token" in res_initial.lower() or "error" in res_initial.lower():
                    res = res_initial
                else:
                    logging.info(
                        "Reboot command sent. Forcing OLT disconnection to allow PC NIC failover (Wait 5s)..."
                    )
                    time.sleep(5)
                    self.olt.disconnect()

                    logging.info("Attempting to reconnect to OLT and waiting for ONT recovery (Max 5 mins)...")
                    start_wait_t = time.time()
                    reconnected = False

                    while time.time() - start_wait_t < 300:
                        if self.olt.connect():
                            reconnected = True
                            logging.info("Successfully reconnected to OLT!")
                            break
                        time.sleep(5)

                    if not reconnected:
                        res = (
                            f"[REBOOT_FAILED] Could not reconnect to OLT after 5 minutes.\n"
                            f"[Reboot Trigger]\n{res_initial}"
                        )
                        logging.error("OLT Reconnection timeout reached!")
                    else:
                        alarm_cmd_template = self.db.get_learned_command(
                            self.olt.olt_profile, "Alarm_Check_Cmd"
                        )

                        if not alarm_cmd_template or "show equipment ont alarm" in alarm_cmd_template:
                            alarm_cmd_template = "show alarm log"
                            self.db.save_learned_command(
                                self.olt.olt_profile, "Alarm_Check_Cmd", alarm_cmd_template
                            )

                        alarm_cmd = alarm_cmd_template.format(
                            port=self.port_full, chanpair=self.chanpair
                        )

                        reboot_success = False
                        reboot_time_sec = 0.0
                        last_alarm_out = ""

                        for _attempt in range(60):
                            last_alarm_out = self.olt.send_command(alarm_cmd, timeout=10)

                            if "^" in last_alarm_out or "invalid token" in last_alarm_out.lower():
                                status_out = self.olt.send_command(
                                    f"show equipment ont status {self.port_full}"
                                )
                                if "up" in status_out.lower() and "admin-up" in status_out.lower():
                                    reboot_success = True
                                    reboot_time_sec = time.time() - start_wait_t
                                    last_alarm_out = (
                                        "[AI-Fallback] Verified recovery via 'show equipment ont status' "
                                        "due to CLI restrictions."
                                    )
                                    break

                            occurred_time = None
                            cleared_time = None
                            is_cleared = False

                            for line in last_alarm_out.splitlines():
                                if (
                                    "ONT is inactive" in line
                                    or "loss of signal" in line
                                    or "loss of frame" in line
                                ) and (self.port_full in line or self.chanpair in line):
                                    match = re.search(
                                        r"(\d{2}/\d{2}/\d{2}\s+\d{2}:\d{2}:\d{2})\s+.*?(occurred|cleared)",
                                        line,
                                        re.IGNORECASE,
                                    )
                                    if match:
                                        timestamp_str, action = match.groups()
                                        try:
                                            dt_obj = datetime.strptime(
                                                timestamp_str, "%y/%m/%d %H:%M:%S"
                                            )
                                            if action.lower() == "occurred":
                                                occurred_time = dt_obj
                                                is_cleared = False
                                            elif action.lower() == "cleared":
                                                cleared_time = dt_obj
                                                is_cleared = True
                                        except ValueError:
                                            continue

                            if occurred_time and is_cleared:
                                reboot_time_sec = (cleared_time - occurred_time).total_seconds()
                                reboot_success = True
                                break
                            elif is_cleared and not occurred_time:
                                reboot_time_sec = time.time() - start_wait_t
                                reboot_success = True
                                break

                            time.sleep(5)

                        if reboot_success:
                            res = (
                                f"[REBOOT_SUCCESS] ONT recovery confirmed via Alarm Log. "
                                f"True Reboot Duration: {reboot_time_sec:.1f} seconds.\n"
                                f"[Reboot Trigger]\n{res_initial}\n[Final Alarm Log]\n{last_alarm_out}"
                            )
                            logging.info(
                                f"ONT Rebooted and Recovery Detected! Actual Reboot Time: {reboot_time_sec:.1f} seconds."
                            )
                        else:
                            res = (
                                f"[REBOOT_FAILED] ONT did not recover within timeout (Alarm 'cleared' not found).\n"
                                f"[Reboot Trigger]\n{res_initial}\n[Last Alarm Log]\n{last_alarm_out}"
                            )
                            logging.error("Reboot timeout reached!")

            else:
                if ";" in sn["command"]:
                    res_lines = []
                    for c in sn["command"].split(";"):
                        c = c.strip()
                        if c:
                            logging.info(f"Executing: {c}")
                            out = self.olt.send_command(c)
                            res_lines.append(out)
                    res = "\n".join(res_lines)
                else:
                    res = self.olt.send_command(sn["command"])

                if sn["type"] == "Software_Info":
                    match = re.search(r"sw-ver-act\s*:\s*(\S+)", res, re.IGNORECASE)
                    if match:
                        self.config["sw_version"] = match.group(1).upper()

            if self._verify(sn["type"], res):
                logging.info(f"[PASS] {sn['type']}")
                self.db.update_test_status(sn["id"], "PASS", res)
            else:
                if sn["type"] == "Speed_Test":
                    logging.error(f"[{sn['type']}] FAILED: Did not meet the target criteria.")
                    self.db.update_test_status(sn["id"], "FAIL", res)
                    continue

                action = self._handle_failure(sn, res)
                if not action:
                    self.db.update_test_status(sn["id"], "FAIL", res)
                    return False
                if action == "SKIP":
                    logging.info(f"[SKIPPED] {sn['type']}")
                    self.db.update_test_status(sn["id"], "N/A", res)
                    continue

                retry_required = True
                break

        if retry_required:
            logging.warning("Test sequence aborted. Initiating re-test in the next cycle.")
            return False
        return True

    def _verify(self, t_type: str, res: str) -> bool:
        if not res:
            if t_type in ["Cleanup"]:
                return True
            return False

        res_lower = res.lower()

        if t_type == "Reboot_Test":
            return "[REBOOT_SUCCESS]" in res
        if t_type == "Speed_Test":
            return "[SPEED_TEST_SUCCESS]" in res

        error_kws = [
            "invalid command",
            "invalid token",
            "unknown command",
            "bad parameter",
            "incomplete command",
        ]
        if any(kw in res_lower for kw in error_kws) or "^" in res:
            return False

        if t_type == "ONT_Discovery_Check":
            serial_clean = self.serial.replace(":", "").lower()
            return serial_clean in res_lower.replace(":", "")

        if t_type == "Registration_Check":
            serial_clean = self.serial.replace(":", "").lower()
            return "pref-ranged" in res_lower and serial_clean in res_lower.replace(":", "")

        if t_type in ["Optics_Signal_Check", "Optics_Temp_Check", "Optics_Voltage_Laser_Check"]:
            success_found = False
            for line in res.split("\n"):
                if self.port_full.lower() in line.lower() or self.port_full.split(":")[-1] in line:
                    tokens = line.strip().split()
                    port_idx = -1
                    for i, t in enumerate(tokens):
                        if self.port_full.lower() in t.lower() or self.port_full.split(":")[-1] in t:
                            port_idx = i
                            break

                    if port_idx != -1 and len(tokens) >= port_idx + 7:
                        rx_sig = tokens[port_idx + 1]
                        tx_sig = tokens[port_idx + 2]
                        temp = tokens[port_idx + 3]
                        voltage = tokens[port_idx + 4]
                        laser = tokens[port_idx + 5]
                        olt_rx = tokens[port_idx + 6]

                        def is_num(val):
                            val_clean = val.lower()
                            if val_clean in [
                                "invalid",
                                "not-appl",
                                "n/a",
                                "none",
                                "-",
                                "inf",
                                "-inf",
                            ]:
                                return False
                            try:
                                float(val)
                                return True
                            except ValueError:
                                return False

                        if t_type == "Optics_Signal_Check":
                            success_found = is_num(rx_sig) and is_num(tx_sig) and is_num(olt_rx)
                        elif t_type == "Optics_Temp_Check":
                            success_found = is_num(temp)
                        elif t_type == "Optics_Voltage_Laser_Check":
                            success_found = is_num(voltage) and is_num(laser)

                        if success_found:
                            break
            return success_found

        if t_type == "Software_Info":
            match = re.search(r"sw-ver-act\s*:\s*(\S+)", res_lower)
            if match:
                val = match.group(1)
                return val not in ["sw-ver-psv", "vendor-id", "unknown"]
            return False

        if t_type == "UNI_Status":
            has_link = "link-status" in res_lower and "up" in res_lower
            has_speed = "config-indicator" in res_lower and any(
                s in res_lower for s in ["10g", "1g", "100m", "10m"]
            )
            return has_link and has_speed

        if t_type == "MAC_Status":
            return bool(re.search(r"([0-9A-F]{2}:){5}[0-9A-F]{2}", res, re.I))

        return True

    def _handle_failure(self, sn: Dict, res: str):
        print(f"\n[FAILURE DETECTED] {sn['type']}")

        res_lower = res.lower() if res else ""
        error_kws = [
            "invalid command",
            "invalid token",
            "unknown command",
            "bad parameter",
            "incomplete command",
        ]
        is_syntax_error = any(kw in res_lower for kw in error_kws) or "^" in res

        if is_syntax_error and sn["type"] != "Speed_Test":
            print("\n[ AI Auto-Detection ] 'Invalid CLI' detected.")
            tmp = gui_ask(f"Enter correct CLI template for {sn['type']} (or type 'exit'):")
            if tmp.lower() == "exit":
                return False
            if tmp and tmp.lower() != "exit":
                self.db.save_learned_command(self.olt.olt_profile, sn["type"], tmp)
                self.db.update_test_case(
                    sn["id"],
                    tmp.format(
                        serial=self.serial,
                        pon=self.pon,
                        ont_id=self.ont_id,
                        port=self.port_full,
                    ),
                    json.dumps(sn["parameters"]),
                )
                return True

        choice = gui_ask(
            f"Error in {sn['type']}. Select Option: [1] Update Cmd [2] Increase Wait [3] Force Pass [4] Abort"
        )
        if choice == "1":
            tmp = gui_ask(f"Enter new template for {sn['type']}:")
            if tmp and tmp.lower() != "exit":
                self.db.save_learned_command(self.olt.olt_profile, sn["type"], tmp)
                self.db.update_test_case(
                    sn["id"],
                    tmp.format(
                        serial=self.serial,
                        pon=self.pon,
                        ont_id=self.ont_id,
                        port=self.port_full,
                    ),
                    json.dumps(sn["parameters"]),
                )
            return True
        elif choice == "2":
            sn["parameters"]["wait"] = sn["parameters"].get("wait", 0) + 10
            self.db.update_test_case(sn["id"], sn["command"], json.dumps(sn["parameters"]))
            return True
        elif choice == "3":
            logging.info(f"Force passing {sn['type']} step.")
            return "SKIP"
        return False
