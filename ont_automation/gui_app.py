import json
import logging
import os
import re
import socket
import subprocess
import sys
import threading
import tkinter as tk
from tkinter import ttk, scrolledtext, messagebox

from ont_automation import gui_helpers
from ont_automation.config import CONFIG_FILE
from ont_automation.database import DatabaseManager
from ont_automation.excel_report import generate_professional_excel_report
from ont_automation.gui_helpers import SafeTextRedirector, TextHandler, gui_ask
from ont_automation.nokia_olt import NokiaOLTConnector
from ont_automation.test_engine import TestAutomationEngine


class AutomationGUI:
    def __init__(self, root):
        gui_helpers.GUI_ROOT = root
        self.root = root
        self.root.title("ONT Automation Tool v2.0")
        self.root.geometry("1100x800")
        self.root.configure(bg="#2E3440")

        self.db = DatabaseManager()
        self.config = self.load_config()
        self.olt = None
        self.setup_ui()
        self.setup_logging()

    def load_config(self):
        if os.path.exists(CONFIG_FILE):
            with open(CONFIG_FILE, "r") as f:
                return json.load(f)
        return {}

    def save_config(self):
        with open(CONFIG_FILE, "w") as f:
            json.dump(self.config, f, indent=4)

    def setup_logging(self):
        redir = SafeTextRedirector(self.log_text)
        sys.stdout = redir
        sys.stderr = redir
        handler = TextHandler(redir)
        handler.setFormatter(logging.Formatter("%(levelname)s: %(message)s"))
        logging.getLogger().addHandler(handler)
        logging.getLogger().setLevel(logging.INFO)

    def get_local_nics(self):
        nics = ["Default (Auto)"]
        try:
            out = subprocess.check_output(
                "wmic nicconfig where IPEnabled=True get Description,IPAddress,MACAddress /format:list",
                shell=True,
                text=True,
            )
            current_nic = {}
            for line in out.splitlines():
                line = line.strip()
                if not line:
                    if current_nic and "IPAddress" in current_nic:
                        ip = current_nic["IPAddress"]
                        mac = current_nic.get("MACAddress", "")
                        desc = current_nic.get("Description", "")
                        nics.append(f"{ip} | {mac} | {desc}")
                        current_nic = {}
                    continue
                if "=" in line:
                    k, v = line.split("=", 1)
                    if k == "IPAddress":
                        m = re.search(r'"([^"]+)"', v)
                        if m:
                            current_nic[k] = m.group(1)
                    else:
                        current_nic[k] = v
            if current_nic and "IPAddress" in current_nic:
                ip = current_nic["IPAddress"]
                mac = current_nic.get("MACAddress", "")
                desc = current_nic.get("Description", "")
                nics.append(f"{ip} | {mac} | {desc}")
        except Exception:
            try:
                for ip in socket.gethostbyname_ex(socket.gethostname())[2]:
                    nics.append(ip)
            except Exception:
                pass

        seen = set()
        unique_nics = []
        for n in nics:
            if n not in seen:
                unique_nics.append(n)
                seen.add(n)
        return unique_nics

    def setup_ui(self):
        style = ttk.Style()
        style.theme_use("clam")

        left_frame = tk.Frame(self.root, bg="#2E3440", width=420)
        left_frame.pack(side=tk.LEFT, fill=tk.Y, padx=10, pady=10)
        left_frame.pack_propagate(False)

        self.notebook = ttk.Notebook(left_frame)
        self.notebook.pack(fill=tk.BOTH, expand=True)

        tab_tgt = ttk.Frame(self.notebook)
        self.notebook.add(tab_tgt, text="Target & Connection")

        info_frame = ttk.LabelFrame(tab_tgt, text="OLT System Info")
        info_frame.pack(fill=tk.X, padx=10, pady=5)
        self.lbl_vendor = ttk.Label(info_frame, text="Vendor: -")
        self.lbl_vendor.pack(anchor="w", padx=5, pady=2)
        self.lbl_model = ttk.Label(info_frame, text="Model: -")
        self.lbl_model.pack(anchor="w", padx=5, pady=2)
        self.lbl_sw_ver = ttk.Label(info_frame, text="SW Version: -")
        self.lbl_sw_ver.pack(anchor="w", padx=5, pady=2)

        self.sv_olt_ip = self.add_entry(tab_tgt, "OLT IP:", self.config.get("olt_ip", ""))
        self.sv_username = self.add_entry(tab_tgt, "Username:", self.config.get("username", "isadmin"))
        self.sv_password = self.add_entry(tab_tgt, "Password:", self.config.get("password", ""), show="*")

        ttk.Separator(tab_tgt, orient="horizontal").pack(fill="x", pady=10)

        self.sv_serial = self.add_entry(tab_tgt, "ONT Serial:", self.config.get("ont_serial", ""))
        self.sv_chanpair = self.add_entry(tab_tgt, "Channel Pair:", self.config.get("chanpair", "1/1/1/3"))
        self.sv_pon = self.add_entry(tab_tgt, "Target PON:", self.config.get("pon", "ng2:3/1"))
        self.sv_ont_id = self.add_entry(tab_tgt, "Target ONT ID:", self.config.get("ont_id", "1"))

        tab_svc = ttk.Frame(self.notebook)
        self.notebook.add(tab_svc, text="Service & VLAN")

        self.cb_ont_type = self.add_combobox(tab_svc, "ONT Type:", ["sfu", "hgu"], self.config.get("ont_type", "sfu"))
        self.sv_lan_ports = self.add_entry(tab_svc, "LAN Ports:", self.config.get("lan_ports", "1"))
        self.sv_max_mac = self.add_entry(tab_svc, "Max MAC:", self.config.get("max_mac", "16"))
        self.sv_bw_profile = self.add_entry(
            tab_svc, "BW Profile Name:", self.config.get("bw_profile", "NG2DATABWUP10000")
        )

        ttk.Separator(tab_svc, orient="horizontal").pack(fill="x", pady=10)

        self.cb_vlan_mode = self.add_combobox(
            tab_svc, "VLAN Mode:", ["untagged", "tagged", "translation"], self.config.get("vlan_mode", "untagged")
        )
        self.sv_vlan_id = self.add_entry(tab_svc, "Service VLAN ID:", self.config.get("vlan_id", "1001"))
        self.sv_cvlan = self.add_entry(tab_svc, "C-VLAN (Translation):", self.config.get("c_vlan", "10"))

        tab_spd = ttk.Frame(self.notebook)
        self.notebook.add(tab_spd, text="Speed Test Config")

        tool_default = self.config.get("throughput_tool", "ask")
        if tool_default not in ("iperf3", "ixia", "ask"):
            tool_default = "ask"
        self.cb_throughput_tool = self.add_combobox(
            tab_spd,
            "Throughput Tool:",
            ["ask", "iperf3", "ixia"],
            tool_default,
        )

        self.sv_iperf_ip = self.add_entry(
            tab_spd, "iPerf3 Server IP:", self.config.get("iperf_server", "103.175.200.43")
        )
        self.sv_iperf_port = self.add_entry(tab_spd, "Server Port:", self.config.get("iperf_port", "5201"))
        self.sv_duration = self.add_entry(tab_spd, "Test Duration (sec):", self.config.get("speed_duration", "30"))
        self.sv_criteria = self.add_entry(tab_spd, "Pass Criteria (Mbps):", self.config.get("speed_pass_mbps", "8000"))
        self.sv_iperf_opts = self.add_entry(tab_spd, "iPerf3 Extra Options:", self.config.get("iperf_options", "-P 8"))

        nic_list = self.get_local_nics()
        self.cb_bind_nic = self.add_combobox(
            tab_spd, "Bind NIC (IP | MAC):", nic_list, self.config.get("bind_nic", "Default (Auto)")
        )

        right_frame = tk.Frame(self.root, bg="#2E3440")
        right_frame.pack(side=tk.RIGHT, fill=tk.BOTH, expand=True, padx=10, pady=10)

        ctrl_frame = tk.Frame(right_frame, bg="#2E3440")
        ctrl_frame.pack(fill=tk.X, pady=(0, 10))

        btn_style = {
            "bg": "#4C566A",
            "fg": "white",
            "font": ("Arial", 10, "bold"),
            "relief": tk.FLAT,
            "padx": 10,
            "pady": 5,
        }

        tk.Button(
            ctrl_frame,
            text="Scan Unprovisioned",
            command=lambda: self.run_thread(self.scan_onts, "unprovisioned"),
            **btn_style,
        ).pack(side=tk.LEFT, padx=5)
        tk.Button(
            ctrl_frame,
            text="Scan Active (UP)",
            command=lambda: self.run_thread(self.scan_onts, "active"),
            **btn_style,
        ).pack(side=tk.LEFT, padx=5)
        tk.Button(
            ctrl_frame,
            text="Provision & Test",
            command=lambda: self.run_thread(self.start_automation, False),
            bg="#A3BE8C",
            fg="white",
            font=("Arial", 10, "bold"),
        ).pack(side=tk.LEFT, padx=5)
        tk.Button(
            ctrl_frame,
            text="Test Only",
            command=lambda: self.run_thread(self.start_automation, True),
            **btn_style,
        ).pack(side=tk.LEFT, padx=5)
        tk.Button(
            ctrl_frame,
            text="Delete ONT",
            command=lambda: self.run_thread(self.delete_ont),
            bg="#BF616A",
            fg="white",
            font=("Arial", 10, "bold"),
        ).pack(side=tk.RIGHT, padx=5)

        self.log_text = scrolledtext.ScrolledText(
            right_frame, bg="#1E1E1E", fg="#D8DEE9", font=("Consolas", 10), state=tk.DISABLED
        )
        self.log_text.pack(fill=tk.BOTH, expand=True)

    def add_entry(self, parent, label_text, default_val, show=None):
        frame = tk.Frame(parent)
        frame.pack(fill=tk.X, padx=10, pady=5)
        tk.Label(frame, text=label_text, width=20, anchor="w").pack(side=tk.LEFT)
        sv = tk.StringVar(value=default_val)
        tk.Entry(frame, textvariable=sv, show=show).pack(side=tk.RIGHT, fill=tk.X, expand=True)
        return sv

    def add_combobox(self, parent, label_text, values, default_val):
        frame = tk.Frame(parent)
        frame.pack(fill=tk.X, padx=10, pady=5)
        tk.Label(frame, text=label_text, width=20, anchor="w").pack(side=tk.LEFT)
        cb = ttk.Combobox(frame, values=values, state="readonly")
        cb.set(default_val)
        cb.pack(side=tk.RIGHT, fill=tk.X, expand=True)
        return cb

    def update_config_from_ui(self):
        self.config["olt_ip"] = self.sv_olt_ip.get()
        self.config["username"] = self.sv_username.get()
        self.config["password"] = self.sv_password.get()
        self.config["protocol"] = "telnet"

        self.config["ont_serial"] = self.sv_serial.get()
        self.config["chanpair"] = self.sv_chanpair.get()
        self.config["pon"] = self.sv_pon.get()
        self.config["ont_id"] = self.sv_ont_id.get()

        self.config["ont_type"] = self.cb_ont_type.get()
        self.config["lan_ports"] = self.sv_lan_ports.get()
        self.config["max_mac"] = self.sv_max_mac.get()
        self.config["bw_profile"] = self.sv_bw_profile.get()
        self.config["vlan_mode"] = self.cb_vlan_mode.get()
        self.config["vlan_id"] = self.sv_vlan_id.get()
        self.config["c_vlan"] = self.sv_cvlan.get()

        self.config["throughput_tool"] = self.cb_throughput_tool.get()
        self.config["iperf_server"] = self.sv_iperf_ip.get()
        self.config["iperf_port"] = self.sv_iperf_port.get()
        self.config["speed_duration"] = self.sv_duration.get()
        self.config["speed_pass_mbps"] = self.sv_criteria.get()
        self.config["iperf_options"] = self.sv_iperf_opts.get()
        self.config["bind_nic"] = self.cb_bind_nic.get()

        self.save_config()

    def update_olt_info_display(self, vendor, model, sw_ver):
        self.lbl_vendor.config(text=f"Vendor: {vendor}")
        self.lbl_model.config(text=f"Model: {model}")
        self.lbl_sw_ver.config(text=f"SW Version: {sw_ver}")
        self.config["olt_vendor"] = vendor
        self.config["olt_model"] = model
        self.save_config()

    def run_thread(self, target_func, *args):
        self.update_config_from_ui()
        threading.Thread(target=target_func, args=args, daemon=True).start()

    def get_connector(self):
        if not self.olt or not self.olt.connection or not self.olt.connection.is_alive():
            self.olt = NokiaOLTConnector(
                self.config["olt_ip"],
                self.config["username"],
                self.config["password"],
                self.config["protocol"],
            )
            if not self.olt.connect():
                logging.error("Failed to connect to OLT. Check IP/Credentials.")
                return None
            self.root.after(
                0, self.update_olt_info_display, self.olt.vendor, self.olt.model, self.olt.sw_version
            )
        return self.olt

    def show_selection_popup(self, title, items, callback):
        if not items:
            messagebox.showinfo("No Results", f"No items found for {title}.")
            return

        def on_select(evt):
            sel = listbox.curselection()
            if sel:
                idx = sel[0]
                callback(items[idx])
                top.destroy()

        top = tk.Toplevel(self.root)
        top.title(title)
        top.geometry("400x300")
        top.transient(self.root)
        top.grab_set()

        listbox = tk.Listbox(top, font=("Consolas", 10))
        listbox.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)

        for item in items:
            listbox.insert(
                tk.END,
                f"Serial: {item['serial']} | Port: {item.get('port', 'N/A')} | Chan: {item['chanpair']}",
            )

        listbox.bind("<Double-1>", on_select)
        tk.Button(top, text="Cancel", command=top.destroy).pack(pady=5)

    def scan_onts(self, scan_type="unprovisioned"):
        olt = self.get_connector()
        if not olt:
            return

        print("\n" + "=" * 50)
        logging.info(f"Starting {scan_type.upper()} ONT Scan...")

        if scan_type == "unprovisioned":
            cmd = self.db.get_learned_command(olt.olt_profile, "Discovery") or "show channel-pair unprovision-onu"
            items = olt.get_unprovisioned_onts(cmd)
        else:
            items = olt.scan_active_onts()

        logging.info(f"Scan complete. Found {len(items)} items.")

        def handle_selection(selected):
            self.sv_serial.set(selected["serial"])
            self.sv_chanpair.set(selected["chanpair"])

            if scan_type == "active":
                pon, ont_id = selected["port"].rsplit("/", 1)
                self.sv_pon.set(pon)
                self.sv_ont_id.set(ont_id)
            else:
                pon = olt.suggest_pon_from_chanpair(selected["chanpair"])
                if pon:
                    self.sv_pon.set(pon)
                ont_id = olt.find_available_ont_id(selected["chanpair"])
                if ont_id:
                    self.sv_ont_id.set(ont_id)

            logging.info(f"Auto-filled target info for {selected['serial']}.")

        self.root.after(
            0, lambda: self.show_selection_popup(f"Select {scan_type.capitalize()} ONT", items, handle_selection)
        )

    def delete_ont(self):
        olt = self.get_connector()
        if not olt:
            return

        port_full = f"{self.config['pon']}/{self.config['ont_id']}"
        confirm = messagebox.askyesno(
            "Confirm Deletion", f"Are you sure you want to delete ONT on port {port_full}?"
        )
        if confirm:
            print("\n" + "=" * 50)
            olt._delete_ont_by_port(port_full, self.db)
            logging.info(f"Deletion complete for {port_full}.")

    def start_automation(self, skip_provisioning=False):
        olt = self.get_connector()
        if not olt:
            return

        port_full = f"{self.config['pon']}/{self.config['ont_id']}"
        print("\n" + "=" * 60)
        logging.info("Starting Automation Sequence...")

        try:
            self.db.add_initial_test_cases(self.config, olt.olt_profile)

            if not skip_provisioning:
                while True:
                    if not olt.register_ont(self.config, self.db):
                        logging.error("Registration aborted.")
                        return
                    if not olt.verify_registration(self.config, self.db):
                        olt._cleanup_failed_provisioning(port_full)
                        if gui_ask("Registration Verification Failed. Retry? (y/n)").lower() == "y":
                            continue
                        return
                    if not olt.provision_service_ont(self.config, self.db):
                        olt._cleanup_failed_provisioning(port_full)
                        if gui_ask("Provisioning Failed. Retry sequence? (y/n)").lower() == "y":
                            continue
                        return
                    break
                self.db.add_ont(
                    self.config["ont_serial"],
                    self.config["pon"],
                    self.config["ont_id"],
                    "PROVISIONED",
                )

            engine = TestAutomationEngine(self.db, olt, self.config)
            for cycle in range(1, 10):
                logging.info(f"\n========== TEST CYCLE {cycle} ==========")
                if engine.execute_tests():
                    cases = self.db.load_test_cases(self.config["ont_serial"])
                    fails = [c for c in cases if c["status"] == "FAIL"]
                    if fails:
                        logging.error(
                            f"TEST CYCLE COMPLETED WITH {len(fails)} FAILURE(S). Check report for details."
                        )
                    else:
                        logging.info("ALL TESTS PASSED SUCCESSFULLY!")
                    break
        except Exception as e:
            logging.error(f"Critical error during automation: {e}")
        finally:
            generate_professional_excel_report(self.config["ont_serial"], self.db, self.config)
            logging.info("Automation Task & Report Generation Completed.")
