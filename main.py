import sqlite3
import logging
import time
import json
import os
import sys
import getpass
import re
import subprocess
import threading
import csv
from datetime import datetime
from typing import List, Dict, Tuple

try:
    from netmiko import ConnectHandler
    from netmiko.exceptions import NetmikoTimeoutException, NetmikoAuthenticationException
except ImportError:
    logging.error("Netmiko library is not installed. Please run: pip install netmiko")
    sys.exit(1)

try:
    import openpyxl
    from openpyxl.styles import Font, Alignment, PatternFill, Border, Side
    from openpyxl.chart import LineChart, Reference
    from openpyxl.chart.marker import Marker
    EXCEL_SUPPORT = True
except ImportError:
    logging.warning("openpyxl is not installed! Excel report generation will be skipped. Run 'pip install openpyxl'")
    EXCEL_SUPPORT = False

# Configure logging
logging.basicConfig(level=logging.INFO, format='%(levelname)s: %(message)s')

CONFIG_FILE = "env_config.json"

def generate_professional_excel_report(serial: str, db, config: Dict):
    """Generates a highly formatted 3-sheet Excel report with live speed graphs."""
    if not EXCEL_SUPPORT: return
    
    cases = db.load_test_cases(serial)
    filename = f"ONT_Test_Result_{serial}_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx"
    
    try:
        wb = openpyxl.Workbook()
        
        bold_font = Font(bold=True)
        title_font = Font(size=16, bold=True, color="FFFFFF")
        hdr_fill = PatternFill(start_color="4F81BD", end_color="4F81BD", fill_type="solid")
        pass_fill = PatternFill(start_color="C6EFCE", end_color="C6EFCE", fill_type="solid")
        fail_fill = PatternFill(start_color="FFC7CE", end_color="FFC7CE", fill_type="solid")
        na_fill = PatternFill(start_color="FFEB9C", end_color="FFEB9C", fill_type="solid")
        nt_fill = PatternFill(start_color="F2F2F2", end_color="F2F2F2", fill_type="solid")
        center_align = Alignment(horizontal="center", vertical="center")
        thin_border = Border(left=Side(style='thin'), right=Side(style='thin'), top=Side(style='thin'), bottom=Side(style='thin'))

        # SHEET 1: SUMMARY
        ws_sum = wb.active
        ws_sum.title = "Summary"
        
        ws_sum.merge_cells('B2:E3')
        title_cell = ws_sum.cell(row=2, column=2, value="ONT Automation Test Report")
        title_cell.font = title_font
        title_cell.alignment = center_align
        title_cell.fill = hdr_fill
        
        total = len(cases)
        c_pass = sum(1 for c in cases if c['status'] == 'PASS')
        c_fail = sum(1 for c in cases if c['status'] == 'FAIL')
        c_na   = sum(1 for c in cases if c['status'] == 'N/A')
        c_nt   = sum(1 for c in cases if c['status'] == 'N/T')
        progress = ((total - c_nt) / total) * 100 if total else 0
        
        info = [
            ("Target Serial", serial),
            ("Target PON / ID", f"{config.get('pon', 'N/A')} / {config.get('ont_id', 'N/A')}"),
            ("Software Version", config.get('sw_version', 'Unknown')),
            ("VLAN Mode", config.get('vlan_mode', 'untagged').capitalize()),
            ("Test Date", datetime.now().strftime('%Y-%m-%d %H:%M:%S')),
            ("Total Tests", total),
            ("PASS", c_pass),
            ("FAIL", c_fail),
            ("N/A (Skipped)", c_na),
            ("N/T (Not Tested)", c_nt),
            ("Progress", f"{progress:.1f}%")
        ]
        
        row_idx = 5
        for key, val in info:
            c_key = ws_sum.cell(row=row_idx, column=2, value=key)
            c_val = ws_sum.cell(row=row_idx, column=3, value=val)
            c_key.font = bold_font
            c_key.border = thin_border
            c_val.border = thin_border
            if key == "PASS": c_val.fill = pass_fill
            elif key == "FAIL": c_val.fill = fail_fill
            elif key == "N/A (Skipped)": c_val.fill = na_fill
            elif key == "N/T (Not Tested)": c_val.fill = nt_fill
            row_idx += 1
            
        ws_sum.column_dimensions['B'].width = 20
        ws_sum.column_dimensions['C'].width = 30

        # SHEET 2: DETAILS
        ws_det = wb.create_sheet(title="Details")
        headers = ["Test Order", "Test Item", "Command / Script", "Result"]
        for col_num, header in enumerate(headers, 1):
            cell = ws_det.cell(row=1, column=col_num, value=header)
            cell.font = bold_font
            cell.fill = PatternFill(start_color="D9D9D9", end_color="D9D9D9", fill_type="solid")
            cell.alignment = center_align
            cell.border = thin_border
            
        for idx, c in enumerate(cases, 2):
            ws_det.cell(row=idx, column=1, value=c['order']).alignment = center_align
            ws_det.cell(row=idx, column=2, value=c['type'])
            ws_det.cell(row=idx, column=3, value=c['command'])
            
            res_cell = ws_det.cell(row=idx, column=4, value=c['status'])
            res_cell.alignment = center_align
            if c['status'] == 'PASS': res_cell.fill = pass_fill
            elif c['status'] == 'FAIL': res_cell.fill = fail_fill
            elif c['status'] == 'N/A': res_cell.fill = na_fill
            else: res_cell.fill = nt_fill
            
            for col in range(1, 5):
                ws_det.cell(row=idx, column=col).border = thin_border
                
        ws_det.column_dimensions['B'].width = 25
        ws_det.column_dimensions['C'].width = 80
        ws_det.column_dimensions['D'].width = 15

        # SHEET 3: LOGS & GRAPH
        ws_log = wb.create_sheet(title="Logs & Graph")
        ws_log.cell(row=1, column=1, value="Test Item").font = bold_font
        ws_log.cell(row=1, column=2, value="Execution Log").font = bold_font
        ws_log.column_dimensions['A'].width = 20
        ws_log.column_dimensions['B'].width = 100
        
        log_row = 2
        up_data = {}
        dl_data = {}
        current_mode = "UPLOAD"
        
        for c in cases:
            ws_log.cell(row=log_row, column=1, value=c['type']).font = bold_font
            log_lines = (c['log'] or "No log captured.").split('\n')
            
            for line in log_lines:
                clean_line = line.strip()
                if not clean_line: continue
                
                clean_line = re.sub(r'[\x00-\x08\x0b-\x0c\x0e-\x1f]', '', clean_line)
                if clean_line.startswith('='):
                    clean_line = "'" + clean_line

                ws_log.cell(row=log_row, column=2, value=clean_line)
                
                if c['type'] == 'Speed_Test':
                    if "--- STARTING DOWNLOAD TEST ---" in clean_line:
                        current_mode = "DOWNLOAD"
                    elif "--- STARTING UPLOAD TEST ---" in clean_line:
                        current_mode = "UPLOAD"
                        
                    if "[SUM]" in clean_line and "sec" in clean_line and "bits/sec" in clean_line:
                        if "sender" not in clean_line and "receiver" not in clean_line:
                            match = re.search(r'\[SUM\]\s+\d+\.\d+-\s*(\d+\.\d+)\s+sec.*?\s+(\d+(?:\.\d+)?)\s+([KMG])bits/sec', clean_line)
                            if match:
                                end_time_float = float(match.group(1))
                                t_sec = int(round(end_time_float))
                                val = float(match.group(2))
                                unit = match.group(3)
                                mbps = val * 1000 if unit == 'G' else (val if unit == 'M' else val / 1000)
                                
                                if current_mode == "UPLOAD":
                                    up_data[t_sec] = max(up_data.get(t_sec, 0), mbps)
                                else:
                                    dl_data[t_sec] = max(dl_data.get(t_sec, 0), mbps)
                log_row += 1
            log_row += 1 

        if up_data or dl_data:
            ws_log.cell(row=1, column=4, value="Time (sec)").font = bold_font
            ws_log.cell(row=1, column=5, value="Upload (Mbps)").font = bold_font
            ws_log.cell(row=1, column=6, value="Download (Mbps)").font = bold_font
            
            all_times = sorted(list(set(list(up_data.keys()) + list(dl_data.keys()))))
            
            r_idx = 2
            if all_times:
                max_time = max(all_times)
                last_up = None
                last_dl = None
                
                for t in range(1, max_time + 1):
                    val_up = up_data.get(t, last_up)
                    val_dl = dl_data.get(t, last_dl)
                    
                    ws_log.cell(row=r_idx, column=4, value=t)
                    ws_log.cell(row=r_idx, column=5, value=val_up)
                    ws_log.cell(row=r_idx, column=6, value=val_dl)
                    
                    if val_up is not None: last_up = val_up
                    if val_dl is not None: last_dl = val_dl
                    r_idx += 1
                
            chart = LineChart()
            chart.title = "iPerf3 Throughput over Time (Up/Down)"
            chart.y_axis.title = 'Throughput (Mbps)'
            chart.x_axis.title = 'Time (sec)'
            chart.width = 20
            chart.height = 10
            
            chart.x_axis.tickLblPos = "low"
            chart.y_axis.tickLblPos = "low"
            chart.dispBlanksAs = "span" 
            
            data = Reference(ws_log, min_col=5, max_col=6, min_row=1, max_row=r_idx-1)
            cats = Reference(ws_log, min_col=4, min_row=2, max_row=r_idx-1)
            
            chart.add_data(data, titles_from_data=True)
            chart.set_categories(cats)
            
            for s in chart.series:
                s.marker = Marker(symbol='circle', size=5)
            
            ws_log.add_chart(chart, "H2")

        wb.save(filename)
        print("\n" + "="*60)
        print(f" [✔] Professional Excel Report Generated: {filename} ")
        print("="*60 + "\n")
        
    except Exception as e:
        logging.error(f"Failed to generate Excel report: {e}")

class DatabaseManager:
    def __init__(self, db_name: str = "olt_physical_automation.db"):
        self.db_name = db_name
        self._initialize_database()

    def _get_conn(self):
        return sqlite3.connect(self.db_name, timeout=20)

    def _initialize_database(self):
        with self._get_conn() as conn:
            cursor = conn.cursor()
            cursor.execute('''CREATE TABLE IF NOT EXISTS ont_inventory (
                id INTEGER PRIMARY KEY AUTOINCREMENT, serial_number TEXT UNIQUE, 
                pon_port TEXT, ont_id TEXT, status TEXT, provisioning_date TIMESTAMP DEFAULT CURRENT_TIMESTAMP)''')
            
            cursor.execute('''CREATE TABLE IF NOT EXISTS test_cases (
                id INTEGER PRIMARY KEY AUTOINCREMENT, target_serial TEXT, execution_order INTEGER, 
                test_type TEXT, command TEXT, parameters TEXT, expected_output TEXT, 
                success_count INTEGER DEFAULT 0, is_evolved BOOLEAN DEFAULT 0)''')
                
            cursor.execute('''CREATE TABLE IF NOT EXISTS learned_commands (
                olt_profile TEXT, test_type TEXT, command_template TEXT, PRIMARY KEY (olt_profile, test_type))''')
            
            try: cursor.execute("ALTER TABLE test_cases ADD COLUMN status TEXT DEFAULT 'N/T'")
            except: pass
            try: cursor.execute("ALTER TABLE test_cases ADD COLUMN execution_log TEXT")
            except: pass
            conn.commit()

    def add_ont(self, serial: str, pon: str, ont_id: str, status: str):
        with self._get_conn() as conn:
            cursor = conn.cursor()
            cursor.execute('INSERT OR REPLACE INTO ont_inventory (serial_number, pon_port, ont_id, status) VALUES (?, ?, ?, ?)', 
                           (serial, pon, ont_id, status))
            conn.commit()

    def get_learned_command(self, profile: str, t_type: str) -> str:
        with self._get_conn() as conn:
            cursor = conn.cursor()
            cursor.execute('SELECT command_template FROM learned_commands WHERE olt_profile = ? AND test_type = ?', (profile, t_type))
            row = cursor.fetchone()
            return row[0] if row else None

    def save_learned_command(self, profile: str, t_type: str, template: str):
        with self._get_conn() as conn:
            cursor = conn.cursor()
            cursor.execute('INSERT OR REPLACE INTO learned_commands (olt_profile, test_type, command_template) VALUES (?, ?, ?)', 
                           (profile, t_type, template))
            conn.commit()

    def update_test_case(self, t_id: int, cmd: str, params: str):
        with self._get_conn() as conn:
            cursor = conn.cursor()
            cursor.execute('UPDATE test_cases SET command = ?, parameters = ? WHERE id = ?', (cmd, params, t_id))
            conn.commit()

    def update_test_status(self, t_id: int, status: str, log: str):
        with self._get_conn() as conn:
            cursor = conn.cursor()
            cursor.execute("UPDATE test_cases SET status = ?, execution_log = ? WHERE id = ?", (status, log, t_id))
            if status == "PASS":
                cursor.execute("UPDATE test_cases SET success_count = success_count + 1 WHERE id = ?", (t_id,))
            conn.commit()

    def load_test_cases(self, serial: str) -> List[Dict]:
        with self._get_conn() as conn:
            cursor = conn.cursor()
            cursor.execute('SELECT id, target_serial, execution_order, test_type, command, parameters, status, execution_log, success_count FROM test_cases WHERE target_serial = ? ORDER BY execution_order ASC', (serial,))
            rows = cursor.fetchall()
            return [{"id": r[0], "serial": r[1], "order": r[2], "type": r[3], "command": r[4], "parameters": json.loads(r[5]), "status": r[6], "log": r[7], "success_count": r[8]} for r in rows]

    def add_initial_test_cases(self, config: Dict, profile: str):
        serial = config['ont_serial']
        pon = config['pon']
        ont_id = config['ont_id']
        chanpair = config.get('chanpair', '')
        port_full = f"{pon}/{ont_id}"
        vendor_id = serial[:4]
        sn_rem = serial[4:]
        serial_formatted = f"{vendor_id}:{sn_rem}"

        with self._get_conn() as conn:
            cursor = conn.cursor()
            cursor.execute("DELETE FROM test_cases WHERE target_serial = ?", (serial,))
                    
            templates = {
                "ONT_Discovery_Check": "show equipment ont status channel-pair {chanpair}", 
                "Registration_Check": "show equipment ont status channel-pair {chanpair}", 
                "Reboot_Test": "admin equipment ont interface {port} reboot with-active-image",
                "Optics_Check": "show equipment ont optics {port}",
                "UNI_Status": "show ethernet ont operational-data {port}/1/1",
                "Software_Info": "show equipment ont interface {port} detail",
                "Speed_Test": "Native iperf3 Execution", 
                "MAC_Status": "show vlan bridge-port-fdb {port}/1/1",
                "Cleanup": "configure equipment ont interface {port} admin-state down ; configure equipment ont no interface {port}"
            }
            
            cursor.execute('SELECT test_type, command_template FROM learned_commands WHERE olt_profile = ?', (profile,))
            learned_dict = {row[0]: row[1] for row in cursor.fetchall()}
            
            mock_cases = []
            for i, (t_type, def_cmd) in enumerate(templates.items(), 1):
                template = learned_dict.get(t_type)
                
                # Auto-Heal incomplete or incorrect templates learned from manual input
                if t_type == "Reboot_Test" and template and "with-active-image" not in template:
                    template = def_cmd
                    cursor.execute('INSERT OR REPLACE INTO learned_commands (olt_profile, test_type, command_template) VALUES (?, ?, ?)', (profile, t_type, def_cmd))
                elif t_type == "UNI_Status" and template and "operational-data" not in template:
                    template = def_cmd
                    cursor.execute('INSERT OR REPLACE INTO learned_commands (olt_profile, test_type, command_template) VALUES (?, ?, ?)', (profile, t_type, def_cmd))
                elif "{port}" in def_cmd and template and "{port}" not in template:
                    template = def_cmd
                    cursor.execute('INSERT OR REPLACE INTO learned_commands (olt_profile, test_type, command_template) VALUES (?, ?, ?)', (profile, t_type, def_cmd))
                elif "{chanpair}" in def_cmd and template and "{chanpair}" not in template:
                    template = def_cmd
                    cursor.execute('INSERT OR REPLACE INTO learned_commands (olt_profile, test_type, command_template) VALUES (?, ?, ?)', (profile, t_type, def_cmd))
                    
                if not template:
                    template = def_cmd
                
                cmd = template.format(
                    serial=serial, 
                    serial_formatted=serial_formatted,
                    pon=pon, 
                    ont_id=ont_id, 
                    port=port_full,
                    chanpair=chanpair
                )
                
                params = {"timeout": 15, "wait": 0}
                if t_type == "Registration_Check": params["wait"] = 5
                elif t_type == "MAC_Status": params["wait"] = 10
                    
                mock_cases.append((serial, i, t_type, cmd, json.dumps(params), "N/T"))
                
            cursor.executemany("INSERT INTO test_cases (target_serial, execution_order, test_type, command, parameters, status) VALUES (?,?,?,?,?,?)", mock_cases)
            conn.commit()

class NokiaOLTConnector:
    def __init__(self, ip: str, user: str, pw: str, proto: str):
        self.proto = proto
        self.device = {'host': ip, 'username': user, 'password': pw, 'global_delay_factor': 2}
        self.connection = None
        self.olt_profile = "Nokia_Default"

    def connect(self) -> bool:
        types = ["nokia_sros_telnet", "alcatel_sros_telnet"] if self.proto == "telnet" else ["nokia_sros", "alcatel_sros"]
        for dt in types:
            self.device['device_type'] = dt
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
            except Exception: pass

    def _discover_profile(self):
        out = self.send_command("show version", 5)
        m = re.search(r'(\d+\.\d+\.\S+)', out)
        self.olt_profile = f"Nokia_OS_{m.group(1)}" if m else "Nokia_Unknown"

    def send_command(self, cmd: str, timeout: int = 15, retry: bool = True) -> str:
        if self.connection:
            try: 
                if not self.connection.is_alive():
                    raise Exception("Connection dead")
                return self.connection.send_command_timing(cmd, read_timeout=timeout)
            except Exception as e:
                if retry:
                    logging.warning(f"Connection dropped during command execution. Attempting Auto-Reconnect...")
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

    def get_unprovisioned_onts(self, discovery_cmd: str) -> List[Dict[str, str]]:
        if self.connection:
            output = self.send_command(discovery_cmd)
            matches = re.findall(r'(\d+(?:/\d+)+)\s+([A-Za-z]{4}[A-Fa-f0-9]{8})', output)
            if matches:
                onts = [{"chanpair": m[0], "serial": m[1]} for m in matches]
                unique_onts = {item['serial']: item for item in onts}.values()
                return list(unique_onts)
            else:
                serials = re.findall(r'([A-Za-z]{4}[A-Fa-f0-9]{8})', output)
                return [{"chanpair": "unknown", "serial": s} for s in set(serials)]
        return []

    def _get_safe_find_cmd(self, db: DatabaseManager) -> str:
        find_cmd = db.get_learned_command(self.olt_profile, "Find_Provisioned") or "show equipment ont status"
        if re.search(r'\d+/\d+/\d+', find_cmd) or "ng2:" in find_cmd:
            find_cmd = "show equipment ont status"
            db.save_learned_command(self.olt_profile, "Find_Provisioned", find_cmd)
        return find_cmd

    def _get_safe_delete_cmd(self, db: DatabaseManager) -> str:
        del_cmd_temp = db.get_learned_command(self.olt_profile, "Delete_Provisioned")
        default_cmd = "configure equipment ont interface {port} admin-state down ; configure equipment ont no interface {port}"
        if not del_cmd_temp or "{port}" not in del_cmd_temp:
            db.save_learned_command(self.olt_profile, "Delete_Provisioned", default_cmd)
            return default_cmd
        return del_cmd_temp

    def _delete_ont_by_port(self, port_full: str, db: DatabaseManager):
        logging.info(f"Aggressively deleting ONT bound to port: {port_full}...")
        del_cmd_temp = self._get_safe_delete_cmd(db)
        
        cmds = [c.strip() for c in del_cmd_temp.split(';') if c.strip()]
        for c in cmds:
            exec_cmd = c.format(port=port_full)
            self.send_command(exec_cmd, timeout=5)
            
        time.sleep(2)
        logging.info(f"Cleanup complete for port {port_full}.")

    def _cleanup_failed_provisioning(self, port_full: str):
        db = DatabaseManager()
        self._delete_ont_by_port(port_full, db)

    def format_serial(self, serial_raw: str) -> str:
        vendor_id = serial_raw[:4]
        sn_rem = serial_raw[4:]
        return f"{vendor_id}:{sn_rem}"

    def register_ont(self, config: Dict, db: DatabaseManager) -> bool:
        port_full = f"{config['pon']}/{config['ont_id']}"
        serial_raw = config['ont_serial']
        serial_formatted = self.format_serial(serial_raw)
        
        reg_key = f"Registration_{config['ont_type'].upper()}"
        learned_cmds = db.get_learned_command(self.olt_profile, reg_key)
        
        if learned_cmds and "admin-state up" not in learned_cmds.lower():
            logging.warning("Learned registration command is missing critical 'admin-state up' step. Reverting to default sequence.")
            learned_cmds = None
        
        if learned_cmds:
            cmd_templates = [c.strip() for c in learned_cmds.split(';') if c.strip()]
        else:
            if config['ont_type'] == 'sfu':
                cmd_templates = [
                    "configure equipment ont interface {port} sw-ver-pland disabled sernum {serial_formatted} fec-up enable sw-dnload-version disabled pref-channel-pair {chanpair}",
                    "configure equipment ont interface {port} admin-state up",
                    "configure equipment ont slot {port}/1 planned-card-type ethernet plndnumdataports {lan_ports} plndnumvoiceports 0",
                    "configure interface port uni:{port}/1/1 admin-up"
                ]
            else:
                cmd_templates = [
                    "configure equipment ont interface {port} sw-ver-pland disabled sernum {serial_formatted} fec-up enable sw-dnload-version disabled pref-channel-pair {chanpair}",
                    "configure equipment ont interface {port} admin-state up",
                    "configure equipment ont slot {port}/14 planned-card-type veip plndnumdataports {lan_ports} plndnumvoiceports 0"
                ]

        logging.info(f"--- [STEP 1] Initiating ONT Registration for {serial_raw} on {port_full} ---")
        error_kws = ["invalid command", "invalid token", "unknown command", "bad parameter", "incomplete command"]
        
        while True:
            success = True
            for template in cmd_templates:
                cmd = template.format(
                    serial_formatted=serial_formatted, port=port_full,
                    chanpair=config.get('chanpair', ''), lan_ports=config.get('lan_ports', '1')
                )
                
                out = self.send_command(cmd, timeout=10)
                out_lower = out.lower()
                if (any(kw in out_lower for kw in error_kws) or "^" in out) and "pattern not detected" not in out_lower:
                    print(f"\n[ REGISTRATION FAILED ] Command rejected: {cmd}")
                    self._cleanup_failed_provisioning(port_full)
                    
                    new_cmds = input("Enter commands (or 'exit' to abort): ").strip()
                    if new_cmds.lower() == 'exit' or not new_cmds: return False
                        
                    db.save_learned_command(self.olt_profile, reg_key, new_cmds)
                    cmd_templates = [c.strip() for c in new_cmds.split(';') if c.strip()]
                    success = False
                    break 
                    
            if success: return True

    def verify_registration(self, config: Dict, db: DatabaseManager) -> bool:
        port_full = f"{config['pon']}/{config['ont_id']}"
        chanpair = config.get('chanpair', '')
        serial_formatted = self.format_serial(config['ont_serial'])

        cmd_template = db.get_learned_command(self.olt_profile, "Reg_Verify_Cmd")
        default_verify_cmd = "show equipment ont status channel-pair {chanpair}"
        
        if not cmd_template or ("{chanpair}" not in cmd_template and "{port}" not in cmd_template):
            db.save_learned_command(self.olt_profile, "Reg_Verify_Cmd", default_verify_cmd)
            cmd_template = default_verify_cmd

        logging.info(f"--- [STEP 2] Verifying ONT Registration on {port_full} ---")

        while True:
            cmd = cmd_template.format(port=port_full, chanpair=chanpair)
            max_retries = 5
            success = False
            
            for attempt in range(1, max_retries + 1):
                last_out = self.send_command(cmd, timeout=5)
                match_line = [line for line in last_out.split('\n') if serial_formatted.lower() in line.lower() and port_full.lower() in line.lower()]
                if match_line:
                    tokens = match_line[0].strip().lower().split()
                    if "up" in tokens:
                        success = True
                        break
                time.sleep(5)
                
            if success: return True
                
            print(f"\n[ VERIFICATION TIMEOUT ] Could not find oper status 'UP' for serial {serial_formatted} on port {port_full}.")
            choice = input("Select an option [1] Retry [2] Change Cmd [3] Force Pass [4] Abort: ").strip()
            if choice == '1': continue
            elif choice == '2':
                new_cmd = input(f"Enter new command template: ").strip()
                if new_cmd: 
                    cmd_template = new_cmd
                    db.save_learned_command(self.olt_profile, "Reg_Verify_Cmd", cmd_template)
            elif choice == '3': return True
            else: return False

    def provision_service_ont(self, config: Dict, db: DatabaseManager) -> bool:
        port_full = f"{config['pon']}/{config['ont_id']}"
        vlan_mode = config.get('vlan_mode', 'untagged').lower()
        prov_key = f"Provisioning_Service_{config['ont_type'].upper()}_{vlan_mode.upper()}"
        
        learned_cmds = db.get_learned_command(self.olt_profile, prov_key)
        
        if learned_cmds and "vlan-id" not in learned_cmds.lower():
            logging.warning(f"Learned provisioning command for {vlan_mode} seems incomplete. Reverting to default sequence.")
            learned_cmds = None
            
        if learned_cmds:
            cmd_templates = [c.strip() for c in learned_cmds.split(';') if c.strip()]
        else:
            base_sfu = [
                "configure qos interface uni:{port}/1/1 queue [0...7] shaper-profile name:StrictPriority",
                "configure qos interface uni:{port}/1/1 upstream-queue [0...7] bandwidth-profile name:{bw_profile} bandwidth-sharing uni-sharing",
                "configure bridge port {port}/1/1 max-unicast-mac {max_mac}"
            ]
            base_hgu = [
                "configure bridge port {port}/14/1 max-unicast-mac {max_mac}"
            ]
            
            if config['ont_type'] == 'sfu':
                cmd_templates = base_sfu
                if vlan_mode == 'untagged':
                    cmd_templates.extend([
                        "configure bridge port {port}/1/1 vlan-id {vlan_id} usacceptframetype untagged",
                        "configure bridge port {port}/1/1 pvid {vlan_id}"
                    ])
                elif vlan_mode == 'tagged':
                    cmd_templates.append("configure bridge port {port}/1/1 vlan-id {vlan_id} tag single-tagged l2fwder-vlan {vlan_id} vlan-scope local")
                elif vlan_mode == 'translation':
                    cmd_templates.append("configure bridge port {port}/1/1 vlan-id {c_vlan} tag single-tagged l2fwder-vlan {vlan_id} vlan-scope local")
            else:
                cmd_templates = base_hgu
                if vlan_mode == 'untagged':
                    cmd_templates.extend([
                        "configure bridge port {port}/14/1 vlan-id {vlan_id} usacceptframetype untagged",
                        "configure bridge port {port}/14/1 pvid {vlan_id}"
                    ])
                elif vlan_mode == 'tagged':
                    cmd_templates.append("configure bridge port {port}/14/1 vlan-id {vlan_id} tag single-tagged l2fwder-vlan {vlan_id} vlan-scope local")
                elif vlan_mode == 'translation':
                    cmd_templates.append("configure bridge port {port}/14/1 vlan-id {c_vlan} tag single-tagged l2fwder-vlan {vlan_id} vlan-scope local")

        logging.info(f"--- [STEP 3] Initiating Service Provisioning ({vlan_mode.upper()}) on {port_full} ---")
        error_kws = ["invalid command", "invalid token", "unknown command", "bad parameter", "incomplete command"]
        
        while True:
            success = True
            for template in cmd_templates:
                cmd = template.format(
                    port=port_full, bw_profile=config.get('bw_profile', ''),
                    max_mac=config.get('max_mac', '128'), vlan_id=config.get('vlan_id', '1001'),
                    c_vlan=config.get('c_vlan', '10')
                )
                
                out = self.send_command(cmd, timeout=10)
                out_lower = out.lower()
                if (any(kw in out_lower for kw in error_kws) or "^" in out) and "pattern not detected" not in out_lower:
                    print(f"\n[ SERVICE PROVISIONING FAILED ] Command rejected: {cmd}")
                    self._cleanup_failed_provisioning(port_full)
                    new_cmds = input("Enter commands (or 'exit' to abort): ").strip()
                    if new_cmds.lower() == 'exit' or not new_cmds: return False 
                    db.save_learned_command(self.olt_profile, prov_key, new_cmds)
                    return False 
                    
            if success: return True

class TestAutomationEngine:
    def __init__(self, db: DatabaseManager, olt: NokiaOLTConnector, config: Dict):
        self.db = db
        self.olt = olt
        self.config = config
        self.serial = config['ont_serial']
        self.pon = config['pon']
        self.ont_id = config['ont_id']
        self.chanpair = config.get('chanpair', '')
        self.port_full = f"{self.pon}/{self.ont_id}"
        
        vendor_id = self.serial[:4]
        sn_rem = self.serial[4:]
        self.serial_formatted = f"{vendor_id}:{sn_rem}"

    def execute_tests(self) -> bool:
        scenarios = self.db.load_test_cases(self.serial)
        retry_required = False
        
        for sn in scenarios:
            logging.info(f"Executing [{sn['order']}] {sn['type']}")
            if "wait" in sn['parameters'] and sn['parameters']['wait'] > 0:
                logging.info(f"Waiting {sn['parameters']['wait']} seconds...")
                time.sleep(sn['parameters']['wait'])
            
            res = ""
            
            if sn['type'] == "Speed_Test":
                server = self.config.get('iperf_server', '').strip()
                port = str(self.config.get('iperf_port', 5201))
                duration = str(self.config.get('speed_duration', 10))
                criteria = float(self.config.get('speed_pass_mbps', 500))
                
                if not server:
                    logging.warning("No iperf3 server specified. Skipping Speed_Test.")
                    res = "[SPEED_TEST_SUCCESS] Skipped (No server)"
                    self.db.update_test_status(sn['id'], "PASS", res)
                    continue
                
                try:
                    overall_up = 0.0
                    overall_dl = 0.0
                    full_log = []
                    
                    base_args = ['iperf3', '-c', server, '-p', port, '-t', duration, '-P', '8']
                    tests = [("UPLOAD", base_args), ("DOWNLOAD", base_args + ['-R'])]
                    
                    for t_name, t_args in tests:
                        logging.info(f"Running Native iperf3 {t_name}: {' '.join(t_args)}")
                        print("\n" + "="*68)
                        print(f"[{t_name}] {'Time':<8} | {'Progress Graph':<42} | {'Throughput':<15}")
                        print("="*68)
                        
                        full_log.append(f"--- STARTING {t_name} TEST ---\n")
                        
                        proc = subprocess.Popen(t_args, stdout=subprocess.PIPE, stderr=subprocess.STDOUT, text=True, bufsize=1)
                        direction_log = []
                        
                        for line in iter(proc.stdout.readline, ''):
                            direction_log.append(line)
                            full_log.append(line)
                            if "[SUM]" in line and "bits/sec" in line and "sender" not in line and "receiver" not in line:
                                match = re.search(r'\[SUM\]\s+\d+\.\d+-\s*(\d+\.\d+)\s+sec.*?\s+(\d+(?:\.\d+)?)\s+([KMG])bits/sec', line)
                                if match:
                                    sec_val = match.group(1)
                                    val = float(match.group(2))
                                    unit = match.group(3)
                                    mbps = val * 1000 if unit == 'G' else (val if unit == 'M' else val / 1000)
                                    
                                    bar_len = 40
                                    max_val = max(criteria * 1.2, 10000)
                                    filled = min(int((mbps / max_val) * bar_len), bar_len)
                                    bar = '█' * filled + '-' * (bar_len - filled)
                                    
                                    print(f"[{sec_val:>5}s] [{bar}] {mbps:8.2f} Mbps")
                                    
                        proc.wait()
                        print("="*68 + "\n")
                        
                        dir_text = "".join(direction_log)
                        if proc.returncode != 0:
                            logging.warning(f"{t_name} Speed Test returned code {proc.returncode}")
                            
                        if t_name == "UPLOAD":
                            sender_matches = re.findall(r'\[SUM\].*?\s+(\d+(?:\.\d+)?)\s+([KMG])bits/sec\s+sender', dir_text)
                            if sender_matches:
                                v, u = float(sender_matches[-1][0]), sender_matches[-1][1]
                                overall_up = v * 1000 if u == 'G' else (v if u == 'M' else v / 1000)
                        else:
                            recv_matches = re.findall(r'\[SUM\].*?\s+(\d+(?:\.\d+)?)\s+([KMG])bits/sec\s+receiver', dir_text)
                            if recv_matches:
                                v, u = float(recv_matches[-1][0]), recv_matches[-1][1]
                                overall_dl = v * 1000 if u == 'G' else (v if u == 'M' else v / 1000)
                                
                        time.sleep(2) 
                        
                    out_text = "".join(full_log)
                    
                    if overall_up >= criteria and overall_dl >= criteria:
                        res = f"[SPEED_TEST_SUCCESS] Upload: {overall_up:.2f} Mbps, Download: {overall_dl:.2f} Mbps\n{out_text}"
                        logging.info(f"Speed test passed criteria! (Up: {overall_up:.2f} Mbps, Down: {overall_dl:.2f} Mbps)")
                    else:
                        res = f"[SPEED_TEST_FAILED] Upload: {overall_up:.2f} Mbps, Download: {overall_dl:.2f} Mbps (Target: {criteria} Mbps)\n{out_text}"
                        logging.warning(f"Speed test throughput below target {criteria} Mbps.")
                
                except FileNotFoundError:
                    res = "[SPEED_TEST_FAILED] iperf3 executable not found in system PATH."
                    logging.error(res)
                except Exception as e:
                    res = f"[SPEED_TEST_FAILED] Unexpected error: {e}"
                    logging.error(res)
                    
            elif sn['type'] == "Reboot_Test":
                logging.info("Sending Reboot Command to OLT...")
                res_initial = self.olt.send_command(sn['command'])
                
                if "^" in res_initial or "invalid token" in res_initial.lower() or "error" in res_initial.lower():
                    res = res_initial 
                else:
                    logging.info("Reboot command sent. Forcing OLT disconnection to allow PC NIC failover (Wait 5s)...")
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
                        res = f"[REBOOT_FAILED] Could not reconnect to OLT after 5 minutes.\n[Reboot Trigger]\n{res_initial}"
                        logging.error("OLT Reconnection timeout reached!")
                    else:
                        alarm_cmd_template = self.db.get_learned_command(self.olt.olt_profile, "Alarm_Check_Cmd")
                        if not alarm_cmd_template:
                            alarm_cmd_template = "show equipment ont alarm {port}" 
                            self.db.save_learned_command(self.olt.olt_profile, "Alarm_Check_Cmd", alarm_cmd_template)
                            
                        alarm_cmd = alarm_cmd_template.format(port=self.port_full, chanpair=self.chanpair)
                        
                        reboot_success = False
                        reboot_time_sec = 0.0
                        last_alarm_out = ""
                        
                        log_pattern = re.compile(r'(\d{2}/\d{2}/\d{2}\s+\d{2}:\d{2}:\d{2})\s+major alarm (occurred|cleared).*ONT is inactive', re.IGNORECASE)
                        
                        for attempt in range(60):
                            last_alarm_out = self.olt.send_command(alarm_cmd, timeout=5)
                            
                            occurred_time = None
                            cleared_time = None
                            
                            for match in log_pattern.finditer(last_alarm_out):
                                timestamp_str, action = match.groups()
                                try:
                                    dt_obj = datetime.strptime(timestamp_str, "%y/%m/%d %H:%M:%S")
                                    if action.lower() == "occurred":
                                        occurred_time = dt_obj
                                    elif action.lower() == "cleared":
                                        cleared_time = dt_obj
                                except ValueError:
                                    continue
                                    
                            if occurred_time and cleared_time and cleared_time > occurred_time:
                                reboot_time_sec = (cleared_time - occurred_time).total_seconds()
                                reboot_success = True
                                break
                            elif cleared_time and not occurred_time:
                                reboot_success = True
                                reboot_time_sec = time.time() - start_wait_t
                                break
                                
                            time.sleep(5)
                            
                        if reboot_success:
                            res = f"[REBOOT_SUCCESS] ONT recovery confirmed via Alarm Log. True Reboot Duration: {reboot_time_sec:.1f} seconds.\n[Reboot Trigger]\n{res_initial}\n[Final Alarm Log]\n{last_alarm_out}"
                            logging.info(f"ONT Rebooted and Recovery Detected! Actual Reboot Time: {reboot_time_sec:.1f} seconds.")
                        else:
                            res = f"[REBOOT_FAILED] ONT did not recover within timeout (Alarm 'cleared' not found).\n[Reboot Trigger]\n{res_initial}\n[Last Alarm Log]\n{last_alarm_out}"
                            logging.error("Reboot timeout reached!")

            else:
                if ';' in sn['command']:
                    res_lines = []
                    for c in sn['command'].split(';'):
                        c = c.strip()
                        if c:
                            logging.info(f"Executing: {c}")
                            out = self.olt.send_command(c)
                            res_lines.append(out)
                    res = "\n".join(res_lines)
                else:
                    res = self.olt.send_command(sn['command'])
                    
                if sn['type'] == "Software_Info":
                    match = re.search(r'sw-ver-act\s*:\s*(\S+)', res, re.IGNORECASE)
                    if match:
                        self.config['sw_version'] = match.group(1).upper()
            
            # Verify and update DB
            if self._verify(sn['type'], res):
                logging.info(f"[PASS] {sn['type']}")
                self.db.update_test_status(sn['id'], "PASS", res)
            else:
                if sn['type'] == "Speed_Test":
                    logging.warning(f"[{sn['type']}] Did not meet the target criteria. Proceeding with remaining tests.")
                    self.db.update_test_status(sn['id'], "FAIL", res)
                    continue
                
                action = self._handle_failure(sn, res)
                if not action: 
                    self.db.update_test_status(sn['id'], "FAIL", res)
                    return False 
                if action == "SKIP":
                    logging.info(f"[SKIPPED] {sn['type']}")
                    self.db.update_test_status(sn['id'], "N/A", res)
                    continue 

                retry_required = True
                break 
                
        if retry_required:
            logging.warning(f"Test sequence aborted. Initiating re-test in the next cycle.")
            return False
        return True

    def _verify(self, t_type: str, res: str) -> bool:
        if not res:
            if t_type in ["Cleanup"]: return True
            return False
            
        res_lower = res.lower()
        
        if t_type == "Reboot_Test" and "[REBOOT_SUCCESS]" in res:
            return True
        if t_type == "Speed_Test" and "[SPEED_TEST_SUCCESS]" in res:
            return True
            
        error_kws = ["invalid command", "invalid token", "unknown command", "bad parameter", "incomplete command"]
        if any(kw in res_lower for kw in error_kws) or "^" in res: 
            return False
            
        if t_type == "ONT_Discovery_Check":
            serial_clean = self.serial.replace(":", "").lower()
            if serial_clean in res_lower.replace(":", ""): return True
            return False

        if t_type == "Registration_Check":
            serial_clean = self.serial.replace(":", "").lower()
            if "pref-ranged" in res_lower and serial_clean in res_lower.replace(":", ""): return True
            return False
            
        if t_type == "Optics_Check": 
            return "rx-signal" in res_lower and "count : 0" not in res_lower
            
        if t_type == "Software_Info":
            match = re.search(r'sw-ver-act\s*:\s*(\S+)', res_lower)
            if match:
                val = match.group(1)
                if val not in ["sw-ver-psv", "vendor-id", "unknown"]: return True
            return False
            
        if t_type == "UNI_Status":
            has_link = "link-status" in res_lower and "up" in res_lower
            has_speed = "config-indicator" in res_lower and any(s in res_lower for s in ['10g', '1g', '100m', '10m'])
            return has_link and has_speed
            
        if t_type == "MAC_Status": 
            return bool(re.search(r'([0-9A-F]{2}:){5}[0-9A-F]{2}', res, re.I))

        return True

    def _handle_failure(self, sn: Dict, res: str):
        print(f"\n[FAILURE DETECTED] {sn['type']}")
        
        res_lower = res.lower() if res else ""
        error_kws = ["invalid command", "invalid token", "unknown command", "bad parameter", "incomplete command"]
        is_syntax_error = any(kw in res_lower for kw in error_kws) or "^" in res
        
        if is_syntax_error and sn['type'] != "Speed_Test":
            print(f"\n[ AI Auto-Detection ] 'Invalid CLI' detected.")
            while True:
                tmp = input(f"Please enter the correct CLI template for {sn['type']} (or type 'exit'): ").strip()
                if tmp.lower() == 'exit': return False
                if tmp:
                    self.db.save_learned_command(self.olt.olt_profile, sn['type'], tmp)
                    self.db.update_test_case(sn['id'], tmp.format(serial=self.serial, pon=self.pon, ont_id=self.ont_id, port=self.port_full), json.dumps(sn['parameters']))
                    return True
                
        print("\nChoose Action for Error Recovery: [1] Update Command [2] Increase Wait [3] Force Pass [4] Abort")
        choice = input("Select an option [1-4]: ").strip()
        if choice == "1":
            tmp = input(f"Enter new template for {sn['type']}: ").strip()
            if tmp:
                self.db.save_learned_command(self.olt.olt_profile, sn['type'], tmp)
                self.db.update_test_case(sn['id'], tmp.format(serial=self.serial, pon=self.pon, ont_id=self.ont_id, port=self.port_full), json.dumps(sn['parameters']))
            return True
        elif choice == "2":
            sn['parameters']['wait'] = sn['parameters'].get('wait', 0) + 10
            self.db.update_test_case(sn['id'], sn['command'], json.dumps(sn['parameters']))
            return True
        elif choice == "3":
            logging.info(f"Force passing {sn['type']} step.")
            return "SKIP"
        return False

if __name__ == "__main__":
    db = DatabaseManager()
    olt = None
    
    try:
        if os.path.exists(CONFIG_FILE):
            with open(CONFIG_FILE, 'r') as f: config = json.load(f)
        else: 
            config = {}

        if not config.get('olt_ip'):
            config = {
                'olt_ip': input("OLT IP: ").strip(),
                'username': input("Username: ").strip(),
                'password': getpass.getpass("Password (hidden): ").strip(),
                'protocol': input("Protocol (ssh/telnet) [telnet]: ").strip().lower() or 'telnet'
            }
        
        olt = NokiaOLTConnector(config['olt_ip'], config['username'], config['password'], config['protocol'])
        if not olt.connect():
            logging.error("Failed to connect to OLT.")
            sys.exit(1)

        skip_provisioning = False

        while True:
            cmd = db.get_learned_command(olt.olt_profile, "Discovery") or "show channel-pair unprovision-onu"
            onts_list = olt.get_unprovisioned_onts(cmd)
            
            print("\nOptions: [1] Provision Unprovisioned ONT [2] Delete Provisioned ONT [3] Test Provisioned ONT [4] Change Discovery Cmd [5] Exit")
            choice = input("Select an option [1-5]: ").strip()
            
            if choice == "1":
                if not onts_list: continue
                for i, item in enumerate(onts_list, 1): print(f"  [{i}] Serial: {item['serial']} | ChanPair: {item['chanpair']}")
                sel_input = int(input("\nSelect Index: ").strip())
                selected_ont = onts_list[sel_input-1]
                config['ont_serial'] = selected_ont['serial']
                config['chanpair'] = selected_ont['chanpair'] if selected_ont['chanpair'] != "unknown" else input(f"Enter Channel-Pair: ").strip()
                
                def_pon = ""
                def_ont_id = "1"
                
                if config['chanpair'] and config['chanpair'] != "unknown":
                    parts = config['chanpair'].split('/')
                    if len(parts) >= 4:
                        def_pon = f"ng2:{parts[3]}/{parts[2]}"
                        
                    logging.info(f"Scanning channel-pair {config['chanpair']} to find an available ONT ID...")
                    out = olt.send_command(f"show equipment ont status channel-pair {config['chanpair']}", timeout=5)
                    used_ids = set()
                    
                    print(f"\n[ Currently Provisioned ONTs on Channel-Pair {config['chanpair']} ]\n{out}\n")
                    
                    matches = re.findall(r'([a-zA-Z0-9:-]+(?:/\d+)+)\s+([A-Za-z]{4}:?[A-Fa-f0-9]{8})', out)
                    for port_str, serial in matches:
                        last_digit = port_str.split('/')[-1]
                        if last_digit.isdigit():
                            used_ids.add(int(last_digit))
                            
                    for i in range(1, 129):
                        if i not in used_ids:
                            def_ont_id = str(i)
                            break
                
                while True:
                    pon_input = input(f"Target PON [{def_pon}]: ").strip() or def_pon
                    if re.match(r'^(?:[a-zA-Z0-9]+:)?\d+(?:/\d+)+$', pon_input):
                        config['pon'] = pon_input
                        break
                    print("  [!] Invalid format! Please enter a valid PON port (e.g., ng2:3/1 or 1/1/1).")
                    
                while True:
                    ont_id_input = input(f"Target ONT ID [{def_ont_id}]: ").strip() or def_ont_id
                    if ont_id_input.isdigit():
                        config['ont_id'] = ont_id_input
                        break
                    print("  [!] Invalid format! ONT ID must be a number (e.g., 1, 10).")
                break
                
            elif choice == "2":
                active_onts = []
                logging.info("Scanning Channels 1/1/1/1 to 1/1/1/8 for Active ONTs (Oper UP)...")
                for i in range(1, 9):
                    out = olt.send_command(f"show equipment ont status channel-pair 1/1/1/{i}", timeout=5)
                    matches = re.findall(r'(\d+(?:/\d+)+)\s+([a-zA-Z0-9:-]+(?:/\d+)+)\s+([A-Za-z]{4}:?[A-Fa-f0-9]{8})\s+(\S+)\s+(\S+)', out)
                    for m in matches:
                        if m[4].lower() == 'up': 
                            active_onts.append({'chanpair': m[0], 'port': m[1], 'serial': m[2].replace(':', '')})
                
                if active_onts:
                    print("\n[ Active Provisioned ONTs (Oper UP) ]")
                    for i, ont in enumerate(active_onts, 1): 
                        print(f"  [{i}] Serial: {ont['serial']} | Port: {ont['port']} | ChanPair: {ont['chanpair']}")
                else:
                    print("\nNo active ONTs (Oper UP) found in 1/1/1/1 ~ 1/1/1/8.")

                del_target = input("\nSelect Index or enter PORT manually (or 'cancel'): ").strip()
                if del_target.lower() == 'cancel' or not del_target: 
                    continue
                    
                target_port = del_target
                if del_target.isdigit() and 1 <= int(del_target) <= len(active_onts):
                    target_port = active_onts[int(del_target)-1]['port']
                    print(f"-> Selected Serial {active_onts[int(del_target)-1]['serial']} on Port {target_port}.")
                elif '/' not in del_target and ':' not in del_target:
                    found = False
                    for ont in active_onts:
                        if del_target.lower() == ont['serial'].lower():
                            target_port = ont['port']
                            found = True
                            print(f"-> Found Serial {ont['serial']} residing on Port {target_port}.")
                            break
                    if not found:
                        print("-> Could not find that Serial in the active list. If it is offline, please enter the full PORT manually.")
                        continue

                olt._delete_ont_by_port(target_port, db)
                time.sleep(2)
                continue
                
            elif choice == "3":
                active_onts = []
                for i in range(1, 9):
                    out = olt.send_command(f"show equipment ont status channel-pair 1/1/1/{i}", timeout=5)
                    matches = re.findall(r'(\d+(?:/\d+)+)\s+([a-zA-Z0-9:-]+(?:/\d+)+)\s+([A-Za-z]{4}:?[A-Fa-f0-9]{8})\s+(\S+)\s+(\S+)', out)
                    for m in matches:
                        if m[4].lower() == 'up': active_onts.append({'chanpair': m[0], 'port': m[1], 'serial': m[2].replace(':', '')})
                
                def_serial, def_pon, def_ont_id, def_chanpair = "", "", "", ""
                if active_onts:
                    for i, ont in enumerate(active_onts, 1): print(f"  [{i}] Serial: {ont['serial']} | Port: {ont['port']} | ChanPair: {ont['chanpair']}")
                    sel_input = input("\nSelect Index: ").strip()
                    if sel_input.isdigit():
                        selected = active_onts[int(sel_input)-1]
                        def_serial, def_chanpair = selected['serial'], selected['chanpair']
                        def_pon, def_ont_id = selected['port'].rsplit('/', 1)

                config['ont_serial'] = input(f"Enter Serial [{def_serial}]: ").strip() or def_serial
                config['chanpair'] = input(f"Enter Channel-Pair [{def_chanpair}]: ").strip() or def_chanpair
                
                while True:
                    pon_input = input(f"Target PON [{def_pon}]: ").strip() or def_pon
                    if re.match(r'^(?:[a-zA-Z0-9]+:)?\d+(?:/\d+)+$', pon_input):
                        config['pon'] = pon_input
                        break
                    print("  [!] Invalid format! Please enter a valid PON port (e.g., ng2:3/1 or 1/1/1).")
                    
                while True:
                    ont_id_input = input(f"Target ONT ID [{def_ont_id}]: ").strip() or def_ont_id
                    if ont_id_input.isdigit():
                        config['ont_id'] = ont_id_input
                        break
                    print("  [!] Invalid format! ONT ID must be a number (e.g., 1, 10).")
                
                skip_provisioning = True
                break
                
            elif choice == "4":
                new_cmd = input("Enter CORRECT Discovery Command: ").strip()
                if new_cmd: db.save_learned_command(olt.olt_profile, "Discovery", new_cmd)
            elif choice == "5":
                sys.exit(0)

        port_full = f"{config['pon']}/{config['ont_id']}"

        if not skip_provisioning:
            config['ont_type'] = input("ONT Type (sfu/hgu) [default: sfu]: ").strip().lower() or 'sfu'
            config['lan_ports'] = input("Number of LAN ports [1]: ").strip() or "1"
            config['max_mac'] = input("Max Unicast MAC learning [128]: ").strip() or "128"
            if config['ont_type'] == 'sfu':
                config['bw_profile'] = input("Enter Bandwidth Profile Name [NG2DATABWUP10000]: ").strip() or "NG2DATABWUP10000"
            net_input = input("Network Environment (real/traffic) [default: real]: ").strip().lower() or 'real'
            config['vlan_id'] = "1001" if net_input == 'real' else "4000"
            
            vlan_mode_input = input("VLAN Mode [1] Untagged [2] Tagged [3] Translation [default: 1]: ").strip()
            if vlan_mode_input == '2':
                config['vlan_mode'] = 'tagged'
            elif vlan_mode_input == '3':
                config['vlan_mode'] = 'translation'
                config['c_vlan'] = input("Enter Customer VLAN ID (C-VLAN) [default: 10]: ").strip() or "10"
            else:
                config['vlan_mode'] = 'untagged'

        speed_setup = input("Configure Speed Test parameters? (y/n) [default: y]: ").strip().lower()
        if speed_setup != 'n':
            config['iperf_server'] = input(f"iperf3 Server IP [{config.get('iperf_server', '')}]: ").strip() or config.get('iperf_server', '')
            config['iperf_port'] = input(f"iperf3 Server Port [{config.get('iperf_port', '5201')}]: ").strip() or config.get('iperf_port', '5201')
            config['speed_duration'] = input(f"Test Duration in sec [{config.get('speed_duration', '10')}]: ").strip() or config.get('speed_duration', '10')
            config['speed_pass_mbps'] = input(f"Pass Criteria (Mbps) [{config.get('speed_pass_mbps', '500')}]: ").strip() or config.get('speed_pass_mbps', '500')

        with open(CONFIG_FILE, 'w') as f: json.dump(config, f, indent=4)
        db.add_initial_test_cases(config, olt.olt_profile)

        if not skip_provisioning:
            while True:
                if not olt.register_ont(config, db): sys.exit(1)
                if not olt.verify_registration(config, db):
                    olt._cleanup_failed_provisioning(port_full)
                    if input("Retry Registration? (y/n): ").lower() == 'y': continue
                    sys.exit(1)
                if not olt.provision_service_ont(config, db):
                    olt._cleanup_failed_provisioning(port_full)
                    if input("Retry sequence? (y/n): ").lower() == 'y': continue
                    sys.exit(1)
                break 
            db.add_ont(config['ont_serial'], config['pon'], config['ont_id'], "PROVISIONED")

        engine = TestAutomationEngine(db, olt, config)
        for cycle in range(1, 10):
            logging.info(f"========== TEST CYCLE {cycle} ==========")
            if engine.execute_tests(): 
                logging.info("ALL TESTS PASSED SUCCESSFULLY!")
                break
                
        generate_professional_excel_report(config['ont_serial'], db, config)

    except KeyboardInterrupt:
        print("\n[!] Program interrupted by user. Exiting safely...")
    except Exception as e:
        logging.error(f"Unexpected Critical Error: {e}")
    finally:
        if olt: olt.disconnect()
        sys.exit(0)