import json
import sqlite3
from typing import Dict, List


class DatabaseManager:
    def __init__(self, db_name: str = "olt_physical_automation.db"):
        self.db_name = db_name
        self._initialize_database()

    def _get_conn(self):
        return sqlite3.connect(self.db_name, timeout=20)

    def _initialize_database(self):
        with self._get_conn() as conn:
            cursor = conn.cursor()
            cursor.execute(
                """CREATE TABLE IF NOT EXISTS ont_inventory (
                id INTEGER PRIMARY KEY AUTOINCREMENT, serial_number TEXT UNIQUE,
                pon_port TEXT, ont_id TEXT, status TEXT, provisioning_date TIMESTAMP DEFAULT CURRENT_TIMESTAMP)"""
            )

            cursor.execute(
                """CREATE TABLE IF NOT EXISTS test_cases (
                id INTEGER PRIMARY KEY AUTOINCREMENT, target_serial TEXT, execution_order INTEGER,
                test_type TEXT, command TEXT, parameters TEXT, expected_output TEXT,
                success_count INTEGER DEFAULT 0, is_evolved BOOLEAN DEFAULT 0)"""
            )

            cursor.execute(
                """CREATE TABLE IF NOT EXISTS learned_commands (
                olt_profile TEXT, test_type TEXT, command_template TEXT, PRIMARY KEY (olt_profile, test_type))"""
            )

            try:
                cursor.execute("ALTER TABLE test_cases ADD COLUMN status TEXT DEFAULT 'N/T'")
            except Exception:
                pass
            try:
                cursor.execute("ALTER TABLE test_cases ADD COLUMN execution_log TEXT")
            except Exception:
                pass
            conn.commit()

    def add_ont(self, serial: str, pon: str, ont_id: str, status: str):
        with self._get_conn() as conn:
            cursor = conn.cursor()
            cursor.execute(
                "INSERT OR REPLACE INTO ont_inventory (serial_number, pon_port, ont_id, status) VALUES (?, ?, ?, ?)",
                (serial, pon, ont_id, status),
            )
            conn.commit()

    def get_learned_command(self, profile: str, t_type: str) -> str:
        with self._get_conn() as conn:
            cursor = conn.cursor()
            cursor.execute(
                "SELECT command_template FROM learned_commands WHERE olt_profile = ? AND test_type = ?",
                (profile, t_type),
            )
            row = cursor.fetchone()
            return row[0] if row else None

    def save_learned_command(self, profile: str, t_type: str, template: str):
        with self._get_conn() as conn:
            cursor = conn.cursor()
            cursor.execute(
                "INSERT OR REPLACE INTO learned_commands (olt_profile, test_type, command_template) VALUES (?, ?, ?)",
                (profile, t_type, template),
            )
            conn.commit()

    def update_test_case(self, t_id: int, cmd: str, params: str):
        with self._get_conn() as conn:
            cursor = conn.cursor()
            cursor.execute("UPDATE test_cases SET command = ?, parameters = ? WHERE id = ?", (cmd, params, t_id))
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
            cursor.execute(
                "SELECT id, target_serial, execution_order, test_type, command, parameters, status, execution_log, success_count "
                "FROM test_cases WHERE target_serial = ? ORDER BY execution_order ASC",
                (serial,),
            )
            rows = cursor.fetchall()
            return [
                {
                    "id": r[0],
                    "serial": r[1],
                    "order": r[2],
                    "type": r[3],
                    "command": r[4],
                    "parameters": json.loads(r[5]),
                    "status": r[6],
                    "log": r[7],
                    "success_count": r[8],
                }
                for r in rows
            ]

    def add_initial_test_cases(self, config: Dict, profile: str):
        serial = config["ont_serial"]
        pon = config["pon"]
        ont_id = config["ont_id"]
        chanpair = config.get("chanpair", "")
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
                "Optics_Signal_Check": "show equipment ont optics {port}",
                "Optics_Temp_Check": "show equipment ont optics {port}",
                "Optics_Voltage_Laser_Check": "show equipment ont optics {port}",
                "UNI_Status": "show ethernet ont operational-data {port}/1/1",
                "Software_Info": "show equipment ont interface {port} detail",
                "Speed_Test": "Native iperf3 Execution",
                "MAC_Status": "show vlan bridge-port-fdb {port}/1/1",
                "Cleanup": "configure equipment ont interface {port} admin-state down ; configure equipment ont no interface {port}",
            }

            cursor.execute(
                "SELECT test_type, command_template FROM learned_commands WHERE olt_profile = ?", (profile,)
            )
            learned_dict = {row[0]: row[1] for row in cursor.fetchall()}

            mock_cases = []
            for i, (t_type, def_cmd) in enumerate(templates.items(), 1):
                template = learned_dict.get(t_type)

                if t_type == "Reboot_Test" and template and "with-active-image" not in template:
                    template = def_cmd
                    cursor.execute(
                        "INSERT OR REPLACE INTO learned_commands (olt_profile, test_type, command_template) VALUES (?, ?, ?)",
                        (profile, t_type, def_cmd),
                    )
                elif t_type == "UNI_Status" and template and "operational-data" not in template:
                    template = def_cmd
                    cursor.execute(
                        "INSERT OR REPLACE INTO learned_commands (olt_profile, test_type, command_template) VALUES (?, ?, ?)",
                        (profile, t_type, def_cmd),
                    )
                elif "{port}" in def_cmd and template and "{port}" not in template:
                    template = def_cmd
                    cursor.execute(
                        "INSERT OR REPLACE INTO learned_commands (olt_profile, test_type, command_template) VALUES (?, ?, ?)",
                        (profile, t_type, def_cmd),
                    )
                elif "{chanpair}" in def_cmd and template and "{chanpair}" not in template:
                    template = def_cmd
                    cursor.execute(
                        "INSERT OR REPLACE INTO learned_commands (olt_profile, test_type, command_template) VALUES (?, ?, ?)",
                        (profile, t_type, def_cmd),
                    )

                if not template:
                    template = def_cmd

                cmd = template.format(
                    serial=serial,
                    serial_formatted=serial_formatted,
                    pon=pon,
                    ont_id=ont_id,
                    port=port_full,
                    chanpair=chanpair,
                )

                params = {"timeout": 15, "wait": 0}
                if t_type == "Registration_Check":
                    params["wait"] = 5
                elif t_type == "MAC_Status":
                    params["wait"] = 10

                mock_cases.append((serial, i, t_type, cmd, json.dumps(params), "N/T"))

            cursor.executemany(
                "INSERT INTO test_cases (target_serial, execution_order, test_type, command, parameters, status) VALUES (?,?,?,?,?,?)",
                mock_cases,
            )
            conn.commit()
