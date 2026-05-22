"""ONT 초기화 — delete ONT and cleanup after failed provisioning."""

import logging
import time
from typing import TYPE_CHECKING

from ont_automation.database import DatabaseManager

if TYPE_CHECKING:
    pass


class ONTCleanupMixin:
    """Mixin for ONT deletion and provisioning rollback."""

    def _get_safe_delete_cmd(self, db: DatabaseManager) -> str:
        del_cmd_temp = db.get_learned_command(self.olt_profile, "Delete_Provisioned")
        default_cmd = (
            "configure equipment ont interface {port} admin-state down ; "
            "configure equipment ont no interface {port}"
        )
        if not del_cmd_temp or "{port}" not in del_cmd_temp:
            db.save_learned_command(self.olt_profile, "Delete_Provisioned", default_cmd)
            return default_cmd
        return del_cmd_temp

    def _delete_ont_by_port(self, port_full: str, db: DatabaseManager):
        logging.info(f"Aggressively deleting ONT bound to port: {port_full}...")
        del_cmd_temp = self._get_safe_delete_cmd(db)

        cmds = [c.strip() for c in del_cmd_temp.split(";") if c.strip()]
        for c in cmds:
            exec_cmd = c.format(port=port_full)
            self.send_command(exec_cmd, timeout=5)

        time.sleep(2)
        logging.info(f"Cleanup complete for port {port_full}.")

    def _cleanup_failed_provisioning(self, port_full: str):
        db = DatabaseManager()
        self._delete_ont_by_port(port_full, db)
