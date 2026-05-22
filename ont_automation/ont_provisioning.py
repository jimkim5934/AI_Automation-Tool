"""ONT Provisioning — service VLAN/QoS configuration on registered ONT."""

import logging
from typing import Dict

from ont_automation.database import DatabaseManager
from ont_automation.gui_helpers import gui_ask


class ONTProvisioningMixin:
    """Mixin for ONT service provisioning CLI sequences."""

    def provision_service_ont(self, config: Dict, db: DatabaseManager) -> bool:
        port_full = f"{config['pon']}/{config['ont_id']}"
        vlan_mode = config.get("vlan_mode", "untagged").lower()
        prov_key = f"Provisioning_Service_{config['ont_type'].upper()}_{vlan_mode.upper()}"

        learned_cmds = db.get_learned_command(self.olt_profile, prov_key)

        if learned_cmds and "vlan-id" not in learned_cmds.lower():
            logging.warning(
                f"Learned provisioning command for {vlan_mode} seems incomplete. Reverting to default sequence."
            )
            learned_cmds = None

        if learned_cmds:
            cmd_templates = [c.strip() for c in learned_cmds.split(";") if c.strip()]
        else:
            base_sfu = [
                "configure qos interface uni:{port}/1/1 queue [0...7] shaper-profile name:StrictPriority",
                "configure qos interface uni:{port}/1/1 upstream-queue [0...7] bandwidth-profile name:{bw_profile} bandwidth-sharing uni-sharing",
                "configure bridge port {port}/1/1 max-unicast-mac {max_mac} pvid-tagging-flag olt",
            ]
            base_hgu = [
                "configure bridge port {port}/14/1 max-unicast-mac {max_mac} pvid-tagging-flag olt"
            ]

            if config["ont_type"] == "sfu":
                cmd_templates = base_sfu
                if vlan_mode == "untagged":
                    cmd_templates.extend(
                        [
                            "configure bridge port {port}/1/1 vlan-id {vlan_id} usacceptframetype untagged",
                            "configure bridge port {port}/1/1 pvid {vlan_id}",
                        ]
                    )
                elif vlan_mode == "tagged":
                    cmd_templates.append(
                        "configure bridge port {port}/1/1 vlan-id {vlan_id} tag single-tagged l2fwder-vlan {vlan_id} vlan-scope local"
                    )
                elif vlan_mode == "translation":
                    cmd_templates.append(
                        "configure bridge port {port}/1/1 vlan-id {c_vlan} tag single-tagged l2fwder-vlan {vlan_id} vlan-scope local"
                    )
            else:
                cmd_templates = base_hgu
                if vlan_mode == "untagged":
                    cmd_templates.extend(
                        [
                            "configure bridge port {port}/14/1 vlan-id {vlan_id} usacceptframetype untagged",
                            "configure bridge port {port}/14/1 pvid {vlan_id}",
                        ]
                    )
                elif vlan_mode == "tagged":
                    cmd_templates.append(
                        "configure bridge port {port}/14/1 vlan-id {vlan_id} tag single-tagged l2fwder-vlan {vlan_id} vlan-scope local"
                    )
                elif vlan_mode == "translation":
                    cmd_templates.append(
                        "configure bridge port {port}/14/1 vlan-id {c_vlan} tag single-tagged l2fwder-vlan {vlan_id} vlan-scope local"
                    )

        logging.info(f"--- [STEP 3] Initiating Service Provisioning ({vlan_mode.upper()}) on {port_full} ---")
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
                    port=port_full,
                    bw_profile=config.get("bw_profile", ""),
                    max_mac=config.get("max_mac", "128"),
                    vlan_id=config.get("vlan_id", "1001"),
                    c_vlan=config.get("c_vlan", "10"),
                )

                out = self.send_command(cmd, timeout=10)
                out_lower = out.lower()
                if (any(kw in out_lower for kw in error_kws) or "^" in out) and "pattern not detected" not in out_lower:
                    print(f"\n[ SERVICE PROVISIONING FAILED ] Command rejected: {cmd}")
                    self._cleanup_failed_provisioning(port_full)
                    new_cmds = gui_ask("Service Provisioning failed. Enter corrective CLI sequence:")
                    if new_cmds.lower() == "exit" or not new_cmds:
                        return False
                    db.save_learned_command(self.olt_profile, prov_key, new_cmds)
                    return False

            if success:
                return True
