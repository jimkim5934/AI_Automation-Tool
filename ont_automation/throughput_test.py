"""Throughput Test — iperf3 or Ixia instrument execution and result parsing."""

import logging
import re
import shlex
import subprocess
import time
from typing import Dict, Tuple

from ont_automation.gui_helpers import gui_ask

TOOL_IPERF3 = "iperf3"
TOOL_IXIA = "ixia"


def get_live_ip(target_mac: str, target_desc: str) -> str:
    """Dynamically retrieves the current IP address of the specified NIC (Windows WMIC)."""
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
                if current_nic:
                    mac = current_nic.get("MACAddress", "")
                    desc = current_nic.get("Description", "")
                    ip = current_nic.get("IPAddress", "")
                    if (target_mac and target_mac == mac) or (target_desc and target_desc == desc):
                        return ip
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
        if current_nic:
            mac = current_nic.get("MACAddress", "")
            desc = current_nic.get("Description", "")
            ip = current_nic.get("IPAddress", "")
            if (target_mac and target_mac == mac) or (target_desc and target_desc == desc):
                return ip
    except Exception as e:
        logging.debug(f"Failed to query wmic for live IP: {e}")
    return ""


def build_iperf_base_args(config: Dict) -> list:
    """Build iperf3 argument list including NIC bind and custom options."""
    server = config.get("iperf_server", "").strip()
    port = str(config.get("iperf_port", 5201))
    duration = str(config.get("speed_duration", 10))
    base_args = ["iperf3", "-c", server, "-p", port, "-t", duration]

    bind_nic_raw = config.get("bind_nic", "Default (Auto)")
    if bind_nic_raw != "Default (Auto)":
        parts = bind_nic_raw.split("|")
        original_ip = parts[0].strip()
        target_mac = parts[1].strip() if len(parts) > 1 else ""
        target_desc = parts[2].strip() if len(parts) > 2 else ""

        valid_ip_found = False

        if not target_mac and not target_desc:
            if original_ip.startswith("169.254."):
                raise Exception(
                    f"Manually entered IP {original_ip} is an APIPA address. Speed test aborted."
                )
            base_args.extend(["-B", original_ip])
            valid_ip_found = True
        else:
            for attempt in range(1, 6):
                live_ip = get_live_ip(target_mac, target_desc)
                if live_ip and not live_ip.startswith("169.254.") and live_ip != "0.0.0.0":
                    logging.info(f"Valid IP detected for bound NIC: {live_ip}")
                    base_args.extend(["-B", live_ip])
                    valid_ip_found = True
                    break
                logging.warning(
                    f"[Attempt {attempt}/5] Bound NIC IP is invalid ({live_ip or 'None'}). Waiting 60 seconds..."
                )
                if attempt < 5:
                    time.sleep(60)

            if not valid_ip_found:
                raise Exception(
                    "Failed to obtain a valid DHCP IP (non-APIPA) for the bound NIC within the timeout."
                )

    custom_opts = config.get("iperf_options", "-P 8").strip()
    if custom_opts:
        try:
            base_args.extend(shlex.split(custom_opts))
        except ValueError:
            base_args.extend(custom_opts.split())

    return base_args


def resolve_throughput_tool(config: Dict) -> str:
    """
    Determine whether to use iperf3 or Ixia.

    Uses config['throughput_tool'] when set to 'iperf3' or 'ixia'.
    Value 'ask' (or unset with legacy configs) prompts the operator at run time.
    """
    tool = (config.get("throughput_tool") or "ask").strip().lower()
    if tool in (TOOL_IPERF3, TOOL_IXIA):
        return tool

    choice = gui_ask(
        "Throughput test instrument:\n"
        "[1] iPerf3\n"
        "[2] Ixia\n\n"
        "Enter 1 or 2 (or 'exit' to abort):"
    )
    if choice.lower() == "exit":
        raise Exception("Throughput test aborted by user.")

    if choice in ("2", "ixia"):
        logging.info("Throughput tool selected: Ixia")
        return TOOL_IXIA

    logging.info("Throughput tool selected: iPerf3")
    return TOOL_IPERF3


def run_ixia_throughput_test(config: Dict) -> Tuple[str, float, float]:
    """Ixia instrument throughput test — placeholder for future implementation."""
    # TODO: Ixia API / chassis session integration
    logging.warning("Ixia throughput test selected but not yet implemented.")
    return "[SPEED_TEST_FAILED] Ixia throughput test is not implemented yet.", 0.0, 0.0


def run_iperf3_throughput_test(config: Dict) -> Tuple[str, float, float]:
    """
    Run upload and download iperf3 tests.

    Returns:
        (result_text, overall_upload_mbps, overall_download_mbps)
    """
    server = config.get("iperf_server", "").strip()
    criteria = float(config.get("speed_pass_mbps", 500))

    if not server:
        return "[SPEED_TEST_SUCCESS] Skipped (No server)", 0.0, 0.0

    overall_up = 0.0
    overall_dl = 0.0
    full_log = []

    base_args = build_iperf_base_args(config)
    tests = [("UPLOAD", base_args), ("DOWNLOAD", base_args + ["-R"])]

    for t_name, t_args in tests:
        logging.info(f"Running Native iperf3 {t_name}: {' '.join(t_args)}")
        print("\n" + "=" * 68)
        print(f"[{t_name}] {'Time':<8} | {'Progress Graph':<42} | {'Throughput':<15}")
        print("=" * 68)

        full_log.append(f"--- STARTING {t_name} TEST ---\n")

        proc = subprocess.Popen(
            t_args, stdout=subprocess.PIPE, stderr=subprocess.STDOUT, text=True, bufsize=1
        )
        direction_log = []

        for line in iter(proc.stdout.readline, ""):
            direction_log.append(line)
            full_log.append(line)
            if "[SUM]" in line and "bits/sec" in line and "sender" not in line and "receiver" not in line:
                match = re.search(
                    r"\[SUM\]\s+\d+\.\d+-\s*(\d+\.\d+)\s+sec.*?\s+(\d+(?:\.\d+)?)\s+([KMG])bits/sec",
                    line,
                )
                if match:
                    sec_val = match.group(1)
                    val = float(match.group(2))
                    unit = match.group(3)
                    mbps = val * 1000 if unit == "G" else (val if unit == "M" else val / 1000)

                    bar_len = 40
                    max_val = max(criteria * 1.2, 10000)
                    filled = min(int((mbps / max_val) * bar_len), bar_len)
                    bar = "█" * filled + "-" * (bar_len - filled)

                    print(f"[{sec_val:>5}s] [{bar}] {mbps:8.2f} Mbps")

        proc.wait()
        print("=" * 68 + "\n")

        dir_text = "".join(direction_log)
        if proc.returncode != 0:
            logging.warning(f"{t_name} Speed Test returned code {proc.returncode}")

        if t_name == "UPLOAD":
            sender_matches = re.findall(
                r"\[SUM\].*?\s+(\d+(?:\.\d+)?)\s+([KMG])bits/sec\s+sender", dir_text
            )
            if sender_matches:
                v, u = float(sender_matches[-1][0]), sender_matches[-1][1]
                overall_up = v * 1000 if u == "G" else (v if u == "M" else v / 1000)
        else:
            recv_matches = re.findall(
                r"\[SUM\].*?\s+(\d+(?:\.\d+)?)\s+([KMG])bits/sec\s+receiver", dir_text
            )
            if recv_matches:
                v, u = float(recv_matches[-1][0]), recv_matches[-1][1]
                overall_dl = v * 1000 if u == "G" else (v if u == "M" else v / 1000)

        time.sleep(2)

    out_text = "".join(full_log)

    if overall_up >= criteria and overall_dl >= criteria:
        res = (
            f"[SPEED_TEST_SUCCESS] Upload: {overall_up:.2f} Mbps, Download: {overall_dl:.2f} Mbps\n{out_text}"
        )
        logging.info(
            f"Speed test passed criteria! (Up: {overall_up:.2f} Mbps, Down: {overall_dl:.2f} Mbps)"
        )
    else:
        res = (
            f"[SPEED_TEST_FAILED] Upload: {overall_up:.2f} Mbps, Download: {overall_dl:.2f} Mbps "
            f"(Target: {criteria} Mbps)\n{out_text}"
        )

    return res, overall_up, overall_dl


def run_throughput_test(config: Dict) -> Tuple[str, float, float]:
    """
    Run throughput test using the configured or selected instrument.

    Returns:
        (result_text, overall_upload_mbps, overall_download_mbps)
    """
    tool = resolve_throughput_tool(config)

    if tool == TOOL_IXIA:
        return run_ixia_throughput_test(config)

    return run_iperf3_throughput_test(config)
