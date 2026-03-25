import sys
import time
import argparse
import traceback
import os

# 현재 실행 중인 폴더(AI_Automation-Tool)를 경로에 추가하여 완벽하게 튜닝된 StcPython.py를 읽어오게 합니다.
sys.path.append(os.path.dirname(os.path.abspath(__file__)))

def print_log(msg):
    """Ensure messages are immediately flushed to the console."""
    print(f"INFO: {msg}", flush=True)

try:
    from StcPython import StcPython
    STC_AVAILABLE = True
except ImportError:
    STC_AVAILABLE = False
    print_log("WARNING: Spirent TestCenter API (StcPython.py) is not found in the current directory.")

def run_stc_traffic(chassis_ip, tx_port_loc, rx_port_loc, vlan_id, frame_size, load_percent, duration):
    if not STC_AVAILABLE:
        print_log("CRITICAL ERROR: Cannot proceed without StcPython module.")
        return False

    stc = StcPython()
    print_log("Initializing Spirent TestCenter API...")
    
    try:
        # 버전 로드 확인
        stc_version = stc.get('system1', 'version')
        print_log(f"Spirent TestCenter Version Loaded: {stc_version}")

        project = stc.get("system1", "children-Project")
        if not project:
            project = stc.create("project", "system1")
            
        print_log(f"Connecting to STC Chassis at {chassis_ip}...")
        stc.perform("ConnectToChassis", ChassisList=chassis_ip)
        
        print_log(f"Reserving TX Port: {tx_port_loc} and RX Port: {rx_port_loc}...")
        tx_port = stc.create("port", project, Location=f"//{chassis_ip}/{tx_port_loc}")
        rx_port = stc.create("port", project, Location=f"//{chassis_ip}/{rx_port_loc}")
        
        stc.perform("AttachPorts", PortList=f"{tx_port} {rx_port}", AutoConnect="TRUE")
        stc.perform("ApplyToMultiple", EndpointList=project)
        
        tx_gen = stc.get(tx_port, "children-generator")
        rx_ana = stc.get(rx_port, "children-analyzer")
        
        print_log(f"Configuring Traffic Stream (VLAN: {vlan_id}, Frame Size: {frame_size}B, Load: {load_percent}%)")
        # Create StreamBlock on TX port
        stream_block = stc.create("streamBlock", tx_port, InsertSig="TRUE", FrameLengthMode="FIXED", FixedFrameLength=frame_size)
        
        # Add VLAN Header if VLAN is specified and not 0
        if vlan_id and str(vlan_id) != "0":
            print_log(f"Applying VLAN ID {vlan_id} to the stream...")
            eth_header = stc.get(stream_block, "children-ethernet:EthernetII")
            stc.create("vlans", eth_header, VlanId=vlan_id)
        
        # Configure Load and Duration
        gen_config = stc.get(tx_gen, "children-GeneratorConfig")
        stc.config(gen_config, DurationMode="SECONDS", Duration=duration, LoadMode="PERCENT_LINE_RATE", FixedLoad=load_percent)
        
        # Subscribe to Results
        stc.perform("ResultsSubscribe", Parent=project, ConfigType="Generator", ResultType="GeneratorPortResults")
        stc.perform("ResultsSubscribe", Parent=project, ConfigType="Analyzer", ResultType="AnalyzerPortResults")
        stc.perform("ApplyToMultiple", EndpointList=project)
        
        print_log("Starting STC Analyzer and Traffic Generator...")
        stc.perform("AnalyzerStart", AnalyzerList=rx_ana)
        stc.perform("GeneratorStart", GeneratorList=tx_gen)
        
        print_log(f">>> Traffic is RUNNING for {duration} seconds. Please wait... <<<")
        time.sleep(duration + 3) # Wait for traffic to finish + buffer
        
        print_log("Stopping STC Traffic...")
        stc.perform("GeneratorStop", GeneratorList=tx_gen)
        stc.perform("AnalyzerStop", AnalyzerList=rx_ana)
        
        tx_results = stc.get(tx_gen, "children-GeneratorPortResults")
        rx_results = stc.get(rx_ana, "children-AnalyzerPortResults")
        
        tx_count = int(stc.get(tx_results, "GeneratorFrameCount"))
        rx_count = int(stc.get(rx_results, "AnalyzerFrameCount"))
        dropped = tx_count - rx_count
        
        print_log("--------------------------------------------------")
        print_log("             STC TRAFFIC TEST RESULTS             ")
        print_log("--------------------------------------------------")
        print_log(f"  TX Frames Transmitted : {tx_count}")
        print_log(f"  RX Frames Received    : {rx_count}")
        print_log(f"  Dropped Frames        : {dropped}")
        print_log("--------------------------------------------------")
        
        stc.perform("DisconnectFromChassis", ChassisList=chassis_ip)
        stc.perform("ResetConfig")
        
        # Verification Logic: TX가 있고, RX가 99% 이상 도달했으면 PASS
        if tx_count > 0 and rx_count >= (tx_count * 0.99):
            print_log("[STC_EXECUTION_SUCCESS] Traffic test passed criteria.")
            return True
        else:
            print_log("[STC_EXECUTION_FAILED] High frame drop or no traffic generated.")
            return False

    except Exception as e:
        print_log(f"STC Execution Error: {e}")
        print_log(traceback.format_exc())
        return False

if __name__ == "__main__":
    parser = argparse.ArgumentParser(description="STC Automated Traffic Script")
    parser.add_argument("--chassis", required=True, help="STC Chassis IP")
    parser.add_argument("--tx_port", required=True, help="TX Port Location (e.g. 1/1)")
    parser.add_argument("--rx_port", required=True, help="RX Port Location (e.g. 1/2)")
    parser.add_argument("--vlan", required=True, help="VLAN ID for the stream")
    parser.add_argument("--framesize", type=int, required=True, help="Frame size in bytes")
    parser.add_argument("--load", type=int, required=True, help="Traffic load percentage")
    parser.add_argument("--duration", type=int, required=True, help="Test duration in seconds")
    
    args = parser.parse_args()
    
    is_success = run_stc_traffic(args.chassis, args.tx_port, args.rx_port, args.vlan, args.framesize, args.load, args.duration)
    sys.exit(0 if is_success else 1)