import sys
import time
import argparse
import subprocess
import urllib.request
import traceback

def print_log(msg):
    """Ensure messages are immediately flushed to the console."""
    print(f"INFO: {msg}", flush=True)

def run_iperf3_test(server_ip, port, duration):
    """Runs an iperf3 client test against the specified server."""
    print_log(f"Starting iperf3 test against {server_ip}:{port} for {duration} seconds...")
    try:
        # Run iperf3 in JSON output mode for easy parsing
        cmd = ['iperf3', '-c', server_ip, '-p', str(port), '-t', str(duration), '-J']
        
        # Capture the output
        result = subprocess.run(cmd, capture_output=True, text=True)
        
        if result.returncode != 0:
            print_log(f"iperf3 failed with error:\n{result.stderr}")
            return False, 0.0, 0.0

        import json
        data = json.loads(result.stdout)
        
        # Extract bits per second and convert to Mbps
        # Check if the server responded properly
        if 'end' in data and 'sum_received' in data['end']:
            bps_sent = data['end']['sum_sent']['bits_per_second']
            bps_recv = data['end']['sum_received']['bits_per_second']
            
            mbps_sent = bps_sent / 1_000_000
            mbps_recv = bps_recv / 1_000_000
            
            print_log(f"iperf3 Test Completed. Upload: {mbps_sent:.2f} Mbps | Download: {mbps_recv:.2f} Mbps")
            return True, mbps_sent, mbps_recv
        else:
            print_log("iperf3 execution succeeded, but expected JSON structure was not found.")
            return False, 0.0, 0.0
            
    except FileNotFoundError:
        print_log("CRITICAL ERROR: 'iperf3' command not found.")
        print_log("Please ensure iperf3 is installed and added to your Windows system PATH.")
        return False, 0.0, 0.0
    except Exception as e:
        print_log(f"iperf3 Execution Error: {e}")
        return False, 0.0, 0.0

def run_openspeedtest_download(server_url, duration):
    """Simulates an OpenSpeedTest HTTP download using a dummy payload."""
    print_log(f"Starting OpenSpeedTest HTTP Download Test against {server_url}...")
    
    # Standard OpenSpeedTest payload endpoint
    download_url = f"{server_url.rstrip('/')}/downloading"
    
    start_time = time.time()
    total_bytes = 0
    
    print_log(f"Downloading payload for max {duration} seconds...")
    try:
        req = urllib.request.Request(download_url)
        with urllib.request.urlopen(req, timeout=10) as response:
            while time.time() - start_time < duration:
                # Read chunks of 1MB
                chunk = response.read(1024 * 1024)
                if not chunk:
                    break
                total_bytes += len(chunk)
                
        elapsed_time = time.time() - start_time
        mbps = (total_bytes * 8) / (elapsed_time * 1_000_000)
        
        print_log(f"OpenSpeedTest Download Completed. Average Speed: {mbps:.2f} Mbps over {elapsed_time:.2f} seconds")
        return True, mbps
        
    except Exception as e:
        print_log(f"OpenSpeedTest Execution Error: {e}")
        return False, 0.0

if __name__ == "__main__":
    parser = argparse.ArgumentParser(description="Real Network Speed Test Script")
    parser.add_argument("--iperf_server", type=str, required=False, default="", help="iperf3 Server IP")
    parser.add_argument("--iperf_port", type=int, required=False, default=5201, help="iperf3 Server Port")
    parser.add_argument("--ost_server", type=str, required=False, default="", help="OpenSpeedTest Server URL (e.g., http://192.168.1.100:3000)")
    parser.add_argument("--duration", type=int, required=True, help="Test duration in seconds per test")
    parser.add_argument("--pass_criteria", type=int, required=False, default=500, help="Pass criteria in Mbps")
    
    args = parser.parse_args()
    
    iperf_success = True
    ost_success = True
    
    overall_sent_mbps = 0.0
    overall_recv_mbps = 0.0
    ost_mbps = 0.0

    print_log("==================================================")
    print_log("         REAL NETWORK SPEED TEST STARTED          ")
    print_log("==================================================")

    # 1. Execute iperf3 if server IP is provided
    if args.iperf_server:
        iperf_success, overall_sent_mbps, overall_recv_mbps = run_iperf3_test(args.iperf_server, args.iperf_port, args.duration)
        if not iperf_success:
            print_log("WARNING: iperf3 test failed or was skipped due to errors.")
    
    # 2. Execute OpenSpeedTest HTTP Download if URL is provided
    if args.ost_server:
        time.sleep(2) # Brief pause between tests
        ost_success, ost_mbps = run_openspeedtest_download(args.ost_server, args.duration)
        if not ost_success:
            print_log("WARNING: OpenSpeedTest HTTP download failed.")

    if not args.iperf_server and not args.ost_server:
        print_log("ERROR: Neither iperf3 server nor OpenSpeedTest server was provided.")
        sys.exit(1)

    print_log("--------------------------------------------------")
    print_log("              FINAL SPEED TEST RESULTS            ")
    print_log("--------------------------------------------------")
    if args.iperf_server:
        print_log(f"  iperf3 Upload   : {overall_sent_mbps:.2f} Mbps")
        print_log(f"  iperf3 Download : {overall_recv_mbps:.2f} Mbps")
    if args.ost_server:
        print_log(f"  OST HTTP Download: {ost_mbps:.2f} Mbps")
    print_log("--------------------------------------------------")

    # Determine Pass/Fail based on the pass criteria
    passed = True
    if args.iperf_server and (overall_recv_mbps < args.pass_criteria and overall_sent_mbps < args.pass_criteria):
        passed = False
    if args.ost_server and ost_mbps < args.pass_criteria:
        passed = False

    if iperf_success and ost_success and passed:
        print_log("[SPEED_TEST_SUCCESS] Throughput meets the defined criteria.")
        sys.exit(0)
    else:
        print_log(f"[SPEED_TEST_FAILED] Throughput is below the target ({args.pass_criteria} Mbps) or a test failed.")
        sys.exit(1)