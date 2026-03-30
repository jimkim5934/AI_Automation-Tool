import sys
import argparse
import subprocess
import re

def print_log(msg):
    """Ensure messages are immediately flushed to the console."""
    print(f"INFO: {msg}", flush=True)

def run_iperf3_test(server_ip, port, duration, pass_criteria):
    """Runs an iperf3 client test against the specified server using 8 parallel streams."""
    print_log(f"Starting iperf3 test against {server_ip}:{port} for {duration} seconds (8 Streams)...")
    try:
        cmd = ['iperf3', '-c', server_ip, '-p', str(port), '-t', str(duration), '-P', '8']
        print_log(f"Executing CLI: {' '.join(cmd)}")
        
        # Draw table header for graph
        print_log("=" * 68)
        print_log(f"{'Time':<8} | {'Progress Graph':<42} | {'Throughput':<15}")
        print_log("=" * 68)
        
        process = subprocess.Popen(cmd, stdout=subprocess.PIPE, stderr=subprocess.STDOUT, text=True, bufsize=1)
        
        output_lines = []
        overall_sent_mbps = 0.0
        overall_recv_mbps = 0.0
        
        for line in iter(process.stdout.readline, ''):
            output_lines.append(line)
            
            # Catch the live [SUM] line to draw the ASCII bar graph
            if "[SUM]" in line and "bits/sec" in line and "sender" not in line and "receiver" not in line:
                match = re.search(r'\[SUM\]\s+\d+\.\d+-\s*(\d+\.\d+)\s+sec.*?\s+(\d+(?:\.\d+)?)\s+([KMG])bits/sec', line)
                if match:
                    sec = match.group(1)
                    val = float(match.group(2))
                    unit = match.group(3)
                    mbps = val * 1000 if unit == 'G' else (val if unit == 'M' else val / 1000)
                    
                    bar_len = 40
                    max_val = max(pass_criteria * 1.2, 10000) # Dynamic scale based on target
                    filled = min(int((mbps / max_val) * bar_len), bar_len)
                    bar = '█' * filled + '-' * (bar_len - filled)
                    
                    # Print the live graph line
                    print(f"INFO: [{sec:>5}s] [{bar}] {mbps:8.2f} Mbps", flush=True)

        process.wait()
        print_log("=" * 68)
        
        if process.returncode != 0:
            print_log("iperf3 failed. Please check server reachability or port.")
            return False, 0.0, 0.0

        output_text = "".join(output_lines)
        
        # Extract final aggregates
        sender_matches = re.findall(r'\[SUM\].*?\s+(\d+(?:\.\d+)?)\s+([KMG])bits/sec\s+sender', output_text)
        recv_matches = re.findall(r'\[SUM\].*?\s+(\d+(?:\.\d+)?)\s+([KMG])bits/sec\s+receiver', output_text)
        
        if sender_matches:
            val, unit = float(sender_matches[-1][0]), sender_matches[-1][1]
            overall_sent_mbps = val * 1000 if unit == 'G' else (val if unit == 'M' else val / 1000)
            
        if recv_matches:
            val, unit = float(recv_matches[-1][0]), recv_matches[-1][1]
            overall_recv_mbps = val * 1000 if unit == 'G' else (val if unit == 'M' else val / 1000)
            
        print_log(f"iperf3 Test Completed. Upload: {overall_sent_mbps:.2f} Mbps | Download: {overall_recv_mbps:.2f} Mbps")
        return True, overall_sent_mbps, overall_recv_mbps
        
    except FileNotFoundError:
        print_log("CRITICAL ERROR: 'iperf3' command not found.")
        return False, 0.0, 0.0
    except Exception as e:
        print_log(f"iperf3 Execution Error: {e}")
        return False, 0.0, 0.0

if __name__ == "__main__":
    parser = argparse.ArgumentParser(description="Real Network Speed Test Script")
    parser.add_argument("--iperf_server", type=str, required=True, help="iperf3 Server IP")
    parser.add_argument("--iperf_port", type=int, required=False, default=5201, help="iperf3 Server Port")
    parser.add_argument("--duration", type=int, required=True, help="Test duration in seconds per test")
    parser.add_argument("--pass_criteria", type=int, required=False, default=500, help="Pass criteria in Mbps")
    
    args = parser.parse_args()

    success, sent_mbps, recv_mbps = run_iperf3_test(args.iperf_server, args.iperf_port, args.duration, args.pass_criteria)

    print_log("--------------------------------------------------")
    print_log("              FINAL SPEED TEST RESULTS            ")
    print_log("--------------------------------------------------")
    print_log(f"  iperf3 Upload   : {sent_mbps:.2f} Mbps")
    print_log(f"  iperf3 Download : {recv_mbps:.2f} Mbps")
    print_log("--------------------------------------------------")

    if success and (sent_mbps >= args.pass_criteria or recv_mbps >= args.pass_criteria):
        print_log("[SPEED_TEST_SUCCESS] Throughput meets the defined criteria.")
        sys.exit(0)
    else:
        print_log(f"[SPEED_TEST_FAILED] Throughput is below the target ({args.pass_criteria} Mbps) or failed.")
        sys.exit(1)