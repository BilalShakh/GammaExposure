import subprocess
import sys

# Run parse_options_data.py with month 1 (MAR 26)
print("Step 1: Extracting options data for MAR 26...")
print("-" * 80)

process = subprocess.Popen(
    [sys.executable, 'parse_options_data.py'],
    stdin=subprocess.PIPE,
    stdout=subprocess.PIPE,
    stderr=subprocess.PIPE,
    text=True
)

# Send choice "1" to select MAR 26
stdout, stderr = process.communicate(input="1\n")
print(stdout)
if stderr:
    print("STDERR:", stderr)

# Run calculate_gex.py
print("\nStep 2: Calculating and graphing GEX...")
print("-" * 80)
process2 = subprocess.Popen(
    [sys.executable, 'calculate_gex.py'],
    stdout=subprocess.PIPE,
    stderr=subprocess.PIPE,
    text=True
)

stdout2, stderr2 = process2.communicate()
print(stdout2)
if stderr2:
    print("STDERR:", stderr2)

print("\nWorkflow complete!")
