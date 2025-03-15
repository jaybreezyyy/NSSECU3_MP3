import os
import subprocess
import pandas as pd

# Define tool paths (Adjust these paths if needed)
TOOLS_DIR = "C:\\Users\\Andrei\\Downloads\\combinedtool\\combinedtool"  # Change to your tools directory
EVTX_CMD = os.path.join(TOOLS_DIR, "EvtxECmd.exe")
MFT_CMD = os.path.join(TOOLS_DIR, "MFTECmd.exe")
RE_CMD = os.path.join(TOOLS_DIR, "RECmd.exe")

# Input files
evtx_file = "sysmon_11_1_lolbas_downldr_desktopimgdownldr.evtx"
mft_file = "SampleMFTrecords1.bin"
registry_file = "NTUSER.DAT"
batch_file = "Kroll_Batch.reb"

# Output directories
output_dir = os.path.join(TOOLS_DIR, "output")
os.makedirs(output_dir, exist_ok=True)  # Ensure output directory exists

# Output file paths
evtx_output = os.path.join(output_dir, "evtx_results.csv")
mft_output = os.path.join(output_dir, "mft_results.csv")
reg_output = os.path.join(output_dir, "reg_results.csv")
final_output = os.path.join(output_dir, "combined_results.csv")


def run_tool(command):
    """Runs an external command and handles errors."""
    try:
        print(f"Running: {command}")
        subprocess.run(command, shell=True, check=True)
    except subprocess.CalledProcessError as e:
        print(f"Error running {command}: {e}")


def process_evtx():
    """Processes EVTX file and ensures output is saved with a specific name."""
    if os.path.exists(evtx_file):
        evtx_command = f'"{EVTX_CMD}" -f "{evtx_file}" --csv "{output_dir}" --csvf "evtx_results.csv"'
        run_tool(evtx_command)


def process_mft():
    """Processes MFT file and ensures output is saved with a specific name."""
    if os.path.exists(mft_file):
        mft_command = f'"{MFT_CMD}" -f "{mft_file}" --csv "{output_dir}" --csvf "mft_results.csv"'
        run_tool(mft_command)


def process_registry():
    """Processes Registry files using RECmd."""
    if os.path.exists(registry_file) and os.path.exists(batch_file):
        reg_command = f'"{RE_CMD}" -f "{registry_file}" --bn "{batch_file}" --csv "{output_dir}" --csvf "reg_results.csv"'
        run_tool(reg_command)


def merge_csv(files, labels, artifact_types, output_file):
    """Merges multiple CSV files into one with source and artifact type labels."""
    dataframes = []
    timestamp_keywords = ["time", "date", "created", "modified", "last", "written", "timestamp"]

    for file, label, artifact in zip(files, labels, artifact_types):
        if os.path.exists(file) and os.path.getsize(file) > 0:
            try:
                print(f"Reading file: {file}")
                df = pd.read_csv(file, low_memory=False)

                if df.empty:
                    print(f"Warning: {file} is empty!")
                    continue  # Skip empty files

                df.insert(0, "Source", label)
                df.insert(1, "Artifact_Type", artifact)

                # Convert timestamp-related columns to UTC+0
                for col in df.columns:
                    col_lower = col.lower()
                    if any(keyword in col_lower for keyword in timestamp_keywords):
                        try:
                            df[col] = pd.to_datetime(df[col], errors='coerce', utc=True)
                            print(f"Converted column '{col}' to UTC+0 in {file}")
                        except Exception as e:
                            print(f"Error converting {col} in {file}: {e}")

                print(f"File {file} has {df.shape[0]} rows and {df.shape[1]} columns")
                dataframes.append(df)

            except Exception as e:
                print(f"Error reading {file}: {e}")

    if dataframes:
        merged_df = pd.concat(dataframes, ignore_index=True, sort=False)
        merged_df.to_csv(output_file, index=False)
        print(f"Combined CSV saved: {output_file}")
    else:
        print("No valid data to merge.")


def check_csv_files(files):
    """Check if CSV files exist and have content."""
    for file in files:
        if os.path.exists(file):
            size = os.path.getsize(file)
            print(f"File '{file}' exists, size: {size} bytes")
        else:
            print(f"File '{file}' does NOT exist!")


def main():
    """Main function to run all processing steps."""
    print("Starting artifact processing...")

    process_evtx()
    process_mft()
    process_registry()

    check_csv_files([evtx_output, mft_output, reg_output])

    print("Merging results...")
    merge_csv(
        [evtx_output, mft_output, reg_output],
        ["EVTX", "MFT", "Registry"],
        ["Event Log", "Master File Table", "Registry Hive"],
        final_output
    )

    print("Processing completed successfully!")


if __name__ == "__main__":
    main()
