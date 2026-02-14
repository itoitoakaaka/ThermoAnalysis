import glob
import os
import pandas as pd
import matplotlib.pyplot as plt
import matplotlib.font_manager as fm
import datetime
import matplotlib.dates as mdates
import numpy as np
from pathlib import Path

# --- Configuration ---
DOWNLOADS_DIR = Path("Downloads")

def setup_japanese_font():
    potential_fonts = ['Hiragino Sans', 'Hiragino Kaku Gothic ProN', 'Arial Unicode MS', 'Meiryo', 'Yu Gothic', 'TakaoPGothic', 'IPAPGothic']
    font_name = None
    for f in potential_fonts:
        try:
            if f in [f.name for f in fm.fontManager.ttflist]:
                font_name = f
                break
        except:
            continue
    
    if font_name:
        plt.rcParams['font.family'] = font_name
    else:
        plt.rcParams['font.family'] = 'sans-serif'

def combine_datetime(row):
    d = row['Date']
    t = row['Time']
    
    if pd.isnull(d) or pd.isnull(t):
        return pd.NaT
    
    if isinstance(d, datetime.datetime) or isinstance(d, pd.Timestamp):
        d_date = d.date()
    else:
        # Try parsing date string if necessary
        return pd.NaT
    
    if isinstance(t, datetime.time):
        return datetime.datetime.combine(d_date, t)
    elif isinstance(t, str):
        try:
            t_time = datetime.datetime.strptime(t, "%H:%M:%S").time()
            return datetime.datetime.combine(d_date, t_time)
        except:
            return pd.NaT
    
    return pd.NaT

def plot_temperature_filtered():
    DOWNLOADS_DIR.mkdir(parents=True, exist_ok=True)
    # Simplified pattern to match 260117*.xlsx
    files = [f for f in glob.glob(str(DOWNLOADS_DIR / "260117*.xlsx")) if "no" in os.path.basename(f).lower()]

    if not files:
        print("No files found matching the pattern.")
        return

    setup_japanese_font()
    
    # Store all data for combined plot
    all_series = []

    for file_path in files:
        filename = os.path.basename(file_path)
        
        try:
            # Read Excel without header to inspect structure
            df = pd.read_excel(file_path, header=None)
            
            # Define Mapping
            # FileNo -> {CapsuleID -> Name}
            # Note: 3-1 means File 3, Capsule 1
            name_mapping = {
                1: {2: "板井", 3: "姜"},
                2: {2: "北田", 3: "伊藤"},
                3: {1: "山本", 3: "高見澤"},
                5: {1: "山口", 2: "藤井"}
            }

            # Extract File No
            import re
            file_no_match = re.search(r'no(\d+)', filename.lower())
            if not file_no_match:
                print(f"Skipping {filename}: Could not determine file number.")
                continue
            file_no = int(file_no_match.group(1))

            if file_no not in name_mapping:
                print(f"Skipping {filename}: File No {file_no} not in mapping.")
                # We can choose to process it without names, but user focused on these.
                # Let's verify if we should just log warning.
                pass

            # Find Capsule IDs in Row 6 (Index 6)
            capsule_row_idx = 6
            if len(df) <= capsule_row_idx:
                print(f"Skipping {filename}: Too few rows for Capsule info.")
                continue

            capsule_row = df.iloc[capsule_row_idx]
            
            # Find indices of "Capsule n-X"
            capsule_indices = []
            for col_idx, val in capsule_row.items():
                if isinstance(val, str) and "Capsule" in val:
                    # Extract ID from "Capsule n-X"
                    match = re.search(r'n[^\d]*(\d+)', val) # n-1, nｰ1 etc.
                    if match:
                        cap_id = int(match.group(1))
                        capsule_indices.append((col_idx, cap_id))
            
            if not capsule_indices:
                print(f"Warning: No Capsule headers found in {filename} row {capsule_row_idx}. Trying Row 7 for Headers directly?")
                # Fallback or strict? User mapping relies on IDs.
                pass

            data_start_idx = 8
            
            for col_idx, cap_id in capsule_indices:
                # Based on observation:
                # Header "Temperature" is usually at col_idx + 3
                # "Sample", "Date", "Hour", "Temperature"
                # So Temp is +3, Date is +1, Hour is +2.
                # Wait, earlier no5 was: Capsule at 0. Temp at 3. (0+3). Correct.
                # no1: Capsule at 4. Temp at 7. (4+3). Correct.
                
                temp_idx = col_idx + 3
                date_idx = col_idx + 1
                time_idx = col_idx + 2
                
                # Verify header in Row 7 just in case
                header_row_check = df.iloc[7]
                if temp_idx >= len(df.columns):
                    continue
                
                # Extract Data
                data_block = df.iloc[data_start_idx:, [date_idx, time_idx, temp_idx]].copy()
                data_block.columns = ['Date', 'Time', 'Temp']
                
                data_block = data_block.dropna(subset=['Date', 'Time', 'Temp'])
                if data_block.empty:
                    continue
                
                data_block['Datetime'] = data_block.apply(combine_datetime, axis=1)
                data_block = data_block.dropna(subset=['Datetime'])
                data_block = data_block.sort_values('Datetime')
                data_block['Temp'] = pd.to_numeric(data_block['Temp'], errors='coerce')
                data_block = data_block.dropna(subset=['Temp'])
                
                # Determine Name
                name = name_mapping.get(file_no, {}).get(cap_id, f"Unknown_{file_no}_{cap_id}")
                
                max_temp = data_block['Temp'].max()
                print(f"Debug: {filename} Capsule {cap_id} ({name}) Max Temp: {max_temp}")
                
                # Filter >= 36.0
                filtered_block = data_block[data_block['Temp'] >= 36.0]
                
                if not filtered_block.empty:
                    all_series.append((name, filtered_block))
                    plot_individual(filtered_block, name, downloads_path)
                else:
                    print(f"Skipping {name}: Max temp {max_temp} < 36.0")


        except Exception as e:
            print(f"Error processing {filename}: {e}")
            import traceback
            traceback.print_exc()



        except Exception as e:
            print(f"Error processing {filename}: {e}")

    # Combined Plot
    if all_series:
        plt.figure(figsize=(20, 6))
        for name, df in all_series:
            plt.plot(df['Datetime'], df['Temp'], label=name)
        
        plt.title("Core Temperature Comparison (>= 36.0°C)")
        plt.xlabel("Time")
        plt.ylabel("Temperature (°C)")
        plt.legend(bbox_to_anchor=(1.05, 1), loc='upper left')
        plt.grid(True)
        plt.tight_layout()
        plt.gca().xaxis.set_major_formatter(mdates.DateFormatter('%H:%M:%S'))
        
        combined_out = os.path.join(downloads_path, "260117_temperature_filtered.png")
        plt.savefig(combined_out)
        print(f"Saved combined plot: {combined_out}")
    else:
        print("No valid series found for combined plot.")

def plot_individual(df, name, output_dir):
    plt.figure(figsize=(10, 6))
    plt.plot(df['Datetime'], df['Temp'], label=name, color='orange')
    plt.title(f"Core Temperature: {name} (>= 36.0°C)")
    plt.xlabel("Time")
    plt.ylabel("Temperature (°C)")
    plt.grid(True)
    
    y_max = df['Temp'].max()
    plt.ylim(36.0, max(y_max + 0.5, 38.0))
    
    plt.gca().xaxis.set_major_formatter(mdates.DateFormatter('%H:%M:%S'))
    
    safe_name = "".join([c for c in name if c.isalnum() or c in (' ', '_', '-', '.')]).strip()
    out_name = os.path.join(output_dir, f"CoreTemp_{safe_name}.png")
    plt.savefig(out_name)
    plt.close()
    print(f"Saved individual plot: {out_name}")


if __name__ == "__main__":
    plot_temperature_filtered()
