import pandas as pd
import matplotlib.pyplot as plt
import matplotlib.font_manager as fm
import matplotlib.dates as mdates
import glob
import os
import re
from datetime import datetime, timedelta
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

# ... (combine_datetime definition) ...

def plot_thermo_unified():
    output_path = DOWNLOADS_DIR / "260117_temperature_unified.png"
    DOWNLOADS_DIR.mkdir(parents=True, exist_ok=True)
    
    files = list(DOWNLOADS_DIR.glob("260117_no*.xlsx")) + list(DOWNLOADS_DIR.glob("260117_No*.xlsx"))
    files = list(set(files))

    if not files:
        print("No 260117_no*.xlsx files found.")
        return

    setup_japanese_font()
    
    # ... (color_map, name_mapping, all_series loop) ...
    # (Inside the loop, ensure file_path handles Path object)
    for file_path in files:
        filename = file_path.name
        print(f"Processing {filename}...")
        
        try:
            df = pd.read_excel(file_path, header=None)
            # ... rest of processing ...
            # ... (all_series plot logic) ...
            pass
        except Exception as e:
            print(f"Error processing {filename}: {e}")

    # Plot Unified
    if all_series:
        plt.figure(figsize=(20, 10))
        for name, data in all_series:
            color = color_map.get(name, 'black')
            plt.plot(data['Datetime'], data['Temp'], label=name, color=color, linewidth=2)
            
        plt.title("Core Temperature (Filtered >= 36.0°C)")
        plt.xlabel("Time")
        plt.ylabel("Temperature (°C)")
        plt.legend(loc='upper right', bbox_to_anchor=(1.1, 1))
        plt.gca().xaxis.set_major_formatter(mdates.DateFormatter('%H:%M'))
        plt.grid(True)
        plt.tight_layout()
        
        plt.savefig(output_path)
        print(f"Saved unified plot to {output_path}")
    else:
        print("No valid data found to plot.")

if __name__ == "__main__":
    plot_thermo_unified()
