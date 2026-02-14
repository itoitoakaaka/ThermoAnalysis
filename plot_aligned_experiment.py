import pandas as pd
import matplotlib.pyplot as plt
import matplotlib.font_manager as fm
import matplotlib.dates as mdates
import glob
import os
import re
from datetime import datetime, timedelta
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

# --- Configuration ---
# Name Mapping: HR_Col_Name -> Kanji Name
NAME_MAP_HR_TO_KANJI = {
    "Fujii": "藤井",
    "Itai": "板井",
    "Ito": "伊藤",
    "Kan": "姜",
    "Kitada": "北田",
    "Takamizawa": "高見澤",
    "Yamaguchi": "山口",
    "Yamamoto": "山本"
}

# Reverse map for convenience
NAME_MAP_KANJI_TO_HR = {v: k for k, v in NAME_MAP_HR_TO_KANJI.items()}

# Color Map (Kanji -> Color Code)
COLOR_MAP = {
    "藤井": "C0",
    "板井": "C1",
    "伊藤": "C2",
    "姜":   "C3",
    "北田": "C4",
    "高見澤": "C5",
    "山口": "C6",
    "山本": "C7"
}

# --- Event Definitions ---
# Format: (TimeStr, [PersonKanji1, PersonKanji2], SuffixLabel)
# Date for these times is assumed to be the date in the files (2025-01-17 based on filename 260117, usually ddmmyy or yymmdd. User file 260117 -> 2026? Or 2017?
# Wait, file `260117_no1.xlsx`.
# HR parsed date might need checking. 
# Looking at Jisedai2026_HR, it has Time only. I assigned a dummy date.
# I need to ensure the DATES match between Event Time and Data Time.
# The previous Jisedai plot used `pd.to_datetime(df['Time'], format='%H:%M:%S')` which attaches 1900-01-01.
# I should use 1900-01-01 for alignment to be safe within the script, or just match the date.
# I'll use 1900-01-01 for the event times to match the HR parsing.
# For Temp data, `260117` suggests 2017-01-26 or 2026-01-17? 
# In `plot_thermo_unified.py`, I did `combine_datetime`. If date is explicit in Excel, it uses it.
# If I use relative time (minutes from start), exact absolute date doesn't matter AS LONG AS it matches the data.
# HR data has NO date in CSV, only Time. So it defaults to 1900-01-01.
# Temp data HAS date in Excel. E.g. `2025-01-17` (maybe?).
# I need to normalize Temp Date to 1900-01-01 or normalize Event/HR to Temp Date.
# Let's normalize everything to a dummy date (e.g. 2000-01-01) just for time comparison, 
# IGNORING the actual date part, assuming all data is within the same day.

EVENTS_EXP1 = [
    ("14:08:12", ["山口", "姜"], "1回目"), # High Effort?
    ("14:13:15", ["北田", "伊藤"], "1回目"),
    ("14:17:35", ["藤井", "山本"], "1回目"), # Low Effort?
    ("14:21:33", ["板井", "高見澤"], "1回目"),
    ("14:28:04", ["山口", "姜"], "2回目"), # Low
    ("14:32:23", ["北田", "伊藤"], "2回目"),
    ("14:36:50", ["藤井", "山本"], "2回目"), # High
    ("14:40:29", ["板井", "高見澤"], "2回目")
]

EVENTS_EXP2 = [
    ("14:52:43", ["山口", "姜"], ""),
    ("14:56:47", ["北田", "伊藤"], ""),
    ("14:59:55", ["藤井", "山本"], ""),
    ("15:05:16", ["板井", "高見澤"], "")
]

def parse_time_to_dummy_datetime(time_str):
    # Returns datetime on 1900-01-01
    return datetime.strptime(time_str, "%H:%M:%S")

def normalize_to_time_only(dt_series):
    # Converts a datetime series to 1900-01-01 + time
    return dt_series.apply(lambda x: x.replace(year=1900, month=1, day=1) if pd.notna(x) else pd.NaT)

# --- Data Loading ---
def load_hr_data():
    path = DOWNLOADS_DIR / "Jisedai2026_HR.csv"
    if not path.exists():
        return pd.DataFrame()
    
    df = pd.read_csv(path)
    # Parse Time. Defaults to 1900-01-01
    df['Datetime'] = pd.to_datetime(df['Time'], format='%H:%M:%S')
    return df

def combine_datetime_excel(row):
    try:
        d = row['Date']
        t = row['Time']
        if pd.isna(d) or pd.isna(t):
            return pd.NaT
        
        # We just want TIME part combined with Dummy Date 1900-01-01
        if isinstance(t, str):
            t_obj = datetime.strptime(t, '%H:%M:%S').time()
        elif isinstance(t, datetime):
            t_obj = t.time()
        elif hasattr(t, 'hour'):
             t_obj = t
        else:
             return pd.NaT
             
        return datetime(1900, 1, 1, t_obj.hour, t_obj.minute, t_obj.second)
    except:
        return pd.NaT

def load_temp_data():
    files = list(DOWNLOADS_DIR.glob("260117_no*.xlsx")) + list(DOWNLOADS_DIR.glob("260117_No*.xlsx"))
    files = list(set(files))
    
    name_mapping = {
        1: {2: "板井", 3: "姜"},
        2: {2: "北田", 3: "伊藤"},
        3: {1: "山本", 3: "高見澤"},
        5: {1: "山口", 2: "藤井"}
    }
    
    all_data = [] # List of (Name, DataFrame)

    for file_path in files:
        filename = os.path.basename(file_path)
        try:
            df = pd.read_excel(file_path, header=None)
            
            # Extract File No
            file_no_match = re.search(r'no(\d+)', filename.lower())
            if not file_no_match: continue
            file_no = int(file_no_match.group(1))

            # Find Capsule
            capsule_row = df.iloc[6]
            capsule_indices = []
            for col_idx, val in capsule_row.items():
                if isinstance(val, str) and "Capsule" in val:
                    match = re.search(r'n[^\d]*(\d+)', val) 
                    if match:
                        cap_id = int(match.group(1))
                        capsule_indices.append((col_idx, cap_id))
            
            data_start_idx = 8
            for col_idx, cap_id in capsule_indices:
                temp_idx = col_idx + 3
                date_idx = col_idx + 1
                time_idx = col_idx + 2
                
                if temp_idx >= len(df.columns): continue

                data_block = df.iloc[data_start_idx:, [date_idx, time_idx, temp_idx]].copy()
                data_block.columns = ['Date', 'Time', 'Temp']
                data_block = data_block.dropna(subset=['Time', 'Temp'])
                
                # Normalize to dummy datetime
                data_block['Datetime'] = data_block.apply(combine_datetime_excel, axis=1)
                data_block = data_block.dropna(subset=['Datetime'])
                data_block['Temp'] = pd.to_numeric(data_block['Temp'], errors='coerce')
                
                # We need ALL data, but user previously filtered >= 36. 
                # For alignment, we probably want the full trace to see trends?
                # User request didn't specify filter this time, but "体温と心拍の波形を出して".
                # Previous task filtered >= 36. Let's keep that to avoid noise/zeros?
                # Or maybe loosen it? If sensor falls off it drops layout.
                # Let's keep >= 35.0 to be safe but exclude complete noise.
                data_block = data_block[data_block['Temp'] >= 30.0] 
                
                name = name_mapping.get(file_no, {}).get(cap_id, None)
                if name:
                    all_data.append((name, data_block))
                    
        except Exception as e:
            print(f"Error loading {filename}: {e}")
            
    return all_data

# --- Plotting ---
def plot_experiment(events, exp_name, hr_df, temp_data_list):
    setup_japanese_font()
    
    fig, axes = plt.subplots(2, 1, figsize=(20, 12), sharex=True)
    ax_hr, ax_temp = axes
    
    # Range: -5 min to +7 min
    # Convert separate lines.
    
    # Track handles for legend to avoid duplicates
    hr_handles, hr_labels = {}, {}
    temp_handles, temp_labels = {}, {}

    for start_time_str, names, suffix in events:
        start_dt = parse_time_to_dummy_datetime(start_time_str)
        
        # Define Window
        start_window = start_dt - timedelta(minutes=5)
        end_window = start_dt + timedelta(minutes=7)
        
        # 1. Plot HR
        for kanji_name in names:
            col_name = NAME_MAP_KANJI_TO_HR.get(kanji_name)
            if col_name and col_name in hr_df.columns:
                # Slice
                mask = (hr_df['Datetime'] >= start_window) & (hr_df['Datetime'] <= end_window)
                segment = hr_df.loc[mask].copy()
                
                if not segment.empty:
                    # Calc relative time in minutes
                    segment['RelTime'] = (segment['Datetime'] - start_dt).dt.total_seconds() / 60.0
                    
                    label = f"{kanji_name}" # Suffix might clutter legend if repetitive. 
                    # If same person appears multiple times, we might want "Yamaguchi (1)" etc?
                    # But keeping consistent color.
                    # Let's plot line.
                    color = COLOR_MAP.get(kanji_name, 'black')
                    
                    # We use solid line for run 1, maybe dashed for run 2? 
                    # Or just overplot? "全員分揃えて" -> Superimposed.
                    # With multiple runs for same person in Exp1, overplotting same color is fine.
                    
                    line, = ax_hr.plot(segment['RelTime'], segment[col_name], color=color, alpha=0.8)
                    
                    if kanji_name not in hr_handles:
                        hr_handles[kanji_name] = line
                        hr_labels[kanji_name] = kanji_name

        # 2. Plot Temp
        for kanji_name in names:
            # Find data for this person
            # temp_data_list is list of (name, df)
            for d_name, d_df in temp_data_list:
                if d_name == kanji_name:
                    # Slice
                    mask = (d_df['Datetime'] >= start_window) & (d_df['Datetime'] <= end_window)
                    segment = d_df.loc[mask].copy()
                    
                    if not segment.empty:
                        segment['RelTime'] = (segment['Datetime'] - start_dt).dt.total_seconds() / 60.0
                        color = COLOR_MAP.get(kanji_name, 'black')
                        
                        line, = ax_temp.plot(segment['RelTime'], segment['Temp'], color=color, alpha=0.8)
                        
                        if kanji_name not in temp_handles:
                            temp_handles[kanji_name] = line
                            temp_labels[kanji_name] = kanji_name

    # Styling
    # HR
    ax_hr.set_title(f"{exp_name} - Heart Rate")
    ax_hr.set_ylabel("HR (bpm)")
    ax_hr.grid(True)
    ax_hr.axvline(0, color='red', linestyle='--', label='Start')
    ax_hr.axvline(2, color='gray', linestyle=':', label='2 min')
    
    # Legend
    # Merge handles
    h_list = list(hr_handles.values())
    l_list = list(hr_labels.values())
    ax_hr.legend(h_list, l_list, loc='upper right')

    # Temp
    ax_temp.set_title(f"{exp_name} - Core Temperature")
    ax_temp.set_ylabel("Temp (°C)")
    ax_temp.set_xlabel("Time from Start (min)")
    ax_temp.grid(True)
    ax_temp.axvline(0, color='red', linestyle='--')
    ax_temp.axvline(2, color='gray', linestyle=':')
    
    # Y-limit for temp (zoom in to valid range)
    ax_temp.set_ylim(36.0, 40.0) # Approx range

    # Save
    out_file = DOWNLOADS_DIR / f"{exp_name}_Aligned.png"
    plt.tight_layout()
    plt.savefig(out_file)
    print(f"Saved {out_file}")

def main():
    hr_df = load_hr_data()
    temp_data = load_temp_data()
    
    if hr_df.empty: 
        print("No HR data loaded.")
        return

    plot_experiment(EVENTS_EXP1, "Experiment1", hr_df, temp_data)
    plot_experiment(EVENTS_EXP2, "Experiment2", hr_df, temp_data)

if __name__ == "__main__":
    main()
