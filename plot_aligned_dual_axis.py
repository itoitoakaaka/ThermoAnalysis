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

# ... (NAME_MAP_HR_TO_KANJI, COLOR_MAP, EVENTS_EXP1, EVENTS_EXP2 definitions) ...

def parse_time_to_dummy_datetime(time_str):
    return datetime.strptime(time_str, "%H:%M:%S")

def combine_datetime_excel(row):
    try:
        t = row['Time']
        if pd.isna(t): return pd.NaT
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

def load_hr_data():
    path = DOWNLOADS_DIR / "Jisedai2026_HR.csv"
    if not path.exists(): return pd.DataFrame()
    df = pd.read_csv(path)
    df['Datetime'] = pd.to_datetime(df['Time'], format='%H:%M:%S')
    return df

def load_temp_data():
    files = list(DOWNLOADS_DIR.glob("260117_no*.xlsx")) + list(DOWNLOADS_DIR.glob("260117_No*.xlsx"))
    files = list(set(files))
    
    name_mapping = {
        1: {2: "板井", 3: "姜"},
        2: {2: "北田", 3: "伊藤"},
        3: {1: "山本", 3: "高見澤"},
        5: {1: "山口", 2: "藤井"}
    }
    
    all_data = [] 

    for file_path in files:
        filename = file_path.name
        try:
            df = pd.read_excel(file_path, header=None)
            file_no_match = re.search(r'no(\d+)', filename.lower())
            if not file_no_match: continue
            file_no = int(file_no_match.group(1))

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
                data_block['Datetime'] = data_block.apply(combine_datetime_excel, axis=1)
                data_block = data_block.dropna(subset=['Datetime'])
                data_block['Temp'] = pd.to_numeric(data_block['Temp'], errors='coerce')
                data_block = data_block[data_block['Temp'] >= 30.0] 
                
                name = name_mapping.get(file_no, {}).get(cap_id, None)
                if name:
                    all_data.append((name, data_block))
        except Exception as e:
            print(f"Error loading {filename}: {e}")
    return all_data

def plot_individual_dual_axis(events, exp_name, hr_df, temp_data_list):
    setup_japanese_font()
    DOWNLOADS_DIR.mkdir(parents=True, exist_ok=True)

    for start_time_str, names, suffix in events:
        start_dt = parse_time_to_dummy_datetime(start_time_str)
        start_window = start_dt - timedelta(minutes=5)
        end_window = start_dt + timedelta(minutes=7)
        
        suffix_clean = suffix.replace(" ", "") if suffix else ""
        
        for kanji_name in names:
            fig, ax1 = plt.subplots(figsize=(10, 6))
            color = COLOR_MAP.get(kanji_name, 'black')
            
            col_name_hr = NAME_MAP_KANJI_TO_HR.get(kanji_name)
            has_hr = False
            
            if col_name_hr and col_name_hr in hr_df.columns:
                mask = (hr_df['Datetime'] >= start_window) & (hr_df['Datetime'] <= end_window)
                segment_hr = hr_df.loc[mask].copy()
                if not segment_hr.empty:
                    segment_hr['RelTime'] = (segment_hr['Datetime'] - start_dt).dt.total_seconds() / 60.0
                    ax1.plot(segment_hr['RelTime'], segment_hr[col_name_hr], color=color, linestyle='-', label='Heart Rate', linewidth=2)
                    has_hr = True

            ax1.set_xlabel('Time from Start (min)')
            ax1.set_ylabel('Heart Rate (bpm)', color=color)
            ax1.tick_params(axis='y', labelcolor=color)
            ax1.axvline(0, color='gray', linestyle='--', alpha=0.5)

            ax2 = ax1.twinx()
            has_temp = False
            
            for d_name, d_df in temp_data_list:
                if d_name == kanji_name:
                    mask = (d_df['Datetime'] >= start_window) & (d_df['Datetime'] <= end_window)
                    segment_temp = d_df.loc[mask].copy()
                    if not segment_temp.empty:
                        segment_temp['RelTime'] = (segment_temp['Datetime'] - start_dt).dt.total_seconds() / 60.0
                        ax2.plot(segment_temp['RelTime'], segment_temp['Temp'], color=color, linestyle=':', label='Temperature', linewidth=2)
                        has_temp = True
            
            ax2.set_ylabel('Core Temp (°C)', color=color)
            ax2.tick_params(axis='y', labelcolor=color)
            
            time_clean = start_time_str.replace(":", "")
            title_suffix = f" ({suffix})" if suffix else ""
            plt.title(f"{kanji_name}{title_suffix} - {exp_name} ({start_time_str})")
            
            fig.tight_layout()
            
            filename = f"Aligned_{exp_name}_{kanji_name}_{time_clean}_{suffix_clean}.png"
            out_path = DOWNLOADS_DIR / filename
            
            if has_hr or has_temp:
                plt.savefig(out_path)
                print(f"Saved {filename}")
            else:
                 print(f"Skipping {filename} (No data)")
            
            plt.close()

def main():
    hr_df = load_hr_data()
    temp_data = load_temp_data()
    
    if hr_df.empty: 
        print("No HR data loaded.")
        return

    plot_individual_dual_axis(EVENTS_EXP1, "Exp1", hr_df, temp_data)
    plot_individual_dual_axis(EVENTS_EXP2, "Exp2", hr_df, temp_data)

if __name__ == "__main__":
    main()
