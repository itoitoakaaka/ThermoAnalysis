import pandas as pd
import matplotlib.pyplot as plt
import matplotlib.font_manager as fm
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

# ... (EXP1_SUBJECTS, EXP1_MAP, EVENTS_EXP2 definitions) ...

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

def load_hr_data_for_subject(kanji_name):
    filename = f"心拍数_{kanji_name}.CSV"
    path = DOWNLOADS_DIR / filename
    if not path.exists():
        return pd.DataFrame()
    
    try:
        with open(path, 'r', encoding='utf-8', errors='replace') as f:
            lines = f.readlines()
            if len(lines) < 4: return pd.DataFrame()
            
            meta_header = lines[0].strip().split(',')
            meta_values = lines[1].strip().split(',')
            
            try:
                date_idx = meta_header.index('Date')
                start_time_idx = meta_header.index('Start time')
                date_str = meta_values[date_idx]
                start_time_str = meta_values[start_time_idx]
                start_dt = datetime.strptime(f"{date_str} {start_time_str}", "%d-%m-%Y %H:%M:%S")
                base_dt = datetime(1900, 1, 1, start_dt.hour, start_dt.minute, start_dt.second)
            except ValueError:
                return pd.DataFrame()

        df = pd.read_csv(path, header=2)
        df['DurationDelta'] = pd.to_timedelta(df['Time'])
        df['Datetime'] = base_dt + df['DurationDelta']
        return df

    except Exception as e:
        print(f"Error loading {filename}: {e}")
        return pd.DataFrame()

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

# ... (calculate_stats definition) ...

def plot_exp1_grid(dummy_hr, temp_data_list):
    setup_japanese_font()
    DOWNLOADS_DIR.mkdir(parents=True, exist_ok=True)
    
    fig, axes = plt.subplots(8, 2, figsize=(15, 30))
    # ... rest of plotting logic ...
    for row_idx, subject in enumerate(EXP1_SUBJECTS):
        trials = ["1", "2"]
        for col_idx, trial in enumerate(trials):
            ax1 = axes[row_idx, col_idx]
            start_time_str = EXP1_MAP.get(subject, {}).get(trial)
            if not start_time_str:
                ax1.text(0.5, 0.5, "No Data", ha='center', va='center')
                continue
            start_dt = parse_time_to_dummy_datetime(start_time_str)
            start_window = start_dt - timedelta(minutes=5)
            end_window = start_dt + timedelta(minutes=7)
            color = COLOR_MAP.get(subject, 'black')
            
            hr_stats_text = ""
            hr_df = load_hr_data_for_subject(subject)
            col_name_hr = "HR (bpm)"
            if not hr_df.empty and col_name_hr in hr_df.columns:
                mask = (hr_df['Datetime'] >= start_window) & (hr_df['Datetime'] <= end_window)
                segment_hr = hr_df.loc[mask].copy()
                if not segment_hr.empty:
                    segment_hr['RelTime'] = (segment_hr['Datetime'] - start_dt).dt.total_seconds() / 60.0
                    ax1.plot(segment_hr['RelTime'], segment_hr[col_name_hr], color=color, linestyle='-', label='HR', linewidth=2, alpha=0.8)
                    pre, during, post = calculate_stats(segment_hr, col_name_hr, start_dt)
                    hr_stats_text = f"HR: {pre:.1f}/{during:.1f}/{post:.1f}"

            ax1.set_ylabel('HR (bpm)', color=color)
            ax1.tick_params(axis='y', labelcolor=color)
            ax1.axvline(0, color='gray', linestyle='--', alpha=0.5)
            ax1.axvline(2, color='gray', linestyle='--', alpha=0.5)

            ax2 = ax1.twinx()
            temp_stats_text = ""
            for d_name, d_df in temp_data_list:
                if d_name == subject:
                    mask = (d_df['Datetime'] >= start_window) & (d_df['Datetime'] <= end_window)
                    segment_temp = d_df.loc[mask].copy()
                    if not segment_temp.empty:
                        segment_temp['RelTime'] = (segment_temp['Datetime'] - start_dt).dt.total_seconds() / 60.0
                        ax2.plot(segment_temp['RelTime'], segment_temp['Temp'], color=color, linestyle=':', label='Temp', linewidth=2, alpha=0.8)
                        pre, during, post = calculate_stats(segment_temp, 'Temp', start_dt)
                        temp_stats_text = f"Temp Avg: {pre:.2f} / {during:.2f} / {post:.2f}"
            
            ax2.set_ylabel('Temp (°C)', color=color)
            ax2.tick_params(axis='y', labelcolor=color)
            stats_title = f"{hr_stats_text}\n{temp_stats_text}"
            ax1.set_title(f"{subject} - Trial {trial} ({start_time_str})\n{stats_title}", fontsize=10)
            ax1.set_xlabel('Time (min)')

    fig.suptitle("Experiment 1: Individual Trials (Pre / During / Post Averages)", fontsize=16)
    fig.tight_layout(rect=[0, 0.03, 1, 0.98])
    out_file = DOWNLOADS_DIR / "Experiment1_Grid_Refined.png"
    plt.savefig(out_file)
    print(f"Saved {out_file}")
    plt.close()

def plot_exp2_grid(dummy_hr, temp_data_list):
    setup_japanese_font()
    DOWNLOADS_DIR.mkdir(parents=True, exist_ok=True)
    n_rows = len(EVENTS_EXP2)
    n_cols = 2
    fig, axes = plt.subplots(n_rows, n_cols, figsize=(15, 4 * n_rows))
    for i, (start_time_str, names, suffix) in enumerate(EVENTS_EXP2):
        start_dt = parse_time_to_dummy_datetime(start_time_str)
        start_window = start_dt - timedelta(minutes=5)
        end_window = start_dt + timedelta(minutes=7)
        for j, subject in enumerate(names):
            if j >= n_cols: break
            ax1 = axes[i, j]
            color = COLOR_MAP.get(subject, 'black')
            hr_stats_text = ""
            hr_df = load_hr_data_for_subject(subject)
            col_name_hr = "HR (bpm)"
            if not hr_df.empty and col_name_hr in hr_df.columns:
                mask = (hr_df['Datetime'] >= start_window) & (hr_df['Datetime'] <= end_window)
                segment_hr = hr_df.loc[mask].copy()
                if not segment_hr.empty:
                    segment_hr['RelTime'] = (segment_hr['Datetime'] - start_dt).dt.total_seconds() / 60.0
                    ax1.plot(segment_hr['RelTime'], segment_hr[col_name_hr], color=color, linestyle='-', linewidth=2, alpha=0.8)
                    pre, during, post = calculate_stats(segment_hr, col_name_hr, start_dt)
                    hr_stats_text = f"HR: {pre:.1f}/{during:.1f}/{post:.1f}"

            ax1.set_ylabel('HR', color=color)
            ax1.tick_params(axis='y', labelcolor=color)
            ax1.axvline(0, color='gray', linestyle='--', alpha=0.5)
            ax1.axvline(2, color='gray', linestyle='--', alpha=0.5)

            ax2 = ax1.twinx()
            temp_stats_text = ""
            for d_name, d_df in temp_data_list:
                if d_name == subject:
                    mask = (d_df['Datetime'] >= start_window) & (d_df['Datetime'] <= end_window)
                    segment_temp = d_df.loc[mask].copy()
                    if not segment_temp.empty:
                        segment_temp['RelTime'] = (segment_temp['Datetime'] - start_dt).dt.total_seconds() / 60.0
                        ax2.plot(segment_temp['RelTime'], segment_temp['Temp'], color=color, linestyle=':', linewidth=2, alpha=0.8)
                        pre, during, post = calculate_stats(segment_temp, 'Temp', start_dt)
                        temp_stats_text = f"Temp: {pre:.2f}/{during:.2f}/{post:.2f}"
            
            ax2.set_ylabel('Temp', color=color)
            ax2.tick_params(axis='y', labelcolor=color)
            stats_str = f"{hr_stats_text}\n{temp_stats_text}"
            ax1.set_title(f"{subject} ({start_time_str})\n{stats_str}", fontsize=10)
            ax1.set_xlabel('Time (min)')

    fig.suptitle("Experiment 2: Overview (Pre / During / Post Avg)", fontsize=16)
    fig.tight_layout(rect=[0, 0.03, 1, 0.98])
    out_file = DOWNLOADS_DIR / "Experiment2_Grid_Refined.png"
    plt.savefig(out_file)
    print(f"Saved {out_file}")
    plt.close()

def main():
    temp_data = load_temp_data()
    plot_exp1_grid(pd.DataFrame(), temp_data)
    plot_exp2_grid(pd.DataFrame(), temp_data)

if __name__ == "__main__":
    main()
