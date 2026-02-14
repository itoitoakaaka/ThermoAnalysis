import pandas as pd
import glob
import os
import re
from datetime import datetime, timedelta
import numpy as np
from pathlib import Path

# --- Configuration ---
# 実行ディレクトリからの相対パス
DOWNLOADS_DIR = Path("Downloads")
# ... (events, name_map etc follow) ...

# ... (EVENTS_EXP1, EVENTS_EXP2, NAME_MAP_KANJI_TO_HR definitions) ...

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
                
                name = name_mapping.get(file_no, {}).get(cap_id, None)
                if name:
                    all_data.append((name, data_block))
        except Exception:
            pass
    return all_data

def process_experiment_data(events, combined_hr_df, combined_temp_df, prefix=""):
    target_seconds = np.arange(-300, 421, 1)
    
    for start_time_str, names, suffix in events:
        start_dt = parse_time_to_dummy_datetime(start_time_str)
        
        for name in names:
            col_label = f"{prefix}{name}_{suffix}" if suffix else f"{prefix}{name}"
            
            hr_series = pd.Series(index=target_seconds, dtype=float)
            hr_source_df = load_hr_data_for_subject(name)
            
            if not hr_source_df.empty and 'HR (bpm)' in hr_source_df.columns:
                hr_source_df['RelSeconds'] = (hr_source_df['Datetime'] - start_dt).dt.total_seconds()
                mask = (hr_source_df['RelSeconds'] >= -310) & (hr_source_df['RelSeconds'] <= 430)
                subset = hr_source_df.loc[mask].copy()
                
                if not subset.empty:
                    subset = subset.set_index('RelSeconds')
                    subset = subset[~subset.index.duplicated(keep='first')]
                    reindexed = subset['HR (bpm)'].reindex(target_seconds, method='nearest', tolerance=2.0)
                    hr_series = reindexed
            
            combined_hr_df[col_label] = hr_series
            pass 

    return

def main():
    target_index = np.arange(-300, 421, 1)
    
    df_export_hr = pd.DataFrame(index=target_index)
    df_export_hr.index.name = 'Seconds_from_Start'
    
    df_export_temp = pd.DataFrame(index=target_index)
    df_export_temp.index.name = 'Seconds_from_Start'
    
    all_temp_data = load_temp_data()
    temp_dict = {}
    for name, df in all_temp_data:
        temp_dict[name] = df

    for start_time_str, names, suffix in EVENTS_EXP1:
        start_dt = parse_time_to_dummy_datetime(start_time_str)
        for name in names:
            col_label = f"Exp1_{name}_{suffix}"
            
            hr_source = load_hr_data_for_subject(name)
            if not hr_source.empty:
                hr_source['RelSeconds'] = (hr_source['Datetime'] - start_dt).dt.total_seconds()
                subset = hr_source[(hr_source['RelSeconds'] >= -300) & (hr_source['RelSeconds'] <= 420)]
                if not subset.empty:
                    subset = subset.set_index('RelSeconds')
                    subset = subset[~subset.index.duplicated()]
                    series = subset['HR (bpm)'].reindex(df_export_hr.index, method='nearest', tolerance=1.5)
                    df_export_hr[col_label] = series
            
            if name in temp_dict:
                temp_source = temp_dict[name].copy()
                temp_source['RelSeconds'] = (temp_source['Datetime'] - start_dt).dt.total_seconds()
                subset = temp_source[(temp_source['RelSeconds'] >= -300) & (temp_source['RelSeconds'] <= 420)]
                if not subset.empty:
                    subset = subset.set_index('RelSeconds')
                    subset = subset[~subset.index.duplicated()]
                    series = np.interp(df_export_temp.index, subset.index, subset['Temp'], left=np.nan, right=np.nan)
                    df_export_temp[col_label] = series

    for start_time_str, names, suffix in EVENTS_EXP2:
        start_dt = parse_time_to_dummy_datetime(start_time_str)
        for name in names:
            col_label = f"Exp2_{name}"
            
            hr_source = load_hr_data_for_subject(name)
            if not hr_source.empty:
                hr_source['RelSeconds'] = (hr_source['Datetime'] - start_dt).dt.total_seconds()
                subset = hr_source[(hr_source['RelSeconds'] >= -300) & (hr_source['RelSeconds'] <= 420)]
                if not subset.empty:
                    subset = subset.set_index('RelSeconds')
                    subset = subset[~subset.index.duplicated()]
                    series = subset['HR (bpm)'].reindex(df_export_hr.index, method='nearest', tolerance=1.5)
                    df_export_hr[col_label] = series

            if name in temp_dict:
                temp_source = temp_dict[name].copy()
                temp_source['RelSeconds'] = (temp_source['Datetime'] - start_dt).dt.total_seconds()
                subset = temp_source[(temp_source['RelSeconds'] >= -300) & (temp_source['RelSeconds'] <= 420)]
                if not subset.empty:
                    subset = subset.set_index('RelSeconds')
                    subset = subset[~subset.index.duplicated()]
                    series = np.interp(df_export_temp.index, subset.index, subset['Temp'], left=np.nan, right=np.nan)
                    df_export_temp[col_label] = series

    def format_seconds(x):
        sign = "-" if x < 0 else ""
        abs_x = int(abs(x))
        m, s = divmod(abs_x, 60)
        return f"{sign}{m}:{s:02d}"

    df_export_hr.index = df_export_hr.index.map(format_seconds)
    df_export_temp.index = df_export_temp.index.map(format_seconds)
    
    df_export_hr.index.name = 'Time'
    df_export_temp.index.name = 'Time'

    out_path = DOWNLOADS_DIR / "Experiment_Data_Aligned.xlsx"
    DOWNLOADS_DIR.mkdir(parents=True, exist_ok=True)
    with pd.ExcelWriter(out_path) as writer:
        df_export_temp.to_excel(writer, sheet_name='Core Temp')
        df_export_hr.to_excel(writer, sheet_name='Heart Rate')
    
    print(f"Saved {out_path}")

if __name__ == "__main__":
    main()
