import glob
import os
import pandas as pd
import matplotlib.pyplot as plt
import matplotlib.font_manager as fm
import datetime
from pathlib import Path

# --- Configuration ---
DOWNLOADS_DIR = Path("Downloads")

def plot_temperature():
    DOWNLOADS_DIR.mkdir(parents=True, exist_ok=True)
    file_pattern = DOWNLOADS_DIR / "260117_no*.xlsx"
    files = glob.glob(str(file_pattern))

    if not files:
        print("No files found matching the pattern.")
        return

    plt.figure(figsize=(20, 6))

    # Try to support Japanese characters in plot
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

    for file_path in files:
        filename = os.path.basename(file_path)
        # Extract name (e.g., "no1" from "260117_no1.xlsx")
        name = os.path.splitext(filename)[0].split('_')[1]

        try:
            # Read Excel, header=None to handle custom layout
            df = pd.read_excel(file_path, header=None)
            
            # The files seem to have data in two blocks.
            # Block 1: Columns 5, 6, 7 (indices 5, 6, 7) -> Date, Time, Temp
            # Block 2: Columns 12, 13, 14 (indices 12, 13, 14) -> Date, Time, Temp
            
            # Extract Block 1
            # Assuming data starts from row 10 based on inspection (indices 9/10), 
            # but let's just drop rows where Date/Time are NaN later to be safe.
            block1 = df.iloc[:, [5, 6, 7]].copy()
            block1.columns = ['Date', 'Time', 'Temp']
            
            # Extract Block 2
            block2 = df.iloc[:, [12, 13, 14]].copy()
            block2.columns = ['Date', 'Time', 'Temp']
            
            # Concatenate blocks
            combined_df = pd.concat([block1, block2], ignore_index=True)
            
            # Clean data
            combined_df = combined_df.dropna(subset=['Date', 'Time', 'Temp'])
            
            # Create datetime column
            # Ensure Date is datetime and Time is time object/string, then combine.
            # Sometimes 'Date' might be an object or timestamp.
            
            def combine_datetime(row):
                d = row['Date']
                t = row['Time']
                
                if pd.isnull(d) or pd.isnull(t):
                    return pd.NaT
                
                # If d is already a timestamp, just combine with t
                if isinstance(d, datetime.datetime) or isinstance(d, pd.Timestamp):
                    d_date = d.date()
                else:
                    # Parse string if needed (unlikely based on read_excel default behavior for dates)
                    return pd.NaT # Or try parsing
                
                if isinstance(t, datetime.time):
                    return datetime.datetime.combine(d_date, t)
                elif isinstance(t, str):
                    try:
                        t_time = datetime.datetime.strptime(t, "%H:%M:%S").time()
                        return datetime.datetime.combine(d_date, t_time)
                    except:
                        return pd.NaT

                return pd.NaT

            combined_df['Datetime'] = combined_df.apply(combine_datetime, axis=1)
            combined_df = combined_df.dropna(subset=['Datetime'])
            combined_df = combined_df.sort_values('Datetime')
            
            # Filter non-numeric Temps (e.g. headers if any slipped through)
            combined_df['Temp'] = pd.to_numeric(combined_df['Temp'], errors='coerce')
            combined_df = combined_df.dropna(subset=['Temp'])

            plt.plot(combined_df['Datetime'], combined_df['Temp'], label=name)
            
        except Exception as e:
            print(f"Error processing {filename}: {e}")
            import traceback
            traceback.print_exc()

    plt.title("Temperature 260117")
    plt.xlabel("Time")
    plt.ylabel("Temperature")
    plt.legend()
    plt.grid(True)
    
    # Format x-axis dates nicely
    import matplotlib.dates as mdates
    plt.gca().xaxis.set_major_formatter(mdates.DateFormatter('%H:%M:%S'))
    
    output_path = os.path.join(downloads_path, "260117_temperature.png")
    plt.savefig(output_path)
    print(f"Plot saved to {output_path}")

if __name__ == "__main__":
    plot_temperature()
