from __future__ import annotations

import argparse
from pathlib import Path

import matplotlib.pyplot as plt
import pandas as pd


def load_signal(path, value_column):
    """Load one physiological signal table."""
    df = pd.read_csv(path)
    required = {"subject_id", "timestamp", value_column}
    missing = required - set(df.columns)
    if missing:
        raise ValueError(f"{path}: missing columns {sorted(missing)}")

    out = df[["subject_id", "timestamp", value_column]].copy()
    out["timestamp"] = pd.to_datetime(out["timestamp"], errors="raise")
    out[value_column] = pd.to_numeric(out[value_column], errors="raise")
    return out.sort_values(["subject_id", "timestamp"])


def load_events(path):
    """Load anonymous event-onset metadata."""
    df = pd.read_csv(path)
    required = {"subject_id", "event_id", "event_time"}
    missing = required - set(df.columns)
    if missing:
        raise ValueError(f"{path}: missing columns {sorted(missing)}")

    out = df[["subject_id", "event_id", "event_time"]].copy()
    out["event_time"] = pd.to_datetime(out["event_time"], errors="raise")
    return out.sort_values(["subject_id", "event_time"])


def align_signal(signal_df, events_df, value_column, before_min=5.0, after_min=7.0):
    """Align one signal to event onset for each anonymous subject/event."""
    if before_min < 0 or after_min < 0:
        raise ValueError("before_min and after_min must be non-negative")

    rows = []
    for event in events_df.itertuples(index=False):
        subject = signal_df[signal_df["subject_id"] == event.subject_id].copy()
        if subject.empty:
            continue

        subject["relative_min"] = (
            subject["timestamp"] - event.event_time
        ).dt.total_seconds() / 60.0

        subject = subject[
            (subject["relative_min"] >= -before_min)
            & (subject["relative_min"] <= after_min)
        ]

        for row in subject.itertuples(index=False):
            rows.append(
                {
                    "subject_id": event.subject_id,
                    "event_id": event.event_id,
                    "relative_min": float(row.relative_min),
                    value_column: float(getattr(row, value_column)),
                }
            )

    columns = ["subject_id", "event_id", "relative_min", value_column]
    return pd.DataFrame(rows, columns=columns)


def plot_aligned(hr_aligned, temp_aligned, output_path):
    """Save aligned heart-rate and temperature traces."""
    fig, axes = plt.subplots(2, 1, figsize=(10, 7), sharex=True)

    if not hr_aligned.empty:
        for _, group in hr_aligned.groupby(["subject_id", "event_id"], sort=True):
            axes[0].plot(group["relative_min"], group["heart_rate"], alpha=0.7)

    axes[0].axvline(0, linestyle="--", linewidth=1)
    axes[0].set_ylabel("Heart rate")
    axes[0].set_title("Aligned heart-rate traces")

    if not temp_aligned.empty:
        for _, group in temp_aligned.groupby(["subject_id", "event_id"], sort=True):
            axes[1].plot(group["relative_min"], group["temperature"], alpha=0.7)

    axes[1].axvline(0, linestyle="--", linewidth=1)
    axes[1].set_ylabel("Temperature")
    axes[1].set_xlabel("Time from event onset (min)")
    axes[1].set_title("Aligned temperature traces")

    fig.tight_layout()
    fig.savefig(output_path, dpi=200)
    plt.close(fig)


def main():
    parser = argparse.ArgumentParser(
        description="Align heart-rate and temperature signals to experimental events."
    )
    parser.add_argument("--hr", required=True, help="Heart-rate CSV path")
    parser.add_argument("--temp", required=True, help="Temperature CSV path")
    parser.add_argument("--events", required=True, help="Event CSV path")
    parser.add_argument("--output", default="output", help="Output directory")
    parser.add_argument("--before-min", type=float, default=5.0)
    parser.add_argument("--after-min", type=float, default=7.0)
    args = parser.parse_args()

    output_dir = Path(args.output)
    output_dir.mkdir(parents=True, exist_ok=True)

    hr = load_signal(args.hr, "heart_rate")
    temp = load_signal(args.temp, "temperature")
    events = load_events(args.events)

    hr_aligned = align_signal(
        hr,
        events,
        "heart_rate",
        before_min=args.before_min,
        after_min=args.after_min,
    )
    temp_aligned = align_signal(
        temp,
        events,
        "temperature",
        before_min=args.before_min,
        after_min=args.after_min,
    )

    hr_aligned.to_csv(output_dir / "heart_rate_aligned.csv", index=False)
    temp_aligned.to_csv(output_dir / "temperature_aligned.csv", index=False)
    plot_aligned(hr_aligned, temp_aligned, output_dir / "aligned_signals.png")


if __name__ == "__main__":
    main()
