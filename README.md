# ThermoAnalysis

Privacy-safe utilities for aligning and visualizing heart-rate and core-temperature time series around experimental events.

## Purpose

This repository is a clean, public methodological demo. It contains no participant names, study schedules, local machine paths, or study data.

The original internal analysis scripts were replaced with a minimal workflow that shows the reusable analysis logic only.

## Workflow

1. Load heart-rate and temperature CSV files.
2. Load a separate event table with anonymous subject IDs and event timestamps.
3. Align each signal to event onset.
4. Export relative-time tables.
5. Plot aligned heart-rate and temperature traces.

## Expected input

Heart-rate CSV:

    subject_id,timestamp,heart_rate
    S01,2026-01-01 14:00:00,72

Temperature CSV:

    subject_id,timestamp,temperature
    S01,2026-01-01 14:00:00,37.1

Event CSV:

    subject_id,event_id,event_time
    S01,E1,2026-01-01 14:05:00

The example values above are fictional and only illustrate the file format.

## Setup

    python3 -m venv .venv
    source .venv/bin/activate
    pip install -r requirements.txt

## Usage

    python align_signals.py \
      --hr heart_rate.csv \
      --temp temperature.csv \
      --events events.csv \
      --output output

## Output

- output/heart_rate_aligned.csv
- output/temperature_aligned.csv
- output/aligned_signals.png

## Validation

Run the unit tests with:

    python -m unittest discover -s tests -v

## Scope

This is a methodological portfolio example for physiological time-series alignment and visualization. It is not a clinical tool and does not contain real participant data.
