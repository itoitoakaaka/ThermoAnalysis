import tempfile
import unittest
from pathlib import Path

import pandas as pd

from align_signals import align_signal, load_events, load_signal


class ThermoAnalysisTests(unittest.TestCase):
    def test_alignment_uses_relative_event_time(self):
        signal = pd.DataFrame(
            {
                "subject_id": ["S01", "S01", "S01"],
                "timestamp": pd.to_datetime(
                    [
                        "2026-01-01 14:04:00",
                        "2026-01-01 14:05:00",
                        "2026-01-01 14:06:00",
                    ]
                ),
                "heart_rate": [70.0, 75.0, 80.0],
            }
        )
        events = pd.DataFrame(
            {
                "subject_id": ["S01"],
                "event_id": ["E1"],
                "event_time": pd.to_datetime(["2026-01-01 14:05:00"]),
            }
        )

        aligned = align_signal(signal, events, "heart_rate", before_min=1, after_min=1)

        self.assertEqual(aligned["relative_min"].tolist(), [-1.0, 0.0, 1.0])
        self.assertEqual(aligned["heart_rate"].tolist(), [70.0, 75.0, 80.0])

    def test_alignment_keeps_subjects_separate(self):
        signal = pd.DataFrame(
            {
                "subject_id": ["S01", "S02"],
                "timestamp": pd.to_datetime(
                    ["2026-01-01 14:00:00", "2026-01-01 14:00:00"]
                ),
                "temperature": [37.0, 38.0],
            }
        )
        events = pd.DataFrame(
            {
                "subject_id": ["S01"],
                "event_id": ["E1"],
                "event_time": pd.to_datetime(["2026-01-01 14:00:00"]),
            }
        )

        aligned = align_signal(signal, events, "temperature")

        self.assertEqual(len(aligned), 1)
        self.assertEqual(aligned.iloc[0]["subject_id"], "S01")
        self.assertEqual(aligned.iloc[0]["temperature"], 37.0)

    def test_loaders_validate_columns(self):
        with tempfile.TemporaryDirectory() as directory:
            root = Path(directory)
            bad_signal = root / "bad_signal.csv"
            bad_events = root / "bad_events.csv"

            pd.DataFrame({"subject_id": ["S01"]}).to_csv(bad_signal, index=False)
            pd.DataFrame({"subject_id": ["S01"]}).to_csv(bad_events, index=False)

            with self.assertRaises(ValueError):
                load_signal(bad_signal, "heart_rate")

            with self.assertRaises(ValueError):
                load_events(bad_events)


if __name__ == "__main__":
    unittest.main()
