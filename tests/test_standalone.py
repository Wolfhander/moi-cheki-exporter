import sys
import tempfile
import unittest
from datetime import datetime
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parents[1]))
from returns_import import OzonReturnsDB
from returns_match import match_returns
from app_utils import parse_russian_date


class StandaloneTests(unittest.TestCase):
    def test_returns_import_has_local_dependencies(self):
        self.assertEqual(parse_russian_date("Заявка от 4 ноября 2023"), datetime(2023, 11, 4))
        self.assertIsNone(parse_russian_date("31 февраля 2024"))
        self.assertTrue(callable(match_returns))

    def test_empty_returns_database_has_numeric_counters(self):
        with tempfile.TemporaryDirectory() as temp:
            db = OzonReturnsDB(str(Path(temp) / "receipts.db"))
            self.assertEqual(db.stats(), {"total": 0, "refunded_count": 0,
                                         "refunded_sum": 0, "matched_count": 0})


if __name__ == "__main__":
    unittest.main()
