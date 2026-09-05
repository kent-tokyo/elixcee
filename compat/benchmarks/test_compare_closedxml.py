"""Statistics checks independent of installed .NET and runtime timings."""
import unittest

from compare_closedxml import summarize


def sample(value):
    return dict.fromkeys(('load_ms', 'mutate_ms', 'save_ms', 'reload_ms', 'total_ms'), value)


class SummaryTests(unittest.TestCase):
    def test_even_count_median_and_nearest_rank_p95(self):
        result = summarize([sample(n) for n in range(1, 31)])
        self.assertEqual(result['total_ms'], {'p50': 15.5, 'p95': 29})

    def test_invalid_or_missing_samples_are_rejected(self):
        for values in ([], [float('nan')], [float('inf')], [-1]):
            with self.subTest(values=values), self.assertRaises(AssertionError):
                summarize([sample(n) for n in values])


if __name__ == '__main__':
    unittest.main()
