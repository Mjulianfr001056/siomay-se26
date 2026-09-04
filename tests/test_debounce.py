"""Tests for reusable UI event debounce behavior."""

import unittest

from utils.debounce import DebounceGate


class FakeClock:
    def __init__(self):
        self.now = 0.0

    def __call__(self):
        return self.now


class DebounceGateTests(unittest.TestCase):
    def setUp(self):
        self.clock = FakeClock()
        self.gate = DebounceGate(0.5, clock=self.clock)

    def test_rejects_clicks_inside_interval(self):
        self.assertTrue(self.gate.accept())

        self.clock.now = 0.1
        self.assertFalse(self.gate.accept())
        self.clock.now = 0.49
        self.assertFalse(self.gate.accept())

    def test_accepts_click_at_interval_boundary(self):
        self.assertTrue(self.gate.accept())

        self.clock.now = 0.5
        self.assertTrue(self.gate.accept())

    def test_rejected_click_does_not_extend_interval(self):
        self.assertTrue(self.gate.accept())

        self.clock.now = 0.4
        self.assertFalse(self.gate.accept())
        self.clock.now = 0.6
        self.assertTrue(self.gate.accept())

    def test_reset_allows_immediate_click(self):
        self.assertTrue(self.gate.accept())
        self.assertFalse(self.gate.accept())

        self.gate.reset()
        self.assertTrue(self.gate.accept())

    def test_negative_interval_is_rejected(self):
        with self.assertRaises(ValueError):
            DebounceGate(-0.1)


if __name__ == "__main__":
    unittest.main()