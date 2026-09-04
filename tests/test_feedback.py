"""Tests for persistent feedback-prompt launch state."""

import json
import tempfile
import unittest
from pathlib import Path
from unittest.mock import patch

from utils.feedback import dismiss_feedback_prompt, record_launch_and_should_prompt


class FeedbackPromptStateTests(unittest.TestCase):
    def setUp(self):
        self.temp_dir = tempfile.TemporaryDirectory()
        self.state_path = Path(self.temp_dir.name) / "feedback.json"

    def tearDown(self):
        self.temp_dir.cleanup()

    def test_prompt_starts_on_second_launch(self):
        self.assertFalse(record_launch_and_should_prompt(self.state_path))
        self.assertTrue(record_launch_and_should_prompt(self.state_path))

    def test_prompt_remains_eligible_until_user_responds(self):
        record_launch_and_should_prompt(self.state_path)
        self.assertTrue(record_launch_and_should_prompt(self.state_path))
        self.assertTrue(record_launch_and_should_prompt(self.state_path))

    def test_dismissal_permanently_disables_automatic_prompt(self):
        record_launch_and_should_prompt(self.state_path)
        dismiss_feedback_prompt(self.state_path)

        self.assertFalse(record_launch_and_should_prompt(self.state_path))
        state = json.loads(self.state_path.read_text(encoding="utf-8"))
        self.assertTrue(state["dismissed"])

    def test_corrupt_state_is_recovered_as_first_launch(self):
        self.state_path.write_text("not json", encoding="utf-8")

        self.assertFalse(record_launch_and_should_prompt(self.state_path))
        state = json.loads(self.state_path.read_text(encoding="utf-8"))
        self.assertEqual(state["launches"], 1)

    @patch("utils.feedback._write_state", side_effect=OSError("read-only"))
    def test_persistence_failure_does_not_show_prompt_or_raise(self, _mock_write):
        self.assertFalse(record_launch_and_should_prompt(self.state_path))
        dismiss_feedback_prompt(self.state_path)