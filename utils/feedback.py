"""Persistent state for the optional SIOMAY user-feedback prompt."""

from __future__ import annotations

import json
import os
from pathlib import Path


FEEDBACK_URL = "http://s.bps.go.id/FeedbackSIOMAY"


def feedback_state_path() -> Path:
    """Return the per-user state location without writing into the app folder."""
    local_app_data = os.environ.get("LOCALAPPDATA")
    base = Path(local_app_data) if local_app_data else Path.home() / ".siomay"
    return base / "SIOMAY" / "feedback.json" if local_app_data else base / "feedback.json"


def _read_state(path: Path) -> dict:
    try:
        data = json.loads(path.read_text(encoding="utf-8"))
    except (OSError, json.JSONDecodeError):
        return {}
    return data if isinstance(data, dict) else {}


def _write_state(path: Path, state: dict) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    temporary_path = path.with_suffix(".tmp")
    temporary_path.write_text(json.dumps(state), encoding="utf-8")
    temporary_path.replace(path)


def record_launch_and_should_prompt(path: Path | None = None) -> bool:
    """Record an app launch and request the prompt from launch two onward.

    The prompt remains eligible until the user explicitly chooses one of its
    actions. Persistence failures are non-fatal and suppress the prompt.
    """
    state_path = path or feedback_state_path()
    state = _read_state(state_path)
    launches = state.get("launches", 0)
    if not isinstance(launches, int) or isinstance(launches, bool) or launches < 0:
        launches = 0
    launches += 1
    state["launches"] = launches
    state["dismissed"] = state.get("dismissed") is True
    try:
        _write_state(state_path, state)
    except OSError:
        return False
    return launches >= 2 and not state["dismissed"]


def dismiss_feedback_prompt(path: Path | None = None) -> None:
    """Permanently dismiss the automatic prompt for the current user."""
    state_path = path or feedback_state_path()
    state = _read_state(state_path)
    state["dismissed"] = True
    try:
        _write_state(state_path, state)
    except OSError:
        pass