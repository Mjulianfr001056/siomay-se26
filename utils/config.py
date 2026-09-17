"""Runtime configuration loaded from process variables or a local ``.env`` file."""

from __future__ import annotations

import ast
import os
import re
import sys
from pathlib import Path


_KEY_RE = re.compile(r"[A-Za-z_][A-Za-z0-9_]*")


def _dotenv_paths():
    """Return likely .env locations for source and packaged executions."""
    candidates = (
        Path(__file__).resolve().parents[1] / ".env",
        Path(sys.executable).resolve().parent / ".env",
        Path.cwd() / ".env",
    )
    return tuple(dict.fromkeys(candidates))


def _read_dotenv(path: Path):
    """Read simple dotenv assignments without modifying ``os.environ``."""
    try:
        lines = path.read_text(encoding="utf-8-sig").splitlines()
    except (OSError, UnicodeError):
        return {}

    values = {}
    for raw_line in lines:
        line = raw_line.strip()
        if not line or line.startswith("#"):
            continue
        if line.startswith("export "):
            line = line[7:].lstrip()
        if "=" not in line:
            continue
        key, raw_value = line.split("=", 1)
        key = key.strip()
        if not _KEY_RE.fullmatch(key):
            continue
        value = raw_value.strip()
        if len(value) >= 2 and value[0] == value[-1] and value[0] in "\"'":
            try:
                value = ast.literal_eval(value)
            except (SyntaxError, ValueError):
                continue
        else:
            value = value.split(" #", 1)[0].rstrip()
        values[key] = value
    return values


def get_config(name: str, default: str = "") -> str:
    """Get one setting, preferring the process environment over ``.env``."""
    if name in os.environ:
        return os.environ[name]
    for path in _dotenv_paths():
        values = _read_dotenv(path)
        if name in values:
            return values[name]
    return default