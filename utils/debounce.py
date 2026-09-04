"""Small event guards for preventing accidental repeated UI actions."""

from __future__ import annotations

import time
from collections.abc import Callable


class DebounceGate:
    """Accept at most one event during each debounce interval."""

    def __init__(
        self,
        interval_seconds: float,
        clock: Callable[[], float] = time.monotonic,
    ) -> None:
        if interval_seconds < 0:
            raise ValueError("Debounce interval must not be negative")
        self._interval_seconds = interval_seconds
        self._clock = clock
        self._last_accepted_at: float | None = None

    def accept(self) -> bool:
        """Return ``True`` for an accepted event and ``False`` for a duplicate."""
        now = self._clock()
        if (
            self._last_accepted_at is not None
            and now - self._last_accepted_at < self._interval_seconds
        ):
            return False
        self._last_accepted_at = now
        return True

    def reset(self) -> None:
        """Allow the next event immediately."""
        self._last_accepted_at = None