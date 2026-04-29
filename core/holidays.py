"""Market-calendar architecture preview.

The private implementation can skip scheduled runs on weekends or official DSE
market holidays. Runtime holiday fetching and parsing are excluded here.
"""

from __future__ import annotations

from dataclasses import dataclass
from datetime import date


@dataclass
class HolidayChecker:
    """Public interface preview for market-closed checks."""

    async def fetch_holidays(self) -> bool:
        raise NotImplementedError("Private implementation omitted from public preview.")

    def is_holiday(self, check_date: date | None = None) -> tuple[bool, str | None]:
        raise NotImplementedError("Private implementation omitted from public preview.")

    async def check_and_exit_if_holiday(self) -> bool:
        raise NotImplementedError("Private implementation omitted from public preview.")


holiday_checker = HolidayChecker()
