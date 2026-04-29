"""Logging architecture preview.

The private project uses a shared logger for terminal output and runtime log
files. The public preview avoids creating local log files on import.
"""

from __future__ import annotations

import logging


logger = logging.getLogger("DSECompanyScraperPreview")
