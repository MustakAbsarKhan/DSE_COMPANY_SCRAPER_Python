"""Project-wide logging setup.

Every module imports the same logger instance from this file. That keeps log
formatting consistent and writes both terminal output and a persistent
`DSECompanyScraper.log` file for later debugging.
"""

import logging as log


def setup_logger():
    """Create and configure the scraper logger."""
    logger = log.getLogger("DSECompanyScraper")
    logger.setLevel(log.INFO)

    # Common log format:
    # 2026-04-27 12:00:00,000 - DSECompanyScraper - INFO - message
    formatter = log.Formatter(
        "%(asctime)s - %(name)s - %(levelname)s - %(message)s"
    )

    # Console handler: shows logs while the scraper is running.
    ch = log.StreamHandler()
    ch.setFormatter(formatter)

    # File handler: keeps a permanent run history for failed URLs, retries, etc.
    fh = log.FileHandler("DSECompanyScraper.log")
    fh.setFormatter(formatter)

    logger.addHandler(ch)
    logger.addHandler(fh)

    return logger


# Initialize the logger once so every module can import and reuse it.
logger = setup_logger()
