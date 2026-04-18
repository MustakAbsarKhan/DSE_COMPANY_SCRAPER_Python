import logging

DOMAIN = "https://www.dsebd.org/"

MAIN_URL = DOMAIN + "by_industrylisting.php"

DEFAULT_IGNORED_SECTORS = [
    "Corporate Bond",
    "Debenture",
    "G-SEC (T.Bond)"
]

logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s [%(levelname)s] %(message)s",
    handlers=[logging.FileHandler("scraper.log"), logging.StreamHandler()],
    DEBUG = True
)