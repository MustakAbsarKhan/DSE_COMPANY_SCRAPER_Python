"""Public configuration preview.

The production scraper uses these values as inputs to the private discovery
pipeline. They are kept here to show the project's configuration boundary.
"""

DOMAIN = "https://www.dsebd.org/"
MAIN_URL = DOMAIN + "by_industrylisting.php"

IGNORED_SECTORS = [
    "Corporate Bond",
    "Debenture",
    "G-SEC (T.Bond)",
]
