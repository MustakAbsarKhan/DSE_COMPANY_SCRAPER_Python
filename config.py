# Base DSE website URL. Other endpoints are built from this value.
DOMAIN = "https://www.dsebd.org/"

# Page that lists all DSE companies grouped by industry/sector.
MAIN_URL = DOMAIN + "by_industrylisting.php"

# These sector names are intentionally skipped because they are not normal
# listed equity companies and can have different page/data structures.
IGNORED_SECTORS = [
    "Corporate Bond",
    "Debenture",
    "G-SEC (T.Bond)"
]
