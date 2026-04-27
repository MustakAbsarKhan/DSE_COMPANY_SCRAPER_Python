import aiohttp
from bs4 import BeautifulSoup as bs
from datetime import datetime, date, timedelta
import asyncio


class HolidayChecker:
    """Fetch DSE holidays and decide whether scraping should run today."""

    def __init__(self):
        # Parsed holiday dates are stored here after fetch_holidays() runs.
        self.holidays = []

        # DSE holiday/trading schedule page.
        self.url = "https://www.dsebd.org/hts.php"

    async def fetch_holidays(self):
        """Fetch the DSE holiday page and parse all listed holiday dates."""
        # Short delay keeps this auxiliary request from firing immediately at
        # program start with the rest of the scraper.
        await asyncio.sleep(2)

        try:
            async with aiohttp.ClientSession(
                headers={"User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36"}
            ) as session:
                async with session.get(self.url, timeout=10) as response:
                    response.raise_for_status()
                    html = await response.text()

            soup = bs(html, "lxml")
            self.holidays = self._parse_holidays(soup)
            return True

        except Exception as error:
            print(f"❌ Failed to fetch holidays: {error}")
            return False

    def _parse_holidays(self, soup):
        """Parse holiday rows from the DSE holiday table."""
        holidays = []

        # The page contains one main table. If DSE changes the layout and no
        # table is found, return an empty list instead of crashing.
        table = soup.find("table")
        if not table:
            return holidays

        # Skip header row. Each remaining row contains a holiday date or range.
        rows = table.find_all("tr")[1:]

        for row in rows:
            cols = row.find_all("td")
            if len(cols) >= 2:
                # Column 2 contains strings like "04 February" or
                # "11-12 February".
                date_str = cols[1].get_text(strip=True)
                holidays.extend(self._parse_date_range(date_str))

        # Debug output showing the parsed holiday dates.
        print(holidays)#test
        return holidays

    def _parse_date_range(self, date_str):
        """Convert one holiday date/range string into date objects."""
        dates = []
        current_year = datetime.now().year

        # Handle ranges like "11-12 February"
        if "-" in date_str:
            parts = date_str.split("-")
            if len(parts) == 2:
                start_str = parts[0].strip() + " " + " ".join(date_str.split()[1:])
                end_str = parts[1].strip() + " " + " ".join(date_str.split()[1:])

                try:
                    start_date = self._parse_single_date(start_str, current_year)
                    end_date = self._parse_single_date(end_str, current_year)

                    if start_date and end_date:
                        current = start_date
                        while current <= end_date:
                            dates.append(current)
                            current = current + timedelta(days=1)
                except:
                    pass
        else:
            # Single date like "04 February"
            parsed_date = self._parse_single_date(date_str, current_year)
            if parsed_date:
                dates.append(parsed_date)

        return dates

    def _parse_single_date(self, date_str, year):
        """Parse a single day/month string into a Python date."""
        try:
            # Remove extra spaces and split
            parts = date_str.split()
            if len(parts) >= 2:
                day = int(parts[0])
                month_name = parts[1]

                # Convert month name to number
                month_names = {
                    'January': 1, 'February': 2, 'March': 3, 'April': 4, 'May': 5, 'June': 6,
                    'July': 7, 'August': 8, 'September': 9, 'October': 10, 'November': 11, 'December': 12
                }

                month = month_names.get(month_name)
                if month:
                    return date(year, month, day)
        except:
            pass
        return None

    def is_holiday(self, check_date=None):
        """Check whether a date is a weekend or listed DSE holiday."""
        if check_date is None:
            check_date = date.today()

        # Check weekends (Friday = 4, Saturday = 5 in Python's weekday())
        if check_date.weekday() in [4, 5]:  # Friday and Saturday
            print("Weekend (Friday/Saturday)")
            return True, "Weekend (Friday/Saturday)"

        # Check holidays
        for holiday in self.holidays:
            if holiday == check_date:
                print("Public Holiday")
                return True, "Public Holiday"

        return False, None

    async def check_and_exit_if_holiday(self):
        """Return True when the scraper should stop because the market is closed."""
        success = await self.fetch_holidays()
        if not success:
            print("⚠️  Warning: Could not fetch holiday data. Proceeding with scraping...")
            return False

        is_holiday, reason = self.is_holiday()

        if is_holiday:
            print(f"🏖️  Today is a {reason}. Market is closed. Exiting gracefully.")
            return True

        print("✅ Today is a trading day. Proceeding with scraping...")
        return False


# Global instance
holiday_checker = HolidayChecker()
