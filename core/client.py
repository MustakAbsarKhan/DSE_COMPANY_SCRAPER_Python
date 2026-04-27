import aiohttp
import asyncio
import random
from core.logger import logger


class AdaptiveThrottle:
    """
    Shared request throttle for the whole scraper.

    It controls two things:
    1. Delay between requests, so DSE is not hit too aggressively.
    2. Maximum concurrent requests, so the scraper can slow down after errors
       and speed up gradually after stable success.
    """

    def __init__(self, initial_concurrency=5, min_concurrency=1, max_concurrency=10, min_delay=0.5, max_delay=3):
        # Delay values are adaptive. The scraper starts at min_delay, then
        # increases delay on failures and decreases it on successful responses.
        self.delay = min_delay
        self.min_delay = min_delay
        self.max_delay = max_delay

        # Concurrency is adaptive too. A failure reduces it quickly; repeated
        # successes raise it slowly.
        self.concurrency = initial_concurrency
        self.min_concurrency = min_concurrency
        self.max_concurrency = max_concurrency

        # active_requests tracks how many requests are currently inside fetch().
        self.active_requests = 0
        self.success_streak = 0

        # Condition lets waiting tasks sleep until another request finishes or
        # the concurrency limit changes.
        self.condition = asyncio.Condition()

    async def acquire(self):
        """Wait until the current adaptive concurrency limit has free space."""
        async with self.condition:
            while self.active_requests >= self.concurrency:
                await self.condition.wait()
            self.active_requests += 1

    async def release(self):
        """Mark one request as finished and wake waiting requests."""
        async with self.condition:
            self.active_requests -= 1
            self.condition.notify_all()

    async def wait_delay(self):
        # Random jitter prevents every concurrent task from firing at exactly
        # the same interval, which is friendlier to the DSE server.
        await asyncio.sleep(self.delay + random.uniform(0, 0.5))

    def success(self):
        # Decrease delay and gently raise concurrency after stable successes.
        self.delay = max(self.min_delay, self.delay * 0.9)
        self.success_streak += 1

        if self.success_streak >= 20 and self.concurrency < self.max_concurrency:
            self.concurrency += 1
            self.success_streak = 0
            logger.info(f"THROTTLE: concurrency increased to {self.concurrency}")

    def failure(self):
        # Increase delay and reduce concurrency quickly when DSE pushes back.
        self.delay = min(self.max_delay, self.delay * 1.5)
        self.success_streak = 0

        if self.concurrency > self.min_concurrency:
            self.concurrency -= 1
            logger.warning(f"THROTTLE: concurrency reduced to {self.concurrency}")


class AsyncClient:
    """Small async HTTP client with shared adaptive throttling and retries."""

    def __init__(self, concurrency=5, min_concurrency=1, max_concurrency=10, retries=3):
        self.throttle = AdaptiveThrottle(
            initial_concurrency=concurrency,
            min_concurrency=min_concurrency,
            max_concurrency=max_concurrency
        )
        self.retries = retries

    async def fetch(self, session, url):
        """Fetch one URL and return HTML text, or None after all retries fail."""
        await self.throttle.acquire()

        try:
            for attempt in range(1, self.retries + 1):
                try:
                    # Respect the current global adaptive delay before each try.
                    await self.throttle.wait_delay()

                    async with session.get(url, timeout=10) as response:
                        response.raise_for_status()
                        text = await response.text()

                        self.throttle.success()
                        logger.info(f"SUCCESS: {url}")

                        return text

                except Exception as error:
                    # Any HTTP/network/timeout error counts as pressure from
                    # the site or connection, so slow down globally.
                    self.throttle.failure()

                    if attempt < self.retries:
                        logger.warning(
                            f"RETRY {attempt}/{self.retries}: {url} | {type(error).__name__}: {error}"
                        )
                        await asyncio.sleep(attempt + random.uniform(0, 1))
                    else:
                        logger.error(
                            f"FAILED: {url} | {type(error).__name__}: {error}"
                        )

            return None

        finally:
            # Always release the concurrency slot, even if the request failed.
            await self.throttle.release()

    async def fetch_all(self, urls):
        """Fetch many URLs concurrently through the same adaptive throttle."""
        async with aiohttp.ClientSession(
            headers={"User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/91.0.4472.124 Safari/537.36"}
        ) as session:

            tasks = [self.fetch(session, url) for url in urls]
            return await asyncio.gather(*tasks)


# One global client is reused by all pipelines. This lets adaptive delay and
# adaptive concurrency learn from the full scraper run instead of restarting in
# every module.
global_client = AsyncClient(concurrency=5, min_concurrency=1, max_concurrency=10, retries=3)
