import aiohttp
import asyncio
import random
from core.logger import logger


class AdaptiveRateLimiter:
    def __init__(self, min_delay=0.5, max_delay=3):
        self.delay = min_delay
        self.min_delay = min_delay
        self.max_delay = max_delay

    async def wait(self):
        # Add random jitter
        await asyncio.sleep(self.delay + random.uniform(0, 0.5))

    def success(self):
        # decrease delay slightly if successful
        self.delay = max(self.min_delay, self.delay * 0.9)

    def failure(self):
        # increase delay if blocked/errors
        self.delay = min(self.max_delay, self.delay * 1.5)


class AsyncClient:
    def __init__(self, concurrency=5):
        self.semaphore = asyncio.Semaphore(concurrency)
        self.rate_limiter = AdaptiveRateLimiter()

    async def fetch(self, session, url):
        async with self.semaphore:
            await self.rate_limiter.wait()

            try:
                async with session.get(url, timeout=10) as res:
                    res.raise_for_status()
                    text = await res.text()

                    self.rate_limiter.success()
                    logger.info(f"SUCCESS: {url}")

                    return text

            except Exception as e:
                self.rate_limiter.failure()
                logger.error(f"FAILED: {url} | {e}")
                return None

    async def fetch_all(self, urls):
        async with aiohttp.ClientSession(
            headers={"User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/91.0.4472.124 Safari/537.36"}
        ) as session:

            tasks = [self.fetch(session, url) for url in urls]
            return await asyncio.gather(*tasks)