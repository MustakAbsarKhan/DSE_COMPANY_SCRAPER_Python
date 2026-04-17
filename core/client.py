import aiohttp
import asyncio
import random
import time
from core.logger import logger


# Implements an adaptive rate limiter that adjusts the delay between requests based on response times and error rates.
class AdaptiveRateLimiter:
    def __init__(self, min_delay=0.5, max_delay=3.0):
        self.delay = min_delay
        self.min_delay = min_delay
        self.max_delay = max_delay
        
    async def wait(self):# Waits for the current delay duration before allowing the next request.
        await asyncio.sleep(self.delay + random.uniform(0, 0.5))  # Add random jitter to avoid patterns
        
    def success(self): # Reduces the delay after a successful request, allowing for faster subsequent requests.
        self.delay = max(self.min_delay, self.delay * 0.9)  # Decrease delay by 10%
        
    def failure(self): # Increases the delay after a failed request, slowing down subsequent requests to avoid overwhelming the server.
        self.delay = min(self.max_delay, self.delay * 1.5)  # Increase delay by 50%
        
        
        
        
# An asynchronous HTTP client that uses the AdaptiveRateLimiter to manage request rates and handle retries with exponential backoff.
class AsyncClient:
    def __init__(self, concurrency=5): # Initializes the AsyncClient with a specified concurrency level and an instance of the AdaptiveRateLimiter.
        self.concurrency = concurrency # Sets the maximum number of concurrent requests allowed.
        self.rate_limiter = AdaptiveRateLimiter() # Initializes the adaptive rate limiter to manage request delays based on response times and error rates.
        self.semaphore = asyncio.Semaphore(concurrency)  # Limit concurrent requests
        
    async def fetch(self, session, url): # Fetches the content of a URL using an aiohttp session, while respecting the adaptive rate limits and handling retries with exponential backoff.
        async with self.semaphore: # Acquires a semaphore to ensure that the number of concurrent requests does not exceed the specified concurrency level.
            await self.rate_limiter.wait() # Waits for the adaptive rate limiter to allow the next request, which may include a delay based on previous response times and error rates.
            
            try:
                async with session.get(url, timeout=10) as response: # Makes an asynchronous GET request to the specified URL using the aiohttp session.
                    response.raise_for_status() # Raises an exception if the response status code indicates an error (e.g., 4xx or 5xx).
                    text = await response.text() # Reads the response content as text.
                    
                    self.rate_limiter.success() # Notifies the adaptive rate limiter of a successful request, which may reduce the delay for subsequent requests.
                    logger.info(f"Successfully fetched: {url}") # Logs a message indicating that the URL was successfully fetched.
                    
                    return text # Returns the fetched content as text.
        
            except Exception as e: # Catches any exceptions that occur during the request, such as timeouts or HTTP errors.
                self.rate_limiter.failure() # Notifies the adaptive rate limiter of a failed request, which may increase the delay for subsequent requests.
                logger.error(f"Error fetching {url}: {e}") # Logs an error message with details about the exception that occurred while fetching the URL.
                return None # Returns None to indicate that the fetch operation was unsuccessful.
            
    
    async def run(self, urls): # Runs the asynchronous fetching of multiple URLs concurrently, while managing the adaptive rate limits and handling retries.
        async with aiohttp.ClientSession() as session: # Creates an aiohttp ClientSession for making HTTP requests.
            tasks = [self.fetch(session, url) for url in urls] # Creates a list of tasks for fetching each URL using the fetch method.
            return await asyncio.gather(*tasks) # Executes the tasks concurrently and waits for all of them to complete, returning their results as a list.