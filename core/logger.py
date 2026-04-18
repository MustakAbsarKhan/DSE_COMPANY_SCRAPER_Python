'''# Configures a logger to output formatted INFO-level messages to both the console and 'DSECompanyScraper.log'.'''
import logging as log

def setup_logger():
    logger = log.getLogger("DSECompanyScraper")
    logger.setLevel(log.INFO)
    
    formatter = log.Formatter(
        "%(asctime)s - %(name)s - %(levelname)s - %(message)s"
        #Log message format example: 2024-06-01 12:00:00,000 - DSECompanyScraper - INFO - Scraping company info
    )
    
    #console handler AKA CH
    ch = log.StreamHandler()
    ch.setFormatter(formatter)
    
    #file handler AKA FH
    fh = log.FileHandler("DSECompanyScraper.log")
    fh.setFormatter(formatter)
    
    logger.addHandler(ch)
    logger.addHandler(fh)
    
    return logger

# Initialize the logger at the module level so it can be imported and used across the project
logger = setup_logger()