import logging
import os

# Set up a default logger for debugging
DEBUG_LOG_FILENAME = 'debug.log'

# Configure logging (file + console)
logging.basicConfig(
    level=logging.DEBUG,
    format='%(asctime)s [%(levelname)s] %(message)s',
    handlers=[
        logging.FileHandler(DEBUG_LOG_FILENAME, encoding='utf-8'),
        logging.StreamHandler()
    ]
)

def debug(msg):
    """Log a debug message."""
    logging.debug(msg)

def info(msg):
    """Log an info message."""
    logging.info(msg)

def warning(msg):
    """Log a warning message."""
    logging.warning(msg)

def error(msg):
    """Log an error message."""
    logging.error(msg)
