
import traceback
from debug_logging import debug

def log_error(message: str, error_log_path: str):
    """Append an error message and traceback to the specified log file."""
    with open(error_log_path, 'a', encoding='utf-8') as f:
        f.write(message + '\n')
        f.write(traceback.format_exc() + '\n')
    # Also log to debug window
    try:
        debug(message)
        debug(traceback.format_exc())
    except Exception:
        pass
