
import traceback
from debug_logging import debug

def log_error(message: str, error_log_path: str):
    """Append an error message and traceback to the specified log file."""
    try:
        with open(error_log_path, 'a', encoding='utf-8') as f:
            f.write(message + '\n')
            f.write(traceback.format_exc() + '\n')
    except Exception as file_ex:
        print("[ERROR LOGGING FAILURE]", message)
        print(traceback.format_exc())
    try:
        debug(message)
        debug(traceback.format_exc())
    except Exception as ex:
        print("[DEBUG WINDOW ERROR]", message)
        print(traceback.format_exc())
