import os
from pathlib import Path

def cache_dir() -> Path:
    cache_dir = os.path.expandvars(r'$LOCALAPPDATA\pdf-preview\cache')
    return Path(cache_dir)

def log_dir() -> Path:
    log_dir = os.path.expandvars(r'$LOCALAPPDATA\pdf-preview\log')
    return Path(log_dir)