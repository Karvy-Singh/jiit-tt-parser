import os
import json

PROG = "jiit-time-table"
CACHE_HOME = os.getenv("XDG_CACHE_HOME") or os.path.expanduser("~/.cache")
CACHE_DIR = os.path.join(CACHE_HOME, PROG)

def get_cache_file(file: str):
    p = os.path.join(CACHE_DIR, file)
    os.makedirs(CACHE_DIR, exist_ok=True)
    return p


def load_faculty_map(path: str):
    with open(path) as f:
        return json.load(f)
