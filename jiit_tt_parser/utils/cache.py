import os 
import json
from jiit_tt_parser.utils import PROG

CACHE_HOME = os.getenv("XDG_CACHE_HOME") or os.path.expanduser("~/.cache")
CACHE_DIR = os.path.join(CACHE_HOME, PROG)
LINKS = {
  "sem1": "https://www.jiit.ac.in/sites/default/files/B%20Tech%20I%20Sem%20odd%202024_17%20July.xlsx",
  "sem3": "https://www.jiit.ac.in/sites/default/files/B%20Tech%20III%20Sem%2018%20JULY%205%20PM.xls",
  "fac1": "https://www.jiit.ac.in/sites/default/files/Faculty%20Abbreviations_0.xlsx",
  "fac2": "https://www.jiit.ac.in/sites/default/files/15.%20Faculty%20Abbreviations.xlsx"
}


def get_cache_file(file: str):
    p = os.path.join(CACHE_DIR, file)
    os.makedirs(CACHE_DIR, exist_ok=True)
    return p

FACULTY_MAP = get_cache_file("faculty.json")

def load_faculty_map(path: str):
    with open(path) as f:
        return json.load(f)
