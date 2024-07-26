from cache import get_cache_file

LINKS = {
  "sem1": "https://www.jiit.ac.in/sites/default/files/B%20Tech%20I%20Sem%20odd%202024_17%20July.xlsx",
  "sem3": "https://www.jiit.ac.in/sites/default/files/B%20Tech%20III%20Sem%2018%20JULY%205%20PM.xls",
  "fac1": "https://www.jiit.ac.in/sites/default/files/Faculty%20Abbreviations_0.xlsx",
  "fac2": "https://www.jiit.ac.in/sites/default/files/15.%20Faculty%20Abbreviations.xlsx"
}

FACULTY_MAP = get_cache_file("faculty.json")
