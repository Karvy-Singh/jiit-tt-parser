import datetime
from typing import Literal, List
import string
import re
import difflib

from openpyxl.styles import colors
from openpyxl.cell import Cell
from openpyxl.worksheet.cell_range import CellRange
from openpyxl.worksheet.worksheet import Worksheet

from jiit_tt_parser.parser.parse_courses import parse_courses
from jiit_tt_parser.parser.parse_electives import parse_electives
from jiit_tt_parser.utils.utils import is_empty_row
from jiit_tt_parser.utils.cache import load_faculty_map, FACULTY_MAP


class Period:
    def __init__(self, start: datetime.time = None, end: datetime.time = None) -> None:
        self.start_time = start or datetime.time(0, 0)
        self.end_time = end or datetime.time(23, 59)

    @classmethod
    def from_string(cls, fmt: str):
        start, end = fmt.split("-")
        start = start.strip(" NO")
        end = end.strip("APM ")

        end_hour, end_min = split_hour_min(end)
        start_hour, start_min = split_hour_min(start)

        if start_hour < 8:
            start_hour += 12
        if end_hour < 9:
            end_hour += 12
        start_time = datetime.time(start_hour, start_min)
        end_time = datetime.time(end_hour, end_min)

        return cls(start_time, end_time)

    def __add__(self, other):
        start_time = other.start_time
        end_time = other.end_time
        if self.start_time < other.start_time:
            start_time = self.start_time

        if self.end_time > other.end_time:
            end_time = self.end_time

        return Period(start_time, end_time)

    def __str__(self):
        return f"{self.start_time.hour}:{str(self.start_time.minute).zfill(2)} - {self.end_time.hour}:{str(self.end_time.minute).zfill(2)}"


class Event:
    def __init__(self, event_string: str):
        self.event_string = event_string
        self.batches: List[str]
        self.event_type: Literal["L", "T", "P", "TALK"]
        self.classroom: str
        self.event: str
        self.eventcode: str
        self.period: Period
        self.day: str
        self.lecturer: List[str]

    @classmethod
    def from_string(
            cls, ev_str, period: Period, day: str, courses: dict,electives:dict, faculties: dict
    ):
        ev_str = ev_str.strip()
        ev = cls(ev_str)
        if "TALK" in ev_str:
            ev.event_type = "TALK"
            ev.batches, ev_str = (
                ev_str[: ev_str.find("(")].split(","),
                ev_str[ev_str.find("(") :],
            )
            ev.batches = [i.strip() for i in ev.batches]
            ev.eventcode = "TALK"
            ev.event = "Talk"
            ev.classroom = ev_str.split("-")[-1].strip()
            ev.lecturer = []

            ev.period = period
            ev.day = day.capitalize()
            return ev
        
        if (ev_str.find("(")==-1 and ev_str not in ["SE","DE 1","HSS-2","OE2","DE 5","DE 6","DE 4"]):
            ev.event_type= "L"
            ev.batches= []
            ev.eventcode= ev_str[:ev_str.find("-")]
            ev.event=""
            if ev_str[:ev_str.find("-")] in electives.keys():
                ev.event= electives[ev.eventcode]
            ev_str= ev_str[ev_str.find("-"):]

        else:
            ev.event_type, ev_str = ev_str[:1], ev_str[1:]
            if ev.event_type not in ["T","P","L","TALK"]:
                return 

            raw_batches = ev_str[: ev_str.find("(")]
            ev_str = ev_str[ev_str.find("(") :]

            ev.batches= parse_batches(raw_batches) 

            ev.eventcode, ev_str = ev_str[1 : ev_str.find(")")], ev_str[ev_str.find(")") :]
            ev.eventcode= ev.eventcode.strip(" -")
            ev.event = courses.get(ev.eventcode)

            if ev.event is None:
                key = ev.eventcode.strip()
                matches = difflib.get_close_matches(key, courses.keys(), n=1, cutoff=0.8)
                if matches:
                    ev.event = courses[matches[0]]
            
        while ev_str and ev_str[0].upper() not in string.ascii_uppercase + string.digits:
            ev_str = ev_str[1:]

        match = re.search(r'[/\\]', ev_str)
        if match:
            sep_index = match.start()
            ev.classroom = ev_str[:sep_index].strip()
            ev_str = ev_str[sep_index:]
        else:
            if ev_str!="MSU":
                ev.classroom = ev_str.strip()
                ev_str = ""
            else:
                ev.classroom= ""

        if ev.classroom.upper() == "EDD" and "/" in ev_str[1:]:
            first, rest = ev_str[1:].split("/", 1)
            ev.classroom = first
            ev_str = "/" + rest
        
        lecturer=""
        if not(contains_number(ev.classroom)) and contains_number(ev_str[1:]):
            #print(f"classroom:{ev.classroom}\nlecturer:{ev_str[1:]}")
            if "TA" not in ev_str[1:]:
                lecturer= ev.classroom
                ev.classroom = ev_str[1:]
            else:
                lecturer= ev_str[1:]
#                 ev.classroom=""
                
            #print(f"NEW:\nclassroom:{ev.classroom}\nlecturer:{lecturer}")

        else:
            if ev_str=="MSU":
                lecturer= ev_str
            else: 
                lecturer = ev_str[1:]

        lec_splitter = ","
        if "/" in lecturer:
            lec_splitter = "/"
        ev.lecturer = lecturer.split(lec_splitter) 
        ev.lecturer = [faculties.get(i.strip()) or i.strip(" -") for i in ev.lecturer]

        
        ev.period = period
        ev.day = day.capitalize()
        return ev

    def __str__(self) -> str:
        # print(repr(self.event_type))
        # print(self.event_string)
        lecture_types = {
            "L": "Lecture",
            "T": "Tutorial",
            "P": "Practical",
            "TALK": "Talk",
        }

        return f"""Event: {lecture_types[self.event_type]}
Time: {self.period}
Day: {self.day}
Batches: {self.batches}
Subject: {self.event or self.eventcode}
Venue: {self.classroom}
Lecturer: {self.lecturer}
"""


def get_time_row(sheet: Worksheet, row, col):
    for i in range(1, row + 1):
        v = sheet.cell(i, 2).value
        if str(v).startswith("9") or str(v).startswith("8"):
            for j in range(2, col + 1):
                if sheet.cell(i, j).value is None:
                    return i, j - 1
            return i, col
    return 2, col


def get_day_row(sheet: Worksheet, row, col, day: str):
    day = day.lower()
    for i in range(1, row + 1):
        v = str(sheet.cell(i, 1).value).lower()
        if day.startswith(v):
            return i

    return -1


def get_periods(sheet: Worksheet, row, col, time_row):
    a = []
    for i in range(2, col + 1):
        p = Period.from_string(s := str(sheet.cell(time_row, i).value))
        a.append(p)
    
    return a


def is_end_of_day(sheet: Worksheet, curr, day):
    if curr >= 300:
        return True

    if day.lower() != "saturday":
        return sheet.cell(curr + 1, 1).value is not None

    # v = sheet.cell(curr, 1).value
    # print(v)
    theme = sheet.cell(curr, 1).fill.start_color.theme
    if theme is not None and theme == 1:
        return True
    return False


def search_merged_cells(merged_cells: list[CellRange], cell: Cell) -> int:
    for c in merged_cells:
        if c.min_row != c.max_row:
            continue

        if (c.min_row == cell.row) and (
            cell.column >= c.min_col and cell.column <= c.max_col
        ):
            return c.max_col

    return None


def parse_day(
    sheet: Worksheet,
    row,
    col,
    start,
    periods: List[Period],
    day: str,
    merged_cells: list[CellRange],
    courses: dict,
    electives: dict,
    faculties: dict,
):
    events = []
    if str(sheet.cell(start, 2).value).startswith("9"):
        start += 1

    for j in range(2, col + 1):
        r = start
        while not is_end_of_day(sheet, r, day):
            c = sheet.cell(r, j)
            if (
                ((v := c.value) is not None)
                and (s := str(v).strip())
                and (s.upper() not in ["LUNCH", "ALL BATCH FREE FOR MEETING"])
                and any(ch.isalpha() for ch in s)
                ):
                ep = periods[j - 2]
                if m := search_merged_cells(merged_cells, c):
                    ep += periods[m - 2]

                #print(periods)
                events.append(Event.from_string(str(v), ep, day, courses,electives, faculties))
            r += 1

    return events


def parse_events(
        sheet: Worksheet,sheet_electives: Worksheet, row: int, col: int, faculty_map_path: str = FACULTY_MAP
) -> List[Event]:
    time_row, col = get_time_row(sheet, row, col)
    periods = get_periods(sheet, row, col, time_row)
    merged_cells = sheet.merged_cells.sorted()
    courses = parse_courses(sheet, row, col)
    electives= parse_electives(sheet_electives)
    faculties = load_faculty_map(faculty_map_path)

    events = []

    days = [
        "monday",
        "tuesday",
        "wednesday",
        "thursday",
        "friday",
        "saturday",
        "sunday",
    ]
    for day in days:
        r = get_day_row(sheet, row, col, day)
        if r < 0:
            continue

        events.extend(
            parse_day(
                sheet, row, col, r, periods, day, merged_cells, courses, electives,faculties
            )
        )

    return events

def split_hour_min(time_str):
    if ":" in time_str or "." in time_str:
        parts = re.split(r'[:.]', time_str)
        hour = int(parts[0])
        minute = int(parts[1].strip("AMP ")) if len(parts) > 1 else 0
    else:
        hour = int(time_str)
        minute = 0
    return hour, minute

def contains_number(s):
    return any(char.isdigit() for char in s)

def parse_batches(raw_batches):
    tokens = re.findall(r'[A-Za-z]+\d+|\d+', raw_batches)
    batches = []
    prefix = None

    for t in tokens:
        if t[0].isdigit():
            if prefix is not None:
                batches.append(f"{prefix}{t}")
        else:
            batches.append(t)
            prefix = re.match(r'[A-Za-z]+', t).group()

    return batches


