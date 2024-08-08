import datetime
from typing import Literal, List
import string

from openpyxl.cell import Cell
from openpyxl.worksheet.cell_range import CellRange
from openpyxl.worksheet.worksheet import Worksheet

from jiit_tt_parser.parser.parse_courses import parse_courses
from jiit_tt_parser.utils.utils import is_empty_row
from jiit_tt_parser.utils.cache import load_faculty_map, FACULTY_MAP


class Period:
    def __init__(
        self, start: datetime.time =  None, end: datetime.time = None
    ) -> None:
        self.start_time = start or datetime.time(0, 0)
        self.end_time = end or datetime.time(23, 59)

    @classmethod
    def from_string(cls, fmt: str):
        start, end = fmt.split("-")
        start = start.strip(" NO")
        end = end.strip("APM ")

        end_hour, end_min = end.split(".")
        end_hour = int(end_hour)
        end_min = int(end_min)

        start_hour = start
        start_min = "0"
        if "." in start:
            start_hour, start_min = start.split(".")

        start_hour = int(start_hour)
        start_min = int(start_min)

        if start_hour < 9:
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
    def from_string(cls, ev_str, period: Period, day: str, courses: dict, faculties: dict):
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


        ev.event_type, ev_str = ev_str[:1], ev_str[1:]
        ev.batches, ev_str = (
            ev_str[: ev_str.find("(")].split(","),
            ev_str[ev_str.find("(") :],
        )
        ev.batches = [i.strip() for i in ev.batches]
        ev.eventcode, ev_str = ev_str[1 : ev_str.find(")")], ev_str[ev_str.find(")") :]
        ev.event = courses.get(ev.eventcode.strip())
        
        while ev_str[0].upper() not in string.ascii_uppercase:
            ev_str = ev_str[1:]
        
        ev.classroom, ev_str = ev_str[: ev_str.find("/")].strip(), ev_str[ev_str.find("/") :]

        lecturer = ev_str[1:]
        lec_splitter = ","
        if '/' in lecturer:
            lec_splitter = "/"
        ev.lecturer = lecturer.split(lec_splitter)
        ev.lecturer = [faculties.get(i.strip()) or i for i in ev.lecturer]

        ev.period = period
        ev.day = day.capitalize()
        return ev

    def __str__(self) -> str:
        # print(repr(self.event_type))
        # print(self.event_string)
        lecture_types = {"L": "Lecture", "T": "Tutorial", "P": "Practical", "TALK": "Talk"}
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
        if str(v).startswith("9"):
            for j in range(2, col + 1):
                if sheet.cell(i, j).value is None:
                    return i, j-1
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


def is_end_of_day(sheet: Worksheet, col, curr, day):
    if day != "saturday":
        return sheet.cell(curr + 1, 1).value is not None

    return is_empty_row(sheet, curr, col)


def search_merged_cells(merged_cells: list[CellRange], cell: Cell) -> int :
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
    faculties: dict
):
    events = []
    if str(sheet.cell(start, 2).value).startswith("9"):
        start += 1

    for j in range(2, col + 1):
        r = start
        while not is_end_of_day(sheet, col, r, day):
            c = sheet.cell(r, j)
            if (
                ((v := c.value) is not None)
                and ("LUNCH" not in str(v))
                and (not str(v).isspace())
            ):
                ep = periods[j - 2]
                if m := search_merged_cells(merged_cells, c):
                    ep += periods[m - 2]
                events.append(Event.from_string(str(v), ep, day, courses, faculties))
            r += 1

    return events


def parse_events(sheet: Worksheet, row: int, col: int) -> List[Event]:
    time_row, col = get_time_row(sheet, row, col)
    periods = get_periods(sheet, row, col, time_row)
    merged_cells = sheet.merged_cells.sorted()
    courses = parse_courses(sheet, row, col)
    faculties = load_faculty_map(FACULTY_MAP)

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

        events.extend(parse_day(sheet, row, col, r, periods, day, merged_cells, courses, faculties))

    return events
