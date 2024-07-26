import openpyxl
import tempfile
import shutil
import os
import sys

# from openpyxl.worksheet.cell_range import CellRange

from utils.utils import cvt_xls_to_xlsx, max_bounds, print_worksheet
# from parser.parse_courses import parse_courses
from parser.parse_events import parse_events



if __name__ == "__main__":
    TT_PATH = sys.argv[1]
    if not TT_PATH.endswith(".xlsx"):
        cvt_xls_to_xlsx(TT_PATH, (TT_PATH:=os.path.join(tempfile.gettempdir(), 'doc.xlsx')))

    PATH = os.path.join(tempfile.gettempdir(), "cec.xlsx")
    shutil.copyfile(TT_PATH, PATH)
    wb = openpyxl.load_workbook(PATH)
    sheet = wb.active
    r, c = max_bounds(sheet)
    # print_worksheet(sheet, r, c)
    # p(sheet, r,c )

    from pprint import pprint
    # pprint(parse_courses(sheet, r, c))
    # pprint(get_faculty_map("./faculty.xlsx", "./ttsem1.xlsx"))

    parse_events(sheet, r, c)
    
    os.remove(PATH)

    if not TT_PATH.endswith(".xlsx"):
        os.remove(TT_PATH)
