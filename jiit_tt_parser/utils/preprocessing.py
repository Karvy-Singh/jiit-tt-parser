from jiit_tt_parser.utils.utils import cvt_xls_to_xlsx, download
from jiit_tt_parser.utils.cache import get_cache_file
from jiit_tt_parser.utils.cache import LINKS
from jiit_tt_parser.parser.parse_faculty import cache_faculty_map, get_faculty_map
import tempfile
import os
import shutil


def cache_tt_xls():
    print("Caching Time Table spreadsheets...")
    tmp_dir = tempfile.gettempdir()
    shutil.copy("bca_sem_1_3_new.xlsx", get_cache_file("bca_sem1_new.xlsx"))   
    for k, v in LINKS.items():
        ext = v.split(".")[-1]
        file_name = f"{k}.{ext}"
        print(f"Downloading {file_name}... ", end="")
        if ext == "xls":
            print("done\nConverting xls to xlsx... ", end="")
            tmp_file = os.path.join(tmp_dir, file_name)
            download(v, tmp_file)
            cvt_xls_to_xlsx(tmp_file, get_cache_file(file_name+"x"))
            os.remove(tmp_file)
        else:
            download(v, get_cache_file(file_name))

        print("done")


def cache_fac():
    print("Checking faculty spreadsheets cache... ", end="")
    assert os.path.exists(sem1:=get_cache_file("sem1.xlsx"))
    assert os.path.exists(fac1:=get_cache_file("fac1.xlsx"))
    assert os.path.exists(fac2:=get_cache_file("fac2.xlsx"))
    assert os.path.exists(bca_sem1:=get_cache_file("bca_sem1.xlsx"))
    print("found")
    
    print("Generating faculty maps... ", end="")
    c = get_faculty_map(fac1, fac2, sem1, bca_sem1)
    cache_faculty_map(c)

    print("done")



if __name__ == "__main__":
    cache_tt_xls()
    print()
    cache_fac()
