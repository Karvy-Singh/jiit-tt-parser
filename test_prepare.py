from jiit_tt_parser.utils.preprocessing import cache_tt_xls, cache_fac
from jiit_tt_parser.utils.cache import ensure_cache_folder


if __name__ == "__main__":
    ensure_cache_folder()
    cache_tt_xls()
    print()
    cache_fac()
