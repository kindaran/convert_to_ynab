import pytest

import convert_to_YNAB_csv

def test_filelist():
    assert len(convert_to_YNAB_csv.getFileList(".")) > 1

    