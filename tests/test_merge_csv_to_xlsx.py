import os
import sys
import pytest
import csv

import pandas as pd

from tests.fixtures import tmp_csv_files

# Добавляем корневую директорию в sys.path
sys.path.insert(0, os.path.abspath(os.path.join(os.path.dirname(__file__), "..")))
from project.auto_exc import merge_csv_to_xlsx


# Тест для объединения CSV
def test_merge_csv_to_xlsx(tmp_csv_files, tmp_path):
    output_path = tmp_path / "combined.xlsx"
    merge_csv_to_xlsx(str(tmp_csv_files), str(output_path))

    assert output_path.exists()
    df = pd.read_excel(output_path)
    assert len(df) == 3
