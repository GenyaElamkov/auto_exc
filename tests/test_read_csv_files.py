import os
import sys
import pytest
import csv


from tests.fixtures import tmp_csv_files

# Добавляем корневую директорию в sys.path
sys.path.insert(0, os.path.abspath(os.path.join(os.path.dirname(__file__), "..")))
from project.auto_exc import read_csv_files


# Тест для объединения CSV
def test_read_csv_files(tmp_csv_files):
    df = read_csv_files(str(tmp_csv_files))
    assert len(df) == 3
    assert list(df["№"]) == [1, 2, 3]
