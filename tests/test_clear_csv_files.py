import os
import sys

from tests.fixtures import tmp_csv_files

# Добавляем корневую директорию в sys.path
sys.path.insert(0, os.path.abspath(os.path.join(os.path.dirname(__file__), "..")))
from src.auto_exc import clear_csv_files


# Тесты для очистки CSV
def test_clear_csv_files(tmp_csv_files):
    clear_csv_files(str(tmp_csv_files))
    assert len(list(tmp_csv_files.glob("*.csv"))) == 0
