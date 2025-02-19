import os
import sys

# Добавляем корневую директорию в sys.path
sys.path.insert(0, os.path.abspath(os.path.join(os.path.dirname(__file__), "..")))

from src.auto_exc import _find_files


# Тест для функций работы с файлами
def test_find_files(tmp_path):
    (tmp_path / "test1.xlsx").touch()
    (tmp_path / "test2.xlsx").touch()
    (tmp_path / "test.txt").touch()

    found = _find_files(str(tmp_path), ".xlsx")
    assert len(found) == 2
    assert all(f.endswith(".xlsx") for f in found)
