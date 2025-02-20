import os
import sys

from tests.fixtures import tmp_excel_file

# Добавляем корневую директорию в sys.path
sys.path.insert(0, os.path.abspath(os.path.join(os.path.dirname(__file__), "..")))
from src.auto_exc import worker


# Тест для worker
def test_worker(tmp_excel_file, tmp_path):
    output_dir = tmp_path / "output"
    output_dir.mkdir()

    result = worker(str(tmp_excel_file), str(output_dir))

    csv_file = output_dir / "test.xlsx.csv"
    assert csv_file.exists()
    assert "Файл создан" in result


# Тест обработки ошибок в worker
def test_worker_error():
    result = worker("invalid_path.xlsx", "invalid_dir")
    assert "Ошибка" in result
