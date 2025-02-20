import os
import sys
import csv

# Добавляем корневую директорию в sys.path
sys.path.insert(0, os.path.abspath(os.path.join(os.path.dirname(__file__), "..")))
from src.auto_exc import Book


# Тест для класса Book
def test_book_save_book_csv(tmp_path):
    book = Book()
    test_data = [{"№": 1, "дата_старт": "2023-01-01"}]
    output_path = tmp_path / "test_output.csv"

    book.save_book_csv(test_data, str(output_path.with_suffix("")))

    assert output_path.exists()
    with open(output_path, "r", encoding="utf-8") as f:
        reader = csv.DictReader(f)
        rows = list(reader)
        assert rows[0]["№"] == "1"
