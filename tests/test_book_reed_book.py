import os
import sys

from tests.fixtures import tmp_excel_file

# Добавляем корневую директорию в sys.path
sys.path.insert(0, os.path.abspath(os.path.join(os.path.dirname(__file__), "..")))
from src.auto_exc import Book


# Тест для класса Book
def test_book_reed_book(tmp_excel_file):
    book = Book()
    data = book.reed_book(str(tmp_excel_file))

    assert len(data) == 1
    assert data[0]["организация"] == "Test Organization"
    assert data[0]["ордер"] == "11111111111111111111"
    assert data[0]["№"] == 1
    assert data[0]["дебет"] == "100.50"
