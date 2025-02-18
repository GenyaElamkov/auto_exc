import os
import sys

# Добавляем корневую директорию в sys.path
sys.path.insert(0, os.path.abspath(os.path.join(os.path.dirname(__file__), "..")))
from project.auto_exc import Book


def test_cuts_numbers():
    book = Book()
    assert book._cuts_numbers("A1") == "A"
    assert book._cuts_numbers("B23") == "B"
    assert book._cuts_numbers("Z9") == "Z"
    assert book._cuts_numbers("AA123") == "AA"
