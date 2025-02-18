import os
import sys
import tempfile

# Добавляем корневую директорию в sys.path
sys.path.insert(0, os.path.abspath(os.path.join(os.path.dirname(__file__), "..")))

from project.auto_exc import Book


def test_create_new_book():
    # Тестовые данные
    data = {
        "A1": "Test1",
        "B1": "Test2",
        "AW1": "Order1",
        "AX1": "Org1",
    }

    # Создаем временный файл для сохранения
    with tempfile.NamedTemporaryFile(suffix=".xlsx", delete=False) as tmp_file:
        tmp_file.close()

        # Создаем книгу и сохраняем данные
        book = Book()
        book.create_new_book(data, tmp_file.name)

        # Проверяем, что файл создан
        assert os.path.exists(tmp_file.name)

        # Удаляем временный файл
        os.unlink(tmp_file.name)
