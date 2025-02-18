import os
import sys
import tempfile

# Добавляем корневую директорию в sys.path
sys.path.insert(0, os.path.abspath(os.path.join(os.path.dirname(__file__), "..")))

from project.auto_exc import find_xlsx_files


def test_find_xlsx_files():
    # Создаем временную директорию с тестовыми файлами
    with tempfile.TemporaryDirectory() as temp_dir:
        # Создаем несколько файлов
        open(os.path.join(temp_dir, "test1.xlsx"), "w").close()
        open(os.path.join(temp_dir, "test2.xlsx"), "w").close()
        open(os.path.join(temp_dir, "test3.txt"), "w").close()

        # Проверяем, что функция находит только .xlsx файлы
        result = find_xlsx_files(temp_dir)
        assert len(result) == 2
        assert all(file.endswith(".xlsx") for file in result)
