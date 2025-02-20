import os
import sys

# Добавляем корневую директорию в sys.path
sys.path.insert(0, os.path.abspath(os.path.join(os.path.dirname(__file__), "..")))

from src.auto_exc import create_directory


# Тест для функций работы с файлами
def test_create_directory(tmp_path):
    new_dir = tmp_path / "new_dir"
    create_directory(str(new_dir))
    assert new_dir.exists()

    # Проверка существующей директории
    create_directory(str(new_dir))  # Не должно вызывать ошибок
