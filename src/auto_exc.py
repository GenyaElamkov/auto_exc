import os
import csv
import time
from typing import Generator, List, Dict, Any

import warnings
import pandas as pd
from colorama import init, Fore

from contextlib import contextmanager
from multiprocessing import Pool, freeze_support


@contextmanager
def timer() -> Generator[None, Any, None]:
    """Измеряет время работы скрипта"""
    start_time = time.perf_counter()
    yield
    end_time = time.perf_counter()
    execution_time = end_time - start_time
    print(f"Скрипт выполнялся {execution_time:.4f} секунд")
    # Пауза, чтобы консоль не закрывалась
    input("\nНажмите Enter для выхода...")


class Book:

    def read_book(self, filename: str) -> List[Dict[str, Any]]:
        """
        Читает данные из Excel-файла и возвращает список словарей с данными.

        :param filename: Путь к Excel-файлу.
        :return: Список словарей с данными из файла.
        """
        
        
        try:
            with warnings.catch_warnings():
                warnings.simplefilter("ignore")
                
                # Чтение Excel-файла с помощью pandas
                df = pd.read_excel(
                    filename,
                    sheet_name=0,
                    header=None,
                    usecols="A:AR",  # Читаем только нужные колонки (A до AR)
                )

            # Получаем значения из конкретных ячеек (например, организация и ордер)
            organization = df.iloc[24, 1]  # B25
            order = "".join(str(int(cell)) for cell in df.iloc[31, 2:22]) if len(df) > 31 else ""  # C32:V32
            # Обрабатываем строки данных
            book_csv = []
            for i in range(43, len(df)):
                row_data = df.iloc[i]  # Получаем строку данных
                
                # Проверяем, есть ли данные в первой колонке (A)
                # Если первая колонка пустая, завершаем обработку
                if pd.isna(row_data[0]):
                    break

                # Формируем словарь с данными
                row = {
                    "№": row_data[0],  # A
                    "дата_старт": row_data[1],  # B
                    "дата_end": row_data[4],  # E
                    "вид": row_data[7],  # H
                    "номер": row_data[9],  # J
                    "дата": row_data[12],  # M
                    "номер_кор": row_data[15],  # P
                    "наименование": row_data[18],  # S
                    "бик": row_data[21],  # V
                    "фио": row_data[24],  # Y
                    "инн": row_data[27],  # AB
                    "кпп": row_data[30],  # AE
                    "номер_счета": row_data[33],  # AH
                    "дебет": row_data[36],  # AK
                    "кредит": row_data[39],  # AN
                    "назначение": row_data[42],  # AQ
                    "ордер": order,
                    "организация": organization,
                }

                book_csv.append(row)

            return book_csv

        except Exception as e:
            print(f"Ошибка при чтении файла {filename}: {e}")
            return []

    def save_book_csv(self, data: list[dict[str]], file_name: str) -> None:
        """Сохраняет данные в файл с раширением .csv"""
        with open(f"{file_name}.csv", "w", newline="", encoding="utf-8") as csv_file:
            writer = csv.DictWriter(csv_file, fieldnames=data[0].keys())
            writer.writeheader()
            writer.writerows(data)


def _find_files(directory: str, extension: str) -> list:
    """Формирует список path где лежат файлы excel"""
    return [
        os.path.join(directory, file)
        for file in os.listdir(directory)
        if file.endswith(extension)
    ]


def create_directory(name: str) -> None:
    """Создает директорию"""
    os.makedirs(name, exist_ok=True)


def read_csv_files(directory: str) -> pd.DataFrame:
    """Читает все CSV файлы в директории и возвращает объединенный DataFrame."""
    combined_data = pd.DataFrame()  # Пустой DataFrame для объединенных данных
    for file_path in _find_files(directory, extension=".csv"):
        df = pd.read_csv(file_path, encoding="utf-8")
        combined_data = pd.concat(
            [combined_data, df], ignore_index=True
        )  # Объединение данных
    return combined_data


def save_to_xlsx(data: pd.DataFrame, output_filename: str) -> None:
    """Сохраняет DataFrame в XLSX файл."""
    with pd.ExcelWriter(output_filename, engine="openpyxl") as writer:
        data.to_excel(writer, index=False, sheet_name="Combined Data")
    print(f"Данные сохранены в файл: {output_filename}")


def merge_csv_to_xlsx(input_directory: str, output_filename: str) -> None:
    """Объединяет все CSV файлы в директории в один XLSX файл."""
    combined_data = read_csv_files(input_directory)
    save_to_xlsx(combined_data, output_filename)  # Сохранение в XLSX


def clear_csv_files(directory: str) -> None:
    """
    Удаляет все CSV файлы в указанной директории.
    :param directory: Путь к директории, в которой нужно очистить CSV файлы.
    """
    if not os.path.exists(directory):
        print(f"Директория {directory} не существует.")
        return

    # Фильтруем только CSV файлы
    files_path = _find_files(directory, extension=".csv")
    if not files_path:
        print(f"В директории {directory} нет CSV файлов для удаления.")
        return

    # Удаляем каждый CSV файл
    for file_path in files_path:
        try:
            os.remove(file_path)
            print(f"Файл {file_path} успешно удален.")
        except Exception as e:
            print(f"Ошибка при удалении файла {file_path}: {e}")


def worker(path: str, name_directory: str) -> str | Exception:
    """
    Обрабатывает файл по указанному пути, сохраняет его данные в формате CSV
    в заданную директорию и удаляет исходный файл (при условии его нахождения
    в текущей рабочей директории
    """
    try:
        bk = Book()
        data = bk.read_book(path)
        path_directory = os.path.join(name_directory, os.path.basename(path))
        bk.save_book_csv(data, file_name=path_directory)

        # Удаляем отработанный файл
        file_name = os.path.basename(path)
        if file_name in os.listdir(os.getcwd()):
            os.remove(file_name)

        return f"Файл создан: {path_directory}"
    except Exception as e:
        return f"Ошибка при обработке файла {path}: {e}"


def processing(name_directory: str) -> None:
    """Общая функция для обработки данных"""
    create_directory(name_directory)
    file_list = _find_files(os.getcwd(), extension=".xlsx")

    print(f"Всего файлов в директории {len(file_list)}. Идет обработка файлов...")
    total_files = len(file_list)
    processed_files = 1

    with Pool(os.cpu_count()) as pool:
        results = []
        results = pool.starmap(worker, [(path, name_directory) for path in file_list])

        for result in results:
            print(f"Обработано файлов: {processed_files}/{total_files}")
            processed_files += 1
            print(result)


def single_file_connection(name_directory: str) -> None:
    """Общая функция для объединения файлов в один xlsx"""
    output_filename = os.path.join(name_directory, "combined_data.xlsx")
    merge_csv_to_xlsx(name_directory, output_filename)

    print("Следующий этап — Очистка файлов")
    clear_csv_files(name_directory)


def main() -> None:
    init()
    print(
        Fore.RED
        + """
        [!] ОБЯЗАТЕЛЬНО СДЕЛАЙТЕ КОПИИ ФАЙЛОВ. 
        [!] ОБРАБОТАННЫЕ ФАЙЛЫ БУДУТ УДАЛЕНЫ
        """
    )
    print(
        Fore.RESET
        + """
    Выберите:
    1. Автоматический режим
    
    Ручной режим:
    2. Обработка файлов
    3. Объединение данных в один файл
    """
    )

    name_directory = "00_Data"
    while True:
        choice = input("\nВыберите цифру (нажмите Enter для подтверждения):")
        if choice == "1":
            processing(name_directory)
            single_file_connection(name_directory)
            print("\n[!] УСПЕШНО. НАСЛАЖДАЙТЕСЬ РЕЗУЛЬТАТОМ")
        elif choice == "2":
            processing(name_directory)
            print(
                "Обработка файлов завершена успешна. Следующий этап — Объединение данных"
            )
        elif choice == "3":
            single_file_connection(name_directory)
        else:
            break


if __name__ == "__main__":
    with timer():
        freeze_support()
        main()
