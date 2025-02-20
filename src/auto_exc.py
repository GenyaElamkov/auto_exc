import os
import csv
import time
import gc
import warnings
import pandas as pd
from colorama import init, Fore

from contextlib import contextmanager
from multiprocessing import Pool, freeze_support

from openpyxl import load_workbook


@contextmanager
def timer():
    """Измеряет время работы скрипта"""
    start_time = time.perf_counter()
    yield
    end_time = time.perf_counter()
    execution_time = end_time - start_time
    print(f"Скрипт выполнялся {execution_time:.4f} секунд")
    # Пауза, чтобы консоль не закрывалась
    input("\nНажмите Enter для выхода...")


class Book:

    def read_book(self, filename: str) -> list[dict[str]]:
        # Убираем предупреждение в консоле
        with warnings.catch_warnings():
            warnings.simplefilter("ignore")

            wb = load_workbook(filename=filename, read_only=True)
            # read_only — режим для отложенной загрузки, экономить оперативную память
            sheet_ranges = wb.active

        organization = sheet_ranges["B25"].value
        order = "".join(
            str(cell.value) for coll in sheet_ranges["C32":"V32"] for cell in coll
        )

        book_csv = []
        id_cell = 43  # Счетчик для row обработки таблицы
        while True:
            id_cell += 1
            id_coll_start = f"A{id_cell}"

            if sheet_ranges[id_coll_start].value is None:
                gc.collect()  # Чистим мусор
                wb.close()  # Закрываем книгу
                break

            row = {
                "№": sheet_ranges[f"A{id_cell}"].value,
                "дата_старт": sheet_ranges[f"B{id_cell}"].value,
                "дата_end": sheet_ranges[f"E{id_cell}"].value,
                "вид": sheet_ranges[f"H{id_cell}"].value,
                "номер": sheet_ranges[f"M{id_cell}"].value,
                "номер_кор": sheet_ranges[f"P{id_cell}"].value,
                "наименование": sheet_ranges[f"S{id_cell}"].value,
                "бик": sheet_ranges[f"V{id_cell}"].value,
                "фио": sheet_ranges[f"Y{id_cell}"].value,
                "инн": sheet_ranges[f"AB{id_cell}"].value,
                "кпп": sheet_ranges[f"AE{id_cell}"].value,
                "номер_счета": sheet_ranges[f"AH{id_cell}"].value,
                "дебет": sheet_ranges[f"AK{id_cell}"].value,
                "кредит": sheet_ranges[f"AN{id_cell}"].value,
                "назначение": sheet_ranges[f"AQ{id_cell}"].value,
                "ордер": order,
                "организация": organization,
            }
            book_csv.append(row)

        return book_csv

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
        df = pd.read_csv(file_path, encoding="utf-8")  # Чтение CSV файла
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
