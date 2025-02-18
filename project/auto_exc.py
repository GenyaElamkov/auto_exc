import os
import gc
import warnings

from multiprocessing import Pool, freeze_support

from openpyxl import Workbook, load_workbook


def find_xlsx_files(directory: str) -> list:
    """Формирует список path где лежат файлы excel"""
    return [
        os.path.join(directory, file)
        for file in os.listdir(directory)
        if file.endswith(".xlsx")
    ]


def create_directory(name: str) -> None:
    os.makedirs(name, exist_ok=True)


class Book:
    def __init__(self) -> None:
        self.counter_row = 1  # Счетчик для row

    def _cuts_numbers(self, text: str) -> str:
        """Обрезает цифры, оставляет буквы для Ячеек"""
        return "".join([char for char in text if not char.isdigit()])

    def reed_book(self, filename: str) -> dict[str]:
        # Убираем предупреждение
        with warnings.catch_warnings():
            warnings.simplefilter("ignore")

            wb = load_workbook(filename=filename, read_only=True)
            # read_only — режим для отложенной загрузки, экономить оперативную память

            sheet_ranges = wb.active

        book = {}
        organization = sheet_ranges["B25"].value

        order = []
        for coll in sheet_ranges["C32":"V32"]:
            for cell in coll:
                order.append(cell.value)
        order = "".join(order)

        id_cell = 43  # Счетчик для row обработки таблицы
        while True:
            id_cell += 1
            id_coll_start = f"A{id_cell}"
            id_coll_end = f"AV{id_cell}"

            if sheet_ranges[id_coll_start].value is not None:
                for coll in sheet_ranges[id_coll_start:id_coll_end]:
                    for cell in coll:
                        if cell.value is not None:
                            coordinate = self._cuts_numbers(cell.coordinate)
                            book[f"{coordinate}{self.counter_row}"] = cell.value
                    book[f"AW{self.counter_row}"] = order
                    book[f"AX{self.counter_row}"] = organization
                self.counter_row += 1
            else:
                gc.collect()  # Чистим мусор
                wb.close()  # Закрываем книгу
                break
        return book

    def create_new_book(self, data: dict[str], file_name: str) -> None:
        wb_new = Workbook()

        sheet = wb_new.active
        sheet.title = "Сводный отчет"
        for k, v in data.items():
            sheet[k] = v
        wb_new.save(file_name)
        wb_new.close()


def worker(path: str, name_directory: str) -> str | None:
    try:
        bk = Book()
        data = bk.reed_book(path)
        path_directory = os.path.join(name_directory, os.path.basename(path))
        bk.create_new_book(data, file_name=path_directory)
        return f"Файл создан: {path_directory}"
    except Exception as e:
        return f"Ошибка при обработке файла {path}: {e}"


def main() -> None:
    name_directory = "00_Data"
    create_directory(name_directory)
    file_list = find_xlsx_files(os.getcwd())
    # file_list = find_xlsx_files(r"C:\Projects\auto_exc\example_files")

    print(f"Всего файлов в директории {len(file_list)}. Идет обработка файлов:")
    total_files = len(file_list)
    processed_files = 1
    with Pool(os.cpu_count()) as pool:
        results = []
        results = pool.starmap(worker, [(path, name_directory) for path in file_list])

        for result in results:
            print(f"Обработано файлов: {processed_files}/{total_files}")
            processed_files += 1
            print(result)

    # Пауза, чтобы консоль не закрывалась
    input("\nНажмите Enter для выхода...")


if __name__ == "__main__":
    freeze_support()
    main()
