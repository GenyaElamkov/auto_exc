import os
from click import pause
import gs
from tqdm import tqdm

from openpyxl import Workbook, load_workbook


def find_xlsx_files(directory: str) -> list:
    """Формирует список path где лежат файлы excel"""
    return [
        f"{directory}\{file}"
        for file in os.listdir(directory)
        if file.endswith(".xlsx")
    ]


class Book:
    def __init__(self) -> None:
        self.counter_row = 1  # Счетчик для row

    def _cuts_numbers(self, text: str):
        """Обрезает цифры, оставляет буквы для Ячеек"""
        return "".join([char for char in text if not char.isdigit()])

    def set_book(self, filename: str) -> dict[str]:
        wb = load_workbook(filename=filename)
        sheet_ranges = wb.active
        order = []
        for coll in sheet_ranges["C32":"V32"]:
            for cell in coll:
                if not cell.value:
                    continue
                order.append(cell.value)
        order = "".join(order)

        id_cell = 43
        counter_col = 0  # Счетчик для col
        book = {}

        while True:
            id_cell += 1
            id_coll_start = f"A{id_cell}"
            id_coll_end = f"AV{id_cell}"

            if sheet_ranges[id_coll_start].value is not None:
                for coll in sheet_ranges[id_coll_start:id_coll_end]:
                    for cell in coll:
                        coordinate = self._cuts_numbers(cell.coordinate)
                        book.setdefault(f"{coordinate}{self.counter_row}", cell.value)
                        counter_col += 1
                    book.setdefault(f"AW{self.counter_row}", order)

                    counter_col = 0
                self.counter_row += 1
            else:
                wb.close()
                break
        return book

    def create_new_book(self, data: dict[str]) -> None:
        wb_new = Workbook()
        sheet = wb_new.active
        sheet.title = "Сводный отчет"
        for k, v in data.items():
            sheet[k] = v

        wb_new.save("report.xlsx")


def main():
    data = {}

    bk = Book()
    # file_list = find_xlsx_files(os.getcwd())
    file_list = find_xlsx_files(r'C:\Projects\auto_exc\data')
    for path in tqdm(file_list):
        data.update(bk.set_book(path))
    bk.create_new_book(data)


if __name__ == "__main__":
    try:
        main()
    except:
        pause(10)