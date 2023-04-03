import openpyxl
import re
import time

from colorama import init, Fore
from openpyxl.styles import PatternFill

init(autoreset=True)

try:
    wb = openpyxl.load_workbook(filename='Table.xlsx')
except Exception:
    print('Не удалось открыть файл!')
    time.sleep(5)
    exit()

sheet = wb['Start']


def get_sheet_row() -> int:

    """
    Функция возвращает длину непустых строк таблицы

    :return sheet_row: Возвращает число непустых строк
    """

    sheet_row = sheet.max_row

    return sheet_row


def replace_string(string: str) -> str:

    """
    Функция форматирует строку (Конкретно здесь - номер). Удаляет лишние символы, добавляет отсутствующие куски номера.

    :param string: Строка для редактирования
    :return: Возвращает переформатированную строку
    """

    # Удаление лишних символов
    string = re.sub(r',', ';', string)
    string = re.sub(r'[^\d;]', '', string)

    list_number = string.split(';')

    for i, phone in enumerate(list_number):

        result = re.match(r'\b\d{6}\b', phone)

        if result:
            list_number[i] = '73952' + phone

        result = re.search(r'^8', phone)
        if result:
            list_number[i] = re.sub(r'^8', '7', phone)

        result = re.search(r'^3', phone)
        if result:
            list_number[i] = '7' + phone

        list_number[i] = re.sub(r'\b\d{1,10}\b', '', list_number[i])

    string = (";".join(list_number))

    string = re.sub(r'^;', '' , string)

    return string


def update_number():

    """
    Функция запускает процесс форматирования номера. Проходит по строкам с номерами и форматирует их. За форматирование
    отвечает функция replace_string(). В конечном итоге сохраняет файл

    :return: Функция ничего не возвращает :/
    """

    columns = []
    column_lenght = get_sheet_row()

    print('Введите колонку с номерами\n'
          'Колонку нужно выбрать по её порядковому номеру\n'
          '** A-1, B-2, C-3, D-4, E-5, F-6, G-7, H-8 **')

    while True:
        column = input('>> ').lower()
        if column == 'далее':
            break
        else:
            columns.append(int(column))
            print(f'\nТекущие колонки: {columns}')
            print(f'{Fore.GREEN}Введите "Далее" если готовы продолжить\n')

    for row in range(2, column_lenght):
        for column in columns:
            cell = sheet.cell(row=row, column=column)
            if cell.value is not None:
                val = str(cell.value)
                val = replace_string(val)
                cell.value = val
            else:
                cell.fill = PatternFill(fill_type='solid', start_color='999999')

    while True:
        try:
            save_excel()
            break
        except:
            print(f'\n{Fore.RED}Возникла ошибка сохранения. Возможно вы не закрыли файл с номерами.\n')
            time.sleep(1)
            for i in range(5, 0, -1):
                print(f'Попытка повторного сохранения через {i} сек.')
                time.sleep(1)
            print('\nПопытка сохранить файл......\n')
            time.sleep(2)


def save_excel():
    wb.save("Table.xlsx")
    print(Fore.GREEN + '\nНомера отформатированы!\n\nВсе пустые ячейки выделены цветом!\n')
