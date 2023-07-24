import json
import time
import requests
import Formatting_numbers
import Support_function

from Excel_work import excel_table
from colorama import init, Fore
from docxtpl import DocxTemplate
from GreenAPI import gr_api

init(autoreset=True)


def main_mailing():
    """
    Главная функция рассылки, которая запускает вспомогательные функции.
    Проверяет каждый номер на его регистрацию в WhatsApp. Неудачные номера или клиентов записывает на новый лист
    """

    #if not gr_api.check_authorization():
        #return

    Formatting_numbers.update_number()  # Запускаем форматирование номеров
    excel_table.create_new_sheet()  # Создаём\очищаем лист с результатами

    max_row = excel_table.get_sheet_row()  # Длина строк
    columns = excel_table.select_column()  # Колонки с номерами

    context_col = excel_table.get_dict_keys_columns()  # Получаем ключи из docx

    doc = DocxTemplate("content.docx")  # Открываем шаблон

    num_row = 0  # Счётчик строки
    path_to_word = 'content.docx'

    for row in excel_table.sheet_main.iter_rows(min_row=2, max_row=max_row, values_only=True):  # Цикл в длину строк

        num_row += 1
        list_failed_numbers = []  # Список для проверки отправки сообщений
        list_full = []  # Список с номерами из всех колонок

        for col in columns:  # Формируем список всех номеров по всем выбранным колонкам
            if row[col] is None:
                continue

            list_numbers = str(row[col]).split(';')
            list_full += list_numbers

        if not list_full:  # Если список пустой, пишем на доп лист и пропускаем итерацию
            excel_table.write_additional_sheet(row=row, num_row=num_row, status=False)
            print(f'{num_row}: Отсутствуют номера в таблице!')
            continue

        if context_col:  # Если ключи в docx есть то формируем словарь с выбранными значениями и пишем в docx_2
            context = {i: row[context_col[i]] for i in context_col}
            doc.render(context)
            doc.save('content_2.docx')
            path_to_word = 'content_2.docx'

        list_full = list(set(list_full))

        for number in list_full:  # Проходимся по каждому номеру и отправляем сообщение, результат false\true пишем в список
            if send_message(number, path_to_word):
                time.sleep(5)
                list_failed_numbers.append(True)
            else:
                list_failed_numbers.append(False)

        if all(x is False for x in list_failed_numbers):  # Если по всем номерам не получилось отправить - пишем на доп лист
            excel_table.write_additional_sheet(row=row, num_row=num_row, status=False)
            print(f'{num_row}: Не получилось отправить сообщение!')
            continue

        excel_table.write_additional_sheet(row=row, num_row=num_row)  # Тоже самое только пишем успех
        print(f'{num_row}: Было отправлено сообщение!')

    print('\nРассылка закончена!\n')


def check_number(phone_number: str) -> bool:
    """
    Функция проверяет Зарегистрирован ли номер в WhatsApp или нет

    :param phone_number: Проверяемый номер телефона

    :return: Возвращает True или False в зависимости от результата. Если номер зарегистрирован, то True,
     в обратном случае False
    """

    url = f"https://api.green-api.com/waInstance{gr_api.id_instance}/checkWhatsapp/{gr_api.api_token}"

    payload = json.dumps({"phoneNumber": phone_number})
    headers = {
        'Content-Type': 'application/json'
    }

    response = requests.request("POST", url, headers=headers, data=payload)


    if response.status_code == 466:
        print(f'{Fore.RED}\nВы исчерпали лимит проверок!\n')
        return False


    if response.status_code == 200 and response.json()['existsWhatsapp'] is False:
            return False

    return True


def send_message(phone_number: str, path_to_word: str) -> bool:
    """
    Функция рассылки сообщения

    :param path_to_word: Путь к файлу, откуда брать данные
    :param phone_number: Номер телефона
    :return: Ничего не возвращает
    """

    if not check_number(phone_number=phone_number):
        return False

    text = Support_function.get_text_from_word(path_to_word)

    gr_api.green_api.sending.sendMessage(f'{phone_number}@c.us', f'{text}')

    return True
