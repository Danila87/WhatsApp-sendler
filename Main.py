import os.path
import time
import Formatting_numbers
import Mailing
import Support_function
import sys
from Support_function import Config
from GreenAPI import gr_api
from docx import Document
from Excel_work import excel_table
from colorama import init, Fore

init(autoreset=True)


class App:
    """
    Класс приложения
    """
    def __init__(self):

        self.create_files()
        self.main()

    @staticmethod
    def create_files():
        """
        Создание файлов
        """
        Config.create_config()

        if not os.path.isfile('content.docx'):
            document = Document()
            document.save('content.docx')
            print('Файл content.docx создан')
            time.sleep(1)

        if not os.path.isfile('content_2.docx'):
            document = Document()
            document.save('content_2.docx')
            print('Файл content_2.docx создан')
            time.sleep(1)

    @staticmethod
    def main():
        """
        Главная функция
        """
        menu_options = {
            1: Formatting_numbers.update_number,
            2: Mailing.main_mailing,
            3: Support_function.print_text_from_word,
            4: gr_api.write_green_api_data,
            5: excel_table.select_column,
            6: excel_table.select_excel,
            7: sys.exit
        }

        while True:

            print('------------------------------------------')
            print('Выберите действие:'
                  '\n1) Отформатировать номера'
                  '\n2) Сделать рассылку'
                  '\n3) Показать отправляемое сообщение'
                  '\n4) Поменять данные GreenAPI'
                  '\n5) Поменять колонки с номерами'
                  '\n6) Выбрать другую таблицу'
                  '\n7) Выход')
            print('------------------------------------------')

            print(f'ТЕХНИЧЕСКАЯ ИНФОРМАЦИЯ\n{excel_table}')

            choice = input('>> ')

            if choice.isdigit():
                func = menu_options.get(int(choice))
                if func:
                    if func == excel_table.select_column:
                        func(clear=True)
                    func()
                else:
                    print(f'{Fore.RED}Вы ввели недопустимое значение. Попробуйте еще раз.')
            else:
                print(f'{Fore.RED}Вы ввели недопустимое значение. Попробуйте еще раз.')


main_app = App()
