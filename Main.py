import Excel_work as ex_work
import Mailing as mailing

from art import *
from colorama import init, Fore

init(autoreset=True)


def main():

    art1 = text2art("REFACTOR")
    print(art1)

    menu_options = {
        1: ex_work.update_number,
        2: mailing.main_mailing,
        3: lambda: None,
        4: exit
    }

    while True:

        print('------------------------------------------')
        print('Выберите действие:'
              '\n1) Отформатировать номера'
              '\n2) Сделать рассылку - Работает не до конца'
              '\n3) Показать отправляемое сообщение - Не работает'
              '\n4) Выход')
        print('------------------------------------------')

        choice = input('>> ')

        if choice.isdigit():
            func = menu_options.get(int(choice))
            if func:
                func()
            else:
                print(f'{Fore.RED}Вы ввели недопустимое значение. Попробуйте еще раз.')
        else:
            print(f'{Fore.RED}Вы ввели недопустимое значение. Попробуйте еще раз.')


if __name__ == "__main__":
    main()
