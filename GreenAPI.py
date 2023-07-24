import requests
import json
from colorama import Fore
from whatsapp_api_client_python import API
from Support_function import Config


class GreenAPI:
    """
    Класс для работы с подключением к GreenAPI
    """
    def __init__(self):

        data_id_api = Config.get_data(Config, title='GreenAPI')

        self.__id_instance = data_id_api[0][1]
        self.__api_token = data_id_api[1][1]

        if not (self.__id_instance and self.__api_token):
            self.write_green_api_data()

        self.__green_api = API.GreenApi(self.__id_instance, self.__api_token)

    def __str__(self):
        return f"""
        Текущие данные
        Id_instance: {self.__id_instance}
        Api_token: {self.__api_token}"""

    def check_authorization(self) -> bool:
        """
        Проверка введенных данных на наличие авторизации на сайте GreenAPI
        """
        url = f"https://api.green-api.com/waInstance{self.__id_instance}/getStateInstance/{self.__api_token}"

        payload = {}
        headers = {}

        response = requests.request("GET", url, headers=headers, data=payload)

        if response.status_code == 401:
            print(f'{Fore.RED}Пользователь не найден\n')
            return False

        if response.status_code == 403:
            print(f'{Fore.RED}\nКод 403\n')

        if response.status_code == 200:
            result = json.loads(response.content)

            if result['stateInstance'] == 'notAuthorized':
                print(f'{Fore.RED}\nПользователь найден но не авторизован\n')
                return False

            if result['stateInstance'] == 'starting':
                print(f'{Fore.RED}\nПользователь в статусе starting\nПопробуйте позже или поменяйте Instance\n')
                return False

        print(f'{Fore.GREEN}\nАвторизация пройдена\n')
        return True

    def write_green_api_data(self):
        """
        Ввод данных для сайта GreenAPI и их последующая запись в конфиг
        """
        while True:

            while True:
                id_instance_inp = input('Введите IdInstance, полученный в GreenAPI >> ')
                if len(id_instance_inp) != 10 or not id_instance_inp.isdigit():
                    print('Id instance должен состоять из 10 цифр ')
                    continue
                break

            while True:
                api_token_inp = input('Введите ApiTokenInstance, полученный в GreenAPI >> ')
                if len(api_token_inp) != 50:
                    print('Api token должен состоять из 50 символов')
                    continue
                break

            self.__id_instance = id_instance_inp
            self.__api_token = api_token_inp

            if not self.check_authorization():
                continue

            self.__green_api = API.GreenApi(self.__id_instance, self.__api_token)

            data_dict = {
                'id_instance': self.__id_instance,
                'api_token_instance': self.__api_token
            }

            Config.write_to_config(Config, title='GreenAPI', data_dict=data_dict)

            break

    @property
    def id_instance(self):
        """
        Возврат id_instance
        """
        return self.__id_instance

    @property
    def api_token(self):
        """
        Возврат Api_token
        """
        return self.__api_token

    @property
    def green_api(self):
        """
        Возврат объекта green_api
        """
        return self.__green_api


gr_api = GreenAPI()
