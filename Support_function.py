import time
import configparser
import re
import os
from docx import Document


class Config:
    """
    Класс для работы с конфигом
    """
    config = configparser.ConfigParser()
    path_config = 'config.ini'
    config.read(path_config)

    @classmethod
    def create_config(cls):
        """
        Создание файла config.ini
        """
        if not os.path.isfile(cls.path_config):

            data_dict = {
                'id_instance': '',
                'api_token_instance': ''
            }

            cls.write_to_config(Config, title='GreenAPI', data_dict=data_dict)

            print(f'Файл {cls.path_config} создан')

            time.sleep(1)

    @staticmethod
    def write_to_config(cls, title: str, data_dict: dict):
        """
        Класс для записи данных в Config
        :param cls: Класс Config
        :param title: Заголовок секции в файле config
        :param data_dict: Словарь с записываемыми значениями
        """
        cls.config[title] = data_dict

        with open('config.ini', 'w') as configfile:
            cls.config.write(configfile)

        print('Данные записаны!')

    @staticmethod
    def get_data(cls, title: str) -> tuple:
        """
        Получение данных из конфига
        :param cls: Класс Config
        :param title: Заголовок секции откуда нужно получить данные
        :return: options: Кортеж значений секции
        """
        if not os.path.isfile(cls.path_config):
            cls.create_config()

        options = cls.config.items(title)

        return options


def get_text_from_word(path_to_word: str) -> str:
    """
    Получение текста из файла
    :param path_to_word: Путь к файлу откуда брать значения
    :return: text: Текст из файла
    """
    doc = Document(path_to_word)

    content = []

    for i in doc.paragraphs:
        content.append(i.text)

    text = ('\n'.join(content))

    return text


def print_text_from_word():
    """
    Печать текста из файла
    """
    print(f"""
ОТПРАВЛЯЕМОЕ СООБЩЕНИЕ
{"*"*20}
{get_text_from_word(path_to_word='content.docx')}
{"*"*20}
        """)
    time.sleep(3)


def get_keys_from_word() -> list:
    """
    Получение всех ключей типа {{}} из файла
    :return: keys: Список ключей
    """
    text = get_text_from_word(path_to_word='content.docx')
    keys = re.findall(r"\{\{(\w+)\}\}", text)
    return keys


def get_image_from_word():
    pass