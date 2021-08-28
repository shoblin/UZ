import sys

import openpyxl as xl
import shutil
import oracle
from datetime import datetime


# Функции для получения дат начала и конца периода
class FutureDate(Exception):
    """
    Обработчик исключения. Дата конца и начала периода,
    должны быть раньше сегодня
    """

    def __init__(self, text):
        self.txt = text


def check_validation_date(str_date: str):
    """
    Переводим str_date в дату, попутно проверяем ее:
        - Что бы дата была раньше сегодня
        - Что бы была в правильном формате
        - Что бы существовала
    Args:
        str_date (str): дата в виде строки, должна быть в формате "ДД-ММ-ГГГГ"
    Returns:
        True или False - в зависимости от того, смогли ли мы преобразовать дату
        date - преобразованую строку в datetime
    """
    try:
        date = datetime.strptime(str_date, '%d-%m-%Y')
        if date > datetime.now():
            raise FutureDate('Даты должны быть в прошлом')
        return True, date
    except ValueError:
        print('Вы ввели дату или не в формате ДД-ММ-ГГГГ.')
        print('Или не существующую дату')
    except FutureDate as mr:
        print(mr)
    print()
    return False, None


def request_date(txt):
    """
    Запрашиваем у пользователя дату в формате ДД-ММ-ГГГГ
    Args:
        txt (str): Текст выдаваемый запросом
             Правильность введеной даты проверяется фунцией
             get_date(str_date)
    Returns:
        date (datetime): Преобразованая дата
    """
    true_date = False
    date = None
    while not true_date:
        print(txt)
        str_date = input('В формате ДД-ММ-ГГГГ или q для выхода:')
        if str_date == 'q':
            sys.exit()
        true_date, date = check_validation_date(str_date)

    return date


def check_difference(first_date, second_date):
    """
    Проверяет не равные ли даты
    Args:
        first_date, second_date: Даты введенные пользователями
    Returns:
        True - если они равны, если это не так False
    """
    print('Две даты отчетного периода должны быь разными.')
    return first_date == second_date


def order_dates(first_date, second_date):
    """
    Сортируем даты, и возращаем их
    Args:
        first_date, second_date:  Даты введенные пользователями
    Returns:
        Возращает отсортированые даты
    """
    return sorted([first_date, second_date])


def new_file_name(template, to_date):
    """
    Create name for new file = region name + month
        region name - we get from template name
        month - number of month
    :param template: Source file - file with template
    :param to_date: date of end  of period
    :return:
        {region_name}_{month:02}_{year:04}.xlsx
    """
    region_name = template.split('_')[-1]
    return f'./{region_name}_{to_date.month:02}_{to_date.year:04}.xlsx'


def copy_template_file(template_file, new_file):
    """
    Копируем шаблон <template_file> в новый файл <new_file>
    Args:
        template_file (str): Имя одно из файлов шаблона, которые находятся в папке Templates
        new_file (str): Имя нового файла сформированого из шаблон
    Returns:
        True - Если копия создана, False если произошла ошибка
    """
    try:
        shutil.copy(template_file, new_file)
        return True

    except shutil.SameFileError:
        # Исключение на будущее, когда можно будет задавать имя нового файла через аргументы
        print('Имя Шаблона и нового файла совпадают')
    except PermissionError:
        # Исключение при недоступноти
        print("У вас нет прав.")
    except Exception:
        # For other errors
        print("Error occurred while copying file.")
    return False


def open_xlsx(file_name):
    wb = xl.load_workbook(file_name)
    ws = wb.worksheets[0]
    return wb, ws


def get_pids(template_file, pid_column=1):
    """
    Получаем points_id с ексел файла. Проверяя колонку под номером <pid_column>
    :param template_file: файл шаблона
    :param pid_column: Колонка с points_id. Default value = 1
    :return pids: diction {point id: номер колонки}
    """
    wb, ws = open_xlsx(template_file)

    # calculate total number of rows and
    # columns in source excel file
    mr, mc = ws.max_row, ws.max_column

    pids = dict()
    for num_row in range(1, mr + 1):
        pid = ws.cell(row=num_row, column=pid_column).value
        if pid:
            pids[pid] = num_row
    wb.close()

    return pids


def fill_xlsx(new_file, pid_row, date_since, date_to, column=1):
    """
    Put date into xlsx file
    :param new_file:
    :param pid_row:
    :param column:
    :return:
    """
    # wb, ws = open_xlsx(new_file)

    conn = oracle.ora_connect()
    ora_prev_data = oracle.ora_get_raw_data(conn, date_since, pid_row)
    ora_current_data = oracle.ora_get_raw_data(conn, date_to, pid_row)

    # for x in range(1, ws.max_row + 1):
    #    ws.cell(row=x, column=column).value = None

    # wb.save(new_file)
    # wb.close()


def create_file_with_data(template_name, since_date, to_date):
    """
    Create xlsx file. And then fill in the file
    :param template_name: Name of template
    :param since_date: Start date of period
    :param to_date: End date of period

    :return: True if
    """


    pid_rows = get_pids(template_file)
    fill_xlsx(new_file, pid_rows, since_date, to_date)
