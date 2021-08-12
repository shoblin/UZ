import openpyxl as xl
import shutil
import logging
import oracle
from datetime import datetime


PIDS_COL = 1
data = {
    1: {'name': 'Ivan', 'surname': 'Ivanov', 'points': 450},
    2: {'name': 'Petro', 'surname': 'Petrov', 'points': 550},
    3: {'name': 'Semen', 'surname': 'Uruk', 'points': 1050},
    4: {'name': 'Valera', 'surname': 'Neo', 'points': 452}
}
row_config = {'name': 2, 'surname': 3, 'points': 4}


# Function for get date
class FutureDate(Exception):
    def __init__(self, text):
        self.txt = text


def get_date(str_date):

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


def ask_date(txt):
    true_date = False
    date = None
    while not true_date:
        print(txt)
        str_date = input('В формате ДД-ММ-ГГГГ: ')
        true_date, date = get_date(str_date)

    return date


# Function for create xlsx file
def get_template_name():
    return None


def copy_template_file(template_file, new_file):
    """
    Copy file <template_file> to <new_file>
    :param template_file: Source file - file with template
    :param new_file: Destination file - new file with date from DB
    :return: NONE
    """
    try:
        shutil.copy(template_file, new_file)
    except shutil.SameFileError:
        logging.error('Source and destination represents the same file.')

        # If there is any permission issue
    except PermissionError:
        logging.error("Permission denied.")

        # For other errors
    except:
        logging.error("Error occurred while copying file.")


def new_file_name(template, to_date):
    """
    Create name for new file = region name + month
        region name - we get from template name
        month - number of month
    :param template: Source file - file with template
    :param to_date: date of end  of period
    :return:
        name for new name = {region_name}_{month}.xlsx
    """
    region_name = template.split('_')[-1]
    return f'./{region_name}_{to_date.month:02}_{to_date.year:04}.xlsx'


def open_xlsx(file_name):
    wb = xl.load_workbook(file_name)
    ws = wb.worksheets[0]
    return wb, ws


def get_pids(template_file, pid_colunm=PIDS_COL):
    """
    Get points id from excel column
    :param template_file: Source file - file with template
    :param pid_colunm: column number. Default value = PIDS_COL
    :return pids: diction {row: point id}
    """
    wb, ws = open_xlsx(template_file)

    # calculate total number of rows and
    # columns in source excel file
    mr, mc = ws.max_row, ws.max_column

    pids = dict()
    for num_row in range(1, mr + 1):
        pid = ws.cell(row=num_row, column=pid_colunm).value
        if pid:
            pids[pid] = num_row

    wb.close()

    return pids


def get_data(pid):
    return data.get(pid)


def fill_xlsx(new_file, pid_row, date_since, date_to, column=PIDS_COL):
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
    logging.basicConfig(filename='uz.log', filemode='a', format='%(asctime)s - %(levelname)s - %(message)s')
    new_file = new_file_name(template_name, to_date)
    template_file = f'./template/{template_name}.xlsx'

    copy_template_file(template_file, new_file)
    pid_rows = get_pids(template_file)
    fill_xlsx(new_file, pid_rows, since_date, to_date)
