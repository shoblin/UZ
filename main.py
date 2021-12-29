import fa
import sys
import oracle as ora


def main():
    # Задача на будущее: прикрутить argparser
    first_date, second_date = None, None
    template_name = 'template_BOE'

    # Даты начала и конца периода отчета вводим с клавиатуры
    while fa.check_dates_difference(first_date, second_date):
        first_date = fa.request_date('Введите дату предыущего отчетного периода')
        second_date = fa.request_date('Введите дату этого отчетного периода')

    print(first_date, second_date)
    # Сортируем, что бы случайно не перепутали даты
    date_since, date_to = fa.order_dates(first_date, second_date)

    # Копируем фаил темплейта в файл <region_name>_<month:02>_<year:04>.xlsx'
    new_file_name = fa.new_file_name(template_name, date_to)
    template_name = f'./template/{template_name}.xlsx'

    if not fa.copy_template_file(template_name, new_file_name):
        print(f'Копия файла {template_name} не создан')
        sys.exit(1)

    # Получаем словарь с points_id: номерами строк
    dict_pids = fa.get_pids(template_name)

    # Подлючаемся к базе данных
    ora_connection = ora.ora_connect()
    cursor = ora_connection.cursor()
    raw_db_data_date_since = ora.ora_get_raw_data(cursor, date_since, dict_pids)
    raw_db_data_date_to = ora.ora_get_raw_data(cursor, date_to, dict_pids)

    test_msg(raw_db_data_date_since, raw_db_data_date_to)

    #fa.create_file_with_data(template_name, date_since, date_to)


def test_msg(raw_db_data_date_since, raw_db_data_date_to):
    print(raw_db_data_date_since)
    print()
    print(raw_db_data_date_to)


if __name__ == '__main__':
    main()