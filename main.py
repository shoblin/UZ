import fa


def main():
    # Задача на будущее: прикрутить argparser

    template_name = 'template_KOE'

    # Даты начала и конца периода отчета
    first_date = fa.request_date('Введите дату предыущего отчетного периода')
    second_date = fa.request_date('Введите дату этого отчетного периода')
    date_since, date_to = fa.order_dates(first_date, second_date)

    new_file_name = fa.new_file_name(template_name, date_to)
    template_name = f'./template/{template_name}.xlsx'

    success_copy = False
    if not fa.copy_template_file(template_name, new_file_name):
        print('Файл не создан')
        return 0

    fa.create_file_with_data(template_name, date_since, date_to)


if __name__ == '__main__':
    main()