import fa


def main():
    # Задача на будущее: прикрутить argparser

    template_name = 'template_KOE'

    new_file_name = fa.new_file_name(template_name, to_date)
    template_name = f'./template/{template_name}.xlsx'

    # Даты начала и конца периода отчета
    date_since = fa.ask_date('Введите дату предыущего отчетного периода')
    date_to = fa.ask_date('Введите дату этого отчетного периода')

    fa.copy_template_file(template_name, new_file_name)

    fa.create_file_with_data(template_name, date_since, date_to)


if __name__ == '__main__':
    main()