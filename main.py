import fa


def main():
    template = 'template_KOE'

    date_since = fa.ask_date('Введите дату предыущего отчетного периода')
    date_to = fa.ask_date('Введите дату этого отчетного периода')

    pids_dict = fa.get_pids(template)
    fa.create_file_with_data(template, date_since, date_to, pids_dict)


if __name__ == '__main__':
    main()