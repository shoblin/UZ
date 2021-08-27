import cx_Oracle as ora
import keyring

SQL_REQUEST = """
SELECT      p.mid, b.date_time, b.erm "a+", b.edm "a-"
FROM        billorg b, points p
WHERE       p.mid = b.mid
            AND (b.date_time> TO_DATE('{0}', 'DD.MM.YYYY')-1) and (b.date_time<= TO_DATE('{0}', 'DD.MM.YYYY'))
            AND b.mid IN {1}
ORDER BY p.name
"""


def ora_connect():
    login = keyring.get_password("oracle", "login")
    password = keyring.get_password("oracle", "password")
    connection = ora.connect(login, password, "ASKUE")
    return connection


def ora_get_raw_data(cursor, date, pids):
    date = f"{date.days:02}.{date.month:02}.{date.years:04}"
    pids = ", ".join([str(pid) for pid in pids.keys()])

    cursor.execute(SQL_REQUEST.format(date, pids))

    res = cursor.fetchall()
    return res


def ora_get_data(cursor):
    raw_data = ora_get_raw_data(cursor)
    return type(raw_data)


def ora_close(connection):
    connection.close()


def main():
    try:
        conn = ora_connect()
        cursor = conn.cursor()
        data = ora_get_data(cursor)
        print(data)
    except ora.DatabaseError as e:
        code, mesg = e.args[0].message[:-1].split(': ', 1)
        print(f'code = {code}')
        print(f'value = ', mesg)
    finally:
        ora_close(conn)


if __name__ == '__main__':
    main()