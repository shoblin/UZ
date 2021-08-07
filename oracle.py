import cx_Oracle as ora
import keyring

SQL_REQUEST = """
SELECT   p.mid, b.date_time, b.erm "a+", b.edm "a-"
FROM   billorg b, points p
WHERE   p.mid = b.mid 
        AND (b.date_time> TO_DATE('{}', 'DD.MM.YYYY')-1) and (b.date_time<= TO_DATE('{}', 'DD.MM.YYYY'))
        AND b.mid IN (SELECT mid
                      FROM   points
                      START WITH mid = 59666
                      CONNECT BY PRIOR mid = parent)
ORDER BY p.name;
"""


def ora_connect():
    login = keyring.get_password("oracle", "login")
    password = keyring.get_password("oracle", "password")
    connection = ora.connect(login, password, "ASKUE")
    return connection


def ora_get_data(cursor):
    cursor.execute(SQL_REQUEST)

    res = cursor.fetchall()
    return res

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