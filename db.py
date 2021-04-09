import sqlite3 as sq
import os


class Database:
    def __init__(self):
        self.__rootpath = os.path.abspath(os.path.dirname(__file__))
        self.__dbpath = os.path.join(self.__rootpath, ".db")
        if not os.path.isdir(self.__dbpath):
            os.makedirs(self.__dbpath)
        self.__dbname = os.path.join(self.__dbpath, "emailclient.db")

    def connect(self):
        try:
            conn = sq.connect(self.__dbname)
            return True, conn
        except Exception as error:
            return False, str(error)

    def execute(self, sql, param=None):
        status, conn = self.connect()
        if status:
            try:
                cursor = conn.cursor()
                if param is None:
                    cursor.execute(sql)
                else:
                    cursor.execute(sql, param)
                returnvalue = cursor.fetchall()
                conn.commit()
                cursor.close()
                conn.close()
                return True, returnvalue
            except Exception as error:
                return False, str(error)
        return False, str(conn)
