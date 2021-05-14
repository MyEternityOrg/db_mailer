import os

# import pyodbc
import pymssql


class MSSQLConnection:
    def __init__(self, server, database, login, password):
        self.__connected = False
        self.__server = server
        self.__database = database
        self.__login = login
        self.__password = password
        self.__driver = "{SQL Server Native Client 11.0}"
        try:
            self.__connection = pymssql.connect(self.__server, self.__login, self.__password, self.__database)
            self.__connected = True
        except Exception as E:
            print(f"Исключительная ситуация при подключении к БД: {E}")
            self.__connected = False
            self.__connection = None

    @property
    def connection(self):
        return self.__connection

    @property
    def connected(self):
        return self.__connected

    def execute(self, query, params=()):
        if self.connection is not None:
            try:
                if len(params) > 0:
                    self.connection.cursor().execute(query, params)
                    self.connection.commit()
                else:
                    self.connection.cursor().execute(query)
                    self.connection.commit()
                return True
            except Exception as E:
                print(f"Исключительная ситуация (execute): {E}")
                return False
        else:
            return False

    def select(self, query, params=()):
        if self.connection is not None:
            try:
                cursor = self.connection.cursor()
                if len(params) > 0:
                    cursor.execute(query, params)
                else:
                    cursor.execute(query)
                for row in cursor:
                    yield [row[el] for el in range(0, len(row))]
            except Exception as E:
                print(f'Исключительная ситуация (select): {E}')
                return ()

    @classmethod
    def file_to_binary_data(cls, filename, delete_after_load=True):
        try:
            with open(filename, "rb") as f:
                binary_data = f.read()
            if delete_after_load:
                os.remove(filename)
            return binary_data
        except Exception as E:
            print(f"Исключительная ситуация при загрузке данных из файла: {E}")
            return 0
