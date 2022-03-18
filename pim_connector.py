import mysql.connector
import pandas as pd


class PimConnector:

    def __init__(self, host, user, password):
        self.host = host
        self.user = user
        self.password = password 
        self.database = self.getLatestDatabase()

        self.conn = mysql.connector.connect(
                host=self.host,
                user=self.user,
                password=self.password,
                database=self.database
        )

    def getLatestDatabase(self):
        server = mysql.connector.connect(
            host=self.host,
            user=self.user,
            password=self.password
        )

        cursor = server.cursor()
        cursor.execute('SHOW DATABASES')

        database = ""
        for db in cursor:
            database = db[0]

        server.close()
        return database

    def get_df_from_table(self, table, filter='', condition=''):
        print('Reading sql database')
        query = f'SELECT * FROM {table}'
        if filter != '' and condition != '':
            query = f'{query} WHERE {filter}={condition}'

        cursor = self.conn.cursor()
        cursor.execute(query)

        sql_data = pd.DataFrame(cursor.fetchall())
        if sql_data.empty:
            return None

        sql_data.columns = cursor.column_names

        return sql_data


if __name__ == '__main__':
    pass