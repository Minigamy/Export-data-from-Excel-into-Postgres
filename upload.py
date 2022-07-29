import tkinter
from tkinter import messagebox

import sys

import psycopg2
import pandas as pd
from sqlalchemy import create_engine
import os



""" Path to file """
path_to_file_excel = sys.argv[1]  # Путь к Excel файлу, полученный из макроса.
path_to_file_dir = path_to_file_excel.rpartition('\\')[0]  # Путь до дирректории файла Excel. Необходим для создания там файла conn_to_postgres.txt.



""" Open the file <conn_to_postgres.txt> and take from there data to connect to the database and path to the XLSM file """
def check_env_variables():
    if os.path.exists(path_to_file_dir + r"\conn_to_postgres.txt"):

        with open(path_to_file_dir + r"\conn_to_postgres.txt", "rt", encoding='utf-8') as file:
            file_source = file.read()

        values = file_source.split("\n")

        values[0] = values[0].replace('USER = "', '')
        values[0] = values[0].replace('"', '')

        values[1] = values[1].replace('PASSWORD = "', '')
        values[1] = values[1].replace('"', '')

        values[2] = values[2].replace('HOSTNAME = "', '')
        values[2] = values[2].replace('"', '')

        values[3] = values[3].replace('DATABASE = "', '')
        values[3] = values[3].replace('"', '')


    else:
        with open(path_to_file_dir + r"\conn_to_postgres.txt", 'w', encoding='utf-8') as f:
            f.write('USER = "postgres"\n')
            f.write('PASSWORD = "postgres"\n')
            f.write('HOSTNAME = "127.0.0.1"\n')
            f.write('DATABASE = "postgres"\n')

            values = ["postgres", "postgres", "127.0.0.1", "postgres", r"C:\Users\myuser\названия точек.xlsm"]

            tkinter.messagebox.showinfo(title='INFO',
                                        message='Для корректной работы программы необходимо заполнить файл conn_to_postgres.txt своими данными!\nДанный файл автоматически создался в папке с Excel файлом.')

    return values


env_values = check_env_variables()



""" Connection data """
USER = env_values[0]
PASSWORD = env_values[1]
HOSTNAME = env_values[2]
DATABASE = env_values[3]



""" Connection string """
conn_string = f'postgresql+psycopg2://{USER}:{PASSWORD}@{HOSTNAME}/{DATABASE}'



""" CREATE TABLE IF NOT EXISTS """
try:
    connection = psycopg2.connect(user=USER,
                                  password=PASSWORD,
                                  host=HOSTNAME,
                                  port="5432",
                                  database=DATABASE)
    cursor = connection.cursor()

    create_table_query = '''CREATE TABLE IF NOT EXISTS endpoint_names
    (endpoint_id serial PRIMARY KEY,
    endpoint_name  TEXT NOT NULL);'''

    cursor.execute(create_table_query)
    connection.commit()
except psycopg2.OperationalError:
    tkinter.messagebox.showerror(title='ERROR', message="Ошибка подключения к БД!")



"""
Connect to DB
Read the Excel file
Create DataFrame
Write DataFrame into Postgres
"""
try:
    db = create_engine(conn_string)
    conn = db.connect()

    data = pd.read_excel(path_to_file_excel, sheet_name='Лист1').to_dict()

    # Создаем DataFrame
    df = pd.DataFrame(data)
    df.to_sql('endpoint_names', con=conn, if_exists='replace', index=False)

    conn.close()
    tkinter.messagebox.showinfo(title='INFO', message='Данные успешно записаны в БД!')
except FileNotFoundError:
    tkinter.messagebox.showerror(title='ERROR',
                                 message="Файл <названия точек.xlsm> не найден!\n Проверьте корректность пути до файла.")
