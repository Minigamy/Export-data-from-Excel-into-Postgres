# Export-data-from-Excel-into-Postgres
Export data from Excel into Postgres with macros VBA

# Описание
Реализован скрипт upload.py и исполняемый файл upload.exe, запускаемый по кнопке "Загрузить" в файле "названия точек.xlsm",
для загрузки/обновления данных из файла в таблицу в БД Postgres. Для удобства пользователя при запуске скрипта в папке с Excel фалом создается файл `conn_to_postgres.txt`, в котором необходимо указать данные для подключения к вашей БД Postgres.

### Для корректной работы макроса VBA в файле Excel в настройке мароса необходимо указать путь до исполняемого файла upload.exe.
***
# Стек
- psycopg2
- pandas
- sqlalchemy
- tkinter
- pyinstaller