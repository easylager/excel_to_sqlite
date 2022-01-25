import os
import sqlite3
import openpyxl
import re
import pyexcel as p

#p.save_book_as(file_name='balance.xls',
               #dest_file_name='balance.xlsx')


def export_to_sqlite():
    '''Экспорт данных из xls в sqlite'''
    # 1. Создание и подключение к базе
    # Получаем текущую папку проекта
    prj_dir = os.path.abspath(os.path.curdir)

    a = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))

    # Имя базы
    base_name = 'auto.sqlite3'

    # метод sqlite3.connect автоматически создаст базу, если ее нет
    connect = sqlite3.connect(prj_dir + '/' + base_name)
    # курсор - это специальный объект, который делает запросы и получает результаты запросов
    cursor = connect.cursor()
    # создание таблицы если ее не существует
    cursor.execute('create table if not exists class (class_id int primary key, title text, in_active float, in_passive float ,debet float, credit float, out_active float, out_passive float);')
    cursor.execute('create table if not exists bank (bank_id int primary key, in_active float, in_passive float, debet float, credit float, out_active float, out_passive float, class_id int, foreign key(class_id) references calss(class_id));')
    cursor.execute('create table if not exists account (account_id int, in_active float, in_passive float ,debet float, credit float, out_active float, out_passive float, class_id int, bank_id int, foreign key(class_id) references class(class_id), foreign key(bank_id) references bank(bank_id));')
    # 2. Работа c xlsx файлом
    # Читаем файл и лист1 книги excel
    file_to_read = openpyxl.load_workbook('balance.xlsx', data_only=True)
    sheet = file_to_read['Sheet1']
    # Цикл по строкам начиная со второй (в первой заголовки)
    for row in range(10, 626):
        # Объявление списка
        data = []
        # Цикл по столбцам от 1 до 4 ( 5 не включая)
        for col in range(1, 8):
            # value содержит значение ячейки с координатами row col
            value = sheet.cell(row, col).value
            # Список который мы потом будем добавлять
            data.append(value)
        if data[1] == None:
            continue
    # 3. Запись в базу и закрытие соединения
        # Вставка данных в поля таблицы
        if re.fullmatch(r'\d\d', str(data[0])):
            print('bank')
            cursor.execute("INSERT INTO bank(bank_id, in_active, in_passive, debet, credit, out_active, out_passive) VALUES (?, ?, ?, ?, ?, ?, ?);", (data[0], data[1], data[2], data[3], data[4], data[5], data[6]))
        elif re.fullmatch(r'\d{4}', str(data[0])):
            print(type(data[0]))
            print('account')
            cursor.execute("INSERT INTO account(account_id, in_active, in_passive, debet, credit, out_active, out_passive, class_id, bank_id) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?);", (data[0], data[1], data[2], data[3], data[4], data[5], data[6], int(data[0][:1]), int(data[0][:2])))
    # сохраняем изменения
    connect.commit()
    # закрытие соединения
    connect.close()


def clear_base():
    '''Очистка базы sqlite'''

    # Получаем текущую папку проекта
    prj_dir = os.path.abspath(os.path.curdir)

    # Имя базы
    base_name = 'auto.sqlite3'

    connect = sqlite3.connect(prj_dir + '/' + base_name)
    cursor = connect.cursor()

    # Запись в базу, сохранение и закрытие соединения
    cursor.execute("DELETE FROM cars")
    connect.commit()
    connect.close()


# Запуск функции
export_to_sqlite()


