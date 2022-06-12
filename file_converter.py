import pyxlsb
import pandas
import xlsxwriter

def file_converter(initial_file, name=False):
    """Функция конвертации файла в формате xlsb (двоичный файл Excel) в формат xlsx (файл Excel)"""
    print('Начата конвертация ...')
    #Чтение двоичного файла Excel с помощью модулей pandas и pyxlsb
    try:
        data = pandas.read_excel(initial_file, engine='pyxlsb', sheet_name=None)
    except FileNotFoundError:
        print('\nФайл не найден.\nПроверьте наличие файла и правильность его имени.')
        return
    # Запись файла Excel с помощью модулей Pandas и xlsxwriter
    if not name: # если имя конвертированного файла не указано, то создается файл с таким же именем
        name = f'{initial_file[:-1]}x'
    try:
        writer = pandas.ExcelWriter(name, engine='xlsxwriter')
    except PermissionError:
        print('\nНет прав на запись файла'
              '\nПроверте права на запись в указанной папке '
              '\nПри наличии файла с таким же именем и форматом закройте файл.')
        return
    for sheet_name in data:
        data[sheet_name].to_excel(writer, sheet_name=sheet_name, index=False, header=False)
    writer.save()
    print('Конвертация закончена')



file = 'Coal indices.xlsb'

file_converter(initial_file=file)