# -*- coding: cp1251 -*-
import csv
import xlwt
import numpy
from tkinter import *
from tkinter import ttk
from tkinter.filedialog import askopenfilename

def get_csv_str (ord_dict):
    print (ord_dict['Description'])
    print (ord_dict['ТУ'])
    print (ord_dict['Quantity'])
    return

def open_dialog():# Выбор файла .csv для обработки
    root = Tk()
    root.withdraw()
    name = askopenfilename(filetypes =(("CSV file", "*.csv"),
                                   ("All Files","*.*")),
                                   title = "Выберите файл...")
    print (name)
    input_file = open(name, "r")
    reader = csv.DictReader(input_file, fieldnames=['Description', 'ТУ',
                                             'DocumentNumber', 'Quantity'])
    t_list = list()
    t_dict = dict()
    for rec in reader:
        t_list.append ([rec['Description'], rec['ТУ']])
        t_dict.update ({(rec['Description'], rec['ТУ'],
                         rec['DocumentNumber']): rec['Quantity']})
    input_file.close()
    return (t_dict)

def xls_config():# Настройка, создание файла xls
    font0 = xlwt.Font()
    font0.name = 'Times New Roman'
    font0.height = 320

    alignment0 = xlwt.Alignment()
    alignment0.shrink_to_fit = True

    style0 = xlwt.XFStyle()
    style0.font = font0
    style0.alignment = alignment0

    wb = xlwt.Workbook()
    ws = wb.add_sheet('6Ц.270.000 ВП')
    return

d_gen = open_dialog()
d_repeat = dict() # Словарь с кол-во повторов
comp_keys = list() # Все деспкрипшены
t_keys = list() # Все дескрипшены с повторами

for i in list(d_gen.keys()):
    t_keys.append (i[0])
    if i[0] not in comp_keys:
        comp_keys.append (i[0])
        
comp_keys.remove('Description')
#comp_keys = tuple(comp_keys)

for i in comp_keys:
    d_repeat.update({i: t_keys.count(i)})  
    
print (d_repeat)



