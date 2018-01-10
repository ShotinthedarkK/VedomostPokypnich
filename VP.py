# -*- coding: cp1251 -*-
import csv
import xlwt
import numpy as np
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
    g_list = list()
    for rec in reader:
        t_list = [rec['Description'], rec['ТУ'],
                   rec['DocumentNumber'], rec['Quantity']]
        g_list.append(t_list)
    input_file.close()
    return (g_list)

def xls_write():# Настройка, создание файла xls
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

    for i in range(len(fin)):
        for j in range(len(fin[i])):
            ws.write(i, j, fin[i] [j], style0)

    wb.save('6Ц.270.000ВП.xls')
    return 

a = np.array(open_dialog())
k = a.shape[0]
while k > 0:
    if a[k-1, 0] == a[k-2, 0]:
        a [k-1, 0] = ''
        a [k-1, 1] = ''
    k = k- 1
print (a)

k = a.shape[0]
i = 1
fin = list()
while i < k:
    fin.append(list(a[i- 1]))
    if (a[i- 1, 0] == '') & (a[i, 0] != '') & (a[i- 1, 2] != '6Ц.270.110 Э3'):
        fin.append(['', '', '', int(a[i- 1, 3])+ int(a[i- 2, 3])])
    elif (a[i- 1, 0] == '') & (a[i, 0] != '') & (a[i, 2] = '6Ц.270.110 Э3'):
        fin.append(['', '', '', int(a[i- 1, 3])+ int(a[i- 2, 3])]) # некорректно
    i = i+ 1
a = fin
print (a)

xls_write()

'''
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
    
print (d_repeat)'''



