# -*- coding: cp1251 -*-
import csv
import xlwt
import numpy
from tkinter import *
from tkinter import ttk
from tkinter.filedialog import askopenfilename

def get_csv_str (ord_dict):
    print (ord_dict['Description'])
    print (ord_dict['��'])
    print (ord_dict['Quantity'])
    return

def open_dialog():# ����� ����� .csv ��� ���������
    root = Tk()
    root.withdraw()
    name = askopenfilename(filetypes =(("CSV file", "*.csv"),
                                   ("All Files","*.*")),
                                   title = "�������� ����...")
    print (name)
    input_file = open(name, "r")
    reader = csv.DictReader(input_file, fieldnames=['Description', '��',
                                             'DocumentNumber', 'Quantity'])
    t_list = list()
    t_dict = dict()
    for rec in reader:
        t_list.append ([rec['Description'], rec['��']])
        t_dict.update ({(rec['Description'], rec['��'],
                         rec['DocumentNumber']): rec['Quantity']})
    input_file.close()
    return (t_dict)

def xls_config():# ���������, �������� ����� xls
    font0 = xlwt.Font()
    font0.name = 'Times New Roman'
    font0.height = 320

    alignment0 = xlwt.Alignment()
    alignment0.shrink_to_fit = True

    style0 = xlwt.XFStyle()
    style0.font = font0
    style0.alignment = alignment0

    wb = xlwt.Workbook()
    ws = wb.add_sheet('6�.270.000 ��')
    return

d_gen = open_dialog()
d_repeat = dict() # ������� � ���-�� ��������
comp_keys = list() # ��� ������������
t_keys = list() # ��� ����������� � ���������

for i in list(d_gen.keys()):
    t_keys.append (i[0])
    if i[0] not in comp_keys:
        comp_keys.append (i[0])
        
comp_keys.remove('Description')
#comp_keys = tuple(comp_keys)

for i in comp_keys:
    d_repeat.update({i: t_keys.count(i)})  
    
print (d_repeat)



