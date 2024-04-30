# -*- coding: utf-8 -*-
"""
Created on Mon Sep 18 03:22:46 2023

@author: eiru_
"""
from openpyxl import load_workbook, Workbook
def sift_vaul(file_path,sheet_name_1,sheet_name_2):#清掃割り振り票の集積をローテーションさせる関数
    wb = load_workbook(filename=file_path)
    sk = wb[sheet_name_1]
    al = wb[sheet_name_2]
    print(sk["A1"].value)
    x = sk["A1"].value
    if x==1:
        al["E18"] = 3
        al["F18"] = 0
        al["G18"] = 0
        al["E19"] = 0
        al["F19"] = 3
        al["G19"] = 0
        al["E20"] = 0
        al["F20"] = 0
        al["G20"] = 3
        al["E21"] = 3
        al["F21"] = 0
        al["G21"] = 0
        al["E22"] = 0
        al["F22"] = 3
        al["G22"] = 0
        sk["A1"] = 2
    elif x==2:
        al["E18"] = 0
        al["F18"] = 3
        al["G18"] = 0
        al["E19"] = 0
        al["F19"] = 0
        al["G19"] = 3
        al["E20"] = 3
        al["F20"] = 0
        al["G20"] = 0
        al["E21"] = 0
        al["F21"] = 3
        al["G21"] = 0
        al["E22"] = 0
        al["F22"] = 0
        al["G22"] = 3
        sk["A1"] = 3
    elif x==3:
        al["E18"] = 0
        al["F18"] = 0
        al["G18"] = 3
        al["E19"] = 3
        al["F19"] = 0
        al["G19"] = 0
        al["E20"] = 0
        al["F20"] = 3
        al["G20"] = 0
        al["E21"] = 0
        al["F21"] = 0
        al["G21"] = 3
        al["E22"] = 3
        al["F22"] = 0
        al["G22"] = 0
        sk["A1"] = 1
    wb.save(filename=file_path)
