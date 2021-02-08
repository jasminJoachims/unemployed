# -*- coding: utf-8 -*-
"""
Created on Fri Jan 29 14:52:45 2021

@author: jasmin
"""

import pandas as pd
from openpyxl import load_workbook
import timeit

start = timeit.default_timer()

file = "C:\\Users\\jasmi\\Institut\\202012_Heft_pol_SGBI_Arbeitslose.xlsx"
file2 = "C:\\Users\\jasmi\\Institut\\202012_Heft_pol_SGBI_ArbeitslosenQuoten.xlsx"

sheet_liste = ["Gesamt", "Frauen", "Männer", "SGB_II", "SGB_III", "schwerbehindert"]
sheet_liste2 = ["Gesamt", "SGB_II", "SGB_III"]

months = ["Januar 2019", "Februar 2019", "März 2019", "April 2019", "Mai 2019", "Juni 2019", "Juli 2019", "August 2019", "September 2019", "Oktober 2019", "November 2019", "Dezember 2019"]
months2 = ["Januar 2020", "Februar 2020", "März 2020", "April 2020", "Mai 2020", "Juni 2020", "Juli 2020", "August 2020", "September 2020", "Oktober 2020", "November 2020", "Dezember 2020"]


col_names = ["Gesamt_Quote", "SGB_II_Quote", "SGB_III_Quote"]
col_names_sorted = ["Monat", "Gesamt", "Frauen", "Männer", "Gesamt_Quote", "SGB_II", "SGB_II_Quote", "SGB_III", "SGB_III_Quote", "schwerbehindert"]

path = "C:\\Users\\jasmi\\Institut\\Frederike2.xlsx"

def read_rows(file, sheet_liste, k):
    rows_2019 = []
    rows_2020 = []

    for i in sheet_liste:
        x = pd.read_excel(file, sheet_name=i, header=None)
        print("Rows read")
        row = x.iloc[k]
        row_2019 = row[3:15]
        row_2020 = row[15:27]
        rows_2019.append(row_2019)
        rows_2020.append(row_2020)
        print("Rows for sheet added")
        
    return rows_2019, rows_2020
  
def transpose_rows(array, sheet_liste):  
    data = pd.DataFrame(array)
    data = data.T
    data.columns = sheet_liste
    
    return data


def create_table(file, file2, sheet_liste, sheet_liste2, k):
    print(k)
    rows_2019, rows_2020 = read_rows(file, sheet_liste, k)
    qrows_2019, qrows_2020 = read_rows(file2, sheet_liste2, k)
    print("All rows added")
    
    rows_2019 = transpose_rows(rows_2019, sheet_liste)
    rows_2020 = transpose_rows(rows_2020, sheet_liste)
    qrows_2019 = transpose_rows(qrows_2019, col_names)
    qrows_2020 = transpose_rows(qrows_2020, col_names)
    
    table_2019 = rows_2019.join(qrows_2019)
    table_2019["Monat"] = months
    
    table_2020 = rows_2020.join(qrows_2020)
    table_2020["Monat"] = months2
    
    table_2019 = table_2019[col_names_sorted]
    table_2020 = table_2020[col_names_sorted]
    
    return table_2019, table_2020

def merge_dataFrames(dataFrame1, dataFrame2):
    
    combined = dataFrame1.append(dataFrame2, ignore_index = True)
    
    return combined

index = [11, 12, 1134, 1137, 2132, 2137, 2592, 3044, 5383, 6533, 8693, 8752, 8755, 9191, 9926, 10359, 10592]

states =["Deutschland", "Schleswig-Holstein", "Hamburg",
         "Niedersachsen", "Bremen", "Nordrhein-Westfalen",
         "Hessen", "Rheinland-Pfalz", "Baden-Württemberg",
         "Bayern", "Saarland", "Berlin", "Brandenburg",
         "Mecklenburg-Vorpommern", "Sachsen", "Sachsen-Anhalt",
         "Thüringen"]

year = ["2019", "2020"]

i = 0


for k in index:    
        state = states[i]
        print(state)
        table_2019, table_2020 = create_table(file, file2, sheet_liste, sheet_liste2, k)
        table = merge_dataFrames(table_2019, table_2020)
        print("Tables of {} created".format(state))
        book = load_workbook(path)
        writer = pd.ExcelWriter(path, engine='openpyxl') 
        writer.book = book
        writer.sheets = dict((ws.title, ws) for ws in book.worksheets)
        
        table.to_excel(writer, state)
        print("Tables of {} saved".format(state))
        i = i+1
        writer.save()
        break


stop = timeit.default_timer()

delta = round((stop-start)/60, 2)

print('Runtime: {} minutes'.format(delta))

