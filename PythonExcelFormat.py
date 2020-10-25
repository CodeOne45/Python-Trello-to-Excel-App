

#f = open("myfile.txt", "w")
import xlrd
import pandas as MyPandas
import openpyxl
import re


titreColonneExcel = "Produit,Titre,Ordre de la tâche,Tâche Charge (heures),Lien Trello"

charListe = ['[', ']', '#', '(', ')']

listeTrello = ["[Titre A] #p BlaBlaBla (xh)", "[Titre A] #p BlaBlaBla (xh)",
               "[Titre A] #p BlaBlaBla (xh)", "[Titre A] #p BlaBlaBla (xh)", "[Titre A] #p BlaBlaBla (xh)"]

newListe = []


def convertTypeExcel(l, newl):
    for i in range(0, len(l)):
        res = re.split('] |#| ', l[i])
        res = [s.replace('[', '') for s in res]
        res = [s.replace('(', '') for s in res]
        res = [s.replace(')', '') for s in res]
        res = [s.replace('h', '') for s in res]
        newl.append(res)


convertTypeExcel(listeTrello, newListe)

print(newListe)


fname = r'C:\Users\malik\OneDrive\Bureau\Info DUT\BoldCode_\PythonJsonToExcel\\BC.Project - VEDEM Maintenance (1).xlsx'

wb = xlrd.open_workbook(fname)
sheet = wb.sheet_by_index(0)
print(sheet.nrows)

for i in range(sheet.nrows):
    for j in range(sheet.ncols):
        print(sheet.cell_value(i, j))

# https://radiusofcircle.blogspot.com/2016/03/how-to-read-data-from-excel-or-spreadsheet-files-using-python.html
