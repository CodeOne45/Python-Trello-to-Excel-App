import simplejson
import re
import pandas as pd
import tkinter as tk
import xlsxwriter
from tkinter import filedialog
from tkinter import messagebox

# Interface---

root = tk.Tk()

canvas1 = tk.Canvas(root, width=300, height=300,
                    bg='lightsteelblue2', relief='raised')
canvas1.pack()

label1 = tk.Label(
    root, text="Covertir JSON --> Excel \n Trello format accepted : \n[Title] #SuenceNumber 'Discription' (ETA h)")
canvas1.create_window(150, 60, window=label1)

# End Interface----

data = None


def getJSON():
    MsgBox = tk.messagebox.showinfo(
        'Execel created',
        "Please make sure that your Trollo card respect this fomat : \n [Title] #SuenceNumber 'Discription' (ETA h)")
    global data
    import_file_path = filedialog.askopenfilename()
    json_data = open(import_file_path, errors='ignore').read()
    json_data = '[' + json_data + ']'
    data = simplejson.loads(json_data)


browseButton_JSON = tk.Button(text="      Import JSON File     ",
                              command=getJSON, bg='green',
                              fg='white', font=('helvetica', 12, 'bold'))
canvas1.create_window(150, 130, window=browseButton_JSON)


def extractTrello():

    global data
    l = [["" for i in range(len(data[0]['cards'])+1)] for j in range(6)]

    l[0][0] = "Status"
    l[1][0] = "Titre"
    l[2][0] = "Sequence"
    l[3][0] = "Discription"
    l[4][0] = "Trello URL"
    l[5][0] = "ETA"

    for i in range(len(data[0]['cards'])):

        # l[1][i+1] = data[0]['cards'][i]['name']
        s = data[0]['cards'][i]['name']
        correctFormat = True

        if len(re.findall(r'\[(.*?)\]', s)) > 0:
            l[1][i+1] = (re.findall(r'\[(.*?)\]', s))[0]
        else:
            correctFormat = False

        if len(re.findall(r"\#(.*?)\'", s)) > 0:
            l[2][i+1] = (re.findall(r"\#(.*?)\'", s))[0]
        else:
            correctFormat = False

        if len(re.findall(r"\'(.*?)\'", s)) > 0:
            l[3][i+1] = (re.findall(r"\'(.*?)\'", s))[0]
        else:
            correctFormat = False

        if len(re.findall(r"\((.*?)\)", s)) > 0:
            res = [st.replace('h', '')
                   for st in (re.findall(r"\((.*?)\)", s))[0]]
            l[5][i+1] = res[0]
        else:
            correctFormat = False

        l[4][i+1] = data[0]['cards'][i]['shortUrl']
        l[6][i+1] = data[0]['cards'][i]['attachments'][0]['date']

        if correctFormat == False:
            MsgBox = tk.messagebox.askquestion(
                'Incorrect format', 'The card ' + data[0]['cards'][i]['id'] +
                ' does not repect the format, continuing the extraction could impact the final excel file ! '
                + '\n The card URL :' + data[0]['cards'][i]['shortUrl']
                + '\n Do you want to continue ?', icon='warning')
            if MsgBox == 'no':
                root.destroy()

        for j in range(len(data[0]['lists'])):
            if data[0]['cards'][i]['idList'] == data[0]['lists'][j]['id']:
                l[0][i+1] = data[0]['lists'][j]['name']

    return l


def getExcel():
    global data
    if data is not None:
        workbook = xlsxwriter.Workbook('Trello.xlsx')
        worksheet = workbook.add_worksheet()

        row = 0

        for col, data in enumerate(extractTrello()):
            worksheet.write_column(row, col, data)
        workbook.close()
        MsgBox = tk.messagebox.showinfo(
            'Execel created',
            'JSON file has been converted to Excel with sucess!')
    else:
        MsgBox = tk.messagebox.showinfo(
            'Execel created',
            'Please upload the JSON file befor the extraction !')


saveAsButton_CSV = tk.Button(text='Convert JSON to Excel', command=getExcel,
                             bg='green', fg='white',
                             font=('helvetica', 12, 'bold'))
canvas1.create_window(150, 180, window=saveAsButton_CSV)


def exitApplication():
    MsgBox = tk.messagebox.askquestion(
        'Exit Application', 'Are you sure you want to exit the application',
        icon='warning')
    if MsgBox == 'yes':
        root.destroy()


exitButton = tk.Button(root, text='       Exit Application     ',
                       command=exitApplication, bg='brown',
                       fg='white', font=('helvetica', 12, 'bold'))
canvas1.create_window(150, 230, window=exitButton)

root.mainloop()
