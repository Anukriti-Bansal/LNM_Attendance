import tkinter
from tkinter import ttk
import tkinter.filedialog
import sys
import pandas as pd
import numpy as np
import csv
from glob import glob
import os

def compile_attendance():
    window = tkinter.Tk()
    window.title('File Explorer')
    window.geometry('700x350')

    label = ttk.Label(window, text='Select multiple files .csv files', font=('Aerial 14')).pack(pady=30)
    btn = tkinter.Button(window, text = 'Select', command=lambda:open_file(window))
#btn.pack(side=TOP, pady=10)
    btn.pack()

    tkinter.Button(window, text = 'Quit', command=window.destroy).pack()
    window.mainloop()


def open_file(window):
    files = tkinter.filedialog.askopenfilenames(parent=window, title='Choose multiple files .csv files to generate attendance record')
    files=list(files)
    print(files)
    print(len(files))
    
    if (len(files) == 0):
        label1 = tkinter.Label(window, text='No file selected', font='Aerial 14', fg='red').pack(pady=30)
    else:
        '''
        list_of_files = []
        for file1 in files:
            dirname,filename = os.path.split(file1)
            list_of_files.append(filename)
        print(list_of_files)
        '''
        column_names = ['Name','Email']
        df = pd.DataFrame(columns=column_names)
        Date=''

        for f in files:
            dirName, filename = os.path.split(f)
            Date = filename[0:10]
            Date = (Date[-2:]+'-'+Date[5:7]+'-'+Date[0:4])
            data = pd.read_csv(f,encoding='latin1')
            data['Email'] = data['Email'].str.upper()
            data['Email'] = data['Email'].str.strip()
            emailID = np.unique(data['Email'])
            Duration = []
            for ID in emailID:
                if ID not in df.values:
                    studentData = data.loc[data['Email']==ID]
                    df = df.append(studentData.iloc[:,:2], ignore_index=True)
        files.sort()
        df.drop_duplicates(subset='Email', inplace=True)

        for f in files: 
            dirName, filename = os.path.split(f)
            Date = filename[0:10]
          #  Date = f[0:10]
            Date = (Date[-2:]+'-'+Date[5:7]+'-'+Date[0:4])
            data = pd.read_csv(f,encoding='latin1')
            data['Email'] = data['Email'].str.upper()
            data['Email'] = data['Email'].str.strip()

            emailID = np.unique(df['Email'])
            Duration = []

            for ID in emailID:
                studentData = data.loc[data['Email']==ID]

                if not studentData.empty:
                    Duration.append(studentData['Duration'].tolist()[0])
                else:
                    Duration.append('A')
            df[Date+'(Duration)'] = Duration

        df.to_csv("Attendance.csv", index=False)

        excelWriter = pd.ExcelWriter('Attendance.xlsx')
        df.to_excel(excelWriter, index=False)
        excelWriter.save()

        label1 = tkinter.Label(window, text='Attendance compiled successfully', font='Aerial 14', fg='green').pack(pady=30)
