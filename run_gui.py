#!/usr/bin/env python
# -*- coding: utf-8 -*-

import os
import logging
import platform
import datetime

import pandas as pd
from pathlib import Path
import tkinter as tk
from tkinter import filedialog

import overall_evaluation
import sort_certificate

gui = tk.Tk()

gui.geometry('555x320')
gui.minsize(555, 320)
gui.title('Wellenwettkampf Auswertung')

inputFilePath = tk.StringVar()
outputPath = tk.StringVar()
customName = tk.StringVar()
exportDetail = tk.IntVar()


class TextHandler(logging.Handler):
    '''This class allows you to log to a Tkinter Text or ScrolledText widget'''
    def __init__(self, text):
        # run the regular Handler __init__
        logging.Handler.__init__(self)
        # Store a reference to the Text it will log to
        self.text = text

    def emit(self, record):
        msg = self.format(record)

        def append():
            self.text.configure(state='normal')
            self.text.insert(tk.END, msg + '\n')
            self.text.configure(state='disabled')
            # Autoscroll to the bottom
            self.text.yview(tk.END)
        # This is necessary because we can't modify the Text from other threads
        self.text.after(0, append)

def getInputFilePath():
    file_selected = filedialog.askopenfilename()
    inputFilePath.set(file_selected)

def getOutputFolderPath():
    folder_selected = filedialog.askdirectory()
    outputPath.set(folder_selected)


frame_Title = tk.Frame(master=gui)
frame_Input = tk.Frame(master=gui)
frame_Output = tk.Frame(master=gui)
frame_OutputName = tk.Frame(master=gui)
frame_DetailExport = tk.Frame(master=gui)
frame_Start = tk.Frame(master=gui)
frame_Log = tk.Frame(master=gui)

frame_Title.pack(fill='x')
frame_Input.pack(fill='x')
frame_Output.pack(fill='x')
frame_OutputName.pack(fill='x')
frame_DetailExport.pack(fill='x')
frame_Start.pack(fill='x')
frame_Log.pack(pady='6', fill='both')
logger_Scrollbar = tk.Scrollbar(master=frame_Log)
logger_Element = tk.Text(master=frame_Log, height=5, state=tk.DISABLED, font='TkFixedFont')
logger_Scrollbar.pack(side='right', fill='y')
logger_Element.pack(side='left', fill='x', expand=True)
logger_Scrollbar.config(command=logger_Element.yview)
logger_Element.config(yscrollcommand=logger_Scrollbar.set)

# Create textLogger
text_handler = TextHandler(logger_Element)

# Add the handler to logger
logger = logging.getLogger()
logger.addHandler(text_handler)


def creation_date(path_to_file):
    '''
    Try to get the date that a file was created, falling back to when it was
    last modified if that isn't possible.
    See http://stackoverflow.com/a/39501288/1709587 for explanation.
    '''
    if platform.system() == 'Windows':
        return os.path.getctime(path_to_file)
    else:
        stat = os.stat(path_to_file)
        try:
            return stat.st_birthtime
        except AttributeError:
            # We're probably on Linux. No easy way to get creation dates here,
            # so we'll settle for when its content was last modified.
            return stat.st_mtime


def generate_overall_evaluation():
    file = inputFilePath.get()
    name = customName.get()

    # Check if path exist
    if not os.path.exists(file):
        logger.error('❌ DATEI EXISTIERT NICHT')
        print('❌ DATEI EXISTIERT NICHT')
        return

    # Check for Excel file
    if not file.endswith(('.xlsx', '.xls')):
        logger.error('❌ KEINE EXCEL DATEI AUSGEWÄHLT')
        print('❌ KEINE EXCEL DATEI AUSGEWÄHLT')
        return

    try:
        wave_table, rescure_table, wave_table_detail, rescure_table_detail = overall_evaluation.calculate(file)
        sort_table = sort_certificate.sort(file)

        op = outputPath.get() if outputPath.get() else Path().resolve()
        fn = f'{name}_Auswertung.xlsx' if name else 'Auswertung.xlsx'
        file_path = os.path.join(op, fn)

        with pd.ExcelWriter(file_path, engine='xlsxwriter') as writer:
            rescure_table.to_excel(writer, sheet_name='Rettungswertung')
            wave_table.to_excel(writer, sheet_name='Wellenwertung')
            if exportDetail.get():
                rescure_table_detail.to_excel(writer, sheet_name='Detail Rettungswertung')
                wave_table_detail.to_excel(writer, sheet_name='Detail Wellenwertung')
            sort_table.to_excel(writer, sheet_name='Urkunden Druck')
        logger.warning(f'✔️  EXPORT ERFOLGREICH!\nGespeichert unter: {file_path}')
        print(f'✔️  EXPORT ERFOLGREICH!\nGespeichert unter: {file_path}')

        file_year = datetime.datetime.fromtimestamp(creation_date(file)).year
        current_year = datetime.date.today().year
        if file_year != current_year:
            logger.error(f'❗HAST DU DIE RICHTIGE DATEI GEWÄHLT?\nDie Datei ist aus dem Jahr {file_year}')

    except Exception as e:
        logger.error('❌ WURDE EINE EXCEL EXPORT AUS JAUSWERTUNG AUSGEWÄHLT?')
        logger.error(f'Fehler: {e}')


def generate_competition_preperation():
    pass


###################
# Export SETTINGS #
###################

mainLabel_tab1 = tk.Label(master=frame_Title, text='Gesamtauswertung für Wellenwettkampf erstellen', font=('Verdana Bold', 15))
mainLabel_tab1.pack(side='top', padx='5', pady='5')

# Input Folder
inputLabel = tk.Label(master=frame_Input, text='JAuswertung Export', width=20)
inputLabel.pack(side='left', padx='5', pady='5')

inputEntry = tk.Entry(master=frame_Input, textvariable=inputFilePath)
inputEntry.pack(side='left', padx='5', pady='5', fill='x', expand=True)

btnInputFind = tk.Button(master=frame_Input, text='Datei auswählen', command=getInputFilePath)
btnInputFind.pack(side='right', padx='5', pady='5')

# Output Folder
outputLabel = tk.Label(master=frame_Output, text='Ausgabe Ordner', width=20)
outputLabel.pack(side='left', padx='5', pady='5')

outputEntry = tk.Entry(master=frame_Output, textvariable=outputPath)
outputEntry.pack(side='left', padx='5', pady='5', fill='x', expand=True)

btnOutputFind = tk.Button(master=frame_Output, text='Ordner auswählen', command=getOutputFolderPath)
btnOutputFind.pack(side='right', padx='5', pady='5')

# Output File Name
nameLabel = tk.Label(master=frame_OutputName, text='Export Name (Option)', width=20)
nameLabel.pack(side='left', padx='5', pady='5')

nameEntry = tk.Entry(master=frame_OutputName, textvariable=customName)
nameEntry.pack(side='left', padx='5', pady='5', fill='x', expand=True)

nameLabelAddition = tk.Label(master=frame_OutputName, text='_Auswertung.xlsx', width=20)
nameLabelAddition.pack(side='left', pady='5')

exportDetailCb = tk.Checkbutton(master=frame_DetailExport, variable=exportDetail, text='Detail Export', textvariable='Detail Export')
exportDetailCb.pack(side='right', padx='5', pady='5')

# START
btnStart = tk.Button(master=frame_Start, text='START', command=generate_overall_evaluation, width=60)
btnStart.pack(side='bottom', padx='5', pady='5', fill='x')

gui.mainloop()
