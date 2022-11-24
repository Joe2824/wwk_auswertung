#!/usr/bin/env python
# -*- coding: utf-8 -*-

import os
import logging
import platform
import datetime

import pandas as pd
from pathlib import Path
import tkinter as tk
from tkinter import ttk
from tkinter import filedialog

import overall_evaluation
import sort_certificate
import competition_preperation



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


class App(tk.Tk):
    def __init__(self):
        super().__init__()

        # configure the root window
        self.geometry('600x400')
        self.minsize(555, 320)
        self.title('Wellenwettkampf Auswertung')

        self.inputFilePath = tk.StringVar()
        self.outputPath = tk.StringVar()
        self.customName = tk.StringVar()
        self.exportDetail = tk.IntVar()

        self.tabControl = ttk.Notebook(master=self)
        self.tab1 = tk.Frame(self.tabControl)
        self.tab2 = tk.Frame(self.tabControl)

        self.frame_Title1 = tk.Frame(master=self.tab1)
        self.frame_Input1 = tk.Frame(master=self.tab1)

        self.frame_Title2 = tk.Frame(master=self.tab2)
        self.frame_Input2 = tk.Frame(master=self.tab2)
        self.frame_Output2 = tk.Frame(master=self.tab2)
        self.frame_OutputName2 = tk.Frame(master=self.tab2)
        self.frame_DetailExport2 = tk.Frame(master=self.tab2)
        self.frame_Start2 = tk.Frame(master=self.tab2)
        self.frame_Log = tk.Frame(master=self)

        self.frame_Title1.pack(fill='x')
        self.frame_Input1.pack(fill='x')

        self.frame_Title2.pack(fill='x')
        self.frame_Input2.pack(fill='x')
        self.frame_Output2.pack(fill='x')
        self.frame_OutputName2.pack(fill='x')
        self.frame_DetailExport2.pack(fill='x')
        self.frame_Start2.pack(fill='x')

        self.layout_elements()
        self.logger_frame()

    def getInputFilePath(self):
        file_selected = filedialog.askopenfilename()
        self.inputFilePath.set(file_selected)

    def getOutputFolderPath(self):
        folder_selected = filedialog.askdirectory()
        self.outputPath.set(folder_selected)

    def creation_date(self, path_to_file):
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

    def generate_overall_evaluation(self):
        file = self.inputFilePath.get()
        name = self.customName.get()

        # Check if path exist
        if not os.path.exists(file):
            self.logger.error('❌ DATEI EXISTIERT NICHT')
            print('❌ DATEI EXISTIERT NICHT')
            return

        # Check for Excel file
        if not file.endswith(('.xlsx', '.xls')):
            self.logger.error('❌ KEINE EXCEL DATEI AUSGEWÄHLT')
            print('❌ KEINE EXCEL DATEI AUSGEWÄHLT')
            return

        try:
            wave_table, rescure_table, wave_table_detail, rescure_table_detail = overall_evaluation.calculate(file)
            sort_table = sort_certificate.sort(file)

            op = self.outputPath.get() if self.outputPath.get() else Path().resolve()
            fn = f'{name}_Auswertung.xlsx' if name else 'Auswertung.xlsx'
            file_path = os.path.join(op, fn)

            with pd.ExcelWriter(file_path, engine='xlsxwriter') as writer:
                rescure_table.to_excel(writer, sheet_name='Rettungswertung')
                wave_table.to_excel(writer, sheet_name='Wellenwertung')
                if self.exportDetail.get():
                    rescure_table_detail.to_excel(writer, sheet_name='Detail Rettungswertung')
                    wave_table_detail.to_excel(writer, sheet_name='Detail Wellenwertung')
                sort_table.to_excel(writer, sheet_name='Urkunden Druck')
            self.logger.warning(f'✔️  EXPORT ERFOLGREICH!\nGespeichert unter: {file_path}')
            print(f'✔️  EXPORT ERFOLGREICH!\nGespeichert unter: {file_path}')

            file_year = datetime.datetime.fromtimestamp(self.creation_date(file)).year
            current_year = datetime.date.today().year
            if file_year != current_year:
                self.logger.error(f'❗HAST DU DIE RICHTIGE DATEI GEWÄHLT?\nDie Datei ist aus dem Jahr {file_year}')

        except Exception as e:
            self.logger.error('❌ WURDE EINE EXCEL EXPORT AUS JAUSWERTUNG AUSGEWÄHLT?')
            self.logger.error(f'Fehler: {e}')

    def generate_competition_preperation(self):
        file = self.inputFilePath.get()

        # Check if path exist
        if not os.path.exists(file):
            self.logger.error('❌ DATEI EXISTIERT NICHT')
            print('❌ DATEI EXISTIERT NICHT')
            return

        # Check for Excel file
        if not file.endswith(('.csv')):
            self.logger.error('❌ KEINE CSV DATEI AUSGEWÄHLT')
            print('❌ KEINE CSV DATEI AUSGEWÄHLT')
            return

        try:
            sanitized_file = competition_preperation.sanitize(file)
            print(sanitized_file)

        except Exception as e:
            self.logger.error('❌ WURDE EIN CSV EXPORT AUS DEM ISC AUSGEWÄHLT?')
            self.logger.error(f'Fehler: {e}')

    def resize_layout(self, what):
        text = self.tabControl.tab(self.tabControl.select(), "text")
        #if text == 'Auswertung':
        #    self.geometry('600x400')
        #else:
        #    self.geometry('900x800')

    def layout_elements(self):
        self.tabControl.add(self.tab1, text='Vorbereitung')
        self.tabControl.add(self.tab2, text='Auswertung')
        self.tabControl.pack(expand=1, fill="both")
        self.tabControl.bind("<<NotebookTabChanged>>", self.resize_layout)

        ###################
        # VORBEREITUNG    #
        ###################
        mainLabel_tab1 = tk.Label(master=self.frame_Title1, text='Vorbereitung Wellenwettkampf', font=('Verdana Bold', 15))
        mainLabel_tab1.pack(side='top', padx='5', pady='5')

        # Input Folder
        inputLabel = tk.Label(master=self.frame_Input1, text='ISC Export', width=20)
        inputLabel.pack(side='left', padx='5', pady='5')

        inputEntry = tk.Entry(master=self.frame_Input1, textvariable=self.inputFilePath)
        inputEntry.pack(side='left', padx='5', pady='5', fill='x', expand=True)

        btnInputFind = tk.Button(master=self.frame_Input1, text='Datei auswählen', command=self.getInputFilePath)
        btnInputFind.pack(side='right', padx='5', pady='5')


        ###################
        # AUSWERTUNG      #
        ###################

        mainLabel_tab2 = tk.Label(master=self.frame_Title2, text='Gesamtauswertung für Wellenwettkampf erstellen', font=('Verdana Bold', 15))
        mainLabel_tab2.pack(side='top', padx='5', pady='5')

        # Input Folder
        inputLabel = tk.Label(master=self.frame_Input2, text='JAuswertung Export', width=20)
        inputLabel.pack(side='left', padx='5', pady='5')

        inputEntry = tk.Entry(master=self.frame_Input2, textvariable=self.inputFilePath)
        inputEntry.pack(side='left', padx='5', pady='5', fill='x', expand=True)

        btnInputFind = tk.Button(master=self.frame_Input2, text='Datei auswählen', command=self.getInputFilePath)
        btnInputFind.pack(side='right', padx='5', pady='5')

        # Output Folder
        outputLabel = tk.Label(master=self.frame_Output2, text='Ausgabe Ordner', width=20)
        outputLabel.pack(side='left', padx='5', pady='5')

        outputEntry = tk.Entry(master=self.frame_Output2, textvariable=self.outputPath)
        outputEntry.pack(side='left', padx='5', pady='5', fill='x', expand=True)

        btnOutputFind = tk.Button(master=self.frame_Output2, text='Ordner auswählen', command=self.getOutputFolderPath)
        btnOutputFind.pack(side='right', padx='5', pady='5')

        # Output File Name
        nameLabel = tk.Label(master=self.frame_OutputName2, text='Export Name (Option)', width=20)
        nameLabel.pack(side='left', padx='5', pady='5')

        nameEntry = tk.Entry(master=self.frame_OutputName2, textvariable=self.customName)
        nameEntry.pack(side='left', padx='5', pady='5', fill='x', expand=True)

        nameLabelAddition = tk.Label(master=self.frame_OutputName2, text='_Auswertung.xlsx', width=20)
        nameLabelAddition.pack(side='left', pady='5')

        exportDetailCb = tk.Checkbutton(master=self.frame_DetailExport2,
                                        variable=self.exportDetail, text='Detail Export', textvariable='Detail Export')
        exportDetailCb.pack(side='right', padx='5', pady='5')

        # START
        btnStart = tk.Button(master=self.frame_Start2, text='START', command=self.generate_overall_evaluation, width=60)
        btnStart.pack(side='bottom', padx='5', pady='5', fill='x')

    def logger_frame(self):
        self.frame_Log.pack(pady='6', fill='both', expand=True)
        self.logger_Scrollbar = tk.Scrollbar(master=self.frame_Log)
        self.logger_Element = tk.Text(master=self.frame_Log, state=tk.DISABLED, font='TkFixedFont')
        self.logger_Scrollbar.pack(side='right', fill='y')
        self.logger_Element.pack(side='left', fill='x', expand=True)
        self.logger_Scrollbar.config(command=self.logger_Element.yview)
        self.logger_Element.config(yscrollcommand=self.logger_Scrollbar.set)

        # Create textLogger
        self.text_handler = TextHandler(self.logger_Element)

        # Add the handler to logger
        self.logger = logging.getLogger()
        self.logger.addHandler(self.text_handler)


if __name__ == "__main__":
    app = App()
    s = ttk.Style()
    app.mainloop()
