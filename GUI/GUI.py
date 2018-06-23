# Import libraries and modules
from PyQt4 import QtGui, QtCore
import sys
import os
import pandas as pd
import xlrd
import numpy as np
import getpass

class MainWindow(QtGui.QMainWindow):

    def __init__(self, parent=None):
        
        super(MainWindow, self).__init__(parent)
        self.resize(600, 500)
        self.setWindowTitle("GUI - KTmfk")
        self.setWindowIcon(QtGui.QIcon("KTmfk-Logo.png"))

        self.setStyleSheet("background-image: url(BACKGROUND.png); background-attachment: fixed")

        openEditor = QtGui.QAction("&Edit", self)
        openEditor.setShortcut("Ctrl+E")
        openEditor.setStatusTip("Open Edit")

        extractAction = QtGui.QAction("&Quit", self)
        extractAction.setShortcut("Ctrl+Q")
        extractAction.setStatusTip("Leave The App")
        extractAction.triggered.connect(self.close_application)

        styleGroup = QtGui.QActionGroup(self, exclusive=True)
        
        plastiqueStyle = QtGui.QAction('Plastique', self ,checkable=True)
        plastiqueStyle.triggered.connect(lambda: self.style_choice('Plastique'))

        windowsStyle = QtGui.QAction('Windows', self ,checkable=True)
        windowsStyle.triggered.connect(lambda: self.style_choice('Windows'))

        windowsXPStyle = QtGui.QAction('WindowsXP', self ,checkable=True)
        windowsXPStyle.triggered.connect(lambda: self.style_choice('WindowsXP'))

        windowsVistaStyle = QtGui.QAction('WindowsVista', self ,checkable=True)
        windowsVistaStyle.triggered.connect(lambda: self.style_choice('WindowsVista'))

        motifStyle = QtGui.QAction('Motif', self ,checkable=True)
        motifStyle.triggered.connect(lambda: self.style_choice('Motif'))

        cdeStyle = QtGui.QAction('CDE', self ,checkable=True)
        cdeStyle.triggered.connect(lambda: self.style_choice('CDE'))

        cleanlooksStyle = QtGui.QAction('Cleanlooks', self ,checkable=True)
        cleanlooksStyle.triggered.connect(lambda: self.style_choice('Cleanlooks'))

        self.statusBar()

        mainMenu = self.menuBar()
        
        fileMenu = mainMenu.addMenu("&File")
        fileMenu.addAction(extractAction)
        
        editorMenu = mainMenu.addMenu("&Edit")

        viewMenu = mainMenu.addMenu("&View")
        styleMenu = viewMenu.addMenu("Display Theme")
        style_1 = styleGroup.addAction(plastiqueStyle)
        styleMenu.addAction(style_1)
        style_2 = styleGroup.addAction(windowsStyle)
        styleMenu.addAction(style_2)
        style_3 = styleGroup.addAction(windowsXPStyle)
        styleMenu.addAction(style_3)
        style_4 = styleGroup.addAction(windowsVistaStyle)
        styleMenu.addAction(style_4)
        style_5 = styleGroup.addAction(motifStyle)
        styleMenu.addAction(style_5)
        style_6 = styleGroup.addAction(cdeStyle)
        styleMenu.addAction(style_6)
        style_7 = styleGroup.addAction(cleanlooksStyle)
        styleMenu.addAction(style_7)

        aboutMenu = mainMenu.addMenu("&Help")
        aboutMenu = mainMenu.addMenu("&About")

        self.MainWindow_home()

    def MainWindow_home(self):

        btn_1 = QtGui.QPushButton("Quit", self)
        btn_1.move(260, 450)
        btn_1.clicked.connect(self.close_application)

        LabelDescription_1 = QtGui.QLabel("Sach- oder Zeichnungsnummer:", self)
        LabelDescription_1.resize(250, 20)
        LabelDescription_1.move(200, 200)
        LabelDescription_1.setStyleSheet("font: bold 10pt")
               
        self.textField_1 = QtGui.QLineEdit(self)
        self.textField_1.resize(200, 30)
        self.textField_1.move(200, 225)

        LabelDescription_2 = QtGui.QLabel("Benennung:", self)
        LabelDescription_2.resize(150, 20)
        LabelDescription_2.move(50, 280)
        LabelDescription_2.setStyleSheet("font: bold 8pt")

        self.textField_2 = QtGui.QLineEdit(self)
        self.textField_2.resize(170, 30)
        self.textField_2.move(50, 305)

        LabelDescription_3 = QtGui.QLabel("Menge:", self)
        LabelDescription_3.resize(50, 20)
        LabelDescription_3.move(255, 280)
        LabelDescription_3.setStyleSheet("font: bold 8pt")

        self.textField_3 = QtGui.QLineEdit(self)
        self.textField_3.resize(80, 30)
        self.textField_3.move(255, 305)

        LabelDescription_4 = QtGui.QLabel("Bezeichnung:", self)
        LabelDescription_4.resize(100, 20)
        LabelDescription_4.move(380, 280)
        LabelDescription_4.setStyleSheet("font: bold 8pt")

        self.textField_4 = QtGui.QLineEdit(self)
        self.textField_4.resize(150, 30)
        self.textField_4.move(380, 305)
        
        btn_4 = QtGui.QPushButton("Enter Item", self)
        btn_4.move(410, 228)
        btn_4.resize(btn_4.minimumSizeHint())
        btn_4.clicked.connect(self.getZeichnungsnummer)

        logo = QtGui.QLabel(self)
        logo.setGeometry(QtCore.QRect(100, 25, 420, 115))
        logo.setPixmap(QtGui.QPixmap("KTmfk.png"))

        self.LabelDescription_5 = QtGui.QLabel(" ", self)
        self.LabelDescription_5.resize(300, 20)
        self.LabelDescription_5.move(202, 255)
        self.LabelDescription_5.setStyleSheet("font: 8pt")

        self.btn_2 = QtGui.QPushButton("Open File", self)
        self.btn_2.move(50, 355)
        self.btn_2.clicked.connect(self.open_new_dialog)

        btn_3 = QtGui.QPushButton("Open Folder",self)
        btn_3.move(50, 390)
        btn_3.clicked.connect(self.folder_open)

        global progress_2
        progress_2 = QtGui.QProgressBar(self)
        progress_2.setGeometry(200, 358, 250, 20)
        
        global progress_3
        progress_3 = QtGui.QProgressBar(self)
        progress_3.setGeometry(200, 394, 250, 20)
        
        self.show()
    
    def getZeichnungsnummer(self):
        
        text, ok = QtGui.QInputDialog.getText(self, 'Zeichnungsnummer Input Dialog', 'Enter "Sachnummer":')

        if ok:
            
            self.textField_1.setText(str(text))
            self.textField_1.setStyleSheet("font: bold")
            
            PythonFile_PATH = r'%s' % os.getcwd().replace('\\','/')
            df = pd.read_csv(str(PythonFile_PATH) + '/AllData.csv', encoding = "ISO-8859-1")

            selected_rows = df.loc[df['Sach- oder Zeichnungsnummer,\r\nNorm-Kurzbezeichnung'] == str(text)]

            if selected_rows.shape[0] == 0:
                
                self.LabelDescription_5.setStyleSheet('QLabel {color: red;}')
                self.LabelDescription_5.setText('"Sachnummer" is wrong or it does not exist in Database!')
                self.textField_2.setText(" ")
                self.textField_3.setText(" ")
                self.textField_4.setText(" ")
                
            
            elif selected_rows.shape[0] == 1:

                self.LabelDescription_5.setText(" ")
                Benennung_list = list(selected_rows['Benennung'])
                Menge_list = list(selected_rows['Menge'])
                Bezeichnung_list = list(selected_rows['Bezeichnung'])
                
                self.textField_2.setText(str(Benennung_list[0]).replace('ue', 'ü').replace('ae', 'ä').replace('oe', 'ö'))
                self.textField_2.setStyleSheet("font: bold")

                self.textField_3.setText(str(Menge_list[0]).replace('ue', 'ü').replace('ae', 'ä').replace('oe', 'ö'))
                self.textField_3.setStyleSheet("font: bold")

                self.textField_4.setText(str(Bezeichnung_list[0]).replace('ue', 'ü').replace('ae', 'ä').replace('oe', 'ö'))
                self.textField_4.setStyleSheet("font: bold")

            elif selected_rows.shape[0] >= 1:

                Benennung_list = list(selected_rows['Benennung'])
                Menge_list = list(selected_rows['Menge'])
                Bezeichnung_list = list(selected_rows['Bezeichnung'])

                self.LabelDescription_5.setStyleSheet('QLabel {color: blue;}')
                self.LabelDescription_5.setText('There is more than one item for this "Sachnummer"!')

                self.textField_2.setText(str(Benennung_list[0]).replace('ue', 'ü').replace('ae', 'ä').replace('oe', 'ö'))
                self.textField_2.setStyleSheet("font: bold")

                self.textField_3.setText(str(Menge_list[0]).replace('ue', 'ü').replace('ae', 'ä').replace('oe', 'ö'))
                self.textField_3.setStyleSheet("font: bold")

                self.textField_4.setText(str(Bezeichnung_list[0]).replace('ue', 'ü').replace('ae', 'ä').replace('oe', 'ö'))
                self.textField_4.setStyleSheet("font: bold")                

    def open_new_dialog(self):
        self.newDialog = TableWindow(self)
    
    def file_open(self):
            
        try:
            userName = getpass.getuser()
            file_path = QtGui.QFileDialog.getOpenFileName(self, 'Open File', 'C:/Users/'+ str(userName) +'/Desktop', ' Excel Files (*.xlsx *.xls)')
            return(file_path)
        
        except:
            
            try:
                file_path = QtGui.QFileDialog.getOpenFileName(self, 'Open File', '', ' Excel Files (*.xlsx *.xls)')
                return(file_path)
            except:
                pass
            
            pass

    def folder_open(self):
        
        try:
            folder_path = QtGui.QFileDialog.getExistingDirectory(self, 'Open Folder').replace('\\','/')
            MainWindow.open_and_preprocessing_folder(self, folder_path)
        except:
            pass

    def open_and_preprocessing_folder(self, folder_path):
        counter_1 = 0
        number_of_files = len(os.listdir(folder_path))
        for filename in os.listdir(folder_path):

            if filename.endswith(".xls") or filename.endswith(".xlsx"):
            
                counter_1 += 1
                xlsx = xlrd.open_workbook(folder_path + '/' + filename, on_demand=True)
                sheet_names = xlsx.sheet_names()
                
                counter_2 = 0
                for sheet in sheet_names:
                    counter_2 += 1
                    df = pd.read_excel(folder_path + '/' + filename, encoding = "ISO-8859-1", sheet_name=sheet)
                    
                    table_1 = df
                    for col in df:
                        if 'Unnamed' in col:
                            table_1.drop([col], axis=1, inplace=True)
                        
                    table_1_columns_name = table_1.columns
                    table_2 = pd.DataFrame(columns=table_1_columns_name)
                    i = -1
                    for index, row in df.iterrows():
                        if np.count_nonzero(pd.isna(row)) != len(row):
                            i += 1
                            table_2.loc[i] = row
                            
                    if counter_2 == 1:
                        
                        for row in range(table_2.shape[0]):
                            for col in table_2.columns:
                                if table_2.at[row,col] == 'Titel, Zusätzlicher Titel':
                                    Titel_location = [row + 1, col]
                                    Titel = table_2.at[Titel_location[0], Titel_location[1]]
                                    break
                                    
                        for row in range(table_2.shape[0]):
                            for col in table_2.columns:
                                if table_2.at[row,col] == 'Zeichnungsnummer':
                                    Zeichnungsnummer_location = [row + 1, col]
                                    Zeichnungsnummer = table_2.at[Zeichnungsnummer_location[0], Zeichnungsnummer_location[1]]
                                    break
                        
                        title_row = [0,
                                     1.0,
                                     'St.',
                                     str(Titel),
                                     str(Zeichnungsnummer),
                                     np.nan,
                                     np.nan,
                                     np.nan]

                    for row in range(table_2.shape[0]):
                        for col in table_2.columns:
                            if table_2.at[row,col] == 'Gesetzlicher Eigentümer':
                                end_of_list_index = row
                                break        
                    table_2_columns_name = table_2.columns
                    table_3 = pd.DataFrame(columns=table_2_columns_name)

                    if counter_2 == 1:
                        table_3.loc[0] = title_row

                    for i in range(end_of_list_index):
                        table_3.loc[i+1] = table_2.loc[i]

                    if counter_2 == 1:
                        final_table = pd.DataFrame(columns=table_2_columns_name)

                    final_table = final_table.append(table_3)

                Bezeichnungsliste = ['(Unter-) baugruppe']
                for Sachnummer in final_table['Sach- oder Zeichnungsnummer,\nNorm-Kurzbezeichnung'][1:]:
                    Bezeichnungsliste.append(TableWindow.Bezeichnung(self, Sachnummer))

                final_table['Bezeichnung'] = Bezeichnungsliste

                selected_rows_baugruppe = final_table.loc[final_table['Bezeichnung'] == '(Unter-) baugruppe']
                table_baugruppe = selected_rows_baugruppe[['Sach- oder Zeichnungsnummer,\nNorm-Kurzbezeichnung', 'Benennung', 'Menge']]

                selected_rows_bauteil = final_table.loc[final_table['Bezeichnung'] == 'Bauteil']
                table_bauteil = selected_rows_bauteil[['Sach- oder Zeichnungsnummer,\nNorm-Kurzbezeichnung', 'Benennung','Menge']]

                selected_rows_zukaufteile = final_table.loc[final_table['Bezeichnung'] == 'Zukaufteile/ Normteile']
                table_zukaufteile = selected_rows_zukaufteile[['Sach- oder Zeichnungsnummer,\nNorm-Kurzbezeichnung', 'Benennung','Menge']]

                foldername = folder_path.split('/')[-1]
                directory_xlsx = folder_path[: - len(foldername)] + 'XLSX'
                
                if not os.path.exists(directory_xlsx):
                    os.makedirs(directory_xlsx)

                writer = pd.ExcelWriter(directory_xlsx + '/' + filename, engine='xlsxwriter')

                table_baugruppe.to_excel(writer, sheet_name= '(Unter-) baugruppe', index=False, encoding = "ISO-8859-1")
                table_bauteil.to_excel(writer, sheet_name= 'Bauteil', index=False, encoding = "ISO-8859-1")
                table_zukaufteile.to_excel(writer, sheet_name= 'Zukaufteil&Normteil', index=False, encoding = "ISO-8859-1")

                writer.save()

        self.completed = 0
        while self.completed < 100:
            self.completed += 0.0002
            progress_3.setValue(self.completed)

        QtGui.QMessageBox.information(self, "Information", "This Folder: XLSX has been added to this path:\n{0}\n{1} Files has been created. ".format(folder_path[: - len(foldername)], counter_1))

    def style_choice(self, text):
        QtGui.QApplication.setStyle(QtGui.QStyleFactory.create(text))        

    def close_application(self):
        choice = QtGui.QMessageBox.question(self, 'Exit',
                                            "Do you want to close the application?",
                                            QtGui.QMessageBox.Yes | QtGui.QMessageBox.No)
        
        if choice == QtGui.QMessageBox.Yes:
                sys.exit()
        else:
            pass     

# Define a class for tabular view of data1; It displays information in table form and also it enables you to store data in excel sheet.
class TableWindow(QtGui.QMainWindow):
    
    def __init__(self, parent=None):
        
        super(TableWindow, self).__init__(parent)
        
        self.resize(500, 730)
        
        self.setWindowTitle("GUI - KTmfk")
        self.setWindowIcon(QtGui.QIcon("KTmfk-Logo.png"))

        extractAction = QtGui.QAction("&Quit", self)
        extractAction.setShortcut("Ctrl+Q")
        extractAction.setStatusTip("Leave The App")
        extractAction.triggered.connect(self.close_application)

        styleGroup = QtGui.QActionGroup(self, exclusive=True)
        
        plastiqueStyle = QtGui.QAction('Plastique', self ,checkable=True)
        plastiqueStyle.triggered.connect(lambda: self.style_choice('Plastique'))

        windowsStyle = QtGui.QAction('Windows', self ,checkable=True)
        windowsStyle.triggered.connect(lambda: self.style_choice('Windows'))

        windowsXPStyle = QtGui.QAction('WindowsXP', self ,checkable=True)
        windowsXPStyle.triggered.connect(lambda: self.style_choice('WindowsXP'))

        windowsVistaStyle = QtGui.QAction('WindowsVista', self ,checkable=True)
        windowsVistaStyle.triggered.connect(lambda: self.style_choice('WindowsVista'))

        motifStyle = QtGui.QAction('Motif', self ,checkable=True)
        motifStyle.triggered.connect(lambda: self.style_choice('Motif'))

        cdeStyle = QtGui.QAction('CDE', self ,checkable=True)
        cdeStyle.triggered.connect(lambda: self.style_choice('CDE'))

        cleanlooksStyle = QtGui.QAction('Cleanlooks', self ,checkable=True)
        cleanlooksStyle.triggered.connect(lambda: self.style_choice('Cleanlooks'))

        self.statusBar()

        mainMenu = self.menuBar()
        
        fileMenu = mainMenu.addMenu("&File")
        fileMenu.addAction(extractAction)
        
        editorMenu = mainMenu.addMenu("&Edit")

        viewMenu = mainMenu.addMenu("&View")
        styleMenu = viewMenu.addMenu("Display Theme")
        style_1 = styleGroup.addAction(plastiqueStyle)
        styleMenu.addAction(style_1)
        style_2 = styleGroup.addAction(windowsStyle)
        styleMenu.addAction(style_2)
        style_3 = styleGroup.addAction(windowsXPStyle)
        styleMenu.addAction(style_3)
        style_4 = styleGroup.addAction(windowsVistaStyle)
        styleMenu.addAction(style_4)
        style_5 = styleGroup.addAction(motifStyle)
        styleMenu.addAction(style_5)
        style_6 = styleGroup.addAction(cdeStyle)
        styleMenu.addAction(style_6)
        style_7 = styleGroup.addAction(cleanlooksStyle)
        styleMenu.addAction(style_7)

        aboutMenu = mainMenu.addMenu("&Help")
        aboutMenu = mainMenu.addMenu("&About")

        try:
            self.TableWindow_home()
        except:
            pass
        
    def TableWindow_home(self):
            
        btn_1 = QtGui.QPushButton("Quit", self)
        btn_1.move(200, 680)
        btn_1.clicked.connect(self.close_application)

        btn_2_ = QtGui.QPushButton("Save", self)
        btn_2_.move(200, 635)
        btn_2_.clicked.connect(lambda: self.open_and_preprocessing_file(filePath))

        filePath = r'%s' % MainWindow.file_open(self).replace('\\','/')

        table_stueckliste = TableWindow.viewTables(self, filePath)

        LabelTable_baugruppe = QtGui.QLabel("(Unter-) baugruppe", self)
        LabelTable_baugruppe.resize(200, 20)
        LabelTable_baugruppe.move(170, 45)
        LabelTable_baugruppe.setStyleSheet("font: bold 11pt Arial Black")

        selected_rows_baugruppe = table_stueckliste.loc[table_stueckliste['Bezeichnung'] == '(Unter-) baugruppe']
        table_baugruppe = selected_rows_baugruppe[['Sach- oder Zeichnungsnummer,\nNorm-Kurzbezeichnung', 'Benennung', 'Menge']]

        data={}
        data['Sachnummer'] = [str(i) for i in list(table_baugruppe['Sach- oder Zeichnungsnummer,\nNorm-Kurzbezeichnung'])]
        data['Benennung'] = [str(i) for i in list(table_baugruppe['Benennung'])]
        data['Menge'] = [str(i) for i in list(table_baugruppe['Menge'])]

        #Create Empty m x n Table; m = row_number, n = column_number
        table = QtGui.QTableWidget(self)
        table.setRowCount(int(table_baugruppe.shape[0]))
        table.setColumnCount(3)
        
        #Enter data onto Table
        horHeaders = []
        for n, key in enumerate(data.keys()):
            horHeaders.append(key)
            for m, item in enumerate(data[key]):
                newitem = QtGui.QTableWidgetItem(item)
                table.setItem(m, n, newitem)
        
        #Add Header
        table.setHorizontalHeaderLabels(horHeaders)        

        #Adjust size of Table
        table.resizeColumnsToContents()
        table.resizeRowsToContents()
        table.resize(405, 150)
        table.move(50, 70)

        LabelTable_bauteil = QtGui.QLabel("Bauteil", self)
        LabelTable_bauteil.resize(200, 20)
        LabelTable_bauteil.move(220, 235)
        LabelTable_bauteil.setStyleSheet("font: bold 11pt Arial Black")

        selected_rows_bauteil = table_stueckliste.loc[table_stueckliste['Bezeichnung'] == 'Bauteil']
        table_bauteil = selected_rows_bauteil[['Sach- oder Zeichnungsnummer,\nNorm-Kurzbezeichnung', 'Benennung','Menge']]

        data={}
        data['Sachnummer'] = [str(i) for i in list(table_bauteil['Sach- oder Zeichnungsnummer,\nNorm-Kurzbezeichnung'])]
        data['Benennung'] = [str(i) for i in list(table_bauteil['Benennung'])]
        data['Menge'] = [str(i) for i in list(table_bauteil['Menge'])]

        #Create Empty m x n Table; m = row_number, n = column_number
        table = QtGui.QTableWidget(self)
        table.setRowCount(int(table_bauteil.shape[0]))
        table.setColumnCount(3)
        
        #Enter data onto Table
        horHeaders = []
        for n, key in enumerate(data.keys()):
            horHeaders.append(key)
            for m, item in enumerate(data[key]):
                newitem = QtGui.QTableWidgetItem(item)
                table.setItem(m, n, newitem)
        
        #Add Header
        table.setHorizontalHeaderLabels(horHeaders)        

        #Adjust size of Table
        table.resizeColumnsToContents()
        table.resizeRowsToContents()
        table.resize(405, 150)
        table.move(50, 260)

        LabelTable_bauteil = QtGui.QLabel("Zukaufteile/ Normteile", self)
        LabelTable_bauteil.resize(200, 20)
        LabelTable_bauteil.move(170, 430)
        LabelTable_bauteil.setStyleSheet("font: bold 11pt Arial Black")

        selected_rows_zukaufteile = table_stueckliste.loc[table_stueckliste['Bezeichnung'] == 'Zukaufteile/ Normteile']
        table_zukaufteile = selected_rows_zukaufteile[['Sach- oder Zeichnungsnummer,\nNorm-Kurzbezeichnung', 'Benennung','Menge']]

        data={}
        data['Sachnummer'] = [str(i) for i in list(table_zukaufteile['Sach- oder Zeichnungsnummer,\nNorm-Kurzbezeichnung'])]
        data['Benennung'] = [str(i) for i in list(table_zukaufteile['Benennung'])]
        data['Menge'] = [str(i) for i in list(table_zukaufteile['Menge'])]

        #Create Empty m x n Table; m = row_number, n = column_number
        table = QtGui.QTableWidget(self)
        table.setRowCount(int(table_zukaufteile.shape[0]))
        table.setColumnCount(3)
        
        #Enter data onto Table
        horHeaders = []
        for n, key in enumerate(data.keys()):
            horHeaders.append(key)
            for m, item in enumerate(data[key]):
                newitem = QtGui.QTableWidgetItem(item)
                table.setItem(m, n, newitem)
        
        #Add Header
        table.setHorizontalHeaderLabels(horHeaders)        

        #Adjust size of Table
        table.resizeColumnsToContents()
        table.resizeRowsToContents()
        table.resize(405, 150)
        table.move(50, 455)

        self.completed = 0
        while self.completed < 100:
            self.completed += 0.0004
            progress_2.setValue(self.completed)
            
        if filePath:
            self.show()

    def Bezeichnung(self, Sachnummer):
        
        if pd.isna(Sachnummer):
            output = '_'
            
        else:
            
            list_1 = str(Sachnummer).split('-')
            list_2 = str(Sachnummer).split('.')
            if len(list_1) >= 3:
                if list_1[0] == 'TSPS':
                    if list_1[-1].isdigit() and len(list_1[-1]) == 2:
                        if sum([int(digit) for digit in list(list_1[-1])]) == 0:
                            output = '(Unter-) baugruppe'
                        elif sum([int(digit) for digit in list(list_1[-1])]) >= 1:
                            output = 'Bauteil'
                        else:
                            output = 'Zukaufteile/ Normteile'
                    elif not list_1[-1].isdigit():
                        output = 'Bauteil'
                    else:
                        output = 'Zukaufteile/ Normteile'
                else:
                    output = 'Zukaufteile/ Normteile'

            elif list_1[0] == 'TSPS':
                output = 'Bauteil'

            elif len(list_2) >= 3:
                if list_2[0].isdigit() and len(list_2[0]) == 2:
                    if list_2[-1].isdigit() and len(list_2[-1]) == 3:
                        if sum([int(digit) for digit in list(list_2[-1])]) == 0:
                            output = '(Unter-) baugruppe'
                        elif sum([int(digit) for digit in list(list_2[-1])]) >= 1:
                            output = 'Bauteil'
                        else:
                            output = 'Zukaufteile/ Normteile'
                    else:
                        output = 'Zukaufteile/ Normteile'
                else:
                    output = 'Zukaufteile/ Normteile'

            else:
                output = 'Zukaufteile/ Normteile'
            
        return(output)

    def viewTables(self, file_path):
        
        xlsx = xlrd.open_workbook(file_path, on_demand=True)
        sheet_names = xlsx.sheet_names()
        counter_1 = 0
        for sheet in sheet_names:
            counter_1 += 1
            df = pd.read_excel(file_path, encoding = "ISO-8859-1", sheet_name=sheet)

            table_1 = df
            for col in df:
                if 'Unnamed' in col:
                    table_1.drop([col], axis=1, inplace=True)

            table_1_columns_name = table_1.columns
            table_2 = pd.DataFrame(columns=table_1_columns_name)
            i = -1
            for index, row in df.iterrows():
                if np.count_nonzero(pd.isna(row)) != len(row):
                    i += 1
                    table_2.loc[i] = row

            if counter_1 == 1:

                for row in range(table_2.shape[0]):
                    for col in table_2.columns:
                        if table_2.at[row,col] == 'Titel, Zusätzlicher Titel':
                            Titel_location = [row + 1, col]
                            Titel = table_2.at[Titel_location[0], Titel_location[1]]
                            break

                for row in range(table_2.shape[0]):
                    for col in table_2.columns:
                        if table_2.at[row,col] == 'Zeichnungsnummer':
                            Zeichnungsnummer_location = [row + 1, col]
                            Zeichnungsnummer = table_2.at[Zeichnungsnummer_location[0], Zeichnungsnummer_location[1]]
                            break

                title_row = [0,
                             1.0,
                             'St.',
                             str(Titel),
                             str(Zeichnungsnummer),
                             np.nan,
                             np.nan,
                             np.nan]

            for row in range(table_2.shape[0]):
                for col in table_2.columns:
                    if table_2.at[row,col] == 'Gesetzlicher Eigentümer':
                        end_of_list_index = row
                        break        
            table_2_columns_name = table_2.columns
            table_3 = pd.DataFrame(columns=table_2_columns_name)

            if counter_1 == 1:
                table_3.loc[0] = title_row

            for i in range(end_of_list_index):
                table_3.loc[i+1] = table_2.loc[i]

            if counter_1 == 1:
                final_table = pd.DataFrame(columns=table_2_columns_name)

            final_table = final_table.append(table_3)
                
        Bezeichnungsliste = ['(Unter-) baugruppe']
        for Sachnummer in final_table['Sach- oder Zeichnungsnummer,\nNorm-Kurzbezeichnung'][1:]:
            Bezeichnungsliste.append(TableWindow.Bezeichnung(self, Sachnummer))

        final_table['Bezeichnung'] = Bezeichnungsliste
       
        return(final_table)

    def open_and_preprocessing_file(self, file_path):
        
        xlsx = xlrd.open_workbook(file_path, on_demand=True)
        sheet_names = xlsx.sheet_names()
        counter_1 = 0
        for sheet in sheet_names:
            counter_1 += 1
            df = pd.read_excel(file_path, encoding = "ISO-8859-1", sheet_name=sheet)
            
            table_1 = df
            for col in df:
                if 'Unnamed' in col:
                    table_1.drop([col], axis=1, inplace=True)
                
            table_1_columns_name = table_1.columns
            table_2 = pd.DataFrame(columns=table_1_columns_name)
            i = -1
            for index, row in df.iterrows():
                if np.count_nonzero(pd.isna(row)) != len(row):
                    i += 1
                    table_2.loc[i] = row
                    
            if counter_1 == 1:
                
                for row in range(table_2.shape[0]):
                    for col in table_2.columns:
                        if table_2.at[row,col] == 'Titel, Zusätzlicher Titel':
                            Titel_location = [row + 1, col]
                            Titel = table_2.at[Titel_location[0], Titel_location[1]]
                            break
                            
                for row in range(table_2.shape[0]):
                    for col in table_2.columns:
                        if table_2.at[row,col] == 'Zeichnungsnummer':
                            Zeichnungsnummer_location = [row + 1, col]
                            Zeichnungsnummer = table_2.at[Zeichnungsnummer_location[0], Zeichnungsnummer_location[1]]
                            break
                
                title_row = [0,
                             1.0,
                             'St.',
                             str(Titel),
                             str(Zeichnungsnummer),
                             np.nan,
                             np.nan,
                             np.nan]

            for row in range(table_2.shape[0]):
                for col in table_2.columns:
                    if table_2.at[row,col] == 'Gesetzlicher Eigentümer':
                        end_of_list_index = row
                        break        
            table_2_columns_name = table_2.columns
            table_3 = pd.DataFrame(columns=table_2_columns_name)

            if counter_1 == 1:
                table_3.loc[0] = title_row

            for i in range(end_of_list_index):
                table_3.loc[i+1] = table_2.loc[i]

            if counter_1 == 1:
                final_table = pd.DataFrame(columns=table_2_columns_name)

            final_table = final_table.append(table_3)

        Bezeichnungsliste = ['(Unter-) baugruppe']
        for Sachnummer in final_table['Sach- oder Zeichnungsnummer,\nNorm-Kurzbezeichnung'][1:]:
            Bezeichnungsliste.append(TableWindow.Bezeichnung(self, Sachnummer))

        final_table['Bezeichnung'] = Bezeichnungsliste

        selected_rows_baugruppe = final_table.loc[final_table['Bezeichnung'] == '(Unter-) baugruppe']
        table_baugruppe = selected_rows_baugruppe[['Sach- oder Zeichnungsnummer,\nNorm-Kurzbezeichnung', 'Benennung', 'Menge']]

        selected_rows_bauteil = final_table.loc[final_table['Bezeichnung'] == 'Bauteil']
        table_bauteil = selected_rows_bauteil[['Sach- oder Zeichnungsnummer,\nNorm-Kurzbezeichnung', 'Benennung','Menge']]

        selected_rows_zukaufteile = final_table.loc[final_table['Bezeichnung'] == 'Zukaufteile/ Normteile']
        table_zukaufteile = selected_rows_zukaufteile[['Sach- oder Zeichnungsnummer,\nNorm-Kurzbezeichnung', 'Benennung','Menge']]

        filename = file_path.split('/')[-1]
        directory_xlsx = file_path[: - len(filename)] + 'XLSX'
        
        if not os.path.exists(directory_xlsx):
            os.makedirs(directory_xlsx)

        writer = pd.ExcelWriter(directory_xlsx + '/' + filename, engine='xlsxwriter')

        table_baugruppe.to_excel(writer, sheet_name= '(Unter-) baugruppe', index=False, encoding = "ISO-8859-1")
        table_bauteil.to_excel(writer, sheet_name= 'Bauteil', index=False, encoding = "ISO-8859-1")
        table_zukaufteile.to_excel(writer, sheet_name= 'Zukaufteil&Normteil', index=False, encoding = "ISO-8859-1")

        writer.save()

        QtGui.QMessageBox.information(self, "Information", "This File: {0}\nhas been added to this path:\n{1} ".format(filename, directory_xlsx))

    def style_choice(self, text):
        QtGui.QApplication.setStyle(QtGui.QStyleFactory.create(text))

    def close_application(self):
        choice = QtGui.QMessageBox.question(self, 'Exit',
                                            "Do you want to close the application?",
                                            QtGui.QMessageBox.Yes | QtGui.QMessageBox.No)
        
        if choice == QtGui.QMessageBox.Yes:
                sys.exit()
        else:
            pass

def run():
    
    app = QtGui.QApplication(sys.argv)
    app.setStyle(QtGui.QStyleFactory.create('Plastique'))
    GUI = MainWindow()
    sys.exit(app.exec_())

run()
