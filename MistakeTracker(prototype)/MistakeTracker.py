
# # # # # # # # # # # # # # # # # # # # # # # # # #
# Name: Austin Kidwell
# Version: 1.0 (Prototype)
# Date: 11/23/2021
# # # # # # # # # # # # # # # # # # # # # # # # # #

from __future__ import print_function

import google_auth_oauthlib
import httplib2
from google.oauth2 import credentials
from googleapiclient import errors
from googleapiclient.discovery import build
from httplib2 import Http
from oauth2client import file as oauth_file, client, tools
import os.path
from google_auth_oauthlib.flow import InstalledAppFlow
from google.auth.transport.requests import Request
from google.oauth2.credentials import Credentials

import os
import sys
import pandas as pd
import pygsheets
from PyQt5 import uic, QtWidgets, QtGui, QtCore
from PyQt5.QtCore import QObject, QThread, pyqtSignal, pyqtSlot, QAbstractTableModel, Qt
from PyQt5.QtWidgets import QFileDialog, QTableView
import datetime
# from google.protobuf import service
from pygsheets import *
from win32com.client import Dispatch


xl = Dispatch("Excel.Application")

xlToLeft = 1
xlToRight = 2
xlUp = 3
xlDown = 4
xlAscending = 1
xlYes = 1

gc = pygsheets.authorize(service_file='mistaketracker-8ece8eccb212.json')

qtCreatorFile = "MTmenu.ui"  # Enter file here.

Ui_MainWindow, QtBaseClass = uic.loadUiType(qtCreatorFile)

SAMPLE_CODE = '''
function activeCell() {
  //Initialize variables
  var sheet = SpreadsheetApp.getActiveSheet();
  var cell = sheet.getCurrentCell();
  var selectedColumn = cell.getColumn();
  var selectedRow = cell.getRow();
  var firstRow = 3;
  var listValues = sheet.getRange("B1:B").getValues();
  var lastRow = listValues.filter(String).length + 1;
  var index;
  //Print values
  Logger.log(`selectedColumn: ${selectedColumn}`);
  Logger.log(`selectedRow: ${selectedRow}`);
  Logger.log(`firstRow: ${firstRow}`);
  Logger.log(`lastRow: ${lastRow}`);
  //Find index to alter and set index to -1 if none selected
  if (selectedColumn >= 10) {
    index = -1;
  }
  else if (selectedRow < firstRow || selectedRow > lastRow) {
    index = -1;
  }
  else {
    index = selectedRow - firstRow;
  }
  //Return results
  Logger.log(`Index: ${index}`);
  return index;
}
'''.strip()

SAMPLE_MANIFEST = '''
{
  "timeZone": "America/New_York",
  "dependencies": {},
  "exceptionLogging": "STACKDRIVER",
  "runtimeVersion": "V8"
}
'''.strip()


class Worker(QObject):  # worker class to be used as additional thread (fix gui freezes)
    finished = pyqtSignal()  # signals to communicate worker thread with main thread
    progress = pyqtSignal(int)

    @pyqtSlot()
    def formatting(self):  # function that worker thread is going to run (main thread used in gui)
        """Long-running task."""
        path = 'mistaketracker.ftr'  # read database
        df = pd.read_feather(path)

        sh = gc.open('Report_Issue')  # open the google spreadsheet (where 'Report_Issue' is the name of my sheet)
        wks = sh[0]  # select the first sheet
        wks.clear()  # allows removed data to be removed from sheet
        wks.set_dataframe(df, (2, 2))  # update the first sheet with df, starting at cell A1.

        hdr = wks.get_values('B2', 'I2', returnas='range')

        rng = DataRange(start='B3', end='I4', worksheet=wks)
        rng.end_addr = 'I' + str(len(df) + 2)
        dte = wks.get_values('G3', 'G' + str(len(df) + 2), returnas='range')

        header_cell = Cell('A1')
        header_cell.color = (0.2, 0.3, 1, 1)  # blue color cell
        header_cell.text_format['fontSize'] = 14
        header_cell.set_text_format('foregroundColor', (1, 1, 1, 1))
        header_cell.set_text_format('bold', True)

        model_cell = Cell('B1')
        date_cell = Cell('F2')
        for i in range(len(df) + 1):
            rng.start_addr = 'B' + str(i + 2)
            dte.start_addr = 'G' + str(i + 2)
            if i % 2 == 0:
                model_cell.color = (0.8, 0.9, 1, 1)
                date_cell.color = (0.8, 0.9, 1, 1)  # light blue color cell
                date_cell.format = (pygsheets.FormatType.DATE, '')
                rng.apply_format(model_cell)
                dte.apply_format(date_cell)
                #rng.update_borders(top=True, right=True, bottom=True, left=True, style='SOLID')#, width=2)
            else:
                model_cell.color = (0.9, 0.9, 0.9, 1)
                date_cell.color = (0.9, 0.9, 0.9, 1)  # gray color cell
                date_cell.format = (pygsheets.FormatType.DATE, '')
                rng.apply_format(model_cell)
                dte.apply_format(date_cell)
                #rng.update_borders(top=True, right=True, bottom=True, left=True, style='SOLID')#, width=1)
            # update progress bar
            self.progress.emit(i)

        hdr.apply_format(header_cell)
        #hdr.update_borders(top=True, right=True, bottom=True, left=True, style='SOLID', width=2, red=0, green=0, blue=0)
        #rng.update_borders(top=True, right=True, bottom=True, left=True, style='SOLID')
        self.finished.emit()


class MistakeTracker(QtWidgets.QMainWindow, Ui_MainWindow):
    def __init__(self):
        QtWidgets.QMainWindow.__init__(self)  # Ui set-up
        Ui_MainWindow.__init__(self)
        self.setupUi(self)
        self.updateBtn2.setEnabled(False)
        self.updateBtn3.setEnabled(False)
        self.editIndex = -1
        # Initialize a QThread object
        self.thread = []  # QThread()
        # Initialize a worker object
        self.worker = []  # Worker()

    def insert(self):
        date_old = str(self.dateEdit.date().toPyDate())                  # get and convert date
        date = datetime.datetime.strptime(date_old, '%Y-%m-%d').strftime('%m/%d/%y')

        try:
            df1 = pd.DataFrame({  # insert data into table
                'NO.': [0],
                'Assembly Num': [int(self.lineEditAssy.text())],
                'Part Num': [int(self.lineEditPart.text())],
                'Description': [str(self.lineEditDesc.text())],
                'Reporter': [str(self.lineEditReport.text())],
                'Date': [date],
                'Responder': [str(self.comboBoxResp.currentText())],
                'Status': [str(self.comboBoxStat.currentText())]
            })
        except Exception as e:  # block wrong data
            print('No/Wrong Values Entered')
            print('Catch: ', e.__class__)
            return 1

        path = 'mistaketracker.ftr'                 # read database
        if os.path.exists(path):
            df = pd.read_feather(path)
            df = df.append(df1, ignore_index=True)
            df.index = range(len(df))
            df.at[df.index, 'NO.'] = df.index
            df.set_index('NO.')
            #df = pd.DataFrame.reset_index(df)  # removes duplicate index
        else:
            df = df1
        pd.DataFrame.to_feather(df, path)  # save database locally

        self.lineEditAssy.setText('')  # clear fields after insert to let user know action took place
        self.lineEditPart.setText('')
        self.lineEditDesc.setText('')
        self.lineEditReport.setText('')

        sh = gc.open('Report_Issue')  # open the google spreadsheet and populate it
        wks = sh[0]
        wks.clear()
        wks.set_dataframe(df, (2, 2))

    def load(self):
        path = 'mistaketracker.ftr'  # read database
        df = pd.read_feather(path)

        #for i in range(len(df)):
        #    row = str(df.iloc[i, 0]) + '-' + str(df.iloc[i, 1]) + '-' + str(df.iloc[i, 2]) + '-' + str(df.iloc[i, 3]) \
        #          + '-' + str(df.iloc[i, 4]) + '-' + str(df.iloc[i, 5]) + '-' + str(df.iloc[i, 6]) + '-' \
        #          + str(df.iloc[i, 7])

        model = pandasModel(df)
        self.tableView.setModel(model)

    def delete(self):
        path = 'mistaketracker.ftr'  # read database
        df = pd.read_feather(path)

        #idx = self.listWidget.currentRow()  # get index and delete item in list
        index = self.tableView.selectionModel().currentIndex()  # get index of selected row
        idx = index.row()
        self.editIndex = idx
        if self.editIndex == -1:            # nothing happens if no selection
            return 1
        df = df.drop([df.index[idx]])       # deletes selection

        df.index = range(len(df))           # reset indexes
        df.at[df.index, 'NO.'] = df.index
        df.set_index('NO.')
        pd.DataFrame.to_feather(df, path)  # save database locally

        sh = gc.open('Report_Issue')  # open the google spreadsheet and populate it
        wks = sh[0]
        wks.clear()
        wks.set_dataframe(df, (2, 2))

        self.load()

    def alter(self):
        path = 'mistaketracker.ftr'  # read database
        df = pd.read_feather(path)

        SCRIPT_ID = '1BsYZ6o2NZsdrB-PUV35pC9B_jIPobJTbq79YU978qwhyahIra5fy3g3j'
        #https://developers.google.com/apps-script/api/how-tos/execute#api_request_examples
        #https://www.benlcollins.com/apps-script/api-tutorial-for-beginners/
        #https://ctrlq.org/google.apps.script/docs/guides/rest/quickstart/python.html

        SCOPES = [
            #'https://www.googleapis.com/auth/script.scriptapp',
            #'https://www.googleapis.com/auth/drive.readonly',
            'https://www.googleapis.com/auth/spreadsheets',
            'https://www.googleapis.com/auth/drive',
            #'https://www.googleapis.com/auth/script.external_request',
            'https://www.googleapis.com/auth/script.projects',
            f'https://script.googleapis.com/v1/scripts/{SCRIPT_ID}:run',
        ]
        #creds = None
        # The file token.json stores the user's access and refresh tokens, and is
        # created automatically when the authorization flow completes for the first
        # time.
        #if os.path.exists('token.json'):
        #    creds = Credentials.from_authorized_user_file('token.json', SCOPES)
        # If there are no (valid) credentials available, let the user log in.
        #print(1)
        #if not creds or not creds.valid:
        #    if creds and creds.expired and creds.refresh_token:
        #        creds.refresh(Request())
        #    else:
        #        print(2)
        #        flow = InstalledAppFlow.from_client_secrets_file(
        #            'client_secret.json', SCOPES)
        #        print(flow)
        #        creds = flow.run_local_server(port=0)
                #creds = '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'
                #creds = flow.run_console()
            # Save the credentials for the next run
        #    print(4)
        #    with open('token.json', 'w') as token:
        #        token.write(creds.to_json())
        #print(5)
        #service = build('script', 'v1', credentials=creds)
        #print(6)

        # Call the Apps Script API
        #try:
            # Create a new project
        #    request = {'title': 'My Script'}
        #    response = service.projects().create(body=request).execute()

            # Upload two files to the project
        #    request = {
        #        'files': [{
        #            'name': 'hello',
        #            'type': 'SERVER_JS',
        #            'source': SAMPLE_CODE
        #        }, {
        #            'name': 'appsscript',
        #            'type': 'JSON',
        #            'source': SAMPLE_MANIFEST
        #        }]
        #    }
        #    response = service.projects().updateContent(
        #        body=request,
        #        scriptId=response['scriptId']).execute()
        #    print('https://script.google.com/d/' + response['scriptId'] + '/edit')
        #except errors.HttpError as error:
            # The API encountered a problem.
        #    print(error.content)

        #service = build('script', 'v1', credentials=creds)
        #print(3)
        #request = {"function": "activeCell"}
        #response = service.scripts().run(body=request, scriptId=SCRIPT_ID).execute()

        #index = wks.cell('K1').value()
        #index = pygsheets.Cell(pos="K1", worksheet=wks)
        #print(response)

        #idx = self.listWidget.currentRow()  # get index and set editing item to selected one
        #row = self.tableView.currentRow
        #self.tableView.cellClicked.connect(self.cell_clicked)
        index = self.tableView.selectionModel().currentIndex()#.selectedRows()
        idx = index.row()
        #for index in sorted(indexes):
        #    print('Row %d is selected' % index.row())
        #    idx = index.row()
        print(idx)
        #idx = -1
        self.editIndex = idx
        if self.editIndex == -1:        # nothing happens if no selection
            return 1

        self.lineEditAssy.setText(str(df.iloc[idx, 1]))
        self.lineEditPart.setText(str(df.iloc[idx, 2]))
        self.lineEditDesc.setText(str(df.iloc[idx, 3]))
        self.lineEditReport.setText(str(df.iloc[idx, 4]))
        date_old = df.iloc[idx, 5]
        date = datetime.datetime.strptime(date_old, '%m/%d/%y')
        self.dateEdit.setDate(date)
        self.comboBoxResp.setCurrentText(str(df.iloc[idx, 6]))
        self.comboBoxStat.setCurrentText(str(df.iloc[idx, 7]))

        self.tabWidget.setCurrentIndex(0)           # set buttons and tab appropriate for edit and update
        self.tabEdit.setEnabled(False)
        self.insertBtn.setEnabled(False)
        self.reportBtn.setEnabled(False)
        self.directoryBtn.setEnabled(False)
        self.updateBtn2.setEnabled(True)
        self.updateBtn3.setEnabled(True)

    def alter2(self):
        path = 'mistaketracker.ftr'  # read database
        df = pd.read_feather(path)

        date_old = str(self.dateEdit.date().toPyDate())  # get and convert date
        date = datetime.datetime.strptime(date_old, '%Y-%m-%d').strftime('%m/%d/%y')

        # update selected value
        df.iloc[self.editIndex] = [self.editIndex, int(self.lineEditAssy.text()), int(self.lineEditPart.text()),
                                   str(self.lineEditDesc.text()), str(self.lineEditReport.text()), str(date),
                                   str(self.comboBoxResp.currentText()), str(self.comboBoxStat.currentText())]

        df.index = range(len(df))               # set index
        df.at[df.index, 'NO.'] = df.index
        df.set_index('NO.')
        pd.DataFrame.to_feather(df, path)  # save database locally

        sh = gc.open('Report_Issue')  # open the google spreadsheet and populate it
        wks = sh[0]
        wks.clear()
        wks.set_dataframe(df, (2, 2))

        #data = SpreadsheetApp.getActiveSheet().getDataRange().getValues()
        #wks.values().clear() #ActiveCell.Select()
        #print(data)

        self.load()
        self.cancel()

    def cancel(self):
        self.tabWidget.setCurrentIndex(1)  # set buttons and tabs for not updating
        self.tabEdit.setEnabled(True)
        self.insertBtn.setEnabled(True)
        self.reportBtn.setEnabled(True)
        self.directoryBtn.setEnabled(True)
        self.updateBtn2.setEnabled(False)
        self.updateBtn3.setEnabled(False)

        self.lineEditAssy.setText('')
        self.lineEditPart.setText('')
        self.lineEditDesc.setText('')
        self.lineEditReport.setText('')

    def reportProgress(self, value):  # Used to update loading bar
        self.progressBar.setValue(value)

    def format(self):
        path = 'mistaketracker.ftr'  # read database
        df = pd.read_feather(path)

        self.progressBar.setMaximum(len(df))
        self.tabWidget.setCurrentIndex(0)

        # Create a QThread object
        self.thread = QThread()
        # Create a worker object
        self.worker = Worker()
        # Move worker to the thread
        self.worker.moveToThread(self.thread)
        # Connect signals and slots
        self.thread.started.connect(self.worker.formatting)
        self.worker.finished.connect(self.thread.quit)
        self.worker.finished.connect(self.worker.deleteLater)
        self.thread.finished.connect(self.thread.deleteLater)
        self.worker.progress.connect(self.reportProgress)
        # Start the thread
        self.thread.start()

        # Disable all buttons except cancel while worker thread is functioning
        self.tabWidget.setEnabled(False)
        self.thread.finished.connect(
            lambda: self.tabWidget.setEnabled(True)
        )

    def index_change(self):
        index = self.comboBoxCol.currentIndex()
        if index == 0 or index == 1:
            self.stackedWidget.setCurrentIndex(0)
            self.stackedWidget_2.setCurrentIndex(0)
        elif index == 2:
            self.stackedWidget.setCurrentIndex(0)
            self.stackedWidget_2.setCurrentIndex(1)
        elif index == 3:
            self.stackedWidget.setCurrentIndex(1)
            self.stackedWidget_2.setCurrentIndex(2)
        elif index == 4:
            self.stackedWidget.setCurrentIndex(1)
            self.stackedWidget_2.setCurrentIndex(3)

    def search(self):
        #self.listWidget_2.clear()
        path = 'mistaketracker.ftr'  # read database
        df = pd.read_feather(path)

        index = self.comboBoxCol.currentIndex()
        if index == 0 or index == 1:                                  # Assembly Num or Part Num
            col = self.comboBoxCol.currentText()
            cond = self.comboBoxCond.currentText()
            try:
                num = int(self.lineEditNum.text())
            except Exception as e:  # block wrong data
                print('No/Wrong Values Entered')
                print('Catch: ', e.__class__)
                return 1
            if cond == 'greater than or equal to':
                self.labelSearch.setText(col + ' >= ' + str(num))
                df1 = df.loc[df[col] >= int(num)]
            elif cond == 'less than or equal to':
                self.labelSearch.setText(col + ' <= ' + str(num))
                df1 = df.loc[df[col] <= int(num)]
            else:
                self.labelSearch.setText(col + ' == ' + str(num))
                df1 = df.loc[df[col] == int(num)]

        elif index == 2:                                # Date
            col = self.comboBoxCol.currentText()
            cond = self.comboBoxCond.currentText()
            date_old = str(self.dateEdit_2.date().toPyDate())  # get and convert date
            date = datetime.datetime.strptime(date_old, '%Y-%m-%d').strftime('%m/%d/%y')
            if cond == 'greater than or equal to':
                self.labelSearch.setText(col + ' >= ' + str(date))
                df1 = df.loc[pd.to_datetime(df[col]) >= pd.to_datetime(date)]
            elif cond == 'less than or equal to':
                self.labelSearch.setText(col + ' <= ' + str(date))
                df1 = df.loc[pd.to_datetime(df[col]) <= pd.to_datetime(date)]
            else:
                self.labelSearch.setText(col + ' == ' + str(date))
                df1 = df.loc[pd.to_datetime(df[col]) == pd.to_datetime(date)]

        elif index == 3:                                # Responder
            col = self.comboBoxCol.currentText()
            cond = self.comboBoxCond2.currentText()
            res = self.comboBoxRe.currentText()
            if cond == 'equal to':
                self.labelSearch.setText(col + ' == ' + res)
                df1 = df.loc[df[col] == res]
            else:
                self.labelSearch.setText(col + ' != ' + res)
                df1 = df.loc[df[col] != res]

        else:                                           # Status
            col = self.comboBoxCol.currentText()
            cond = self.comboBoxCond2.currentText()
            sta = self.comboBoxSt.currentText()
            if cond == 'equal to':
                self.labelSearch.setText(col + ' == ' + sta)
                df1 = df.loc[df[col] == sta]
            else:
                self.labelSearch.setText(col + ' != ' + sta)
                df1 = df.loc[df[col] != sta]

        self.labelResults.setText('Results: ' + str(len(df1)))
        model = pandasModel(df1)
        self.tableView_2.setModel(model)
        #for i in range(len(df1)):
        #    row = str(df1.iloc[i, 0]) + '-' + str(df1.iloc[i, 1]) + '-' + str(df1.iloc[i, 2]) + '-' + str(
        #        df1.iloc[i, 3]) + '-' + str(df1.iloc[i, 4]) + '-' + str(df1.iloc[i, 5]) + '-' + str(
        #        df1.iloc[i, 6]) + '-' + str(df1.iloc[i, 7])
        #    self.listWidget_2.insertItem(i, row)

    def report(self):
        #xl.Workbooks("TimingSpace.xlsx").Worksheets("Sheet1").Activate()
        pathDir = self.lineEditDir.text() + '/'
        pathDir = pathDir.replace('/', '\\')
        if pathDir == '\\':
            pathDir = 'c:\\Python27\\'
        #path = 'c:\\Adams\\Python\\MistakeTracker\\'
        #xl.Range("B3").Select()
        #ShelveName = str(xl.ActiveCell.Value)
        xlname = "QualityTracking.xlsx"
        #CURR_DIR = os.path.dirname(os.path.realpath(__file__)) + '/' + xlname
        #print(CURR_DIR)
        xlnamelong = pathDir + xlname
        #filename = pathDir + "QualityTracking"
        if os.path.isfile(xlnamelong):  # generate workbook
            xl.Workbooks.Open(xlnamelong, ReadOnly=False)
            xl.Workbooks(xlname).Worksheets("Sheet1").Activate()
            xl.Visible = 1
        else:
            xl.Workbooks.Add()
            xl.Worksheets("Sheet1").Activate()
            xl.Visible = 1
        #GwareData.database = shelve.open(filename, writeback=False)
        #GwareData.RawDataList = GwareData.database['RawDataList']
        #RawDict = [obj.__dict__ for obj in GwareData.RawDataList]
        #RawDict = sorted(RawDict, key=lambda i: i['name'])  # sort list of dict by name
        #projName = ['Project Name:', ShelveName]
        header = ['NO.', 'Assembly Num', 'Part Num', 'Description', 'Reporter', 'Date', 'Responder', 'Status']
        path = 'mistaketracker.ftr'  # read database
        df = pd.read_feather(path)
        self.progressBar.setMaximum(len(df))

        xl.Range("B2").Select()
        #WriteRowFromSelected(projName)  # write project name
        for entry in header:
            # print(entry)
            xl.ActiveCell.Value = entry
            xl.ActiveCell.Offset(1, 2).Select()
        for i in range(len(df)):
            xlspot = i + 3
            xl.Range("B" + str(xlspot)).Select()
            for entry in df.iloc[i]:
                # print(entry)
                xl.ActiveCell.Value = entry
                xl.ActiveCell.Offset(1, 2).Select()
            self.progressBar.setValue(i + 1)
            #WriteColumnFromSelected(initProfile)  # write column headers for profile
            #xl.Range("C" + str(xlspot)).Select()
            #xl.ActiveCell.Value = RawDict[i]['name']  # write profile names
        #numNodes = []
        #for i in range(len(RawDict)):  # get list containing # of nodes per profile
        #    templist = []
        #    temp = len(RawDict[i]['RawData'][0])
        #    for j in range(1, temp + 1):  # name nodes for importing
        #        templist.append('Node ' + str(j))
        #    numNodes.append(templist)
        #for i in range(len(RawDict)):  # write nodes
        #    xlspot = 6 + (8 * i)
        #    xl.Range("C" + str(xlspot)).Select()
        #    WriteRowFromSelected(numNodes[i])
        #self.WriteXLRowData(RawDict, 7, 0)  # write times
        #self.WriteXLRowData(RawDict, 8, 1)  # write displcement
        #self.WriteXLRowData(RawDict, 9, 2)  # write velocity
        #self.WriteXLRowData(RawDict, 10, 3)  # write acceleration
        self.FormatBordersColors()
        xl.ActiveWorkbook.SaveAs(xlnamelong)  # save file
        #wb.SaveAs(Filename=xlnamelong, FileFormat=6)
                  #win32com.client.constants.xlWorkbookNormal, None, None, False, False,
                  #win32com.client.constants.xlNoChange, win32com.client.constants.xlOtherSessionChanges)

    #def WriteXLRowData(self, RawDict, startLine, dataNum):  # writes rows of data for raw data
    #    for i in range(len(RawDict)):  # write nodes
    #        xlspot = startLine + (8 * i)
    #        xl.Range("C" + str(xlspot)).Select()
    #        WriteRowFromSelected(RawDict[i]['RawData'][dataNum])

    def FormatBordersColors(self):  # handles font size, borders, colors
        path = 'mistaketracker.ftr'  # read database
        df = pd.read_feather(path)

        xl.Range("B2:I2").Font.Size = 22  # title font size
        xl.Range("B2:I2").Interior.ColorIndex = 5  # title background color
        xl.Range("B2:I2").Font.ColorIndex = 2   # title font color

        for i in range(1, 5):  # creates borders
            xl.Range("B2:I" + str(len(df) + 2)).Borders(i).Weight = 4
        #    for j in range(len(numNodes)):
        #        xlspot = 5 + (8 * j)
        #        xl.Range(xl.Cells(xlspot, 2), xl.Cells(xlspot + 5, len(numNodes[j]) + 2)).Borders(i).Weight = 4
        for k in range(len(df)):  # creates font size and background color
            xlspot = k + 3
            xl.Range(xl.Cells(xlspot, 2), xl.Cells(xlspot, 9)).Font.Size = 16  # font size
            #xl.Range(xl.Cells(xlspot + 2, 2), xl.Cells(xlspot + 5, 2)).Font.Size = 18
            #xl.Range(xl.Cells(xlspot + 2, 3), xl.Cells(xlspot + 5, len(numNodes[k]) + 2)).Font.Size = 16
            #xl.Cells(xlspot, 2).Interior.ColorIndex = 15  # cell colors
            #xl.Cells(xlspot, 3).Interior.ColorIndex = 43
            if k % 2 == 0:
                xl.Range(xl.Cells(xlspot, 2), xl.Cells(xlspot, 9)).Interior.ColorIndex = 15
            else:
                xl.Range(xl.Cells(xlspot, 2), xl.Cells(xlspot, 9)).Interior.ColorIndex = 34
            #xl.Range(xl.Cells(xlspot + 3, 2), xl.Cells(xlspot + 3, len(numNodes[k]) + 2)).Interior.ColorIndex = 35
            #xl.Range(xl.Cells(xlspot + 4, 2), xl.Cells(xlspot + 4, len(numNodes[k]) + 2)).Interior.ColorIndex = 34
            #xl.Range(xl.Cells(xlspot + 5, 2), xl.Cells(xlspot + 5, len(numNodes[k]) + 2)).Interior.ColorIndex = 38
            #xl.Range(xl.Cells(xlspot + 1, 3), xl.Cells(xlspot + 1, len(numNodes[k]) + 2)).Interior.ColorIndex = 40
        xl.Columns.AutoFit()  # ensures data fits in excel boxes

    def SelectWorkingDirectory(self):
        _OutputFolder = str(QFileDialog.getExistingDirectory(self, "Select Directory"))
        self.lineEditDir.setText(_OutputFolder)


class pandasModel(QAbstractTableModel):         # Used to import pandas data frame into tableView

    def __init__(self, data):
        QAbstractTableModel.__init__(self)
        self._data = data

    def rowCount(self, parent=None):
        return self._data.shape[0]

    def columnCount(self, parent=None):
        return self._data.shape[1]

    def data(self, index, role=Qt.DisplayRole):
        if index.isValid():
            if role == Qt.DisplayRole:
                return str(self._data.iloc[index.row(), index.column()])
        return None

    def headerData(self, col, orientation, role):
        if orientation == Qt.Horizontal and role == Qt.DisplayRole:
            return self._data.columns[col]
        return None

#df = pd.DataFrame({
#    'CompanyName': ['Google', 'Microsoft', 'SpaceX', 'Amazon', 'Samsung'],
#    'Founders': ['Larry Page, Sergey Brin', 'Bill Gates, Paul Allen', 'Elon Musk',
#                 'Jeff Bezos', 'Lee Byung-chul'],
#    'Founded': [1998, 1975, 2002, 1994, 1938],
#    'Number of Employees': ['103,459', '144,106', '6,500', '647,500', '320,671']
#})

#df.iloc[3] = ['YouTube', 'Chad Hurley, Steve Chen, Jawed Karim', 2005, '2,800']

#df = df.drop([df.index[5]])

#df = df[df.Founders != 'Chad Hurley, Steve Chen, Jawed Karim']

#print(df.loc[2])

#df.at[1, 'Number of Employees'] = '200,000'

#path2 = 'mistaketracker.csv'
#df.to_csv(path2, index=False)


if __name__ == "__main__":
    app = QtWidgets.QApplication(sys.argv)
    window = MistakeTracker()
    window.show()
    sys.exit(app.exec_())
