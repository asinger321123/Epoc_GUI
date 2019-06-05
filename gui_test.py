import sys
import utils as utils
from PyQt4 import QtCore, QtGui, uic
from PyQt4.QtGui import *
from PyQt4.QtCore import *
from PyQt4.QtWebKit import QWebView
# from PyQt4.QtGui import QLineEdit, QCheckBox, QListWidget, QPushButton, QPlainTextEdit, QComboBox, QLabel, QGroupBox, QStringList
from openpyxl import Workbook, load_workbook
import re
import os
from os import listdir
from os.path import isfile, join, isdir
import sys
import xlrd
import csv
from shutil import copyfile
from fuzzywuzzy import fuzz
from fuzzywuzzy import process
import json
import subprocess
# from subprocess import call
import psycopg2
import itertools
import resource_rc
import glob
from collections import defaultdict
import winsound
from threading import Thread
import sqlite3
# from colorama import init, Fore, Back, Style

userhome = os.path.expanduser('~')
desktop = userhome + '\\Desktop\\'
downloads = userhome + '\\Downloads\\'
newest = max(os.listdir(downloads), key=lambda f: os.path.getmtime("{}/{}".format(downloads, f)))
userID = os.path.expanduser('~').split("""\\""")[-1]
strList = ['MD', 'DO', 'Allergy and Immunology', 'Anesthesiology', 'Cardiology', 'Dermatology', 'Emergency Medicine', 'Endocrinology', 'Family Practice', 'Gastroenterology', 'Geriatrics', 'Hematology', 'Infectious Disease', 'Internal Medicine, General', 'Neonatology', 'Nephrology', 'Neurology', 'Neurosurgery', 'Obstetrics/Gynecology', 'Occupational Medicine', 'Oncology', 'Ophthalmology', 'Oral and Maxillofacial Surgery (OFMS)', 'Orthopaedic Surgery', 'Other', 'Otorhinolaryngology', 'Pathology', 'Pediatrics, General', 'Pediatrics, Subspecialty', 'Pediatrics, Surgical', 'Perinatology', 'Physical Medicine/Rehabilitation', 'Preventive Medicine', 'Psychiatry', 'Pulmonology', 'Radiation Oncology', 'Radiology, General', 'Radiology, Vascular and Interventional', 'Rheumatology', 'Student', 'Surgery, Cardiac/Thoracic', 'Surgery, General', 'Surgery, Plastic', 'Surgery, Transplant', 'Surgery, Vascular', 'Urology', 'Chiropractor', 'Clinical Research Associate', 'CNM', 'CNS', 'CRNA', 'Dental Hygienist', 'Dentist', 'Dietician/Nutritionist', 'EMT/Paramedic', 'Genetic Counselor', 'Geneticist', 'Healthcare IT', 'Hospital or Practice Administrator', 'Lab Director', 'Lab Technician', 'Law Enforcement', 'LPN', 'Managed Care - Director of Pharmacy', 'Managed Care - Medical Director', 'Managed Care - Other', 'Managed Care - Provider Relations', 'Medical Assistant', 'Medical Librarian', 'ND', 'NP', 'Optician', 'Optometrist', 'PA', 'Pharmacist - Community', 'Pharmacist - Hospital', 'Pharmacy Technician', 'Physical Therapist', 'Podiatrist', 'Psychologist', 'Research Scientist', 'Residency Coordinator', 'Respiratory Therapist', 'RN', 'Veterinarian']
# drugImport = utils.importDrugs()
downloadFilesDirectory = [f for f in listdir(downloads) if isfile(join(downloads, f))]
downloadFilesDirectory.sort(key=lambda x: os.stat(os.path.join(downloads, x)).st_mtime, reverse=True)

qtCreatorFile = os.path.join(desktop, 'Ewok','new_andrew3.ui') # Enter file here.
Ui_MainWindow, QtBaseClass = uic.loadUiType(qtCreatorFile)

uifile_2 = os.path.join(desktop, 'Ewok','fileWidget.ui')
uifile_3 = os.path.join(desktop, 'Ewok','pivotTabe_ui.ui')
uifile_4 = os.path.join(desktop, 'Ewok','AdditionalSDA.ui')
uifile_5 = os.path.join(desktop, 'Ewok','AdditionalBDA.ui')

form_2, base_2 = uic.loadUiType(uifile_2)
form_3, base_3 = uic.loadUiType(uifile_3)
form_4, base_4 = uic.loadUiType(uifile_4)
form_5, base_5 = uic.loadUiType(uifile_5)

class Completer(QtGui.QCompleter):
    def __init__(self, *args, **kwargs):
        super(Completer, self).__init__(*args, **kwargs)

        self.setCaseSensitivity(QtCore.Qt.CaseInsensitive)
        self.setCompletionMode(0)
        self.setWrapAround(False)

    # Add texts instead of replace
    def pathFromIndex(self, index):
        path = QtGui.QCompleter.pathFromIndex(self, index)
        lst = str(self.widget().text()).split(',')
        if len(lst) > 1:
            path = '%s, %s' % (','.join(lst[:-1]), path)
            path = str(path).replace('Internal Medicine, Internal Medicine, General', 'Internal Medicine, General').replace('Pediatrics, Pediatrics, General', 'Pediatrics, General').replace('Pediatrics, Pediatrics, Subspecialty', 'Pediatrics, Subspecialty').replace('Pediatrics, Pediatrics, Surgical', 'Pediatrics, Surgical').replace('Radiology, Radiology, General', 'Radiology, General').replace('Radiology, Radiology, Vascular and Interventional', 'Radiology, Vascular and Interventional').replace('Surgery, Surgery, Cardiac/Thoracic', 'Surgery, Cardiac/Thoracic').replace('Surgery, Surgery, General', 'Surgery, General').replace('Surgery, Surgery, Plastic', 'Surgery, Plastic').replace('Surgery, Surgery, Transplant', 'Surgery, Transplant').replace('Surgery, Surgery, Vascular', 'Surgery, Vascular')
        return path

    def splitPath(self, path):
        path = str(path.split(',')[-1]).lstrip(' ')
        return [path]

class TextEdit(QLineEdit):

    def __init__(self, parent):
        super(TextEdit, self).__init__(parent)
        self._completer = Completer(strList, self)
        self.setCompleter(self._completer)

        sizePolicy = QtGui.QSizePolicy(QtGui.QSizePolicy.MinimumExpanding, QtGui.QSizePolicy.Fixed)
        # sizePolicy.setHorizontalStretch(0)
        self.setSizePolicy(sizePolicy)
        # self.setMaximumWidth(500)
        # self.setMinimumSize(QtCore.QSize(134, 0))
        self.setFixedWidth(160)

class TextEdit2(QLineEdit):

    def __init__(self, parent):
        super(TextEdit2, self).__init__(parent)
        self._completer = Completer(strList, self)
        self.setCompleter(self._completer)

        sizePolicy = QtGui.QSizePolicy(QtGui.QSizePolicy.MinimumExpanding, QtGui.QSizePolicy.Fixed)
        # sizePolicy.setHorizontalStretch(0)
        self.setSizePolicy(sizePolicy)
        # self.setMaximumWidth(500)
        # self.setMinimumSize(QtCore.QSize(134, 0))
        self.setFixedWidth(134)

class SDA_Widget(base_4, form_4):

    def __init__(self):
        super(base_4, self).__init__()
        self.setupUi(self)

        self.newSdaOcc = ""
        self.newSdaSpec = ""

        self.sdaOccAdd = TextEdit(self.findChild(QLineEdit, 'sdaOcc_line_edit2'))
        self.sdaSpecAdd = TextEdit(self.findChild(QLineEdit, 'sdaSpec_line_edit2'))
        self.sdaOkButton = self.findChild(QPushButton, 'submitSDA_pushButton')
        self.sdaOnlyCheckBox = self.findChild(QCheckBox, 'sdaOnly_checkBox2')

        self.sdaOkButton.pressed.connect(self.writeSDAConfig)
        self.sdaOkButton.released.connect(self.close)
        self.sdaOkButton.released.connect(window.showSDAToolTip)

    # def setSdaDefaults(self):
        self.sdaOccAdd.setText('MD, DO, NP, PA')
        self.sdaOnlyCheckBox.setEnabled(False)


    def writeSDAConfig(self):
        self.sdaConfig = dict()
        self.occList = str(self.sdaOccAdd.text()).split(", ")
        self.specList = str(self.sdaSpecAdd.text()).split(", ")
        self.sdaOcc2 = ", ".join('"{0}"'.format(w) for w in self.occList)
        self.sdaSpec2 = ", ".join('"{0}"'.format(w) for w in self.specList)

        self.checkTotalSdas()
        self.addSdaOcc = 'additonalSdaOcc'+'_'+str(window.sdainc)
        self.addSdaSpec = 'additonalSdaSpec'+'_'+str(window.sdainc)

        self.sdaConfig['totalAdditionalSDAs'] = window.sdainc
        self.sdaConfig[self.addSdaOcc] = str(self.sdaOcc2)
        self.sdaConfig[self.addSdaSpec] = str(self.sdaSpec2).replace('"Internal Medicine", "General"', '"Internal Medicine, General"').replace('"Pediatrics", "General"', '"Pediatrics, General"').replace('"Pediatrics", "Subspecialty"', '"Pediatrics, Subspecialty"').replace('"Pediatrics", "Surgical"', '"Pediatrics, Surgical"').replace('"Radiology", "General"', '"Radiology, General"').replace('"Radiology", "Vascular and Interventional"', '"Radiology, Vascular and Interventional"').replace('"Surgery", "Cardiac/Thoracic"', '"Surgery, Cardiac/Thoracic"').replace('"Surgery", "General"', '"Surgery, General"').replace('"Surgery", "Plastic"', '"Surgery, Plastic"').replace('"Surgery", "Transplant"', '"Surgery, Transplant"').replace('"Surgery", "Vascular"', '"Surgery, Vascular"')
        self.setSDACountLabel()
        # with open(desktop+'\\'+'sdaConfig.json', 'w') as outfile:
        #     json.dump(self.sdaConfig, outfile, indent=2, sort_keys=True, encoding='utf8')

        with open(desktop+'\\'+'sdaConfig.json', 'r') as infile:
            data = json.loads(infile.read(), encoding='utf8')

            data.update(self.sdaConfig)

        with open(desktop+'\\'+'sdaConfig.json', 'w') as outfile:
            json.dump(data, outfile, indent=2, sort_keys=True, encoding='utf8')

    def checkTotalSdas(self):
        with open(desktop+'\\'+'sdaConfig.json', 'r') as infile:
            config = json.loads(infile.read(), encoding='utf8')
            # if config['additonalSdaOcc'+'_'+str(window.sdainc)]:
            if 'additonalSdaOcc'+'_'+str(window.sdainc) in config:
                window.sdainc += 1

                # window.sdaLabel.setText('Additional SDAs: '+ window.sdainc)

    def setSDACountLabel(self):
        window.sdaLabel.setText('Additional SDAs: '+ str(window.sdainc))

class BDA_Widget(base_5, form_5):

    def __init__(self):
        super(base_5, self).__init__()
        self.setupUi(self)

        self.newBdaOcc = ""
        self.newBdaSpec = ""

        self.bdaOccAdd = TextEdit2(self.findChild(QLineEdit, 'bdaOcc_line_edit2'))
        self.bdaSpecAdd = TextEdit2(self.findChild(QLineEdit, 'bdaSpec_line_edit2'))
        self.bdaOkButton = self.findChild(QPushButton, 'submitBDA_pushButton')
        self.bdaLookUpsAdd = self.findChild(QLineEdit, 'bdaLookUps_line_edit2')
        self.bdaNumLookUpsAdd = self.findChild(QLineEdit, 'bdaNumLookUps_line_edit2')
        self.deDupAdd = self.findChild(QCheckBox, 'deDup_check_box2')
        self.drugListAdd = self.findChild(QPlainTextEdit, 'drugList_text_edit2')
        self.bdaOnlyAdd = self.findChild(QCheckBox, 'bdaOnly_checkBox2')
        self.therapyClassBoxAdd = self.findChild(QCheckBox, 'therapyClass_checkBox2')


        self.bdaOkButton.pressed.connect(self.writeBDAConfig)
        self.bdaOkButton.released.connect(self.close)
        self.bdaOkButton.released.connect(window.showBDAToolTip)

        self.bdaOccAdd.setText('MD, DO, NP, PA')
        self.bdaLookUpsAdd.setText('31')
        self.bdaNumLookUpsAdd.setText('1')
        self.deDupAdd.setEnabled(False)
        self.bdaOnlyAdd.setEnabled(False)


    def writeBDAConfig(self):
        isTherapyChecked = self.therapyClassBoxAdd.isChecked()

        self.bdaConfig = dict()
        self.occList = str(self.bdaOccAdd.text()).split(", ")
        self.specList = str(self.bdaSpecAdd.text()).split(", ")
        self.bdaOcc2 = ", ".join('"{0}"'.format(w) for w in self.occList)
        self.bdaSpec2 = ", ".join('"{0}"'.format(w) for w in self.specList)


        self.checkTotalBdas()
        self.addBdaOcc = 'additonalBdaOcc'+'_'+str(window.bdainc)
        self.addBdaSpec = 'additonalBdaSpec'+'_'+str(window.bdainc)
        self.addDedup = 'additonalBdaDedup'+'_'+str(window.bdainc)
        self.addBdaNumLookUps = 'additonalBdaLookUps'+'_'+str(window.bdainc)
        self.addBdaLookUpPeriod = 'additonalBdaLookUpPeriod'+'_'+str(window.bdainc)
        self.addTherapyChecked = 'additonalBdatherapyChecked'+'_'+str(window.bdainc)
        self.addDrugList = 'additonalBdaDrugList'+'_'+str(window.bdainc)

        self.bdaConfig['totalAdditionalBDAs'] = window.bdainc
        self.bdaConfig[self.addBdaOcc] = str(self.bdaOcc2)
        self.bdaConfig[self.addBdaSpec] = str(self.bdaSpec2).replace('"Internal Medicine", "General"', '"Internal Medicine, General"').replace('"Pediatrics", "General"', '"Pediatrics, General"').replace('"Pediatrics", "Subspecialty"', '"Pediatrics, Subspecialty"').replace('"Pediatrics", "Surgical"', '"Pediatrics, Surgical"').replace('"Radiology", "General"', '"Radiology, General"').replace('"Radiology", "Vascular and Interventional"', '"Radiology, Vascular and Interventional"').replace('"Surgery", "Cardiac/Thoracic"', '"Surgery, Cardiac/Thoracic"').replace('"Surgery", "General"', '"Surgery, General"').replace('"Surgery", "Plastic"', '"Surgery, Plastic"').replace('"Surgery", "Transplant"', '"Surgery, Transplant"').replace('"Surgery", "Vascular"', '"Surgery, Vascular"')
        self.setBDACountLabel()

        if self.deDupAdd.isChecked():
            self.bdaConfig[self.addDedup] = str('Yes')
        if not self.deDupAdd.isChecked():
            self.bdaConfig[self.addDedup] = str('No')

        self.bdaConfig[self.addBdaNumLookUps] = str(self.bdaNumLookUpsAdd.text())
        self.bdaConfig[self.addBdaLookUpPeriod] = str(self.bdaLookUpsAdd.text())
        if isTherapyChecked:
            self.bdaConfig[self.addTherapyChecked] = str('Yes')
        if not isTherapyChecked:
            self.bdaConfig[self.addTherapyChecked] = str('No')
        if self.drugListAdd.toPlainText():
            self.bdaConfig[self.addDrugList] = str(self.createDrugListBDA())


        #Load BDA Configs and append them to the master config
        with open(desktop+'\\'+'bdaConfig.json', 'r') as infile:
            data = json.loads(infile.read(), encoding='utf8')
            data.update(self.bdaConfig)

        with open(desktop+'\\'+'bdaConfig.json', 'w') as outfile:
            json.dump(data, outfile, indent=2, sort_keys=True, encoding='utf8')

    def createDrugListBDA(self):
        self.drugContents = self.drugListAdd.toPlainText()
        commaCount = 0
        for char in self.drugContents:
            if char == ',':
                commaCount += 1
        if commaCount > 0:
            self.drugContents = str(self.drugContents).split(", ")
            self.drugContents = "\n".join(self.drugContents)
            return self.drugContents
        else:
            return self.drugContents

    def checkTotalBdas(self):
        with open(desktop+'\\'+'bdaConfig.json', 'r') as infile:
            config = json.loads(infile.read(), encoding='utf8')
            # if config['additonalSdaOcc'+'_'+str(window.bdainc)]:
            if 'additonalBdaOcc'+'_'+str(window.bdainc) in config:
                window.bdainc += 1

                # window.sdaLabel.setText('Additional SDAs: '+ window.bdainc)

    def setBDACountLabel(self):
        window.bdaLabel.setText('Additional BDAs: '+ str(window.bdainc))



class PivotPage(base_3, form_3):
    def __init__(self):
        super(base_3, self).__init__()
        self.setupUi(self)
        self.selectSeg1 = ""
        self.selectSeg2 = ""

        self.pivotSegments = self.findChild(QListWidget, 'pivotSegments_listWidget')
        self.pivSeg1 = self.findChild(QListWidget, 'seg1_listWidget')
        self.pivSeg2 = self.findChild(QListWidget, 'seg2_listWidget')
        self.pivOkButton = self.findChild(QPushButton, 'setPivot_pushButton')

        self.pivOkButton.clicked.connect(self.setPivotSegs)
        self.pivOkButton.clicked.connect(self.closeWindow)

        if self.pivotSegments:
            self.pivotSegments.clear()
        colList = utils.fetchColumns()
        for i in colList:
            self.pivotSegments.addItem(i)

    def setPivotSegs(self):
        for s1 in xrange(self.pivSeg1.count()):
            self.selectSeg1 = str(self.pivSeg1.item(s1).text())
            print self.selectSeg1
        for s2 in xrange(self.pivSeg2.count()):
            self.selectSeg2 = str(self.pivSeg2.item(s2).text())
            print self.selectSeg2

    def closeWindow(self):
        self.close()
        window.highlightPivotSegs()

class FilePage(base_2, form_2):
    def __init__(self):
        super(base_2, self).__init__()
        self.setupUi(self)


        self.selectMatchFile = ""
        self.selectSuppFile = ""
        self.suppMatchType = ""
        self.suppSasFileInput = ""
        # self.fileDict = defaultdict(list)

        self.downloadFiles = self.findChild(QListWidget, 'files_listWidget')
        self.listMatchFile = self.findChild(QListWidget, 'listMatchFile_listWidget')
        self.suppMatchFile = self.findChild(QListWidget, 'suppressionFile_listWidget')
        self.setFilesButton = self.findChild(QPushButton, 'setFIles_pushButton')
        self.standardSuppCheck = self.findChild(QCheckBox, 'suppStandardCheck_checkBox')
        self.exactSuppCheck = self.findChild(QCheckBox, 'suppExactCheck_checkBox')
        self.suppModel = self.downloadFiles.model()
        self.suppSheetCountLabel = self.findChild(QLabel, 'suppSheetCount_Label')
        self.suppSasFile = self.findChild(QLineEdit, 'suppSAS_lineEdit')


        self.setFilesButton.pressed.connect(self.setMatchFile)
        self.setFilesButton.pressed.connect(self.setSuppFile)
        self.setFilesButton.released.connect(self.refreshFrom2Files)
        self.setFilesButton.released.connect(self.close)
        self.standardSuppCheck.toggled.connect(self.suppFileCheckboxesStandard)
        self.exactSuppCheck.toggled.connect(self.suppFileCheckboxesExact)
        self.suppModel.rowsRemoved.connect(self.suppSheetCount)



        for i in downloadFilesDirectory:
            self.downloadFiles.addItem(i)

    def suppSheetCount(self):
        for i in xrange(self.suppMatchFile.count()):
            self.selectSuppFile = str(self.suppMatchFile.item(i).text())
        sheetFile = self.selectSuppFile
        filename, extension = os.path.splitext(os.path.join(downloads, sheetFile))
        if extension == '.xlsx':
            wb = xlrd.open_workbook(downloads + sheetFile, on_demand=True)
            totalSheets = len(wb.sheet_names())
            self.suppSheetCountLabel.setText('Sheet Count: '+str(totalSheets))
            if totalSheets == 1:
                self.suppSheetCountLabel.setStyleSheet('color: green')
            else:
                self.suppSheetCountLabel.setStyleSheet('color: red')
        else:
            self.suppSheetCountLabel.setText('Sheet Count: N/A')
            self.suppSheetCountLabel.setStyleSheet('color: black')        

    def setMatchFile(self):
        self.selectMatchFile = ""
        for i in xrange(self.listMatchFile.count()):
            self.selectMatchFile = str(self.listMatchFile.item(i).text())
            print 'Your Match File is', self.selectMatchFile
            window.setWindowTitle("File Loaded: "+self.selectMatchFile)
            window.loadedFile = self.selectMatchFile
            window.countSheets()
            # return self.selectMatchFile
            # self.fileDict['matchFile'].append(self.selectMatchFile)

    def setSuppFile(self):
        for i in xrange(self.suppMatchFile.count()):
            self.selectSuppFile = str(self.suppMatchFile.item(i).text())
            print 'Your Suppression File is', self.selectSuppFile
        if not self.selectSuppFile:
            self.suppSasFileInput = str(self.suppSasFile.text())
            print 'Your using a _supp_####: ', self.suppSasFileInput
        if self.standardSuppCheck.isChecked():
            self.suppMatchType = "Standard"
        else:
            self.suppMatchType = "Exact"
        # return self.selectSuppFile
            # self.fileDict['suppFile'].append(self.selectSuppFile)

    def refreshFrom2Files(self):
        #remove files and clear current column names
        if os.path.exists(os.path.join(downloads, 'csvFile.csv')):
            os.remove(os.path.join(downloads, 'csvFile.csv'))
        if os.path.exists(os.path.join(downloads, 'target.csv')):
            os.remove(os.path.join(downloads, 'target.csv'))
        if os.path.exists(os.path.join(downloads, 'target_mod.csv')):
            os.remove(os.path.join(downloads, 'target_mod.csv'))
        if window.sourceSegs:
            window.sourceSegs.clear()
            window.finalSegs.clear()
        utils.checkExtension2(self.selectMatchFile)
        utils.removeChar()
        utils.incDupColumns()

        #populates the listWidget
        colList = utils.fetchColumns()
        for i in colList:
            window.sourceSegs.addItem(i)
        window.returnUserTable() 
        window.countFullFile()
        window.resetEditTab()
        window.highlightSourceSegs()
        window.detectMatchType()



    def suppFileCheckboxesStandard(self):
        isStandardSuppChecked = self.standardSuppCheck.isChecked()

        if isStandardSuppChecked:
            self.standardSuppCheck.setStyleSheet("color: green; font-weight: bold")
            self.exactSuppCheck.setChecked(False)
        elif not isStandardSuppChecked:
            self.exactSuppCheck.setChecked(True)
            self.standardSuppCheck.setStyleSheet("color: black")

    def suppFileCheckboxesExact(self):
        isExactSuppChecked = self.exactSuppCheck.isChecked()

        if isExactSuppChecked:
            self.exactSuppCheck.setStyleSheet("color: green; font-weight: bold")
            self.standardSuppCheck.setChecked(False)
        elif not isExactSuppChecked:
            self.standardSuppCheck.setChecked(True)
            self.exactSuppCheck.setStyleSheet("color: black")

 
class MyApp(QtGui.QMainWindow, Ui_MainWindow):
    def __init__(self):
        QtGui.QMainWindow.__init__(self)
        Ui_MainWindow.__init__(self)
        self.setupUi(self)
        self.setWindowTitle("File Loaded: "+newest)
        # self.config = dict()
        self.sdainc = 1
        self.bdainc = 1

        #clean masterConfig File
        with open(desktop+'\\'+'config.json', 'r') as infile:
            configCheck = json.loads(infile.read(), encoding='utf8')
            for element in configCheck.keys():
                # print element
                if element.startswith('additonal'):
                    configCheck.pop(element, None)

        with open(desktop+'\\'+'config2.json', 'w') as outfile:
            json.dump(configCheck, outfile, indent=2, sort_keys=True, encoding='utf8')

        os.remove(desktop+'\\'+'config.json')
        os.rename(desktop+'\\'+'config2.json', desktop+'\\'+'config.json')


        #clean sdaConfig File
        with open(desktop+'\\'+'sdaConfig.json', 'w') as firstFile:
            data = {'totalAdditionalSDAs': 0}
            # data['totalAdditionalSDAs'] = 0
            json.dump(data, firstFile, indent=2, sort_keys=True, encoding='utf8')

        with open(desktop+'\\'+'sdaConfig.json', 'r') as infile:
            configCheck = json.loads(infile.read(), encoding='utf8')
            for element in configCheck.keys():
                # print element
                if element.startswith('additonal'):
                    configCheck.pop(element, None)
                if element.startswith('totalAdditional'):
                    configCheck[element] = 0

        with open(desktop+'\\'+'sdaConfig2.json', 'w') as outfile:
            json.dump(configCheck, outfile, indent=2, sort_keys=True, encoding='utf8')

        os.remove(desktop+'\\'+'sdaConfig.json')
        os.rename(desktop+'\\'+'sdaConfig2.json', desktop+'\\'+'sdaConfig.json')

        #clean bdaConfig File
        with open(desktop+'\\'+'bdaConfig.json', 'w') as firstFile:
            data = {'totalAdditionalBDAs': 0}
            # data['totalAdditionalSDAs'] = 0
            json.dump(data, firstFile, indent=2, sort_keys=True, encoding='utf8')

        with open(desktop+'\\'+'bdaConfig.json', 'r') as infile:
            configCheck = json.loads(infile.read(), encoding='utf8')
            for element in configCheck.keys():
                # print element
                if element.startswith('additonal'):
                    configCheck.pop(element, None)
                if element.startswith('totalAdditional'):
                    configCheck[element] = 0

        with open(desktop+'\\'+'bdaConfig2.json', 'w') as outfile:
            json.dump(configCheck, outfile, indent=2, sort_keys=True, encoding='utf8')

        os.remove(desktop+'\\'+'bdaConfig.json')
        os.rename(desktop+'\\'+'bdaConfig2.json', desktop+'\\'+'bdaConfig.json')


        self.dateButton = self.findChild(QToolButton, 'dateButton')
        self.calendarLine = self.findChild(QLineEdit, 'date_lineEdit')

        #QTabWidget OBJECTS
        self.tabWidget = self.findChild(QTabWidget, 'tabWidget')
        self.listMatchTab = self.findChild(QWidget, 'listMatchTab')
        self.targetingTab = self.findChild(QWidget, 'targetingTab')

        #Menubar
        self.mainMenu = self.findChild(QMenuBar, 'menuBar')
        self.fileMenu = self.findChild(QMenu, 'menuFile')
        self.changeLogMenu = self.findChild(QAction, 'actionChange_Log')

        #LISTMATCH TAB OBJECTS
        self.sdaOcc = TextEdit(self.findChild(QLineEdit, 'sdaOcc_line_edit'))
        # self.sdaOcc = TextEdit(parent=self.findChild(QLineEdit, 'sdaOcc_line_edit'))
        self.outputFolder = self.findChild(QLineEdit, 'outputFolder')
        self.sdaSpec = TextEdit(self.findChild(QLineEdit, 'sdaSpec_line_edit'))
        self.sdaAddOn = self.findChild(QCheckBox, 'sdaAddOn_check_box')
        self.bdaAddOn = self.findChild(QCheckBox, 'bdaAddOn_check_box')
        self.bdaOcc = TextEdit2(self.findChild(QLineEdit, 'bdaOcc_line_edit'))
        self.bdaSpec = TextEdit2(self.findChild(QLineEdit, 'bdaSpec_line_edit'))
        self.bdaLookUps = self.findChild(QLineEdit, 'bdaLookUps_line_edit')
        self.bdaNumLookUps = self.findChild(QLineEdit, 'bdaNumLookUps_line_edit')
        self.deDup = self.findChild(QCheckBox, 'deDup_check_box')
        self.sourceSegs = self.findChild(QListWidget, 'sourceSegs_list_widget')
        self.finalSegs = self.findChild(QListWidget, 'finalSegs_list_widget')
        self.addSegButton = self.findChild(QPushButton, 'addSegButton_push_btn')
        self.removeSegButton = self.findChild(QPushButton, 'removeSegButton_push_btn')
        self.drugList = self.findChild(QPlainTextEdit, 'drugList_text_edit')
        self.segBox = self.findChild(QCheckBox, 'segBox_check_box')
        self.runProgramButton = self.findChild(QPushButton, 'runProgramButton')
        self.brandName = self.findChild(QLineEdit, 'brandName_line_edit')
        self.caseOptions = self.findChild(QGroupBox, 'caseOptionsGroupBox')
        self.addOns = self.findChild(QGroupBox, 'addOnGroupBox')
        self.segmentationInfo = self.findChild(QGroupBox, 'segmentGroupBox')
        self.presalesManuName = self.findChild(QLineEdit, 'presalesManuName_line_edit')
        self.standardMatch = self.findChild(QCheckBox, 'standardMatch_check_box')
        self.exactMatch = self.findChild(QCheckBox, 'exactMatch_check_box')
        self.listMatchDetails = self.findChild(QGroupBox, 'listMatchGroupBox')
        self.emailRecip = self.findChild(QLineEdit, 'emailTo_line_edit')
        self.caseNumber = self.findChild(QLineEdit, 'caseNumber_line_edit')
        self.matchTypeSelection = self.findChild(QComboBox, 'emailMatchType_comboBox')
        self.saleExecutive = self.findChild(QLineEdit, 'saleExecutive_line_edit')
        self.postgresTable = self.findChild(QLineEdit, 'postgresTable_line_edit')
        self.docAlert = self.findChild(QCheckBox, 'docAlert_check_box')
        self.epocQuiz = self.findChild(QCheckBox, 'epocQuiz_check_box')
        self.myInIt = self.findChild(QLineEdit, 'myInit_line_edit')
        self.requesterInIt = self.findChild(QLineEdit, 'requestInit_line_edit')
        self.refreshButton = self.findChild(QPushButton, 'refreshButton')
        self.previousSettings = self.findChild(QCheckBox, 'previousSettings_check_box')
        self.tableNameWarning = self.findChild(QLabel, 'tableLabel')
        self.suppMatchPath = self.findChild(QLineEdit, 'suppFilePath_lineEdit')
        self.suppCheck = self.findChild(QCheckBox, 'suppression_checkBox')
        self.bdaOnly = self.findChild(QCheckBox, 'bdaOnly_checkBox')
        self.sdaOnly = self.findChild(QCheckBox, 'sdaOnly_checkBox')
        self.refreshLabel = self.findChild(QLabel, 'refreshLabel_QLabel')
        self.fuzzyBox = self.findChild(QCheckBox, 'fuzzy_checkBox')
        # self.testTextEdit = MyHighlighter(self.findChild(QTextEdit, 'Test_textEdit'), "classic")
        self.loadAction = self.findChild(QAction, 'actionLoad_List')
        self.therapyClassBox = self.findChild(QCheckBox, 'therapyClass_checkBox')
        self.sheetCount = self.findChild(QLabel, 'sheetCount_Label')
        self.dataCap = self.findChild(QLineEdit, 'dataCap_lineEdit')
        self.pivotTableCheck = self.findChild(QCheckBox, 'pivotCheck_checkBox')
        self.addSDAButton = self.findChild(QPushButton, 'AdditionalSDA_pushButton')
        self.addBDAButton = self.findChild(QPushButton, 'AdditionalBDA_pushButton')
        self.sdaLabel = self.findChild(QLabel, 'addSDA_Label')
        self.bdaLabel = self.findChild(QLabel, 'addBDA_Label')
        self.suppSDAOnly = self.findChild(QCheckBox, 'supressSDA_checkBox')
        self.suppBDAOnly = self.findChild(QCheckBox, 'supressBDA_checkBox')

        #Targeting Objects
        self.targetManuName = self.findChild(QComboBox, 'targetManuName_comboBox')
        self.targetNumber = self.findChild(QLineEdit, 'targetNumber_line_edit')
        self.brandTargetName = self.findChild(QLineEdit, 'brandName_line_edit_2')
        self.postgresTargetTable = self.findChild(QLineEdit, 'postgresTable_line_edit_2')
        self.myTargetInIt = self.findChild(QLineEdit, 'myInit_line_edit_2')
        self.segmentList = self.findChild(QCheckBox, 'segmentList_checkBox')
        self.keepSeg = self.findChild(QCheckBox, 'keepSeg_checkBox')
        self.segVariable = self.findChild(QLineEdit, 'segVariable_lineEdit')
        self.segValues = self.findChild(QLineEdit, 'segValues_lineEdit')
        self.sdaTargetNum = self.findChild(QLineEdit, 'sdaTarget_lineEdit')
        self.bdaTargetNum = self.findChild(QLineEdit, 'bdaTarget_lineEdit')
        self.cmiCompass = self.findChild(QCheckBox, 'cmi_compass_checkBox')
        self.tableNameWarningTargeting = self.findChild(QLabel, 'targetTableName')
        self.backFillBox = self.findChild(QCheckBox, 'backFill_checkBox')

        #Edit List Objects
        self.uniqueValuesList = self.findChild(QListWidget, 'uniqueValues_listWidget')
        self.getUniqueValuesButton = self.findChild(QPushButton, 'getUniqueValues_pushButton')
        self.renameValueEdit = self.findChild(QLineEdit, 'renameValue_lineEdit')
        self.renameButton = self.findChild(QPushButton, 'renameButton_pushButton')
        self.newColumnName = self.findChild(QLineEdit, 'newColumnEdit_lineEdit')
        self.renameColumnButton = self.findChild(QPushButton, 'renameColumn_pushButton')
        self.createFilter = self.findChild(QPushButton, 'createFilter_pushButton')
        self.clearSelections = self.findChild(QPushButton, 'resetEditTab_pushButton')
        self.addColumnText = self.findChild(QLineEdit, 'addColumn_lineEdit')
        self.addColumnButton = self.findChild(QPushButton, 'addColumn_pushButton')
        self.deleteColumnButton = self.findChild(QPushButton, 'deleteColumn_pushButton')
        self.concatButton = self.findChild(QPushButton, 'concat_pushButton')
        self.concatenateList = self.findChild(QListWidget, 'concatList_listWidget')
        self.fullListCount = self.findChild(QLabel, 'itemCount_QLabel')
        self.uniqueValuesList.setContextMenuPolicy(QtCore.Qt.CustomContextMenu)

        #console Objects
        self.consoleOutput = self.findChild(QTextEdit, 'console_textEdit')

        # self.frwon = self.findChild(QWebView, 'frwon_webView')
        # self.frwon.load(QUrl("http://172.18.13.10/analytics/saw.dll?bieehome"))
        # self.frwon.load(QUrl("https://docs.google.com/spreadsheets/d/1go5cjJ7ORoJCR4bNV2pJZU3nkwHAJbOAor0hBbgp9ag/edit#gid=353317686"))

        # self.frwon.load(QUrl("https://epocrates.my.salesforce.com/500?fcf=00Ba0000009bM2m"))
        # self.frwon.show()

        self.dataSharingDetails = self.findChild(QTableView, 'target_tableView')
        self.model = QtGui.QStandardItemModel(self)
        self.model.setHorizontalHeaderLabels(self.getHeaders())

        # self.dataSharingDetails = QtWidgets.QTableView(self)
        self.dataSharingDetails.setModel(self.model)
        self.dataSharingDetails.horizontalHeader().setStretchLastSection(True)

        # self.dataSharingDetails.setSpan(0, 0, 1, 17)
        # self.dataSharingDetails.show()
        # self.dataSharingDetails.setHorizonalHeaderLabels(self.getHeaders())

        self.sourceSegs.setSelectionMode(QtGui.QAbstractItemView.ExtendedSelection)
        self.finalSegs.setSelectionMode(QtGui.QAbstractItemView.ExtendedSelection)
        self.uniqueValuesList.setSelectionMode(QtGui.QAbstractItemView.ExtendedSelection)

        #On startup attach call backs and set default toggle options for enabled widgets
        self.setDefaults()
        self.attachCallbacks()
        self.getPostgresTables()
        #On start up build csv files needed to populate the Segmentation column names
        # utils.csv_from_excel()
        utils.removeChar()
        utils.incDupColumns()

        #populates the listWidget
        if self.sourceSegs:
            self.sourceSegs.clear()
        colList = utils.fetchColumns()
        for i in colList:
            self.sourceSegs.addItem(i)

        self.highlightSourceSegs()

        self.refreshButton.pressed.connect(self.refreshFile)
        self.refreshButton.released.connect(self.refreshSuccessful)
        self.docAlert.toggled.connect(self.docalertCallback)
        self.epocQuiz.toggled.connect(self.epocquizCallback)
        self.segBox.toggled.connect(self.segmentCallBack)
        self.standardMatch.toggled.connect(self.standardCallback)
        self.exactMatch.toggled.connect(self.exactCallback)
        
        # self.tabWidget.currentChanged.connect(self.targetCaseCallback)
        self.sdaAddOn.toggled.connect(self.sdaCallback)
        self.bdaAddOn.toggled.connect(self.bdaCallback)
        self.fuzzyBox.toggled.connect(self.fuzzyCallback)
        self.addSegButton.clicked.connect(self.addSeg)
        self.removeSegButton.clicked.connect(self.removeSeg)
        self.runProgramButton.clicked.connect(self.runProgram)
        self.previousSettings.toggled.connect(self.usePreviousConfig)
        self.dateButton.clicked.connect(self.showCalWid)
        self.tabWidget.currentChanged.connect(self.bdaCallback)
        self.tabWidget.currentChanged.connect(self.sdaCallback)
        self.tabWidget.currentChanged.connect(self.returnUserTable)
        self.targetManuName.currentIndexChanged.connect(self.dataSharingClients)
        self.targetManuName.currentIndexChanged.connect(self.highlightCMI)
        self.getUniqueValuesButton.clicked.connect(self.getUniqueSegmentValues)
        self.renameButton.clicked.connect(self.renameValue)
        self.renameButton.clicked.connect(self.bajabThread)
        self.renameColumnButton.clicked.connect(self.renameColumn)
        self.cmiCompass.toggled.connect(self.dataSharingClients)
        self.cmiCompass.toggled.connect(self.cmiCompasCallback)
        self.postgresTable.textChanged.connect(self.returnUserTable)
        self.postgresTargetTable.textChanged.connect(self.returnUserTable)
        self.createFilter.clicked.connect(self.filterList)
        self.addColumnButton.clicked.connect(self.addColumn)
        self.deleteColumnButton.clicked.connect(self.deleteColumn)
        self.tabWidget.currentChanged.connect(self.editListCallback)
        self.cmiCompass.toggled.connect(self.highlightCMI)
        # self.suppCheck.toggled.connect(self.suppressionCallback)
        self.tabWidget.currentChanged.connect(self.showFrwon)
        self.clearSelections.clicked.connect(self.resetEditTab)
        self.concatButton.clicked.connect(self.concatColumns)
        self.sourceSegs.itemClicked.connect(self.previewConcat)
        self.tabWidget.currentChanged.connect(self.countFullFile)
        self.uniqueValuesList.itemClicked.connect(self.countSelectedValues)
        self.sdaOnly.toggled.connect(self.sdaOnlyCallback)
        self.bdaOnly.toggled.connect(self.sdaOnlyCallback)
        self.tabWidget.currentChanged.connect(self.sdaOnlyCallback)
        self.presalesManuName.textChanged.connect(self.gskCallback)
        self.segmentList.toggled.connect(self.segmentedListCallback)
        self.uniqueValuesList.customContextMenuRequested.connect(self.listItemRightClicked)
        self.suppCheck.toggled.connect(self.showFileSelection)
        # self.pivotTableCheck.toggled.connect(self.showPivotSelection)
        self.loadAction.triggered.connect(self.getfiles)
        self.sdaOnly.toggled.connect(self.set_SDA_BDA_Only_colors)
        self.bdaOnly.toggled.connect(self.set_SDA_BDA_Only_colors)
        self.suppMatchPath.textChanged.connect(self.set_SDA_BDA_Only_colors)
        self.sdaBdaNone = [self.caseNumber, self.requesterInIt, self.presalesManuName, self.saleExecutive, self.brandName, self.myInIt]
        for obj in self.sdaBdaNone:
            obj.textChanged.connect(self.set_SDA_BDA_Only_colors)

        self.pivotTableCheck.toggled.connect(self.pivotWarning)
        self.segBox.toggled.connect(self.threadSong)
        self.sdaAddOn.toggled.connect(self.exitsThread)
        self.runProgramButton.released.connect(self.trashThread)

        self.addSDAButton.pressed.connect(self.additionalSDA)
        self.addBDAButton.pressed.connect(self.additionalBDA)

        self.sdaLabel.setToolTip(self.showSDAToolTip())
        self.bdaLabel.setToolTip(self.showBDAToolTip())
        self.therapyClassBox.toggled.connect(self.sda_bda_options)
        self.suppSDAOnly.toggled.connect(self.sda_bda_options)
        self.suppBDAOnly.toggled.connect(self.sda_bda_options)

        self.loadedFile = newest
        self.countSheets()
        print 'Gui Started'

    def showSDAToolTip(self):
        with open(desktop+'\\'+'sdaConfig.json', 'r') as infile:
            configCheck = json.loads(infile.read(), encoding='utf8')
            totalCodesBuilt = 1
            toolTipString = ""
            totalAdditionalSDAs = int(configCheck['totalAdditionalSDAs'])
            underlineText = "<span style=\"text-decoration:underline\" >"
            endSpan = "</span>"
            pageBreak = ""
            if configCheck['totalAdditionalSDAs'] == 0:
                return 'There are 0 Additonal Sda\'s Add Ons'
            else:
                while totalCodesBuilt <= totalAdditionalSDAs:
                    occupation2 = str(configCheck['additonalSdaOcc_'+str(totalCodesBuilt)]).replace('"', '')
                    dispOcc = '<b>Occupations:</b> {}{}'.format(occupation2, '\n')
                    specialty2 = str(configCheck['additonalSdaSpec_'+str(totalCodesBuilt)]).replace('"', '')
                    dispSpec = '<b>Specialties:</b> {}{}'.format(specialty2, '\n')
                    occLabel = "<b>Additional SDA {}:</b>{}".format(str(totalCodesBuilt),'\n')
                    newLine = '\n'
                    totalCodesBuilt += 1
                    toolTipString += '<p>'+underlineText+occLabel+endSpan+'<br>'
                    toolTipString += dispOcc+'<br>'
                    toolTipString += dispSpec+'<br>'+'</p>'
                    # toolTipString += newLine+'<p>'
                self.sdaLabel.setToolTip(toolTipString)

    def showBDAToolTip(self):
        with open(desktop+'\\'+'bdaConfig.json', 'r') as infile:
            configCheck = json.loads(infile.read(), encoding='utf8')
            totalCodesBuilt = 1
            toolTipString = ""
            totalAdditionalBDAs = int(configCheck['totalAdditionalBDAs'])
            underlineText = "<span style=\"text-decoration:underline\" >"
            endSpan = "</span>"
            if configCheck['totalAdditionalBDAs'] == 0:
                return 'There are 0 Additonal Bda\'s Add Ons'
            else:
                while totalCodesBuilt <= totalAdditionalBDAs:
                    occLabel = "<b>Additional BDA {}</b>{}".format(str(totalCodesBuilt),'\n')
                    occupation2 = str(configCheck['additonalBdaOcc_'+str(totalCodesBuilt)]).replace('"', '')
                    dispOcc = '<b>Occupations:</b> {}{}'.format(occupation2, '\n')
                    specialty2 = str(configCheck['additonalBdaSpec_'+str(totalCodesBuilt)]).replace('"', '')
                    dispSpec = '<b>Specialties:</b> {}{}'.format(specialty2, '\n')
                    lookUps = str(configCheck['additonalBdaLookUps_'+str(totalCodesBuilt)])
                    dispLookUps = '<b># of lookUps: </b>{}{}'.format(lookUps, '\n')
                    lookUpPeriod = str(configCheck['additonalBdaLookUpPeriod_'+str(totalCodesBuilt)])
                    dispLookUpPeriod = '<b>LookUp Period: </b>{}{}'.format(lookUpPeriod, '\n')

                    drugs = str(configCheck['additonalBdaDrugList_'+str(totalCodesBuilt)])
                    joinDrugs = "".join(drugs)
                    splitDrugs = joinDrugs.split('\n')
                    totalDrugs = len(splitDrugs)
                    print totalDrugs

                    shortenedList = []
                    myDrugs = ""

                    count15 = False
                    if totalDrugs >= 15:
                        drugCount = 1
                        for d in splitDrugs:
                            if drugCount <= 15:
                                myDrugs += d+'<br>'
                                drugCount += 1
                            else:
                                myDrugs += '+ '+str(totalDrugs - 15)+' More Drugs . . .'
                                break
                    
                        print myDrugs

                        drugs = myDrugs
                        dispDrugs = '<b>Drug List: </b>{}{}</p>'.format(drugs, '\n')
                    else:
                        drugs = str(configCheck['additonalBdaDrugList_'+str(totalCodesBuilt)]).replace('\n', '<br>')
                        dispDrugs = '<b>Drug List: </b>{}{}</p>'.format(drugs, '\n')
                
                    # newLine = '\n'
                    totalCodesBuilt += 1
                    toolTipString += '<p>'+underlineText+occLabel+endSpan+'<br>'
                    toolTipString += dispOcc+'<br>'
                    toolTipString += dispSpec+'<br>'
                    toolTipString += dispLookUps+'<br>'
                    toolTipString += dispLookUpPeriod+'<br>'
                    toolTipString += dispDrugs
                    # toolTipString += newLine
                self.bdaLabel.setToolTip(toolTipString)



        # self.sdaOcc.show()
        # self.sdaSpec.show()
        # self.bdaOcc.show()
        # self.bdaSpec.show()
        # self.drugList.show()
        # self.playHello() 

    def bajabThread(self):
        thread1 = Thread(target = self.playBajab)
        thread1.start()

    def playBajab(self):
        winsound.PlaySound(os.path.join(desktop, 'Ewok\\Sounds\\Bajab.wav'), winsound.SND_FILENAME)

    def playHello(self):
        winsound.PlaySound(os.path.join(desktop, 'Ewok\\Sounds\\Hello.wav'), winsound.SND_FILENAME)


    def threadSong(self):
        if self.segBox.isChecked():
            thread1 = Thread(target = self.playSong)
            thread1.start()

    def playSong(self):
        if self.segBox.isChecked():
            winsound.PlaySound(os.path.join(desktop, 'Ewok\\Sounds\\Uhhh.wav'), winsound.SND_FILENAME)

    def exitsThread(self):
        thread1 = Thread(target = self.playExits)
        thread1.start()

    def playExits(self):
        if self.sdaAddOn.isChecked():
                winsound.PlaySound(os.path.join(desktop, 'Ewok\\Sounds\\skipExits.wav'), winsound.SND_FILENAME)

    def goyThread(self):
        thread1 = Thread(target = self.playGoy)
        thread1.start()

    def playGoy(self):
        winsound.PlaySound(os.path.join(desktop, 'Ewok\\Sounds\\Goy.wav'), winsound.SND_FILENAME)

    def trashThread(self):
        thread1 = Thread(target = self.playTrash)
        thread1.start()

    def playTrash(self):
        winsound.PlaySound(os.path.join(desktop, 'Ewok\\Sounds\\Trash.wav'), winsound.SND_FILENAME)

    def highlightPivotSegs(self):
        self.seg1 = self.pivotWindow.selectSeg1
        self.seg2 = self.pivotWindow.selectSeg2
        notInFinal = []
        foundPivot = []
        pivotList = [self.seg1, self.seg2]


        for item in xrange(self.finalSegs.count()):
            col = self.finalSegs.item(item)
            # col.setBackground(QColor(213, 233, 221,255))
            self.pivotValue = str(self.finalSegs.item(item).text())
            if self.pivotValue in [self.seg1, self.seg2]:
                foundPivot.append(self.pivotValue)
                col.setBackground(QColor(213, 233, 221,255))
            else:
                notInFinal.append(self.pivotValue)
            # if self.seg1 in notInFinal:
            #     seg2Bad = self.

        for item in pivotList:
            if item not in foundPivot:
                msg = QMessageBox()
                msg.setIcon(QMessageBox.Warning)
                msg.setText("You Dumbass You selected a Pivot Segment \"{}\" that's not even in your Final Segments List".format(item))

                msg.setStandardButtons(QMessageBox.Ok)

                msg.exec_()             

    def pivotWarning(self):
        isPivotChecked = self.pivotTableCheck.isChecked()
        self.finalSegments = ""
        if isPivotChecked:
            if len(self.finalSegs) == 0:
                msg = QMessageBox()
                msg.setIcon(QMessageBox.Warning)
                msg.setText("Please add source segments first to proceed to Pivot Options")

                msg.setStandardButtons(QMessageBox.Ok)

                msg.exec_()
            else:
                self.showPivotSelection()
        else:
            self.pivotTableCheck.setStyleSheet("color: black")

    def showPivotSelection(self):
        isPivotChecked = self.pivotTableCheck.isChecked()
        if isPivotChecked:
            self.pivotTableCheck.setStyleSheet("color: green; font-weight: bold")
            self.pivotWindow = PivotPage()
            self.pivotWindow.show()

    def additionalSDA(self):
        self.addSDA = SDA_Widget()
        self.addSDA.show()

    def additionalBDA(self):
        self.addBDA = BDA_Widget()
        self.addBDA.show()

    def countSheets(self):
        sheetFile = self.loadedFile
        filename, extension = os.path.splitext(os.path.join(downloads, sheetFile))
        if extension == '.xlsx':
            wb = xlrd.open_workbook(downloads + sheetFile, on_demand=True)
            totalSheets = len(wb.sheet_names())
            self.sheetCount.setText('Sheet Count: '+str(totalSheets))
            if totalSheets == 1:
                self.sheetCount.setStyleSheet('color: green')
            else:
                self.sheetCount.setStyleSheet('color: red')
        else:
            self.sheetCount.setText('Sheet Count: N/A')
            self.sheetCount.setStyleSheet('color: black')


    def getfiles(self):
        dlg = QFileDialog()
        dlg.setDirectory(downloads)
        dlg.setFileMode(QFileDialog.AnyFile)
        dlg.setFilter("List Match File (*.txt *.xlsx *.csv)")
        filenames = QStringList()
        
        if dlg.exec_():
            filenames = dlg.selectedFiles()
            filenames = str(filenames[0]).rsplit('/', 1)[-1]
            self.setWindowTitle("File Loaded: "+filenames)
            self.loadedFile = filenames
            self.countSheets()
            # f = open(filenames[0], 'r')
            if os.path.exists(os.path.join(downloads, 'csvFile.csv')):
                os.remove(os.path.join(downloads, 'csvFile.csv'))
            if os.path.exists(os.path.join(downloads, 'target.csv')):
                os.remove(os.path.join(downloads, 'target.csv'))
            if os.path.exists(os.path.join(downloads, 'target_mod.csv')):
                os.remove(os.path.join(downloads, 'target_mod.csv'))
            if self.sourceSegs:
                self.sourceSegs.clear()
                self.finalSegs.clear()
            utils.checkExtension2(filenames)
            utils.removeChar()
            utils.incDupColumns()

            colList = utils.fetchColumns()
            for i in colList:
                self.sourceSegs.addItem(i)
            self.returnUserTable() 
            self.countFullFile()
            self.resetEditTab()
            self.highlightSourceSegs()
            self.detectMatchType()


    def showFrwon(self):
        fileName = os.path.join(desktop, 'Ewok', 'targetSpecs.csv')
        with open(fileName, "rb") as fileInput:
            reader = csv.reader(fileInput)
            dsharingHeaders = reader.next()
            for row in reader:    
                items = [QtGui.QStandardItem(field) for field in row]
                self.model.appendRow(items)

        self.dataSharingDetails.setSpan(0, 0, 17, 1)
        self.dataSharingDetails.setSpan(0, 2, 17, 1)
        self.dataSharingDetails.setSpan(0, 4, 17, 1)

        self.dataSharingDetails.setSpan(17, 0, 22, 1)
        self.dataSharingDetails.setSpan(17, 2, 22, 1)
        self.dataSharingDetails.setSpan(17, 4, 22, 1)
        self.dataSharingDetails.show()


    def getHeaders(self):
        fileName = os.path.join(desktop, 'Ewok', 'targetSpecs.csv')
        with open(fileName, "rb") as fileInput:
            self.reader = csv.reader(fileInput)
            self.dsharingHeaders = self.reader.next()
        return self.dsharingHeaders

    def listItemRightClicked(self, QPos): 
        self.listMenu= QtGui.QMenu()
        menu_item = self.listMenu.addAction("Push to Targeting Form")
        menu_item.triggered.connect(self.menuItemClicked)
        parentPosition = self.uniqueValuesList.mapToGlobal(QtCore.QPoint(0, 0))        
        self.listMenu.move(parentPosition + QPos)
        self.listMenu.show()

    def menuItemClicked(self):
        self.segValues.clear()
        self.segVariable.clear()
        selectedSegment = [index.row() for index in self.sourceSegs.selectedIndexes()]
        segValueList = []
        for i in xrange(self.uniqueValuesList.count()):
            segValueList.append(str(self.uniqueValuesList.item(i).text()))
        for i in selectedSegment:
            self.segVariable.setText(str(self.sourceSegs.item(i).text()))

        valueString = ", ".join(segValueList)
        self.segValues.setText(valueString.replace(', ', ' '))


    def highlightSourceSegs(self):
        totalBlanks = 0
        blankColumnsList = []
        for i in xrange(self.sourceSegs.count()):
            blankColumnsList.append(str(self.sourceSegs.item(i).text()))
            for val in blankColumnsList:
                if val == '':
                    totalBlanks += 1
            if len(self.sourceSegs) < 4 or totalBlanks > 1:
                self.sourceSegs.setStyleSheet('background-color: rgb(247, 243, 117,255);')
                if self.sourceSegs.item(i).text() == 'No_File_Found':
                    self.sourceSegs.setStyleSheet('background-color: rgb(223, 138, 138, 255);')

            else:
                self.sourceSegs.setStyleSheet("")

    def detectMatchType(self):
        foundColumns = []
        for i in xrange(self.sourceSegs.count()):
            checkCol =  str(self.sourceSegs.item(i).text()).lower()
            col = self.sourceSegs.item(i)

            if checkCol == 'npi' or checkCol == 'npi_id' or re.search('.+ npi .+', checkCol) or re.search('.+_npi_.+', checkCol) or re.search('^npi.+', checkCol) or re.search('.+ npi', checkCol) or re.search('.+ npi', checkCol):
                if 'npi' not in foundColumns:
                    foundColumns.append('npi')

            elif checkCol == 'me' or checkCol == 'me_id' or checkCol == 'meded' or checkCol == 'menum' or re.search('.+ me .+', checkCol) or  re.search('.+_me_.+', checkCol) or re.search('^me .+', checkCol) or re.search('.+ me', checkCol) or re.search('.+ me', checkCol) or re.search('^me_.+', checkCol):
                if 'me' not in foundColumns:
                    foundColumns.append('me')

            elif checkCol == 'fname' or re.search('^first.+name', checkCol) or re.search('.+first.+name', checkCol) or re.search('.+fname', checkCol) or re.search('.+first', checkCol) or re.search('.+frst.+', checkCol):
                if 'fname' not in foundColumns:
                    foundColumns.append('fname')

            elif re.search(r'^lname|^last.+name|.+last|.+last.+name', checkCol):
                if 'lname' not in foundColumns:
                    foundColumns.append('lname')

            elif checkCol == 'zip' or checkCol == 'Postal' or (re.search('^zip.+', checkCol) and (checkCol != 'zip_4' or checkCol != 'zip4')) or re.search('^postal.+', checkCol) or re.search('.+_zip', checkCol) or re.search('.+ zip', checkCol) or re.search('.+_postal', checkCol) or re.search('.+ zip', checkCol) or re.search('.+ postal', checkCol):
                if 'zip' not in foundColumns:
                    foundColumns.append('zip')
        # print foundColumns
        if len(foundColumns) >= 1:
            if 'npi' in foundColumns and 'me' not in foundColumns and ('fname' not in foundColumns or 'lname' not in foundColumns or 'zip' not in foundColumns):
                self.matchTypeSelection.setCurrentIndex(3)
            elif 'me' in foundColumns and 'npi' not in foundColumns and ('fname' not in foundColumns or 'lname' not in foundColumns or 'zip' not in foundColumns):
                self.matchTypeSelection.setCurrentIndex(2)
            elif 'me' in foundColumns and 'npi' in foundColumns and ('fname' not in foundColumns or 'lname' not in foundColumns or 'zip' not in foundColumns):
                self.matchTypeSelection.setCurrentIndex(6)
            elif len(foundColumns) == 5:
                self.matchTypeSelection.setCurrentIndex(5)
            elif len(foundColumns) == 3:
                if 'fname' in foundColumns and 'lname' in foundColumns and 'zip' in foundColumns:
                    self.matchTypeSelection.setCurrentIndex(7)
            elif len(foundColumns) == 4:
                if 'npi' in foundColumns and 'fname' in foundColumns and 'lname' in foundColumns and 'zip' in foundColumns:
                    self.matchTypeSelection.setCurrentIndex(4)
                elif 'me' in foundColumns and 'fname' in foundColumns and 'lname' in foundColumns and 'zip' in foundColumns:
                    self.matchTypeSelection.setCurrentIndex(1)
        else:
            self.matchTypeSelection.setCurrentIndex(0)



    def addColumn(self):
        with open(downloads + 'csvFile.csv', 'rb') as inputFile, open(downloads + 'csvFileTemp.csv', 'wb') as outputFile:
            reader = csv.reader(inputFile)
            csvwriter = csv.writer(outputFile, lineterminator='\n')
            header = reader.next()
            header.append(str(self.addColumnText.text()))
            csvwriter.writerow(header)
            for row in reader:
                row.append("")
                csvwriter.writerow(row)

        os.chdir(downloads)
        os.remove(os.path.join(downloads, 'csvFile.csv'))
        os.rename(os.path.join(downloads, 'csvFileTemp.csv'), os.path.join(downloads, 'csvFile.csv'))
        if self.sourceSegs:
            self.sourceSegs.clear()
        with open(downloads + 'csvFile.csv', 'rb') as f:
            segReader = csv.reader(f)
            i = segReader.next()
            columns = [row for row in i]
            for item in columns:
                self.sourceSegs.addItem(item)
        self.highlightSourceSegs()
        self.highlightCMI()
        self.detectMatchType()

    def deleteColumn(self):
        selectedSegment = sorted([index.row() for index in self.sourceSegs.selectedIndexes()], reverse=True)
        # for item in selectedSegment:
            # itemIndex = item
        with open(downloads + 'csvFile.csv', 'rb') as inputFile, open(downloads + 'csvFileTemp.csv', 'wb') as outputFile:
            reader = csv.reader(inputFile)
            csvwriter = csv.writer(outputFile, lineterminator='\n')
            for row in reader:
                for item in selectedSegment:
                    del row[item]
                csvwriter.writerow(row)

        os.chdir(downloads)
        os.remove(os.path.join(downloads, 'csvFile.csv'))
        os.rename(os.path.join(downloads, 'csvFileTemp.csv'), os.path.join(downloads, 'csvFile.csv'))
        if self.sourceSegs:
            self.sourceSegs.clear()
        with open(downloads + 'csvFile.csv', 'rb') as f:
            segReader = csv.reader(f)
            i = segReader.next()
            columns = [row for row in i]
            for item in columns:
                self.sourceSegs.addItem(item)
        self.highlightSourceSegs()
        self.highlightCMI()
        self.detectMatchType()


    def getUniqueSegmentValues(self):
        self.uniqueValuesList.clear()
        uniqueList2 = []
        selectedSegment = sorted([index.row() for index in self.sourceSegs.selectedIndexes()], reverse=True)
        for self.segItem in selectedSegment:
            self.segIndex = self.segItem
            self.segItem = str(self.sourceSegs.currentItem().text())
            with open(downloads + 'csvFile.csv', 'rb') as f:
                reader = csv.DictReader(f)
                for value in reader:
                    segmentValue = str(value[self.segItem])
                    if segmentValue not in uniqueList2:
                        uniqueList2.append(segmentValue)
        uniqueList = sorted(uniqueList2)
        for i in uniqueList:
            self.uniqueValuesList.addItem(i)

    def renameValue(self):
        with open(downloads + 'csvFileTemp.csv', 'wb') as outFile:
            row_writer = csv.writer(outFile, lineterminator='\n')
            # selectedValue = sorted([index.row() for index in self.uniqueValuesList.selectedIndexes()], reverse=True)
            # for selectedItem in selectedValue:
            #     selectedItem = str(self.uniqueValuesList.currentItem().text())
            selectedItem = sorted([index.row() for index in self.uniqueValuesList.selectedIndexes()], reverse=True)
            selectedItems = []
            for i in selectedItem:
                selectedItems.append(self.uniqueValuesList.item(i).text())

            with open(downloads + 'csvFile.csv', 'rb') as f:
                reader = csv.reader(f)
                first_row = reader.next()
                row_writer.writerow(first_row)
                for row in reader:
                    for n, segVal in enumerate(row):
                        if segVal in selectedItems:
                            row[n] = str(self.renameValueEdit.text())
                    row_writer.writerow(row)
        os.chdir(downloads)
        os.remove(os.path.join(downloads, 'csvFile.csv'))
        os.rename(os.path.join(downloads, 'csvFileTemp.csv'), os.path.join(downloads, 'csvFile.csv'))
        self.getUniqueSegmentValues()

    def returnSelectedValues(self):
        selectedList = []
        unselectedList = []
        selectedSegments = sorted([index.row() for index in self.uniqueValuesList.selectedIndexes()], reverse=True)
        for i in selectedSegments:
            selectedList.append(self.uniqueValuesList.item(i).text())
        return selectedList

    def filterList(self):
        with open(downloads + 'csvFile.csv', 'rb') as inputFile, open(downloads + 'csvFileTemp.csv', 'wb') as outputFile:
            reader = csv.reader(inputFile)
            csvwriter = csv.writer(outputFile, lineterminator='\n')
            headers = reader.next()
            csvwriter.writerow(headers)
            selVals = self.returnSelectedValues()
            filtered = itertools.ifilter(lambda p: p[self.segIndex] in selVals, reader)
            csvwriter.writerows(filtered)

        os.chdir(downloads)
        os.remove(os.path.join(downloads, 'csvFile.csv'))
        os.rename(os.path.join(downloads, 'csvFileTemp.csv'), os.path.join(downloads, 'csvFile.csv'))
        self.getUniqueSegmentValues()
        self.countFullFile()

    def returnSourceIndexes(self):
        selectedList = []
        for item in selectedList:
            selectedList.remove(item)
        selectedSegments = [index.row() for index in self.sourceSegs.selectedIndexes()]
        for i in selectedSegments:
            selectedList.append(i)
        return selectedList

    def concatColumns(self):
        self.concatenateList.clear()
        concatList = self.returnSourceIndexes()
        col1 = concatList[0]
        col2 = concatList[1]
        with open(downloads + 'csvFile.csv', 'rb') as inputFile, open(downloads + 'csvFileTemp.csv', 'wb') as outputFile:
            reader = csv.reader(inputFile)
            csvwriter = csv.writer(outputFile, lineterminator='\n')
            for r in reader:
                csvwriter.writerow(r+[r[col1] + '_' + r[col2]])
        os.chdir(downloads)
        os.remove(os.path.join(downloads, 'csvFile.csv'))
        os.rename(os.path.join(downloads, 'csvFileTemp.csv'), os.path.join(downloads, 'csvFile.csv'))
        if self.sourceSegs:
            self.sourceSegs.clear()
        with open(downloads + 'csvFile.csv', 'rb') as f:
            segReader = csv.reader(f)
            i = segReader.next()
            columns = [row for row in i]
            for item in columns:
                self.sourceSegs.addItem(item)
        self.concatenateList.addItem("Concatenation Complete")
        self.concatenateList.setStyleSheet('background-color: rgb(213, 233, 221,255);')

    def getFieldnames(self):
        fieldNames = []
        for i in xrange(self.sourceSegs.count()):
            if str(self.sourceSegs.item(i).text()) not in self.notFound:
                fieldNames.append(str(self.sourceSegs.item(i).text()))
        for i in xrange(self.finalSegs.count()):
            fieldNames.append(str(self.finalSegs.item(i).text()))

        return fieldNames

    def previewConcat(self):
        # fieldNames = []
        with open(downloads + 'csvFile.csv', 'rb') as f, open(downloads + 'csvFileTemp.csv', 'w') as out:
            csvDict = csv.DictReader(f)
            # for i in xrange(self.sourceSegs.count()):
            #     fieldNames.append(str(self.sourceSegs.item(i).text()))
            writer = csv.DictWriter(out, fieldnames=self.getFieldnames())
            writer.writeheader()
            for row in csvDict:
                writer.writerow(row)
        os.remove(os.path.join(downloads, 'csvFile.csv'))
        os.rename(os.path.join(downloads, 'csvFileTemp.csv'), os.path.join(downloads, 'csvFile.csv'))
        concatList = self.returnSourceIndexes()
        if len(concatList) > 1:
            col1 = concatList[0]
            col2 = concatList[1]
            concatList2 = []
            selectedSegments = [index.row() for index in self.sourceSegs.selectedIndexes()]
            if len(selectedSegments) == 2:
                with open(downloads + 'csvFile.csv', 'rb') as f:
                    reader = csv.reader(f)
                    for value in reader:
                        segmentValue = str(value[col1] + "_" + value[col2])
                        concatList2.append(segmentValue)
                uniqueList = concatList2
                for i in uniqueList:
                    self.concatenateList.addItem(i)
            else:
                self.concatenateList.clear()
                self.concatenateList.setStyleSheet('')
        else:
            self.concatenateList.clear()
            self.concatenateList.setStyleSheet('')

    def renameColumn(self):
        selectedValue = sorted([index.row() for index in self.sourceSegs.selectedIndexes()], reverse=True)
        for selectedItem in selectedValue:
            selectedItem = str(self.sourceSegs.currentItem().text())
            with open(downloads + 'csvFile.csv', 'rb') as inputFile, open(downloads + 'csvFileTemp.csv', 'wb') as outputFile:
                newtext = str(self.newColumnName.text())
                r = csv.reader(inputFile)
                headers = r.next()
                for index, col in enumerate(headers):
                    headerVal = str(col)
                    if headerVal == selectedItem:
                        headers[index] = newtext


                w = csv.writer(outputFile, lineterminator='\n')
                w.writerow(headers)
                for row in r:
                    w.writerow(row)
            os.chdir(downloads)
            os.remove(os.path.join(downloads, 'csvFile.csv'))
            os.rename(os.path.join(downloads, 'csvFileTemp.csv'), os.path.join(downloads, 'csvFile.csv'))
            if self.sourceSegs:
                self.sourceSegs.clear()
            with open(downloads + 'csvFile.csv', 'rb') as f:
                segReader = csv.reader(f)
                i = segReader.next()
                columns = [row for row in i]
                for item in columns:
                    self.sourceSegs.addItem(item)

        self.highlightCMI()
        self.detectMatchType()

    def countFullFile(self):
        with open(downloads + 'csvFile.csv', 'rb') as f:
            reader = csv.reader(f)
            data = list(reader)
            fileCount = len(data)
            self.fullListCount.setText("Count: " + str(fileCount - 1))

    def countSelectedValues(self):
        selectedItems = self.returnSelectedValues()
        selectedStrings = []
        for i in selectedItems:
            selectedStrings.append(str(i))
        with open(downloads + 'csvFile.csv', 'rb') as f:
            f = csv.reader(f)
            total = 0
            for lines in f:
                for item in lines:
                    if item in selectedStrings:
                        total += 1
            selectedTotal = total
            self.fullListCount.setText("Count: " + str(total))


    def highlightCMI(self):
        isCmiCompassChecked = self.cmiCompass.isChecked()
        cmiList = ['npi', 'address1', 'campaign_type', 'city', 'fname', 'lname', 'me', 'zip', 'clientid', 'compasid', 'middle_name', 'segment1', 'specialty', 'state_code', 'tier', 'segment2', 'segment3']
        foundCMIColumns = []
        self.notFound = []
        possilbeCMIColumns = []
        possilbeCMIColumnsFinal = []
        self.columnIndex = []
        self.finalSegs.setEnabled(False)
        if isCmiCompassChecked:
            for i in xrange(self.sourceSegs.count()):
                checkCol =  str(self.sourceSegs.item(i).text()).lower()
                col = self.sourceSegs.item(i)
                if checkCol == 'state' or checkCol == 'state_code':
                    foundCMIColumns.append('state_code')
                    self.columnIndex.append(i)
                    col.setBackground(QColor(213, 233, 221,255))
                elif checkCol == 'address_1' or checkCol == 'address 1' or checkCol == 'address1' or checkCol == 'addr1' or checkCol == 'addressline1':
                    foundCMIColumns.append('address1')
                    self.columnIndex.append(i)
                    col.setBackground(QColor(213, 233, 221,255))
                elif checkCol == 'compasid' or checkCol == 'compas_id' or checkCol == 'compas id' or re.search('^compasid+.', checkCol):
                    foundCMIColumns.append('compasid')
                    self.columnIndex.append(i)
                    col.setBackground(QColor(213, 233, 221,255))
                elif checkCol == 'middle_name' or re.search('^middle.+name', checkCol) or checkCol == 'middlename':
                    foundCMIColumns.append('middle_name')
                    self.columnIndex.append(i)
                    col.setBackground(QColor(213, 233, 221,255))
                elif checkCol == 'city':
                    foundCMIColumns.append('city')
                    self.columnIndex.append(i)
                    col.setBackground(QColor(213, 233, 221,255))
                elif checkCol == 'tier':
                    foundCMIColumns.append('tier')
                    self.columnIndex.append(i)
                    col.setBackground(QColor(213, 233, 221,255))
                elif checkCol == 'client_id' or checkCol == 'client_id_1' or checkCol == 'clientid':
                    foundCMIColumns.append('clientid')
                    self.columnIndex.append(i)
                    col.setBackground(QColor(213, 233, 221,255))
                elif checkCol == 'campaign_type' or checkCol == 'campaign type':
                    foundCMIColumns.append('campaign_type')
                    self.columnIndex.append(i)
                    col.setBackground(QColor(213, 233, 221,255))
                elif checkCol == 'segment1':
                    foundCMIColumns.append('segment1')
                    self.columnIndex.append(i)
                    col.setBackground(QColor(213, 233, 221,255))
                elif checkCol == 'segment2':
                    foundCMIColumns.append(checkCol)
                    self.columnIndex.append(i)
                    col.setBackground(QColor(213, 233, 221,255))
                elif checkCol == 'segment3':
                    foundCMIColumns.append('segment3')
                    self.columnIndex.append(i)
                    col.setBackground(QColor(213, 233, 221,255))
# (checkCol != 'segment1' or checkCol != 'segment2' or checkCol != 'segment3') and (re.search('^segment.+', checkCol) or re.search('.+segment.+', checkCol) or re.search('.+segment', checkCol))
                elif checkCol == 'segment' or re.search('^segment.+', checkCol) or re.search('.+segment.+', checkCol) or re.search('.+segment', checkCol):
                    possilbeCMIColumns.append(checkCol)
                    col.setBackground(QColor(247, 243, 117,255))
                    if 'segment1' in foundCMIColumns and ('segment2' not in foundCMIColumns and 'segment3' not in foundCMIColumns):
                        if len(possilbeCMIColumns) > 1:
                            possilbeCMIColumnsFinal.append('segment2')
                            possilbeCMIColumnsFinal.append('segment3')
                            self.columnIndex.append(i)
                            col.setBackground(QColor(247, 243, 117,255))
                    if 'segment2' in foundCMIColumns and ('segment1' not in foundCMIColumns and 'segment3' not in foundCMIColumns):
                        if len(possilbeCMIColumns) > 1:
                            possilbeCMIColumnsFinal.append('segment1')
                            possilbeCMIColumnsFinal.append('segment3')
                            self.columnIndex.append(i)
                            col.setBackground(QColor(247, 243, 117,255))
                    if 'segment3' in foundCMIColumns and ('segment1' not in foundCMIColumns and 'segment2' not in foundCMIColumns):
                        if len(possilbeCMIColumns) > 1:
                            possilbeCMIColumnsFinal.append('segment1')
                            possilbeCMIColumnsFinal.append('segment2')
                            self.columnIndex.append(i)
                            col.setBackground(QColor(247, 243, 117,255))
                    if ('segment1' in foundCMIColumns and 'segment2' in foundCMIColumns) and 'segment3' not in foundCMIColumns:
                        if len(possilbeCMIColumns) > 0:
                            possilbeCMIColumnsFinal.append('segment3')
                            self.columnIndex.append(i)
                            col.setBackground(QColor(247, 243, 117,255))
                    if ('segment2' in foundCMIColumns and 'segment3' in foundCMIColumns) and 'segment1' not in foundCMIColumns:
                        if len(possilbeCMIColumns) > 0:
                            possilbeCMIColumnsFinal.append('segment1')
                            self.columnIndex.append(i)
                            col.setBackground(QColor(247, 243, 117,255))
                    if ('segment1' in foundCMIColumns and 'segment3' in foundCMIColumns) and 'segment2' not in foundCMIColumns:
                        if len(possilbeCMIColumns) > 0:
                            possilbeCMIColumnsFinal.append('segment2')
                            self.columnIndex.append(i)
                            col.setBackground(QColor(247, 243, 117,255))
                elif checkCol == 'specialty' or checkCol == 'speciality' or re.search('^specialty.+', checkCol) or re.search('.+specialty.+', checkCol):
                    foundCMIColumns.append('specialty')
                    self.columnIndex.append(i)
                    col.setBackground(QColor(213, 233, 221,255))
                elif checkCol == 'npi' or checkCol == 'npi_id' or re.search('.+ npi .+', checkCol) or re.search('.+_npi_.+', checkCol) or re.search('^npi.+', checkCol) or re.search('.+ npi', checkCol) or re.search('.+ npi', checkCol):
                    foundCMIColumns.append('npi')
                    self.columnIndex.append(i)
                    col.setBackground(QColor(213, 233, 221,255))
                elif checkCol == 'me' or checkCol == 'me_id' or checkCol == 'meded' or checkCol == 'menum' or checkCol == 'menumber' or re.search('.+ me .+', checkCol) or  re.search('.+_me_.+', checkCol) or re.search('^me .+', checkCol) or re.search('.+ me', checkCol) or re.search('.+ me', checkCol) or re.search('^me_.+', checkCol):
                    foundCMIColumns.append('me')
                    self.columnIndex.append(i)
                    col.setBackground(QColor(213, 233, 221,255))
                elif checkCol == 'fname' or checkCol == 'firstname' or re.search('^first.+name', checkCol) or re.search('.+first.+name', checkCol) or re.search('.+fname', checkCol) or re.search('.+first', checkCol) or re.search('.+frst.+', checkCol):
                    foundCMIColumns.append('fname')
                    self.columnIndex.append(i)
                    col.setBackground(QColor(213, 233, 221,255))
                elif re.search(r'^lname|^last.+name|^lastname|.+last|.+last.+name', checkCol):
                    foundCMIColumns.append('lname')
                    self.columnIndex.append(i)
                    col.setBackground(QColor(213, 233, 221,255))
                elif checkCol == 'zip' or checkCol == 'Postal' or (re.search('^zip.+', checkCol) and (checkCol != 'zip_4' or checkCol != 'zip4')) or re.search('^postal.+', checkCol) or re.search('.+_zip', checkCol) or re.search('.+ zip', checkCol) or re.search('.+_postal', checkCol) or re.search('.+ zip', checkCol) or re.search('.+ postal', checkCol):
                    foundCMIColumns.append('zip')
                    self.columnIndex.append(i)
                    col.setBackground(QColor(213, 233, 221,255))
            for item in cmiList:
                if item not in possilbeCMIColumnsFinal and item not in foundCMIColumns:
                    self.notFound.append(item)
                    self.sourceSegs.addItem(item)
            for i in xrange(self.sourceSegs.count()):
                checkCol =  str(self.sourceSegs.item(i).text()).lower()
                col = self.sourceSegs.item(i)
                if checkCol in self.notFound:
                    col.setBackground(QColor(223, 138, 138, 255))
                    self.columnIndex.append(i)
            #datacheck to see if yellow segments exist. If so it disables the run button until addressed
            # for i in xrange(self.sourceSegs.count()):
            #     checkCol =  str(self.sourceSegs.item(i).text()).lower()
            #     col = self.sourceSegs.item(i)
            #     rgbValue = str(col.background().color().getRgb())
            #     if rgbValue == '(247, 243, 117, 255)':
            #         self.runProgramButton.setEnabled(False)
            #         # print 'i found a possible segment column'
            #         break
            #     else:
            #         self.runProgramButton.setEnabled(True)
                    # print 'Were good to run'
        if not isCmiCompassChecked:
            for i in xrange(self.sourceSegs.count()):
                col = self.sourceSegs.item(i)
                col.setBackground(QColor('white'))
            deleteIndex = []
            for i in xrange(self.sourceSegs.count()):
                colText = str(self.sourceSegs.item(i).text()).lower()
                if colText in self.notFound:
                    deleteIndex.append(i)
            for index in deleteIndex:
                self.sourceSegs.takeItem(index)
                for i in xrange(self.sourceSegs.count()):
                    colText = str(self.sourceSegs.item(i).text()).lower()
                    if colText in self.notFound:
                        deleteIndex.append(i)
 
    def dataSharingClients(self):
        isCMIChecked = self.cmiCompass.isChecked()
        self.cmiCompass.setEnabled(True)
        if str(self.targetManuName.currentText()) == 'Merck':
            # self.segmentList.setEnabled(False)
            # self.keepSeg.setEnabled(False)
            # self.segValues.setEnabled(False)
            # self.segVariable.setEnabled(False)
            self.segBox.setChecked(True)
            self.segBox.setEnabled(False)
            self.addSegButton.setEnabled(False)
            self.removeSegButton.setEnabled(False)
            self.sourceSegs.setEnabled(True)
            self.finalSegs.setEnabled(False)
            if self.sourceSegs:
                self.sourceSegs.clear()
            if self.finalSegs:
                self.finalSegs.clear()
            merckList = utils.fetchColumns()
            for i in merckList:
                if i not in ['Customer_ID', 'Campaign_ID', 'Tactic_ID', 'Wave_ID', 'Product_ID']:
                    self.sourceSegs.addItem(i)
                else:
                    self.finalSegs.addItem(i)
        if str(self.targetManuName.currentText()) == 'Boehringer':
            # self.segmentList.setEnabled(False)
            # self.keepSeg.setEnabled(False)
            # self.segValues.setEnabled(False)
            # self.segVariable.setEnabled(False)
            # self.cmiCompass.setChecked(False)
            # self.cmiCompass.setEnabled(False)
            self.segBox.setChecked(True)
            self.segBox.setEnabled(False)
            self.addSegButton.setEnabled(False)
            self.removeSegButton.setEnabled(False)
            self.sourceSegs.setEnabled(True)
            self.finalSegs.setEnabled(False)
            # self.cmiCompass.setChecked(False)
            # self.cmiCompass.setEnabled(False)
            if self.sourceSegs:
                self.sourceSegs.clear()
            if self.finalSegs:
                self.finalSegs.clear()
            merckList = utils.fetchColumns()
            for i in merckList:
                self.sourceSegs.addItem(i)
        if str(self.targetManuName.currentText()) == 'Novartis':
            # self.segmentList.setEnabled(False)
            # self.keepSeg.setEnabled(False)
            # self.segValues.setEnabled(False)
            # self.segVariable.setEnabled(False)
            self.segBox.setChecked(True)
            self.segBox.setEnabled(False)
            self.addSegButton.setEnabled(False)
            self.removeSegButton.setEnabled(False)
            self.sourceSegs.setEnabled(True)
            self.finalSegs.setEnabled(False)
            if self.finalSegs:
                self.finalSegs.clear()
            self.finalSegs.addItem('CL_ME')
            self.finalSegs.addItem('CL_NPI')
            self.finalSegs.addItem('NOV_ID')
            self.finalSegs.addItem('MDM_ID')
            self.finalSegs.addItem('vendor_contact_event_id')
            self.finalSegs.addItem('fulfillment_kit_code')

        if str(self.targetManuName.currentText()) == 'AstraZeneca':
            # self.segmentList.setEnabled(False)
            # self.keepSeg.setEnabled(False)
            # self.segValues.setEnabled(False)
            # self.segVariable.setEnabled(False)
            self.segBox.setChecked(True)
            self.segBox.setEnabled(False)
            self.addSegButton.setEnabled(False)
            self.removeSegButton.setEnabled(False)
            self.sourceSegs.setEnabled(True)
            self.finalSegs.setEnabled(False)
            if self.finalSegs:
                self.finalSegs.clear()
            self.finalSegs.addItem('IMSDR')
            self.finalSegs.addItem('HCP_AZ_CUSTOM_ID')

        if str(self.targetManuName.currentText()) == 'Biogen':
            # self.segmentList.setEnabled(False)
            # self.keepSeg.setEnabled(False)
            # self.segValues.setEnabled(False)
            # self.segVariable.setEnabled(False)
            self.segBox.setChecked(True)
            self.segBox.setEnabled(False)
            self.addSegButton.setEnabled(False)
            self.removeSegButton.setEnabled(False)
            self.sourceSegs.setEnabled(True)
            self.finalSegs.setEnabled(False)
            if self.finalSegs:
                self.finalSegs.clear()

            for i in xrange(self.sourceSegs.count()):
                if str(self.sourceSegs.item(i).text()).lower() == 'vnid':
                    self.finalSegs.addItem('veeva_network_id')
                if str(self.sourceSegs.item(i).text()).lower() == 'veeva_network_id':
                    self.finalSegs.addItem('veeva_network_id')


        if str(self.targetManuName.currentText()) == 'GSK':
            # self.segmentList.setEnabled(False)
            # self.keepSeg.setEnabled(False)
            # self.segValues.setEnabled(False)
            # self.segVariable.setEnabled(False)
            self.segBox.setChecked(True)
            self.segBox.setEnabled(False)
            self.addSegButton.setEnabled(False)
            self.removeSegButton.setEnabled(False)
            self.sourceSegs.setEnabled(True)
            self.finalSegs.setEnabled(False)
            if self.finalSegs:
                self.finalSegs.clear()
            self.finalSegs.addItem('GSK_Metadata_Tag')
        if str(self.targetManuName.currentText()) == 'Amgen':
            # self.segmentList.setEnabled(False)
            # self.keepSeg.setEnabled(False)
            # self.segValues.setEnabled(False)
            # self.segVariable.setEnabled(False)
            self.segBox.setChecked(True)
            self.segBox.setEnabled(False)
            self.addSegButton.setEnabled(False)
            self.removeSegButton.setEnabled(False)
            # self.sourceSegs.setEnabled(True)
            # self.finalSegs.setEnabled(False)
            if self.sourceSegs:
                self.sourceSegs.clear()
            if self.finalSegs:
                self.finalSegs.clear()
            merckList = utils.fetchColumns()
            for i in merckList:
                if i not in ['CMA_ID', 'PRIMARY_KEYCODE', 'ME_NUMBER', 'NPI_NUMBER', 'FIRST_NAME', 'LAST_NAME', 'ADDRESS1', 'CITY', 'ZIP', 'STATE_CD', 'EMAIL_ADDRESS']:
                    self.sourceSegs.addItem(i)
                else:
                    self.finalSegs.addItem(i)
        if str(self.targetManuName.currentText()) == 'Sanofi-Aventis':
            # self.segmentList.setEnabled(False)
            # self.keepSeg.setEnabled(False)
            # self.segValues.setEnabled(False)
            # self.segVariable.setEnabled(False)
            self.segBox.setChecked(True)
            self.segBox.setEnabled(False)
            self.sourceSegs.setEnabled(False)
            self.finalSegs.setEnabled(False)
            if self.sourceSegs:
                self.sourceSegs.clear()
            if self.finalSegs:
                self.finalSegs.clear()
            merckList = utils.fetchColumns()
            for i in merckList:
                if i not in ['cct_id', 'npi', 'me', 'fname', 'middle_name', 'lname', 'address1', 'city', 'zip', 'state', 'email_address', 'country', 'brand', 'tactic', 'vendor']:
                    self.sourceSegs.addItem(i)
                else:
                    i = i.replace('me', 'me_number')
                    i = i.replace('npi', 'npi_number')
                    i = i.replace('fname', 'first_name')
                    i = i.replace('lname', 'last_name')
                    i = i.replace('zip', 'zip2')
                    self.finalSegs.addItem(i)
        if str(self.targetManuName.currentText()) == 'Lilly':
            # self.segmentList.setEnabled(False)
            # self.keepSeg.setEnabled(False)
            # self.segValues.setEnabled(False)
            # self.segVariable.setEnabled(False)
            self.segBox.setChecked(True)
            self.segBox.setEnabled(False)
            self.addSegButton.setEnabled(False)
            self.removeSegButton.setEnabled(False)
            self.sourceSegs.setEnabled(True)
            self.finalSegs.setEnabled(False)
            if self.sourceSegs:
                self.sourceSegs.clear()
            if self.finalSegs:
                self.finalSegs.clear()
            merckList = utils.fetchColumns()
            for i in merckList:
                self.sourceSegs.addItem(i)

        if str(self.targetManuName.currentText()) not in ['Merck', 'Amgen', 'Boehringer', 'AstraZeneca', 'GSK', 'Sanofi-Aventis', 'Lilly', 'Novartis', 'Biogen']:
            isSegmentChecked = self.segBox.isChecked()
            # self.segmentList.setEnabled(True)
            # self.keepSeg.setEnabled(True)
            # self.segValues.setEnabled(True)
            # self.segVariable.setEnabled(True)
            self.segBox.setEnabled(True)
            self.segBox.setChecked(False)
            self.standardMatch.setEnabled(True)
            self.exactMatch.setChecked(False)
            self.sourceSegs.setEnabled(isSegmentChecked)
            self.finalSegs.setEnabled(isSegmentChecked)
            if self.finalSegs:
                self.finalSegs.clear()
            if self.sourceSegs:
                self.sourceSegs.clear()
            colList = utils.fetchColumns()
            for i in colList:
                self.sourceSegs.addItem(i)
            self.segmentCallBack()
            self.sdaOnlyCallback()
            # self.refreshFile()

    def showCalWid(self):
        self.calendar = QtGui.QCalendarWidget()
        self.calendar.setMinimumDate(QtCore.QDate(1900, 1, 1))
        self.calendar.setMaximumDate(QtCore.QDate(3000, 1, 1))
        self.calendar.setGridVisible(True)
        self.calendar.clicked.connect(self.updateDate)
        self.calendar.setWindowFlags(QtCore.Qt.FramelessWindowHint)
        self.calendar.setStyleSheet('background: white; color: black')
        self.calendar.setGridVisible(True)
        pos = QtGui.QCursor.pos()
        self.calendar.setGeometry(pos.x(), pos.y(),300, 200)
        self.calendar.show()

    # def windowChange(self):
    #     self.fileWindow = FilePage()
    #     self.fileWindow

    def showFileSelection(self):
        isSuppChecked = self.suppCheck.isChecked()
        if isSuppChecked:
            self.suppCheck.setStyleSheet("color: green; font-weight: bold")
            self.fileWindow = FilePage()
            self.fileWindow.show()
        else:
            self.suppCheck.setStyleSheet("color: black")


    def updateDate(self,*args):
        getDate = self.calendar.selectedDate().toString('MM-dd-yy')
        self.calendarLine.setText(getDate)
        self.calendar.deleteLater()

    def formatTableName(self):
        s = ""
        formatedTableName = str(self.postgresTable.text()).replace(' ', '_').lower()
        s += formatedTableName
        return formatedTableName

    def formatTableNameTarget(self):
        s = ""
        formatedTableName = str(self.postgresTargetTable.text()).replace(' ', '_').lower()
        s += formatedTableName
        return s

    def returnUserTable(self):
        # currentTableName = self.postgresTable.setText(self.formatTableName())
        targetingTableName = str(self.postgresTargetTable.setText(self.formatTableNameTarget()))
        if self.tabWidget.currentIndex() == 0:
            if self.formatTableName():
                if self.formatTableName() in self.getPostgresTables():
                    self.postgresTable.setText(self.formatTableName())
                    self.tableNameWarning.setText('TABLE NAME EXISTS')
                    self.tableNameWarning.setStyleSheet('color: red')
                    self.runProgramButton.setEnabled(False)
                elif len(self.formatTableName()) == 0:
                    self.tableNameWarning.setText('')
                    self.runProgramButton.setEnabled(True)
                else:
                    self.postgresTable.setText(self.formatTableName())
                    self.tableNameWarning.setText('Table Name Acceptable')
                    self.tableNameWarning.setStyleSheet('color: green')
                    self.runProgramButton.setEnabled(True)
        else:
            if self.formatTableNameTarget():
                if self.formatTableNameTarget() in self.getPostgresTables():
                    self.postgresTargetTable.setText(self.formatTableNameTarget())
                    self.tableNameWarningTargeting.setText('TABLE NAME EXISTS')
                    self.tableNameWarningTargeting.setStyleSheet('color: red')
                    self.runProgramButton.setEnabled(False)
                elif len(self.formatTableNameTarget()) == 0:
                    self.tableNameWarning.setText('')
                    self.runProgramButton.setEnabled(True)
                else:
                    self.postgresTargetTable.setText(self.formatTableNameTarget())
                    self.tableNameWarningTargeting.setText('Table Name Acceptable')
                    self.tableNameWarningTargeting.setStyleSheet('color: green')
                    self.runProgramButton.setEnabled(True)


    def getPostgresTables2(self):
        postgresTables = []
        with open(os.path.join(desktop, 'TheEagleHasLanded.csv'), 'rb') as passFile:
            reader = csv.DictReader(passFile)
            for item in reader:
                password = item['password']
                conn_string = "host='localhost' dbname='postgres' port='5432' user='postgres' password='{password}'".format(password=password)
                conn = psycopg2.connect(conn_string)
                cursor = conn.cursor()
                sql = """select table_name from information_schema.tables where table_schema = 'public'"""
                cursor.execute(sql)
                conn.commit()
                results = [r[0] for r in cursor.fetchall()]
                for table in results:
                    postgresTables.append(table)

                return postgresTables

    def getPostgresTables(self):
        sqlLiteTables = []
        conn = sqlite3.connect(os.path.join(desktop, 'Ewok', 'epocrates_tables.db'))
        # print 'connected to sqlite'
        cursor = conn.cursor()
        sql = """SELECT name as count FROM sqlite_master WHERE type='table';"""
        cursor.execute(sql)
        conn.commit()
        results = [r[0] for r in cursor.fetchall()]
        # print results
        for row in results:
            sqlLiteTables.append(row)
        return sqlLiteTables

    def usePreviousConfig(self):
        isPreviousChecked = self.previousSettings.isChecked()
        if isPreviousChecked:
            self.previousSettings.setStyleSheet("color: green; font-weight: bold")
            with open(os.path.join(desktop, 'config.json'), 'r') as infile:
                config = json.loads(infile.read(), encoding='utf8')
                if config['caseType'] == 'listMatch':
                    self.tabWidget.setCurrentIndex(0)
                    self.caseNumber.setText(config['caseno'])
                    self.brandName.setText(config['Brand'])
                    self.saleExecutive.setText(config['SE'])
                    self.presalesManuName.setText(config['Manu'])
                    self.postgresTable.setText(config['tableName'])
                    self.requesterInIt.setText(config['reIn'])
                    self.myInIt.setText(config['yourIn'])
                    self.matchTypeSelection.setCurrentIndex(config['mtypeIndex'])
                    self.emailRecip.setText(config['email'])
                if config['caseType'] == 'Targeting':
                    self.tabWidget.setCurrentIndex(1)
                    if config['segmentListChecked'] == 'y':
                        self.segmentList.setChecked(True)
                    if config['keep_seg'] == 'Yes':
                        self.keepSeg.setChecked(True)
                    if config['cmi_compass_client'] == 'Y':
                        self.cmiCompass.setChecked(True)
                    self.targetManuName.setCurrentIndex(config['targetManuIndex'])
                    self.brandTargetName.setText(config['Brand'])
                    self.targetNumber.setText(config['targetNum'])
                    self.postgresTargetTable.setText(config['tableName'])
                    self.segVariable.setText(config['segVariable'])
                    self.segValues.setText(config['varValues'])
                    self.calendarLine.setText(config['date'])
                    self.myTargetInIt.setText(config['yourIn'])
                if config['listProduct'] == 'DocAlert':
                    self.docAlert.setChecked(True)
                if config['listProduct'] == 'Epoc_Quiz':
                    self.epocQuiz.setChecked(True)
                if config['listMatchType'] == 'Standard':
                    self.standardMatch.setChecked(True)
                if config['listMatchType'] == 'Exact':
                    self.exactMatch.setChecked(True)
                if config['listMatchType'] == 'Standard_Seg':
                    self.standardMatch.setChecked(True)
                    self.segBox.setChecked(True)
                    self.sourceSegs.clear()
                    self.finalSegs.clear()
                    allSegs = utils.fetchColumns()
                    for seg in allSegs:
                        if seg not in config['finalSegs']:
                            self.sourceSegs.addItem(seg)
                        else:
                            self.finalSegs.addItem(seg)

                if config['listMatchType'] == 'Exact_Seg':
                    self.exactMatch.setChecked(True)
                    self.segBox.setChecked(True)
                    self.sourceSegs.clear()
                    self.finalSegs.clear()
                    allSegs = utils.fetchColumns()
                    for seg in allSegs:
                        if seg not in config['finalSegs']:
                            self.sourceSegs.addItem(seg)
                        else:
                            self.finalSegs.addItem(seg)
                if config['sDa'] == 'y':
                    self.sdaOcc.setEnabled(True)
                    self.sdaSpec.setEnabled(True)
                    self.sdaAddOn.setChecked(True)
                    self.sdaOcc.setText(str(config['SDA_Occ']).replace('"', ''))
                    self.sdaSpec.setText(str(config['SDA_Spec']).replace('"', ''))
                    self.sdaTargetNum.setText(config['SDA_Target'])
                if config['bDa'] == 'y':
                    self.bdaAddOn.setChecked(True)
                    self.bdaOcc.setText(str(config['occupation']).replace('"', ''))
                    self.bdaSpec.setText(str(config['specialty']).replace('"', ''))
                    self.bdaLookUps.setText(config['lookUpPeriod'])
                    self.bdaNumLookUps.setText(config['totalLookUps'])
                    self.bdaTargetNum.setText(config['BDA_Target'])
                    self.drugList.setPlainText(config['drugList'])


        else:
            self.setDefaults()
            self.previousSettings.setStyleSheet("color: black")
            #sdaresets
            self.sdaAddOn.setChecked(False)
            self.bdaAddOn.setChecked(False)

    def refreshFile(self):
        if os.path.exists(os.path.join(downloads, 'csvFile.csv')):
            os.remove(os.path.join(downloads, 'csvFile.csv'))
        if os.path.exists(os.path.join(downloads, 'target.csv')):
            os.remove(os.path.join(downloads, 'target.csv'))
        if os.path.exists(os.path.join(downloads, 'target_mod.csv')):
            os.remove(os.path.join(downloads, 'target_mod.csv'))
        if os.path.exists(os.path.join(downloads, 'csvFile1.csv')):
            os.remove(os.path.join(downloads, 'csvFile1.csv'))
        if os.path.exists(os.path.join(downloads, 'target_mod_temp2.csv')):
            os.remove(os.path.join(downloads, 'target_mod_temp2.csv'))
        self.loadedFile = newest
        self.setWindowTitle("File Loaded: "+self.loadedFile)
        self.countSheets()
        if self.sourceSegs:
            self.sourceSegs.clear()
            self.finalSegs.clear()
        utils.checkExtension()
        utils.removeChar()
        utils.incDupColumns()

        #populates the listWidget
        colList = utils.fetchColumns()
        for i in colList:
            self.sourceSegs.addItem(i)
        self.returnUserTable() 
        self.countFullFile()
        self.resetEditTab()
        self.highlightSourceSegs()
        self.detectMatchType()
        # self.playGoy()
        # self.refreshSuccessful()

    def refreshSuccessful(self):
        self.refreshLabel.setText("Refresh Successful")

    def clearLabel(self):
        self.refreshLabel.clear()

    def closeEvent(self, *args, **kwargs):
        super(QtGui.QMainWindow, self).closeEvent(*args, **kwargs)
        print "you just closed the pyqt window!!! you are awesome!!!"
        if os.path.exists(os.path.join(downloads, 'csvFile.csv')):
            os.remove(os.path.join(downloads, 'csvFile.csv'))
        if os.path.exists(os.path.join(downloads, 'target.csv')):
            os.remove(os.path.join(downloads, 'target.csv'))
        if os.path.exists(os.path.join(downloads, 'target_mod.csv')):
            os.remove(os.path.join(downloads, 'target_mod.csv'))
        if os.path.exists(os.path.join(downloads, 'csvFile1.csv')):
            os.remove(os.path.join(downloads, 'csvFile1.csv'))
        if os.path.exists(os.path.join(downloads, 'target_mod_temp2.csv')):
            os.remove(os.path.join(downloads, 'target_mod_temp2.csv'))

    def runProgram(self):
        isSdaChecked = self.sdaOnly.isChecked()
        isBdaChecked = self.bdaOnly.isChecked()
        self.writeConfigFile()
        showError = False

        for i in xrange(self.sourceSegs.count()):
            checkCol =  str(self.sourceSegs.item(i).text()).lower()
            col = self.sourceSegs.item(i)
            rgbValue = str(col.background().color().getRgb())
            if rgbValue == '(247, 243, 117, 255)':
                msg = QMessageBox()
                msg.setIcon(QMessageBox.Warning)
                msg.setText("There Are Yellow Possible CMI Segment Columns. Please Rename Column(s) to Run Code")
                msg.setStandardButtons(QMessageBox.Ok)
                msg.exec_()
                showError = True
                break
        if showError == False:   
            if self.config['suppressionApplied'] == 'N':
                if str(self.targetManuName.currentText()) == 'Amgen':
                    subprocess.call(['python.exe', 'G:\\Communicator Ops\\Epocrates\\Python Files\\DataSharing\\Amgen.py', 'config.json'])
                if str(self.targetManuName.currentText()) == 'AstraZeneca':
                    subprocess.call(['python.exe', 'G:\\Communicator Ops\\Epocrates\\Python Files\\DataSharing\\AZ.py', 'config.json'])
                if str(self.targetManuName.currentText()) == 'Boehringer':
                    subprocess.call(['python.exe', 'G:\\Communicator Ops\\Epocrates\\Python Files\\DataSharing\\Boehringer.py', 'config.json'])
                if str(self.targetManuName.currentText()) == 'GSK':
                    subprocess.call(['python.exe', 'G:\\Communicator Ops\\Epocrates\\Python Files\\DataSharing\\GSK.py', 'config.json'])
                if str(self.targetManuName.currentText()) == 'Novartis':
                    subprocess.call(['python.exe', 'G:\\Communicator Ops\\Epocrates\\Python Files\\DataSharing\\Novartis.py', 'config.json'])
                if str(self.targetManuName.currentText()) == 'Sanofi-Aventis':
                    subprocess.call(['python.exe', 'G:\\Communicator Ops\\Epocrates\\Python Files\\DataSharing\\Sanofi.py', 'config.json'])
                if str(self.targetManuName.currentText()) == 'Biogen':
                    subprocess.call(['python.exe', 'G:\\Communicator Ops\\Epocrates\\Python Files\\DataSharing\\Biogen.py', 'config.json'])
                if str(self.targetManuName.currentText()) == 'Merck':
                    subprocess.call(['python.exe', 'G:\\Communicator Ops\\Epocrates\\Python Files\\DataSharing\\Merck.py', 'config.json'])
                if self.config['caseType'] == 'Targeting' and str(self.targetManuName.currentText()) not in ['Merck', 'Amgen', 'Boehringer', 'AstraZeneca', 'GSK', 'Sanofi-Aventis', 'Novartis', 'Biogen']:
                    subprocess.call(['python.exe', os.path.join(desktop, 'Ewok','theOne_GUI.py'), 'config.json'])
                if self.config['caseType'] == 'listMatch':
                    if not isSdaChecked and not isBdaChecked:
                        subprocess.call(['python.exe', os.path.join(desktop, 'Ewok','theOne_GUI.py'), 'config.json'])
                    elif isSdaChecked and not isBdaChecked:
                        if self.config['matchedFile'] != "":
                            subprocess.call(['python.exe', 'G:\\Communicator Ops\\Epocrates\\Python Files\\Official Codes\\SDA.py', 'config.json'])
                        else:
                            subprocess.call(['python.exe', os.path.join(desktop, 'Ewok','theOne_GUI.py'), 'config.json'])
                    elif not isSdaChecked and isBdaChecked:
                        if self.config['matchedFile'] != "":
                            if int(self.config['totalLookUps']) == 1:
                                if str(self.config['specialty']) != '""':
                                    subprocess.call(['python.exe', 'G:\\Communicator Ops\\Epocrates\\Python Files\\Official Codes\\BDA_Multi.py', 'config.json'])
                                else:
                                    subprocess.call(['python.exe', 'G:\\Communicator Ops\\Epocrates\\Python Files\\Official Codes\\BDA_Multi_Occ_Only.py', 'config.json'])
                            else:
                                if str(self.config['specialty']) != '""':
                                    subprocess.call(['python.exe', 'G:\\Communicator Ops\\Epocrates\\Python Files\\Official Codes\\BDA_Multi.py', 'config.json'])
                                else:
                                    subprocess.call(['python.exe', 'G:\\Communicator Ops\\Epocrates\\Python Files\\Official Codes\\BDA_Multi_Occ_Only.py', 'config.json'])
                        else:
                            subprocess.call(['python.exe', os.path.join(desktop, 'Ewok','theOne_GUI.py'), 'config.json'])
                    elif isSdaChecked and isBdaChecked:
                        if self.config['matchedFile'] != "":
                            if int(self.config['totalLookUps']) == 1:
                                if str(self.config['specialty']) != '""':
                                    subprocess.call(['python.exe', 'G:\\Communicator Ops\\Epocrates\\Python Files\\Official Codes\\SDA.py', 'config.json'])
                                    subprocess.call(['python.exe', 'G:\\Communicator Ops\\Epocrates\\Python Files\\Official Codes\\BDA_Multi.py', 'config.json'])
                                else:
                                    subprocess.call(['python.exe', 'G:\\Communicator Ops\\Epocrates\\Python Files\\Official Codes\\SDA.py', 'config.json'])
                                    subprocess.call(['python.exe', 'G:\\Communicator Ops\\Epocrates\\Python Files\\Official Codes\\BDA_Multi_Occ_Only.py', 'config.json'])
                            else:
                                if str(self.config['specialty']) != '""':
                                    subprocess.call(['python.exe', 'G:\\Communicator Ops\\Epocrates\\Python Files\\Official Codes\\SDA.py', 'config.json'])
                                    subprocess.call(['python.exe', 'G:\\Communicator Ops\\Epocrates\\Python Files\\Official Codes\\BDA_Multi.py', 'config.json'])
                                else:
                                    subprocess.call(['python.exe', 'G:\\Communicator Ops\\Epocrates\\Python Files\\Official Codes\\SDA.py', 'config.json'])
                                    subprocess.call(['python.exe', 'G:\\Communicator Ops\\Epocrates\\Python Files\\Official Codes\\BDA_Multi_Occ_Only.py.py', 'config.json'])
                        else:
                            subprocess.call(['python.exe', os.path.join(desktop, 'Ewok','theOne_GUI.py'), 'config.json'])
            if self.config['suppressionApplied'] == 'Y':
                if str(self.targetManuName.currentText()) == 'Amgen':
                    subprocess.call(['python.exe', 'G:\\Communicator Ops\\Epocrates\\Python Files\\DataSharing\\Amgen.py', 'config.json'])
                if str(self.targetManuName.currentText()) == 'AstraZeneca':
                    subprocess.call(['python.exe', 'G:\\Communicator Ops\\Epocrates\\Python Files\\DataSharing\\AZ.py', 'config.json'])
                if str(self.targetManuName.currentText()) == 'Boehringer':
                    subprocess.call(['python.exe', 'G:\\Communicator Ops\\Epocrates\\Python Files\\DataSharing\\Boehringer.py', 'config.json'])
                if str(self.targetManuName.currentText()) == 'GSK':
                    subprocess.call(['python.exe', 'G:\\Communicator Ops\\Epocrates\\Python Files\\DataSharing\\GSK.py', 'config.json'])
                if str(self.targetManuName.currentText()) == 'Novartis':
                    subprocess.call(['python.exe', 'G:\\Communicator Ops\\Epocrates\\Python Files\\DataSharing\\Novartis.py', 'config.json'])
                if str(self.targetManuName.currentText()) == 'Sanofi-Aventis':
                    subprocess.call(['python.exe', 'G:\\Communicator Ops\\Epocrates\\Python Files\\DataSharing\\Sanofi.py', 'config.json'])
                if str(self.targetManuName.currentText()) == 'Biogen':
                    subprocess.call(['python.exe', 'G:\\Communicator Ops\\Epocrates\\Python Files\\DataSharing\\Biogen.py', 'config.json'])
                if str(self.targetManuName.currentText()) == 'Merck':
                    subprocess.call(['python.exe', 'G:\\Communicator Ops\\Epocrates\\Python Files\\DataSharing\\Merck.py', 'config.json'])
                if self.config['caseType'] == 'Targeting' and str(self.targetManuName.currentText()) not in ['Merck', 'Amgen', 'Boehringer', 'AstraZeneca', 'GSK', 'Sanofi-Aventis', 'Novartis','Biogen']:
                    subprocess.call(['python.exe', os.path.join(desktop, 'Ewok','theOne_GUI.py'), 'config.json'])
                if self.config['caseType'] == 'listMatch':
                    if not isSdaChecked and not isBdaChecked:
                        subprocess.call(['python.exe', os.path.join(desktop, 'Ewok','theOne_GUI.py'), 'config.json'])
                    elif isSdaChecked and not isBdaChecked:
                        if self.config['matchedFile'] != "":
                            subprocess.call(['python.exe', 'G:\\Communicator Ops\\Epocrates\\Python Files\\Official Codes\\SDA.py', 'config.json'])
                        else:
                            subprocess.call(['python.exe', os.path.join(desktop, 'Ewok','theOne_GUI.py'), 'config.json'])
                    elif not isSdaChecked and isBdaChecked:
                        if self.config['matchedFile'] != "":
                            if int(self.config['totalLookUps']) == 1:
                                if str(self.config['specialty']) != '""':
                                    subprocess.call(['python.exe', 'G:\\Communicator Ops\\Epocrates\\Python Files\\Official Codes\\BDA_Multi.py', 'config.json'])
                                else:
                                    subprocess.call(['python.exe', 'G:\\Communicator Ops\\Epocrates\\Python Files\\Official Codes\\BDA_Multi_Occ_Only.py', 'config.json'])
                            else:
                                if str(self.config['specialty']) != '""':
                                    subprocess.call(['python.exe', 'G:\\Communicator Ops\\Epocrates\\Python Files\\Official Codes\\BDA_Multi.py', 'config.json'])
                                else:
                                    subprocess.call(['python.exe', 'G:\\Communicator Ops\\Epocrates\\Python Files\\Official Codes\\BDA_Multi_Occ_Only.py', 'config.json'])
                        else:
                            subprocess.call(['python.exe', os.path.join(desktop, 'Ewok','theOne_GUI.py'), 'config.json'])
                    elif isSdaChecked and isBdaChecked:
                        if self.config['matchedFile'] != "":
                            if int(self.config['totalLookUps']) == 1:
                                if str(self.config['specialty']) != '""':
                                    subprocess.call(['python.exe', 'G:\\Communicator Ops\\Epocrates\\Python Files\\Official Codes\\SDA.py', 'config.json'])
                                    subprocess.call(['python.exe', 'G:\\Communicator Ops\\Epocrates\\Python Files\\Official Codes\\BDA_Multi.py', 'config.json'])
                                else:
                                    subprocess.call(['python.exe', 'G:\\Communicator Ops\\Epocrates\\Python Files\\Official Codes\\SDA.py', 'config.json'])
                                    subprocess.call(['python.exe', 'G:\\Communicator Ops\\Epocrates\\Python Files\\Official Codes\\BDA_Multi_Occ_Only.py', 'config.json'])
                            else:
                                if str(self.config['specialty']) != '""':
                                    subprocess.call(['python.exe', 'G:\\Communicator Ops\\Epocrates\\Python Files\\Official Codes\\SDA.py', 'config.json'])
                                    subprocess.call(['python.exe', 'G:\\Communicator Ops\\Epocrates\\Python Files\\Official Codes\\BDA_Multi.py', 'config.json'])
                                else:
                                    subprocess.call(['python.exe', 'G:\\Communicator Ops\\Epocrates\\Python Files\\Official Codes\\SDA.py', 'config.json'])
                                    subprocess.call(['python.exe', 'G:\\Communicator Ops\\Epocrates\\Python Files\\Official Codes\\BDA_Multi_Occ_Only.py.py', 'config.json'])
                        else:
                            subprocess.call(['python.exe', os.path.join(desktop, 'Ewok','theOne_GUI.py'), 'config.json'])
                if self.config['suppSASFile'] == '':
                    subprocess.call(['python.exe', os.path.join(desktop, 'Ewok','suppFile.py'), 'config.json'])
            self.returnUserTable()
            self.clearLabel()
            # self.playTrash()

    def writeConfigFile(self):
        self.config = dict()
        isSegmentChecked = self.segBox.isChecked()
        isSdaChecked = self.sdaAddOn.isChecked()
        isBdaChecked = self.bdaAddOn.isChecked()
        isDedupChecked = self.deDup.isChecked()
        isSegListChecked = self.segmentList.isChecked()
        isKeepSegChecked = self.keepSeg.isChecked()
        isCmiCompassChecked = self.cmiCompass.isChecked()
        isFuzzyBoxChecked = self.fuzzyBox.isChecked()
        isTherapyChecked = self.therapyClassBox.isChecked()
        isSuppSdaChecked = self.suppSDAOnly.isChecked()
        isSuppBdaChecked = self.suppBDAOnly.isChecked()

        # if self.listMatchCase.isChecked():
        if self.tabWidget.currentIndex() == 0:
            self.config['caseType'] = 'listMatch'
            self.config['Manu'] = str(self.presalesManuName.text())
            self.config['Brand'] = str(self.brandName.text())
            self.config['yourIn'] = str(self.myInIt.text())
            self.config['reIn'] = str(self.requesterInIt.text())
            self.config['caseno'] = str(self.caseNumber.text())
            self.config['SE'] = str(self.saleExecutive.text())
            self.config['email'] = str(self.emailRecip.text())
            self.config['mtype'] = str(self.matchTypeSelection.currentText())
            self.config['mtypeIndex'] = self.matchTypeSelection.currentIndex()
            self.config['tableName'] = str(self.postgresTable.text()).lower().replace(' ', '_')
            if self.standardMatch.isChecked() and isSegmentChecked:
                self.config['listMatchType'] = 'Standard_Seg'
                self.config['finalSegs'] = list()
                for i in range(self.finalSegs.count()):
                    self.config['finalSegs'].append(str(self.finalSegs.item(i).text()))
            if self.exactMatch.isChecked() and isSegmentChecked:
                self.config['listMatchType'] = 'Exact_Seg'
                self.config['finalSegs'] = list()
                for i in range(self.finalSegs.count()):
                    self.config['finalSegs'].append(str(self.finalSegs.item(i).text()))
            if not isSegmentChecked:
                self.config['finalSegs'] = []
            if self.deDup.isChecked():
                self.config['deDup'] = str('Yes')
            if not isDedupChecked:
                self.config['deDup'] = str('No')
        if self.tabWidget.currentIndex() == 1:
            if isSuppSdaChecked:
                self.config['suppSDAOnly'] = 'Y'
            if not isSuppSdaChecked:
                self.config['suppSDAOnly'] = 'N'
            if isSuppBdaChecked:
                self.config['suppBDAOnly'] = 'Y'
            if not isSuppBdaChecked:
                self.config['suppBDAOnly'] = 'N'
            if self.backFillBox.isChecked():
                self.config['backFill'] = 'Yes'
            if not self.backFillBox.isChecked():
                self.config['backFill'] = 'No'
            if self.deDup.isChecked():
                self.config['deDup'] = str('Y')
            if not isDedupChecked:
                self.config['deDup'] = str('N')
            self.config['caseType'] = str('Targeting')
            self.config['Manu'] = str(self.targetManuName.currentText())
            self.config['targetManuIndex'] = self.targetManuName.currentIndex()
            self.config['Brand'] = str(self.brandTargetName.text())
            self.config['targetNum'] = str(self.targetNumber.text())
            self.config['yourIn'] = str(self.myTargetInIt.text())
            self.config['tableName'] = str(self.postgresTargetTable.text())
            self.config['date'] = str(self.calendarLine.text())
            if isSegmentChecked:
                self.config['dSharing'] ='Y'
                self.config['finalSegs'] = list()
                if not self.cmiCompass.isChecked():
                    for i in range(self.finalSegs.count()):
                        self.config['finalSegs'].append(str(self.finalSegs.item(i).text()))
                else:
                    self.config['finalSegs'] = []
            if not isSegmentChecked:
                self.config['dSharing'] ='N'
                self.config['finalSegs'] = []
            if (self.standardMatch.isChecked() or self.exactMatch.isChecked()) and not isSegListChecked:
                self.config['segmentListChecked'] = 'n'
                self.config['segVariable'] = ''
                self.config['varValues'] = ''
                if self.standardMatch.isChecked():
                    if self.sdaOnly.isChecked() or self.bdaOnly.isChecked():
                        self.config['listMatchType'] = 'None'
                    if not self.sdaOnly.isChecked() and not self.bdaOnly.isChecked():
                        self.config['listMatchType'] = 'Standard'
                if self.exactMatch.isChecked():
                    self.config['listMatchType'] = 'Exact'
                    if self.sdaOnly.isChecked() or self.bdaOnly.isChecked():
                        self.config['listMatchType'] = 'None'
            if (self.standardMatch.isChecked() or self.exactMatch.isChecked()) and isSegListChecked:
                self.config['segmentListChecked'] = 'y'
                self.config['segVariable'] = str(self.segVariable.text())
                self.config['varValues'] = str(self.segValues.text())
                self.segsList = str(self.segValues.text()).split(' ')
                self.neededSegs = ", ".join('"{0}"'.format(w) for w in self.segsList[0:-1])
                self.backFillSeg = '"'+self.segsList[-1]+'"'
                self.config['neededSegs'] = str(self.neededSegs)
                self.config['backFillSeg'] = str(self.backFillSeg)

                # print self.neededSegs
                # print self.backFillSeg

                if self.standardMatch.isChecked():
                    self.config['listMatchType'] = 'Standard_Seg'
                if self.exactMatch.isChecked():
                    self.config['listMatchType'] = 'Exact_Seg'
            if str(self.dataCap.text()) == '':
                self.config['dataCap'] = '""'
            if str(self.dataCap.text()) != '':
                self.config['dataCap'] = str(self.dataCap.text())
            # if self.suppCheck.isChecked():
            #     self.config['matchFileName'] = self.fileWindow.selectMatchFile
            #     self.config['suppFileName'] = self.fileWindow.selectSuppFile
            #     self.config['suppMatchType'] = self.fileWindow.suppMatchType
            if self.keepSeg.isChecked():
                self.config['keep_seg'] = 'Yes'
            if not isKeepSegChecked:
                self.config['keep_seg'] = 'No'
            if self.cmiCompass.isChecked():
                self.config['cmi_compass_client'] = 'Y'
            if not isCmiCompassChecked:
                self.config['cmi_compass_client'] = 'N'

        if self.suppCheck.isChecked():
            self.config['matchFileName'] = self.fileWindow.selectMatchFile
            self.config['suppFileName'] = self.fileWindow.selectSuppFile
            self.config['suppMatchType'] = self.fileWindow.suppMatchType
            self.config['suppSASFile'] = self.fileWindow.suppSasFileInput
        if self.pivotTableCheck.isChecked():
            self.config['createPivotTable'] = 'Y'
            self.config['pivotSeg1'] = self.pivotWindow.selectSeg1
            self.config['pivotSeg2'] = self.pivotWindow.selectSeg2
        if not self.pivotTableCheck.isChecked():
            self.config['createPivotTable'] = 'N'
        if self.docAlert.isChecked():
            self.config['listProduct'] = str('DocAlert')
        if self.epocQuiz.isChecked():
            self.config['listProduct'] = str('Epoc_Quiz')
        if self.standardMatch.isChecked() and not isSegmentChecked and not isSegListChecked:
            self.config['listMatchType'] = str('Standard')
            if self.sdaOnly.isChecked() or self.bdaOnly.isChecked():
                self.config['listMatchType'] = 'None'
        if self.standardMatch.isChecked() and not isSegmentChecked and isSegListChecked:
            self.config['listMatchType'] = str('Standard_Seg')
            if self.sdaOnly.isChecked() or self.bdaOnly.isChecked():
                self.config['listMatchType'] = 'None'
        if self.exactMatch.isChecked() and not isSegmentChecked:
            self.config['listMatchType'] = 'Exact'
            if self.sdaOnly.isChecked() or self.bdaOnly.isChecked():
                self.config['listMatchType'] = 'None'
        if isFuzzyBoxChecked:
            self.config['listMatchType'] = 'Fuzzy'
        if self.sdaAddOn.isChecked():
            self.occList = str(self.sdaOcc.text()).split(", ")
            self.specList = str(self.sdaSpec.text()).split(", ")
            self.sdaOcc2 = ", ".join('"{0}"'.format(w) for w in self.occList)
            self.sdaSpec2 = ", ".join('"{0}"'.format(w) for w in self.specList)
            self.config['sDa'] = str('y')
            self.config['SDA_Occ'] = str(self.sdaOcc2)
            self.config['SDA_Spec'] = str(self.sdaSpec2).replace('"Internal Medicine", "General"', '"Internal Medicine, General"').replace('"Pediatrics", "General"', '"Pediatrics, General"').replace('"Pediatrics", "Subspecialty"', '"Pediatrics, Subspecialty"').replace('"Pediatrics", "Surgical"', '"Pediatrics, Surgical"').replace('"Radiology", "General"', '"Radiology, General"').replace('"Radiology", "Vascular and Interventional"', '"Radiology, Vascular and Interventional"').replace('"Surgery", "Cardiac/Thoracic"', '"Surgery, Cardiac/Thoracic"').replace('"Surgery", "General"', '"Surgery, General"').replace('"Surgery", "Plastic"', '"Surgery, Plastic"').replace('"Surgery", "Transplant"', '"Surgery, Transplant"').replace('"Surgery", "Vascular"', '"Surgery, Vascular"')
        if not isSdaChecked:
            self.config['sDa'] = str('n')
        if self.bdaAddOn.isChecked():
            self.config['bDa'] = str('y')
            self.bdaOccList = str(self.bdaOcc.text()).split(", ")
            self.bdaSpecList = str(self.bdaSpec.text()).split(", ")
            self.bdaOcc2 = ", ".join('"{0}"'.format(w) for w in self.bdaOccList)
            self.bdaSpec2 = ", ".join('"{0}"'.format(w) for w in self.bdaSpecList)
            self.config['occupation'] = str(self.bdaOcc2)
            self.config['specialty'] = str(self.bdaSpec2).replace('"Internal Medicine", "General"', '"Internal Medicine, General"').replace('"Pediatrics", "General"', '"Pediatrics, General"').replace('"Pediatrics", "Subspecialty"', '"Pediatrics, Subspecialty"').replace('"Pediatrics", "Surgical"', '"Pediatrics, Surgical"').replace('"Radiology", "General"', '"Radiology, General"').replace('"Radiology", "Vascular and Interventional"', '"Radiology, Vascular and Interventional"').replace('"Surgery", "Cardiac/Thoracic"', '"Surgery, Cardiac/Thoracic"').replace('"Surgery", "General"', '"Surgery, General"').replace('"Surgery", "Plastic"', '"Surgery, Plastic"').replace('"Surgery", "Transplant"', '"Surgery, Transplant"').replace('"Surgery", "Vascular"', '"Surgery, Vascular"')
            self.config['totalLookUps'] = str(self.bdaNumLookUps.text())
            self.config['lookUpPeriod'] = str(self.bdaLookUps.text())
            if isTherapyChecked:
                self.config['therapyChecked'] = str('Yes')
            else:
                self.config['therapyChecked'] = str('No')
        if not isTherapyChecked:
            self.config['therapyChecked'] = str('No')
        if isBdaChecked and self.drugList.toPlainText():
            self.config['drugList'] = str(self.createDrugList())
        if not isBdaChecked:
            self.config['bDa'] = str('n')
        if isSdaChecked:
            self.config['SDA_Target'] = str(self.sdaTargetNum.text())
        if isBdaChecked:
            self.config['BDA_Target'] = str(self.bdaTargetNum.text())
        sourceList = []
        for i in xrange(self.sourceSegs.count()):
            checkCol =  str(self.sourceSegs.item(i).text()).lower()
            sourceList.append(checkCol)
        for item in sourceList:
            if str(item) == 'full_name' or str(item) == 'fullname' or item == 'prescriber_name' or re.search('^full.+name', item) or re.search('.+full.+name', item):
                self.config['foundFullName'] = 'y'
            # else: 
            #     self.config['foundFullName'] = 'n'
        if self.suppCheck.isChecked():
            self.config['suppressionApplied'] = 'Y'
            self.config['suppMatchFile'] = str(self.suppMatchPath.text())
        if not self.suppCheck.isChecked():
            self.config['suppressionApplied'] = 'N'
        if self.sdaOnly.isChecked():
            self.config['sdaOnly'] = 'Y'
            if self.tabWidget.currentIndex() == 0:
                self.config['matchedFile'] = str(self.suppMatchPath.text())
        if not self.sdaOnly.isChecked():
            self.config['sdaOnly'] = 'N'
        if self.bdaOnly.isChecked():
            self.config['bdaOnly'] = 'Y'
            if self.tabWidget.currentIndex() == 0:
                self.config['matchedFile'] = str(self.suppMatchPath.text())
        if not self.bdaOnly.isChecked():
            self.config['bdaOnly'] = 'N'
        # print self.filePage.fileDict
        self.config['loadedFile'] = str(self.loadedFile)

        #write all the new setting first when the RUN BUTTON is clicked
        with open(desktop+'\\'+'config.json', 'w') as outfile1:
            json.dump(self.config, outfile1, indent=2, sort_keys=True, encoding='utf8')

        #This config is built dynamically in the gui and then read to be added to the main config.json file
        with open(desktop+'\\'+'sdaConfig.json', 'r') as infileSDA:
            sdaData = json.loads(infileSDA.read(), encoding='utf8')

        #This config is built dynamically in the gui and then read to be added to the main config.json file    
        with open(os.path.join(desktop, 'bdaConfig.json'), 'r') as infileBDA:
            bdaData = json.loads(infileBDA.read(), encoding='utf8')

        #Read the Newly written config file from above and load its data as masterFile and then update it with the other 2 config files
        with open(desktop+'\\'+'config.json', 'r') as infile:
            masterData = json.loads(infile.read(), encoding='utf8')

            masterData.update(sdaData)
            masterData.update(bdaData)

        #re write the final file one last time with all the new data
        with open(desktop+'\\'+'config.json', 'w') as outfile2:
            json.dump(masterData, outfile2, indent=2, sort_keys=True, encoding='utf8')

    def createDrugList(self):
        self.drugContents = self.drugList.toPlainText()
        commaCount = 0
        for char in self.drugContents:
            if char == ',':
                commaCount += 1
        if commaCount > 0:
            self.drugContents = str(self.drugContents).split(", ")
            self.drugContents = "\n".join(self.drugContents)
            return self.drugContents
        else:
            return self.drugContents

    #Function to add source segments to final segments list
    def addSeg(self):
        # sort rows in descending order in order to compensate shifting due to takeItem
        rows = sorted([index.row() for index in self.sourceSegs.selectedIndexes()],
                      reverse=True)
        for row in rows:
            # assuming the other listWidget is called listWidget_2
            self.finalSegs.addItem(self.sourceSegs.takeItem(row))

    #Function to remove segments from final segment list
    def removeSeg(self):
        # sort rows in descending order in order to compensate shifting due to takeItem
        rows = sorted([index.row() for index in self.finalSegs.selectedIndexes()],
                      reverse=True)
        for row in rows:
            # assuming the other listWidget is called listWidget_2
            self.sourceSegs.addItem(self.finalSegs.takeItem(row))

    def resetEditTab(self):
        self.sourceSegs.clearSelection()
        self.uniqueValuesList.clearSelection()
        self.uniqueValuesList.clear()
        self.renameValueEdit.clear()
        self.newColumnName.clear()
        self.addColumnText.clear()
        self.concatenateList.clear()

    def sdaCallback(self):
        isSdaChecked = self.sdaAddOn.isChecked()
        self.sdaAddOn.setStyleSheet("color: green; font-weight: bold")
        self.sdaOcc.setEnabled(isSdaChecked)
        self.addSDAButton.setEnabled(isSdaChecked)
        self.sdaOcc.setText('MD, DO, NP, PA')
        self.sdaSpec.setEnabled(isSdaChecked)
        self.sdaOnly.setEnabled(isSdaChecked)
        self.suppSDAOnly.setEnabled(isSdaChecked)
        if self.tabWidget.currentIndex() == 1:
            self.sdaTargetNum.setEnabled(isSdaChecked)
        if self.tabWidget.currentIndex() == 0:
            self.sdaTargetNum.setEnabled(False)
            self.sdaTargetNum.clear()
        if self.sdaOnly.isChecked():
            self.segBox.setEnabled(False)
            self.segBox.setChecked(False)
        if not isSdaChecked:
            self.sdaAddOn.setStyleSheet("color: black")
            self.sdaOcc.clear()
            self.sdaSpec.clear()
            self.sdaOnly.setChecked(False)
            self.segBox.setEnabled(True)
            self.suppSDAOnly.setChecked(False)



    def gskCallback(self):
        brandLower = str(self.presalesManuName.text()).lower()
        if brandLower == 'gsk':
            self.exactMatch.setChecked(True)

    def sda_bda_options(self):
        isSuppSdaChecked = self.suppSDAOnly.isChecked()
        isSuppBdaChecked = self.suppBDAOnly.isChecked()
        istherapyChecked = self.therapyClassBox.isChecked()
        if isSuppSdaChecked:
            self.suppSDAOnly.setStyleSheet("color: green; font-weight: bold")
        if isSuppBdaChecked:
            self.suppBDAOnly.setStyleSheet("color: green; font-weight: bold")
        if istherapyChecked:
            self.therapyClassBox.setStyleSheet("color: green; font-weight: bold")
        if not isSuppSdaChecked:
            self.suppSDAOnly.setStyleSheet("color: black")
        if not isSuppBdaChecked:
            self.suppBDAOnly.setStyleSheet("color: black")
        if not istherapyChecked:
            self.therapyClassBox.setStyleSheet("color: black")



    def sdaOnlyCallback(self):
        isSdaOnlyChecked = self.sdaOnly.isChecked()
        isBdaOnlyChecked = self.bdaOnly.isChecked()
        isSegmentListChecked = self.segmentList.isChecked()
        if isSdaOnlyChecked or isBdaOnlyChecked:
            if isSdaOnlyChecked:
                self.sdaOnly.setStyleSheet("color: green; font-weight: bold")
            if isBdaOnlyChecked:
                self.bdaOnly.setStyleSheet("color: green; font-weight: bold")

        if not isSdaOnlyChecked:
            self.sdaOnly.setStyleSheet("color: black")
        if not isBdaOnlyChecked:
            self.bdaOnly.setStyleSheet("color: black")


    def bdaCallback(self):
        isBdaChecked = self.bdaAddOn.isChecked()
        self.bdaAddOn.setStyleSheet("color: green; font-weight: bold")
        self.bdaOcc.setEnabled(isBdaChecked)
        self.bdaOcc.setText('MD, DO, NP, PA')
        self.bdaSpec.setEnabled(isBdaChecked)
        self.bdaLookUps.setEnabled(isBdaChecked)
        self.bdaOnly.setEnabled(isBdaChecked)
        self.addBDAButton.setEnabled(isBdaChecked)
        self.bdaLookUps.setText('31')
        self.bdaNumLookUps.setEnabled(isBdaChecked)
        self.bdaNumLookUps.setText('1')
        self.drugList.setEnabled(isBdaChecked)
        self.deDup.setEnabled(isBdaChecked)
        self.therapyClassBox.setEnabled(isBdaChecked)
        self.suppBDAOnly.setEnabled(isBdaChecked)
        if self.tabWidget.currentIndex() == 1:
            self.bdaTargetNum.setEnabled(isBdaChecked)
        if self.tabWidget.currentIndex() == 0:
            self.bdaTargetNum.setEnabled(False)
            self.bdaTargetNum.clear()
        if not isBdaChecked:
            self.bdaOcc.clear()
            self.bdaSpec.clear()
            self.bdaLookUps.clear()
            self.bdaNumLookUps.clear()
            self.drugList.clear()
            self.bdaTargetNum.clear()
            self.bdaOnly.setChecked(False)
            self.deDup.setChecked(False)
            self.therapyClassBox.setChecked(False)
            self.suppBDAOnly.setChecked(False)
            self.bdaAddOn.setStyleSheet("color: black")

    def set_SDA_BDA_Only_colors(self):
        sdaBdaNone = [self.caseNumber, self.requesterInIt, self.presalesManuName, self.saleExecutive, self.brandName, self.myInIt]
        isSdaOnlyChecked = self.sdaOnly.isChecked()
        isBdaOnlyChecked = self.bdaOnly.isChecked()
        if isSdaOnlyChecked or isBdaOnlyChecked:
            self.matchTypeSelection.setCurrentIndex(0)
            self.matchTypeSelection.setEnabled(False)
            self.matchTypeSelection.setStyleSheet("background-color: rgb(223, 138, 138, 255)")
            self.postgresTable.setEnabled(False)
            self.postgresTable.setStyleSheet("background-color: rgb(223, 138, 138, 255)")
            if self.suppMatchPath.text() == "":
                self.caseNumber.setEnabled(True)
                self.requesterInIt.setEnabled(True)
                self.saleExecutive.setEnabled(True)
                self.brandName.setEnabled(True)
                self.presalesManuName.setEnabled(True)
                self.myInIt.setEnabled(True)
                for item in sdaBdaNone:
                    if item.text() != "":
                        self.suppMatchPath.setEnabled(False)
                        # self.matchTypeSelection.setEnabled(False)
                        # self.postgresTable.setEnabled(False)
                        self.caseNumber.setStyleSheet("background-color: rgb(213, 233, 221,255);")
                        self.requesterInIt.setStyleSheet("background-color: rgb(213, 233, 221,255)")
                        self.saleExecutive.setStyleSheet("background-color: rgb(213, 233, 221,255)")
                        # self.matchTypeSelection.setStyleSheet("background-color: rgb(223, 138, 138, 255)")
                        self.brandName.setStyleSheet("background-color: rgb(213, 233, 221,255)")
                        self.presalesManuName.setStyleSheet("background-color: rgb(213, 233, 221,255)")
                        self.myInIt.setStyleSheet("background-color: rgb(213, 233, 221,255)")
                        self.emailRecip.setStyleSheet("background-color: rgb(213, 233, 221,255)")
                        break

                    elif item.text() == "":
                        self.suppMatchPath.setEnabled(True)
                        self.caseNumber.setStyleSheet("")
                        self.requesterInIt.setStyleSheet("")
                        self.saleExecutive.setStyleSheet("")
                        # self.matchTypeSelection.setStyleSheet("")
                        self.brandName.setStyleSheet("")
                        self.presalesManuName.setStyleSheet("")
                        self.myInIt.setStyleSheet("")
                        self.suppMatchPath.setStyleSheet("")
                        self.emailRecip.setStyleSheet("")

            elif self.suppMatchPath.text() != "":
                self.caseNumber.setStyleSheet("")
                self.caseNumber.setEnabled(False)
                self.requesterInIt.setStyleSheet("")
                self.requesterInIt.setEnabled(False)
                self.saleExecutive.setStyleSheet("")
                self.saleExecutive.setEnabled(False)
                # self.matchTypeSelection.setStyleSheet("")
                self.brandName.setStyleSheet("")
                self.brandName.setEnabled(False)
                self.presalesManuName.setStyleSheet("")
                self.presalesManuName.setEnabled(False)
                self.myInIt.setStyleSheet("")
                self.myInIt.setEnabled(False)
                self.suppMatchPath.setStyleSheet("")
                self.emailRecip.setStyleSheet("")
                    
        if not isBdaOnlyChecked and not isSdaOnlyChecked:
            self.caseNumber.setStyleSheet("")
            self.caseNumber.setEnabled(True)
            self.requesterInIt.setStyleSheet("")
            self.requesterInIt.setEnabled(True)
            self.saleExecutive.setStyleSheet("")
            self.saleExecutive.setEnabled(True)
            self.matchTypeSelection.setEnabled(True)
            self.matchTypeSelection.setStyleSheet("")
            self.brandName.setStyleSheet("")
            self.brandName.setEnabled(True)
            self.presalesManuName.setStyleSheet("")
            self.presalesManuName.setEnabled(True)
            self.myInIt.setStyleSheet("")
            self.myInIt.setEnabled(True)
            self.suppMatchPath.setStyleSheet("")
            self.emailRecip.setStyleSheet("")   
            self.postgresTable.setEnabled(True)
            self.postgresTable.setStyleSheet("")


    def segmentCallBack(self):
        isSegmentChecked = self.segBox.isChecked()
        if isSegmentChecked:
            self.segBox.setStyleSheet("color: green; font-weight: bold")
        self.sourceSegs.setEnabled(isSegmentChecked)
        self.finalSegs.setEnabled(isSegmentChecked)
        if not isSegmentChecked:
            self.segBox.setStyleSheet("color: black")
            if self.finalSegs:
                while self.finalSegs:
                    for index in xrange(self.finalSegs.count()): 
                        self.sourceSegs.addItem(self.finalSegs.takeItem(index))

    def standardCallback(self):
        isStandardChecked = self.standardMatch.isChecked()
        isFuzzyBoxChecked = self.fuzzyBox.isChecked()
        isExactChecked = self.exactMatch.isChecked()
        if isStandardChecked:
            self.standardMatch.setStyleSheet("color: green; font-weight: bold")
            self.exactMatch.setChecked(False)
            self.exactMatch.setStyleSheet("color: black")
            self.fuzzyBox.setChecked(False)
            self.fuzzyBox.setStyleSheet("color: black")
        elif not isStandardChecked and not isExactChecked and not isFuzzyBoxChecked:
            self.exactMatch.setChecked(True)
            self.standardMatch.setStyleSheet("color: black")

    def exactCallback(self):
        isStandardChecked = self.standardMatch.isChecked()
        isFuzzyBoxChecked = self.fuzzyBox.isChecked()
        isExactChecked = self.exactMatch.isChecked()
        if isExactChecked:
            self.exactMatch.setStyleSheet("color: green; font-weight: bold")
            self.standardMatch.setChecked(False)
            self.standardMatch.setStyleSheet("color: black")
            self.fuzzyBox.setChecked(False)
            self.fuzzyBox.setStyleSheet("color: black")
        elif not isStandardChecked and not isExactChecked and not isFuzzyBoxChecked:
            self.exactMatch.setStyleSheet("color: black")
            self.standardMatch.setChecked(True)

    def fuzzyCallback(self):
        isStandardChecked = self.standardMatch.isChecked()
        isFuzzyBoxChecked = self.fuzzyBox.isChecked()
        isExactChecked = self.exactMatch.isChecked()
        if isFuzzyBoxChecked:
            self.fuzzyBox.setStyleSheet("color: green; font-weight: bold")
            self.standardMatch.setChecked(False)
            self.standardMatch.setStyleSheet("color: black")
            self.exactMatch.setChecked(False)
            self.exactMatch.setStyleSheet("color: black")
        elif not isStandardChecked and not isExactChecked and not isFuzzyBoxChecked:
            self.fuzzyBox.setStyleSheet("color: black")
            self.standardMatch.setChecked(True)

    def docalertCallback(self):
        isDocChecked = self.docAlert.isChecked()
        if isDocChecked:
            self.docAlert.setStyleSheet("color: green; font-weight: bold")
            self.epocQuiz.setChecked(False)
        elif not isDocChecked:
            self.epocQuiz.setChecked(True)
            self.docAlert.setStyleSheet("color: black")

    def epocquizCallback(self):
        isEpocChecked = self.epocQuiz.isChecked()
        if isEpocChecked:
            self.epocQuiz.setStyleSheet("color: green; font-weight: bold")
            self.docAlert.setChecked(False)
        elif not isEpocChecked:
            self.docAlert.setChecked(True)
            self.epocQuiz.setStyleSheet("color: black")

    def cmiCompasCallback(self):
        isCMIChecked = self.cmiCompass.isChecked()
        if isCMIChecked:
            self.segBox.setChecked(True)
        else:
            self.segBox.setChecked(False)

    def editListCallback(self):
        isSegmentChecked = self.segBox.isChecked()
        if self.tabWidget.currentIndex() == 2:
            self.addSegButton.setEnabled(False)
            self.removeSegButton.setEnabled(False)
            self.segBox.setEnabled(False)
            self.sourceSegs.setEnabled(True)
            self.runProgramButton.setEnabled(False)
        else:
            self.addSegButton.setEnabled(True)
            self.removeSegButton.setEnabled(True)
            self.runProgramButton.setEnabled(True)
            self.segBox.setEnabled(True)
            self.sourceSegs.setEnabled(isSegmentChecked)

    def segmentedListCallback(self):
        isSegmentListChecked = self.segmentList.isChecked()
        if isSegmentListChecked:
            self.segmentList.setStyleSheet("color: green; font-weight: bold")
            self.segVariable.setEnabled(True)
            self.segValues.setEnabled(True)
            self.keepSeg.setEnabled(True)
        else:
            self.segVariable.setEnabled(False)
            self.segVariable.clear()
            self.segValues.setEnabled(False)
            self.segValues.clear()
            self.keepSeg.setEnabled(False)
            self.keepSeg.setChecked(False)
            self.segmentList.setStyleSheet("color: black")

    #set all defaults on start up
    def setDefaults(self):
        self.tabWidget.setCurrentIndex(0)
        self.suppMatchPath.setEnabled(False)
        self.sdaOnly.setEnabled(False)
        self.bdaOnly.setEnabled(False)
        self.drugList.setEnabled(False)
        self.sourceSegs.setEnabled(False)
        self.finalSegs.setEnabled(False)
        self.docAlert.setChecked(True)
        self.standardMatch.setChecked(True)
        self.sdaOcc.setEnabled(False)
        self.sdaSpec.setEnabled(False)
        self.bdaOcc.setEnabled(False)
        self.bdaSpec.setEnabled(False)
        self.bdaLookUps.setEnabled(False)
        self.bdaNumLookUps.setEnabled(False)
        self.bdaTargetNum.setEnabled(False)
        self.sdaTargetNum.setEnabled(False)
        self.deDup.setEnabled(False)
        self.segBox.setChecked(False)
        self.brandName.clear()
        self.caseNumber.clear()
        self.brandName.clear()
        self.saleExecutive.clear()
        self.presalesManuName.clear()
        self.postgresTable.clear()
        self.requesterInIt.clear()
        self.myInIt.clear()
        self.matchTypeSelection.setCurrentIndex(0)
        self.emailRecip.setText(userID)
        self.targetManuName.setCurrentIndex(0)
        self.brandTargetName.clear()
        self.targetNumber.clear()
        self.postgresTargetTable.clear()
        self.segVariable.clear()
        self.segValues.clear()
        self.calendarLine.clear()
        self.myTargetInIt.clear()
        self.keepSeg.setEnabled(False)
        self.segVariable.setEnabled(False)
        self.segValues.setEnabled(False)
        self.addSDAButton.setEnabled(False)
        self.addBDAButton.setEnabled(False)

    def attachCallbacks(self):
        self.refreshFile()
        self.docalertCallback()
        self.epocquizCallback()
        self.segmentCallBack()
        self.standardCallback()
        self.exactCallback()
        self.fuzzyCallback()
        # self.targetCaseCallback()
        self.sdaCallback()
        self.bdaCallback()
        self.addSeg()
        self.removeSeg()
        self.dataSharingClients()
        self.getUniqueSegmentValues()
        self.renameValue()
        self.renameColumn()
        self.detectMatchType()


if __name__ == "__main__":
    app = QtGui.QApplication(sys.argv)
    window = MyApp()
    QtGui.QApplication.setStyle(QtGui.QStyleFactory.create('plastique'))

    window.show()
    sys.exit(app.exec_())
    
    