from openpyxl import Workbook, load_workbook
import re
import os
import sys
import xlrd
import csv
from shutil import copyfile
from fuzzywuzzy import fuzz
from fuzzywuzzy import process
import psycopg2
import unicodecsv
import pprint
from collections import Counter
import datetime
from os import listdir
from os.path import isfile, join, isdir
import itertools

userhome = os.path.expanduser('~')
desktop = userhome + '\\Desktop'
downloads = userhome + '\\Downloads'
newest = max(os.listdir(downloads), key=lambda f: os.path.getmtime("{}/{}".format(downloads, f)))

headerFile = os.path.join(downloads, 'Header_NNI_EMailSuppression_HCP_Header.xlsx')
myDatFile = []

def csv_from_excel():
	# newest = max(os.listdir(downloads), key=lambda f: os.path.getmtime("{}/{}".format(downloads, f)))
	wb = xlrd.open_workbook(headerFile)
	sh = wb.sheet_by_index(0)
	your_csv_file = open (downloads + '\\headerPipe.csv', 'wb')
	wr = unicodecsv.writer(your_csv_file, encoding='utf8', delimiter='|')
	# reader = csv.reader(open(downloads + '\\headerPipe.csv', 'r'))

	for rownum in range(sh.nrows):
		# print ", ".join(map(str, sh.row_values(rownum)))
		# if ", ".join(map(str, sh.row_values(rownum))).strip().strip(", ") != "":
			wr.writerow(sh.row_values(rownum))

	your_csv_file.close()


def findDat():
    onlyFiles = [f for f in listdir(downloads) if isfile(join(downloads, f))]
    for files in onlyFiles:
        filename, extension = os.path.splitext(files)
        filePart = filename.rsplit('\\')[-1]
        if extension == '.DAT':
            datFile = filePart+extension
            if filePart not in myDatFile:
                myDatFile.append(filePart)
            else:
                pass
        else:
            pass

    for item in myDatFile:
        return item


def changeDat():
    if os.path.exists(os.path.join(downloads, findDat()+'.DAT')):
        os.rename(os.path.join(downloads, findDat()+'.DAT'), os.path.join(downloads, findDat()+'.csv'))
    else:
        pass

def writeFinalFile():
    delCols = [26, 27, 28, 29, 30]
    with open(os.path.join(downloads, 'headerPipe.csv'), 'r') as inHeader, open(os.path.join(downloads, findDat()+'.csv'), 'r') as dataFile, open(os.path.join(downloads, 'novoSupp.txt'), 'w') as outFile:
        headerReader = csv.reader(inHeader, delimiter='|')
        dataReader = csv.reader(dataFile, delimiter='|')
        writer = csv.writer(outFile, delimiter='|', lineterminator='\n')

        header = headerReader.next()
        header.extend(['de1', 'del2', 'del3', 'del4', 'del5'])
        # print header
        writer.writerow(header)

        for row in dataReader:
            if row[10] == 'Y':
                writer.writerow(row)

    # with open(os.path.join(downloads, 'novoSupp1.csv'), 'r') as inFile1, open(os.path.join(downloads, 'novoSupp.txt'), 'w') as outFile1:
    #     reader = csv.reader(inFile1, delimiter='|')
    #     writer = csv.writer(outFile1, delimiter='|')

    #     header2 = reader.next()
    #     writer.writerow(header2)

    #     for row in reader:
    #         if row[10] == 'Y':
    #             writer.writerow(row)


csv_from_excel()
changeDat()
writeFinalFile()