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


userhome = os.path.expanduser('~')
downloads = userhome + '\\Downloads\\'
csvFile = 'target.csv'
# newest = max(os.listdir(downloads), key=lambda f: os.path.getmtime("{}/{}".format(downloads, f)))


def fetchColumns():
	with open(downloads + 'csvFile.csv', 'rb') as f:
		reader = csv.reader(f)
		i = reader.next()
		columns = [row for row in reader]

		return i

def checkExtension():
	newest = max(os.listdir(downloads), key=lambda f: os.path.getmtime("{}/{}".format(downloads, f)))
	filename, extension = os.path.splitext(os.path.join(downloads, newest))
	# print extension
	if extension == '.xlsx':
		csv_from_excel()
	elif extension == '.txt':
		with open(os.path.join(downloads, newest), 'rb') as f:
			f = f.read()
			totalpipecount = 0
			totaltabcount = 0
			for char in f:
				if char == '\t':
					totaltabcount += 1
				if char == '|':
					totalpipecount += 1
			if totalpipecount > totaltabcount:
				pipe_to_csv()
			if totaltabcount > totalpipecount:
				tab_to_csv()
	elif extension == '.csv':
		copyfile(os.path.join(downloads, newest), os.path.join(downloads, 'target.csv'))
	else:
		with open(os.path.join(downloads, 'target.csv'), 'wb') as out:
			writer = csv.writer(out, lineterminator='\n')
			writer.writerows([['No File Found']])


def pipe_to_csv():
	newest = max(os.listdir(downloads), key=lambda f: os.path.getmtime("{}/{}".format(downloads, f)))
	with open(os.path.join(downloads, newest), 'rb') as f, open(os.path.join(downloads, 'target.csv'), 'wb') as out:
		pipereader = csv.reader(f, delimiter='|')
		csvwriter = csv.writer(out, delimiter=',', lineterminator='\n')
		for row in pipereader:
			csvwriter.writerow(row)

def tab_to_csv():
	newest = max(os.listdir(downloads), key=lambda f: os.path.getmtime("{}/{}".format(downloads, f)))
	with open(os.path.join(downloads, newest), 'rb') as f, open(os.path.join(downloads, 'target.csv'), 'wb') as out:
		tabreader = csv.reader(f, delimiter='\t')
		csvwriter = csv.writer(out, delimiter=',', lineterminator='\n')
		for row in tabreader:
			csvwriter.writerow(row)

def countSheets():
	totalSheets = len(wb.sheet_names())
	return

def csv_from_excel():
	newest = max(os.listdir(downloads), key=lambda f: os.path.getmtime("{}/{}".format(downloads, f)))
	wb = xlrd.open_workbook(downloads + newest)
	sh = wb.sheet_by_index(0)
	your_csv_file = open (downloads + 'target.csv', 'wb')
	wr = unicodecsv.writer(your_csv_file, encoding='utf8', lineterminator='\n')
	reader = csv.reader(open(downloads + 'target.csv', 'r'))

	for rownum in range(sh.nrows):
		# print ", ".join(map(str, sh.row_values(rownum)))
		# if ", ".join(map(str, sh.row_values(rownum))).strip().strip(", ") != "":
			wr.writerow(sh.row_values(rownum))

	your_csv_file.close()

def csv_from_excel2(test=None):
	matched = test
	wb = xlrd.open_workbook(downloads + matched)
	sh = wb.sheet_by_index(0)
	your_csv_file = open (downloads + 'target.csv', 'wb')
	wr = unicodecsv.writer(your_csv_file, encoding='utf8', lineterminator='\n')

	for rownum in range(sh.nrows):
		# print ", ".join(map(str, sh.row_values(rownum)))
		# if ", ".join(map(str, sh.row_values(rownum))).strip().strip(", ") != "":
		wr.writerow(sh.row_values(rownum))

	your_csv_file.close()

def checkExtension2(test=None):
	matched = test
	filename, extension = os.path.splitext(os.path.join(downloads, matched))
	# print extension
	if extension == '.xlsx':
		csv_from_excel2(matched)
	elif extension == '.txt':
		with open(os.path.join(downloads, matched), 'rb') as f:
			f = f.read()
			totalpipecount = 0
			totaltabcount = 0
			for char in f:
				if char == '\t':
					totaltabcount += 1
				if char == '|':
					totalpipecount += 1
			if totalpipecount > totaltabcount:
				pipe_to_csv2(matched)
			if totaltabcount > totalpipecount:
				tab_to_csv2(matched)
	elif extension == '.csv':
		copyfile(os.path.join(downloads, matched), os.path.join(downloads, 'target.csv'))
	else:
		with open(os.path.join(downloads, 'target.csv'), 'wb') as out:
			writer = csv.writer(out, lineterminator='\n')
			writer.writerows([['No File Found']])

def tab_to_csv2(test=None):
	matched = test
	with open(os.path.join(downloads, matched), 'rb') as f, open(os.path.join(downloads, 'target.csv'), 'wb') as out:
		tabreader = csv.reader(f, delimiter='\t')
		csvwriter = csv.writer(out, delimiter=',', lineterminator='\n')
		for row in tabreader:
			csvwriter.writerow(row)

def pipe_to_csv2(test=None):
	matched = test
	with open(os.path.join(downloads, matched), 'rb') as f, open(os.path.join(downloads, 'target.csv'), 'wb') as out:
		pipereader = csv.reader(f, delimiter='|')
		csvwriter = csv.writer(out, delimiter=',', lineterminator='\n')
		for row in pipereader:
			csvwriter.writerow(row)

def removeChar():
	inputFile = open(downloads + csvFile, 'r')
	outputFile = open(downloads + 'csvFile1.csv', 'wb')
	conversion = '-/%$# @<>+*?&)('
	numbers = '0123456789'
	newtext = '_'

	index = 0
	for line in inputFile:
		if index == 0:
			for c in conversion:
				line = line.replace(c, newtext)
		outputFile.write(line)
		index += 1

def incDupColumns():
	with open(downloads + 'csvFile1.csv', 'r') as myFile, open(downloads + 'csvFile.csv', 'wb') as myOut:
		reader = csv.reader(myFile)
		headers = reader.next()
		headersList = []
		visited = []
		inc = 1
		for header in headers:
			headersList.append(header)
		for i, x in enumerate(headersList):
			if x not in visited:
				visited.append(headersList[i])
			else:
				dup = x +'_'+str(inc)
				if dup not in visited:
					visited.append(x+'_'+str(inc))
				else:
					inc += 1
					visited.append(x+'_'+str(inc))

		w = csv.writer(myOut, lineterminator='\n')
		w.writerow(visited)
		for row in reader:
			w.writerow(row)

def cmiCompasCheck():
	cmiColumns = []
	with open(downloads + 'csvFile.csv', 'rb') as f:
		reader = csv.reader(f)
		i = reader.next()
		columns = [row for row in reader]
		for col in columns:
			cellVal = str(col).lower().replace('/', '_').replace('-', '_')
			if cellVal == 'state' or cellVal == 'state_code':
				cmiColumns.append(cellVal)
			elif cellVal == 'address_1' or cellVal == 'address 1' or cellVal == 'address1':
				cmiColumns.append(cellVal)
			elif cellVal == 'client_id' or cellVal == 'client_id_1':
				cmiColumns.append(cellVal)
			elif cellVal == 'segment1' or cellVal == 'segment2' or cellVal == 'segment' or re.search('^segment.+', cellVal) or re.search('.+segment.+', cellVal):
				cmiColumns.append(cellVal)

	return cmiColumns


def importDrugs():
	drugComplete = []
	with open("P:\\Epocrates Analytics\\Drug Compare\\Master Drug List\\drugs.csv", 'rb') as inFile:
		reader = csv.DictReader(inFile)
		for item in reader:
			drugs = item['drugs']
			drugComplete.append(drugs)
	return drugComplete


def main():
	fetchColumns()
	checkExtension()
	pipe_to_csv()
	tab_to_csv()
	csv_from_excel2(test)
	checkExtension2(test)
	pipe_to_csv2(test)
	tab_to_csv2(test)
	csv_from_excel()
	removeChar()
	incDupColumns()
	cmiCompasCheck()


if __name__ == "__main__":
	main()