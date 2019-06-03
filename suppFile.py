import pandas
from openpyxl import Workbook, load_workbook
import re
import numpy
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
import json

listText = 'supp.txt'
colList = []
finalDrugs = []
unmatchedDrugs = []
segmentList = []
segmentListSingle = []
colLength = None
csvFile = 'target.csv'
csvFileMod = 'target_mod.csv'
csvFileFinal = 'target_final.csv'
csvFileModTemp2 = 'target_mod_temp2.csv'
dt = str(datetime.datetime.now().strftime("%Y%m%d"))

args = sys.argv[1:]

userhome = os.path.expanduser('~')
downloads = userhome + '\\Downloads\\'
desktop = userhome + '\\Desktop\\'
newest = max(os.listdir(downloads), key=lambda f: os.path.getctime("{}/{}".format(downloads, f)))

with open(os.path.join(desktop, 'TheEagleHasLanded.csv'), 'rb') as passFile:
	reader = csv.DictReader(passFile)
	for item in reader:
		password = item['password']

if len(args) > 0:
	with open(os.path.join(desktop, args[0]), 'r') as infile:
		config = json.loads(infile.read(), encoding='utf8')

		listMatchType = str(config['suppMatchType'])
		if str(config['caseType']) == 'Targeting':
			date = str(config['date'])
		brand = str(config['Brand'])
		manu = str(config['Manu'])
		yourIn = str(config['yourIn'])
		if str(config['caseType']) == 'listMatch':
			reIn = str(config['reIn'])
			name = "{requester}_{manufact}_{brand}_{dt}_{initials}".format(requester=reIn, manufact=manu, brand=brand, dt=dt, initials=yourIn)
		tableName = str(config['tableName'])+'_supp'

def csv_from_excel():
	newest = str(config['suppFileName'])
	wb = xlrd.open_workbook(downloads + newest)
	sh = wb.sheet_by_index(0)
	your_csv_file = open (downloads + 'target.csv', 'wb')
	wr = unicodecsv.writer(your_csv_file, encoding='utf8')

	for rownum in range(sh.nrows):
		wr.writerow(sh.row_values(rownum))

	your_csv_file.close()

def checkExtension():
	newest = str(config['suppFileName'])
	filename, extension = os.path.splitext(os.path.join(downloads, newest))
	print extension
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
	else:
		with open(os.path.join(downloads, 'target.csv'), 'wb') as out:
			writer = csv.writer(out)
			writer.writerows([['No File Found']])


def pipe_to_csv():
	newest = str(config['suppFileName'])
	with open(os.path.join(downloads, newest), 'rb') as f, open(os.path.join(downloads, 'target.csv'), 'wb') as out:
		pipereader = csv.reader(f, delimiter='|')
		csvwriter = csv.writer(out, delimiter=',')
		for row in pipereader:
			csvwriter.writerow(row)

def tab_to_csv():
	newest = str(config['suppFileName'])
	with open(os.path.join(downloads, newest), 'rb') as f, open(os.path.join(downloads, 'target.csv'), 'wb') as out:
		tabreader = csv.reader(f, delimiter='\t')
		csvwriter = csv.writer(out, delimiter=',')
		for row in tabreader:
			csvwriter.writerow(row)


def getMain():
	with open (downloads + 'csvFile.csv', 'rb') as inFile2, open(downloads + csvFileModTemp2, 'wb') as targetFile:
		r = csv.reader(inFile2)
		headers = r.next()
		for index, col in enumerate(headers):
			cellVal = str(col).lower().replace('/', '_').replace('-', '_')
			#Regular Expression Rules. We can add new Rules as we need to this list to cover more common cases.
			if cellVal == 'npi' or cellVal == 'npi_id' or re.search('.+ npi .+', cellVal) or re.search('.+_npi_.+', cellVal) or re.search('^npi.+', cellVal) or re.search('.+ npi', cellVal) or re.search('.+ npi', cellVal):
				print cellVal, ': ',  'I found a NPI Number'
				headers[index] = 'npi'
			elif cellVal == 'me' or cellVal == 'me_id' or cellVal == 'meded' or re.search('.+ me .+', cellVal) or  re.search('.+_me_.+', cellVal) or re.search('^me .+', cellVal) or re.search('.+ me', cellVal) or re.search('.+ me', cellVal) or re.search('^me_.+', cellVal):
				print cellVal, ': ', 'I found a ME Number'
				headers[index] = 'me'
			elif cellVal == 'fname' or re.search('^first.+name', cellVal) or re.search('.+first.+name', cellVal) or re.search('.+fname', cellVal) or re.search('.+first', cellVal) or re.search('.+frst.+', cellVal):
				print cellVal, ': ', 'I found a First Name'
				headers[index] = 'fname'
			elif re.search(r'lname|^l.+name|.+last|.+lname|.+last.+name', cellVal):
				print cellVal, ': ',  'I found a Last Name'
				headers[index] = 'lname'
			elif cellVal == 'zip_4' or cellVal == 'zip4':
				headers[index] = 'whatever'
			elif cellVal == 'Group' or cellVal == 'group':
				headers[index] = '_group'
			elif cellVal == 'zip' or cellVal == 'Postal' or (re.search('^zip.+', cellVal) and (cellVal != 'zip_4' or cellVal != 'zip4')) or re.search('^postal.+', cellVal) or re.search('.+_zip', cellVal) or re.search('.+ zip', cellVal) or re.search('.+_postal', cellVal) or re.search('.+ zip', cellVal) or re.search('.+ postal', cellVal):
				print cellVal, ': ',  'I found a Zip/Postal Code'
				headers[index] = 'zip'
			elif cellVal.startswith(('0','1','2','3','4','5','6','7','8','9')):
				print 'I added an underscore to: ', cellVal
				headers[index] = '_' + headers[index]

		w = csv.writer(targetFile)
		w.writerow(headers)
		for row in r:
			w.writerow(row)

	with open(downloads + csvFileModTemp2, 'r') as myFile, open(downloads + csvFileMod, 'wb') as myOut:
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

		w = csv.writer(myOut)
		w.writerow(visited)
		for row in reader:
			w.writerow(row)

def postgresConn():
	conn_string = "host='localhost' dbname='postgres' port='5432' user='postgres' password='{password}'".format(password=password)
	conn = psycopg2.connect(conn_string)
	cursor = conn.cursor()
	print "connected!\n"
	print "Checking if {} table exists. If soo ill drop dat jawn . . . ".format(tableName)

	sqlDrop = """DROP TABLE IF EXISTS {};
			     DROP TABLE IF EXISTS temp_table;""".format(tableName)
	cursor.execute(sqlDrop)
	conn.commit()

	sql = "select load_csv_file('{tableName}', '{downloads}{csvFile}', {get_Cols});".format(tableName=tableName, downloads=downloads, csvFile=csvFileMod, get_Cols = get_Cols())
	cursor.execute(sql)
	conn.commit()

	addMissing = """DO $$
				    BEGIN
				  		BEGIN
				  			ALTER Table {tableName} ADD COLUMN me text;
			  			Exception
			  				WHEN duplicate_column THEN RAISE NOTICE 'Column already exists';
		  				END;
		  				BEGIN
				  			ALTER Table {tableName} ADD COLUMN npi text;
			  			Exception
			  				WHEN duplicate_column THEN RAISE NOTICE 'Column already exists';
		  				END;
		  				BEGIN
				  			ALTER Table {tableName} ADD COLUMN fname text;
			  			Exception
			  				WHEN duplicate_column THEN RAISE NOTICE 'Column already exists';
		  				END;
		  				BEGIN
				  			ALTER Table {tableName} ADD COLUMN lname text;
			  			Exception
			  				WHEN duplicate_column THEN RAISE NOTICE 'Column already exists';
		  				END;
		  				BEGIN
				  			ALTER Table {tableName} ADD COLUMN zip text;
			  			Exception
			  				WHEN duplicate_column THEN RAISE NOTICE 'Column already exists';
		  				END;
	  				END;
  				$$""".format(tableName=tableName)
	cursor.execute(addMissing)
	conn.commit()

	strip = """UPDATE {tableName}
				SET zip = REPLACE("zip", '.0', '');

				UPDATE {tableName}
				SET npi = REPLACE("npi", '.0', '');

				UPDATE {tableName}
				SET me = REPLACE("me", '.0', '');""".format(tableName=tableName)


	cursor.execute(strip)
	conn.commit()

	if listMatchType =='Standard':
		exportSeg = """COPY (
							SELECT me, npi, fname, lname, zip FROM {tableName}
						 )
					 TO '{downloads}supp.txt' DELIMITER '	'  CSV HEADER;""".format(tableName=tableName, downloads=downloads)
		cursor.execute(exportSeg)
		conn.commit()

	elif listMatchType =='Exact':
		exportSeg = """COPY (
							SELECT me, npi FROM {tableName}
						 )
					 TO '{downloads}supp.txt' DELIMITER '	'  CSV HEADER;""".format(tableName=tableName, downloads=downloads)
		cursor.execute(exportSeg)
		conn.commit()


def removeChar():
	inputFile = open(downloads + csvFile, 'r')
	outputFile = open(downloads + 'csvFile.csv', 'wb')
	conversion = '-/%$# @<>+*?&'
	newtext = '_'

	index = 0
	for line in inputFile:
		if index == 0:
			for c in conversion:
				line = line.replace(c, newtext)

		outputFile.write(line)
		index += 1



def get_Cols():
	with open(downloads + 'target.csv', 'rb') as f:
		reader = csv.reader(f)
		i = reader.next()
		columns = [row for row in reader]

		colLength = len(i)
		return colLength

def get_cols_names():
	with open(downloads + 'target.csv', 'rb') as f:
		reader = csv.reader(f)
		i = reader.next()
		columns = [row for row in reader]

		return i


def removeFiles():
	# os.remove(os.path.join(downloads, 'target.txt'))
	os.remove(os.path.join(downloads, 'target_mod.csv'))
	os.remove(os.path.join(downloads, 'target.csv'))
	os.remove(os.path.join(downloads, 'csvFile.csv'))
	if str(config['caseType']) == 'Targeting':
		targetFolder = manu+" "+brand
		os.chdir("P:\\Epocrates Analytics\\TARGETS\\{date}{slashes}{targetFolder}".format(date = date, slashes = "\\", targetFolder=targetFolder))
		if os.path.exists("P:\\Epocrates Analytics\\TARGETS\\{date}{slashes}{targetFolder}\\Supp".format(date = date, slashes = "\\", targetFolder=targetFolder)):
			print 'Supp Folder already exist. Skipping creating this folder'
		if not os.path.exists("P:\\Epocrates Analytics\\TARGETS\\{date}{slashes}{targetFolder}\\Supp".format(date = date, slashes = "\\", targetFolder=targetFolder)):
			os.mkdir('Supp')
		copyfile(downloads + listText, "P:\\Epocrates Analytics\\TARGETS\\{date}{slashes}{targetFolder}\\Supp\\supp.txt".format(date = date, slashes = "\\", targetFolder=targetFolder))
	else:
		os.chdir("""P:\\Epocrates Analytics\\List Match\\List Match Folder\\{folderName}""".format(folderName = name))
		if os.path.exists("P:\\Epocrates Analytics\\List Match\\List Match Folder\\{folderName}\\Supp".format(slashes = "\\", folderName=name)):
			print 'Supp Folder already exist. Skipping creating this folder'
		if not os.path.exists("P:\\Epocrates Analytics\\List Match\\List Match Folder\\{folderName}\\Supp".format(slashes = "\\", folderName=name)):
			os.mkdir('Supp')
		copyfile(downloads + listText, "P:\\Epocrates Analytics\\List Match\\List Match Folder\\{folderName}\\Supp\\supp.txt".format(slashes = "\\", folderName=name))
	os.remove(os.path.join(downloads, listText))
	if os.path.exists(os.path.join(downloads, 'csvFile1.csv')):
		os.remove(os.path.join(downloads, 'csvFile1.csv'))
	if os.path.exists(os.path.join(downloads, csvFileModTemp2)):
		os.remove(os.path.join(downloads, csvFileModTemp2))

# def copyTarget():
# 	copyfile(downloads + listText, outFileFinal)

	print ''
	print ''
	print 'Suppression Program completed'

# csv_from_excel()

checkExtension()
removeChar()
get_Cols()
get_cols_names()
getMain()
postgresConn()
removeFiles()
