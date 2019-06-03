from openpyxl import Workbook, load_workbook
import re
import os
import sys
import xlrd
import csv
from shutil import copyfile
from fuzzywuzzy import fuzz
from fuzzywuzzy import process
# import psycopg2
import unicodecsv
# import pprint
from collections import Counter
import datetime
import json
from subprocess import *
# import subprocess
from os import listdir
from os.path import isfile, join
from colorama import init, Fore, Back, Style
from termcolor import colored
import time
import pandas
import sqlite3
import random

init(autoreset=True)

args = sys.argv[1:]

listText = 'target.txt'
colList = []
finalDrugs = []
unmatchedDrugs = []
segmentList = []
segmentListSingle = []
fullNameList = []
colLength = None
csvFile = 'target.csv'
csvFileMod = 'target_mod.csv'
csvFileModTemp = 'target_mod_temp.csv'
csvFileModTemp2 = 'target_mod_temp2.csv'
csvFileFinal = 'target_final.csv'
dt = str(datetime.datetime.now().strftime("%Y%m%d"))

cmiCompasSegmentation = "npi, address1, campaign_type, city, cl_fname, cl_lname, cl_me, cl_zip, clientid, compasid, middle_name, segment1, specialty, state_code, tier, segment2, segment3"
cmiCompasSQL = "address1, campaign_type, city, fname as cl_fname, lname as cl_lname, me as cl_me, zip as cl_zip, clientid, compasid, middle_name, segment1, specialty, state_code, tier, segment2, segment3"
cmiList = ['npi', 'address1', 'campaign_type', 'city', 'cl_fname', 'cl_lname', 'cl_me', 'cl_zip', 'clientid', 'compasid', 'middle_name', 'segment1', 'specialty', 'state_code', 'tier', 'segment2', 'segment3']

userhome = os.path.expanduser('~')
desktop = userhome + '\\Desktop\\'
# configFile = os.path.join(desktop, 'config.json')

randomList = []
with open('G:\\Communicator Ops\\Epocrates\\Python Files\\Import Phrases\\importPhrases.csv', 'r') as myPhrases:
	reader = csv.reader(myPhrases)
	for row in reader:
		for item in row:
			randomList.append(item)



# randomList = ['Indubitably', 'The first thing you gotta do is say . . . WOULD YOU LOOK AT THAT?', 'Tootles', 'This burrito is sooooo Filling', 'Have you seen Broke Back Rhonda? It Rhond-tastic!', 'How much wood would a woodchuck chuck if a woodchuck could chuck wood?']

with open(os.path.join(desktop, 'TheEagleHasLanded.csv'), 'rb') as passFile:
	reader = csv.DictReader(passFile)
	for item in reader:
		password = item['password']

if len(args) > 0:
	with open(os.path.join(desktop, args[0]), 'r') as infile:
		config = json.loads(infile.read(), encoding='utf8')
		foundFullName = 'n'
		if 'foundFullName' in config:
			foundFullName = config['foundFullName']
		caseType = str(config['caseType'])
		therapyClass = str(config['therapyChecked'])
		sDa_only = str(config['sdaOnly'])
		bDa_only = str(config['bdaOnly'])
		# finalSDATotal = int(config["totalAdditionalSDAs"]) + 1
		createPivotTable = str(config['createPivotTable'])
		if createPivotTable == 'Y':
			pivSeg1 = str(config['pivotSeg1']).lower()
			pivSeg2 = str(config['pivotSeg2']).lower()
		if createPivotTable == 'N':
			pivSeg1 = ""
			pivSeg2 = ""
		if caseType == 'listMatch':
			suppApplied = str(config['suppressionApplied'])
			listProduct = str(config['listProduct'])
			listMatchType = str(config['listMatchType'])
			manu = str(config['Manu'])
			if listMatchType == 'Standard_Seg' or listMatchType == 'Exact_Seg':
				addSeg1 = config['finalSegs']
				addSeg = ", ".join(addSeg1).lower().replace('Group', '_group').replace('group', '_group').replace('GROUP', '_group')
				finalSeg = addSeg.split(', ')
				for seg in finalSeg:
					if seg.startswith(('0','1','2','3','4','5','6','7','8','9')):
						seg = '_' + seg
					segmentList.append(str(seg).replace(' ', '_'))
					splitList = ", ".join(segmentList)
			brand = str(config['Brand'])
			yourIn = str(config['yourIn'])
			reIn = str(config['reIn'])
			sDa = str(config['sDa'])
			name = "{requester}_{manufact}_{brand}_{dt}_{initials}".format(requester=reIn, manufact=manu, brand=brand, dt=dt, initials=yourIn)
			if sDa == 'y':
				finalSDATotal = str(int(config["totalAdditionalSDAs"]) + 1)
				SDA_Occ = str(config['SDA_Occ'])
				SDA_Spec = str(config['SDA_Spec'])
				SDA_Occ2 = SDA_Occ.replace('"', '')
			if sDa == 'n':
				finalSDATotal = '0'
				SDA_Occ = '""'
				SDA_Spec = '""'
				SDA_Occ2 = ''
			bDa = str(config['bDa'])
			if (bDa == 'y' and sDa == 'n') or (bDa == 'y' and sDa == 'y'):
				finalBDATotal = str(int(config["totalAdditionalBDAs"]) + 1)
				deDup = str(config['deDup'])
				lookUpPeriod = str(config['lookUpPeriod'])
				displayPeriod = str(int(lookUpPeriod)-1)
				totalLookUps = str(config['totalLookUps'])
				occupation = str(config['occupation'])
				specialty = str(config['specialty'])
				occupation2 = occupation.replace('"', '')
				drugList = str(config['drugList'])
				drugsnocomma = str(config['drugList']).replace("\n", ", ").replace("'", '')
				# print drugsnocomma
				# drugsnocomma = ", ".join(drugList2)
			if bDa == 'n':
				finalBDATotal = '0'
				deDup = ""
				lookUpPeriod = ""
				totalLookUps = ""
				occupation = '""'
				specialty = '""'
				drugList = ""
				occupation2 = ''
				drugsnocomma = ''
			caseno = str(config['caseno'])
			mtype = str(config['mtype'])
			SE = str(config['SE'])
			email = str(config['email'])
			tableName = str(config['tableName'])
		elif caseType == 'Targeting':
			suppSDAOnly = str(config['suppSDAOnly'])
			suppBDAOnly = str(config['suppBDAOnly'])
			backFill = str(config['backFill'])
			dataCap = str(config['dataCap'])
			suppApplied = str(config['suppressionApplied'])
			listProduct = str(config['listProduct'])
			listMatchType = str(config['listMatchType'])
			date = str(config['date'])
			sDa_only = str(config['sdaOnly'])
			bDa_only = str(config['bdaOnly'])
			if bDa_only  == 'N' and sDa_only == 'N':
				if listMatchType == 'Standard' or listMatchType == 'Exact':
					targetNum = str(config['targetNum'])
					segVariable = str(config['segVariable']).lower()
					varValues = ''
					neededValues = ''
					backFillValue = ''
				if listMatchType == 'Standard_Seg' or listMatchType == 'Exact_Seg':
					targetNum = str(config['targetNum'])
					segVariable = str(config['segVariable']).lower()
					varValues = str(config['varValues'])
					neededValues = str(config['neededSegs'])
					backFillValue = str(config['backFillSeg'])

					print backFillValue
			if bDa_only  == 'Y' or sDa_only == 'Y':
				listMatchType == 'None'
				neededValues = ''
				backFillValue = ''
			dSharing = str(config['dSharing'])
			if dSharing == 'Y' and config['cmi_compass_client'] == 'N':
				addSeg1 = config['finalSegs']
				addSeg = ", ".join(addSeg1).lower().replace('Group', '_group').replace('group', '_group').replace('GROUP', '_group')
				if config['segVariable'] != "":
					finalSeg = addSeg.split(', ')
					finalSeg.append(segVariable)
					for seg in finalSeg:    
						segmentList.append(str(seg).replace(' ', '_'))
						splitList = ", ".join(segmentList)
				
				if segVariable == '':
					finalSeg = addSeg.split(', ')
					for seg in finalSeg:    
						segmentList.append(str(seg).replace(' ', '_'))
						splitList = ", ".join(segmentList)
			if dSharing == 'Y' and config['cmi_compass_client'] == 'Y':
				addSeg = cmiCompasSegmentation
				segVariable = str(config['segVariable'])
				if config['segVariable'] != "":
					finalSeg = addSeg.split(', ')
					finalSeg.append(segVariable)
					for seg in finalSeg:    
						segmentList.append(str(seg).replace(' ', '_'))
						splitList = ", ".join(segmentList)
				
				if segVariable == '':
					finalSeg = addSeg.split(', ')
					for seg in finalSeg:    
						segmentList.append(str(seg).replace(' ', '_'))
						splitList = ", ".join(segmentList)
			if dSharing == 'N' and bDa_only  == 'N' and sDa_only == 'N':
				dSharing = 'N'
				finalSeg = str(config['segVariable']).lower().replace(' ', '_')
				keep_seg = 'No'
				if config['segVariable'] != "":
					splitList = finalSeg

			if dSharing == 'N' and (bDa_only  == 'Y' or sDa_only == 'Y'):
				dSharing = 'N'
				finalSeg = ''
				keep_seg = 'No'
			keep_seg = str(config['keep_seg'])
			manu = str(config['Manu'])
			brand = str(config['Brand'])
			yourIn = str(config['yourIn'])
			sDa = str(config['sDa'])
			name = "T_{manu}_{brand}_{dt}_{initials}".format(manu=manu, brand=brand, dt=dt, initials=yourIn)
			if sDa == 'y':
				sDa_only = sDa_only
				SDA_Occ = str(config['SDA_Occ'])
				SDA_Spec = str(config['SDA_Spec'])
				SDA_Occ2 = SDA_Occ.replace('"', '')
				SDA_Target = str(config['SDA_Target'])
			if sDa == 'n':
				sDa_only = sDa_only
				SDA_Occ = '""'
				SDA_Spec = '""'
				SDA_Occ2 = ''
				SDA_Target = ''
			bDa = str(config['bDa'])
			if (bDa == 'y' and sDa == 'n') or (bDa == 'y' and sDa == 'y'):
				bDa_only = bDa_only
				deDup = str(config['deDup'])
				lookUpPeriod = str(config['lookUpPeriod'])
				displayPeriod = str(int(lookUpPeriod)-1)
				totalLookUps = str(config['totalLookUps'])
				occupation = str(config['occupation'])
				specialty = str(config['specialty'])
				BDA_Target = str(config['BDA_Target'])
				occupation2 = occupation.replace('"', '')
				drugList = str(config['drugList'])
				drugsnocomma = str(config['drugList']).replace("\n", ", ")
				# print drugsnocomma
				# drugsnocomma = ", ".join(drugList2)
			if bDa == 'n':
				bDa_only = bDa_only
				deDup = ""
				lookUpPeriod = ""
				totalLookUps = ""
				occupation = '""'
				specialty = '""'
				BDA_Target = ''
				drugList = ""
				occupation2 = ''
				drugsnocomma = ''
			tableName = str(config['tableName'])
			# print listMatchType
			# print sDa_only
			# print bDa_only

else:

	#Take user input for which type of case it is. List Match or Targeting Case
	caseType = raw_input('Which Type of Case Are You Running [listMatch or Targeting]: ')

	#if user chooses List Match then will be prompted the following Variables for input
	if caseType == 'listMatch':
		listProduct = raw_input('DocAlert or Epoc_Quiz: ')
		listMatchType = raw_input('Choose Either Standard, Exact, Standard_Seg, Exact_Seg: ')
		# name = raw_input('Enter List Match Folder Here: ')
		if listMatchType == 'Standard_Seg' or listMatchType == 'Exact_Seg':
			addSeg = raw_input('Enter Segments Here Here (Comma Separate Multiple Segments): ').lower().replace('Group', '_group').replace('group', '_group').replace('GROUP', '_group')
			finalSeg = addSeg.split(', ')
			for seg in finalSeg:
				if seg.startswith(('0','1','2','3','4','5','6','7','8','9')):
					seg = '_' + seg
				segmentList.append(str(seg).replace(' ', '_'))
				splitList = ", ".join(segmentList)
		# code = raw_input('Enter Sas Code you want to run [listMatch or segmentation]: ')
		manu = raw_input('Enter Manufacture Name Here: ')
		brand = raw_input('Enter Drug Name: ')
		yourIn = raw_input('Enter Your Initials Here: ')
		reIn = raw_input('Enter Requesters Initials Here: ')
		sDa = raw_input('Is there a Specialty Add On y/n: ')
		name = "{requester}_{manufact}_{brand}_{dt}_{initials}".format(requester=reIn, manufact=many, brand=brand, dt=dt, initials=yourIn)
		if sDa == 'y':
			SDA_Occ = raw_input('Enter Occupation(s) Here (Wrap in Quotations and use comma to separate): ')
			SDA_Spec = raw_input('Enter Specialty(s) Here (Wrap in Quotations and use comma to separate): ')
			SDA_Occ2 = SDA_Occ.replace('"', '')
		if sDa == 'n':
			SDA_Occ = '""'
			SDA_Spec = '""'
			SDA_Occ2 = ''
			SDA_Target = ''
		bDa = raw_input('Is there a Behavioral y/n: ')
		if bDa == 'y' and sDa == 'n':
			deDup = 'No'
			lookUpPeriod = raw_input('Enter Lookup Period Here: ')
			displayPeriod = str(int(lookUpPeriod)-1)
			totalLookUps = raw_input('Enter # of look ups here: ')
			occupation = raw_input('Enter Occupation(s) Here (Wrap in Single Quotes and use comma to separate): ')
			specialty = raw_input('Enter Specialty(s) Here (Wrap in Quotations and use comma to separate): ')
			occupation2 = occupation.replace('"', '')
			drugList = str(config['drugList'])
			drugList2 = str(config['drugList2'])
			itemDict = {i:drugList2.count(i) for i in drugList2}
			for key, val in itemDict.items():
				if val > 1:
					print''
					print''
					print key, 'is duplicated. Please review drugs and remove duplicates'

			drugList = "\n".join(drugList)
			drugsnocomma = ", ".join(drugList2)
		if bDa == 'y' and sDa == 'y':
			deDup = raw_input('Dedup From SDA (Yes/No): ')
			lookUpPeriod = raw_input('Enter Lookup Period Here: ')
			displayPeriod = str(int(lookUpPeriod)-1)
			totalLookUps = raw_input('Enter # of look ups here: ')
			occupation = raw_input('Enter Occupation(s) Here (Wrap in Single Quotes and use comma to separate): ')
			specialty = raw_input('Enter Specialty(s) Here (Wrap in Quotations and use comma to separate): ')
			occupation2 = occupation.replace('"', '')
			drugList = []
			drugList2 = []
			try:
				while True:
					drugList.append(raw_input('Enter or Paste Drugs Here (Press Ctrl + Z After Entering/Pasting Then hit Enter): '))
			except EOFError:
				pass

			for item in drugList:
				drugList2.append(item)
			itemDict = {i:drugList2.count(i) for i in drugList2}
			for key, val in itemDict.items():
				if val > 1:
					print''
					print''
					print key, 'is duplicated. Please review drugs and remove duplicates'

			drugList = "\n".join(drugList)
			drugsnocomma = ", ".join(drugList2)
		
		if bDa == 'n':
			deDup = ""
			lookUpPeriod = ""
			totalLookUps = ""
			occupation = '""'
			specialty = '""'
			drugList = ""
			occupation2 = ''
			drugsnocomma = ''

		print "EMAIL INPUTS"
		print "-------------------------------------------------------------------"

		caseno = raw_input("Enter Case Number Here: ")
		# manu = raw_input("Enter Manufacturer Here: ")
		mtype = raw_input("1 = ME + NPI + 3-Point Match / 2 = ME + NPI Exact Match: ")
		if mtype == "1":
			mtype = "ME + NPI + 3-Point Match"
		elif mtype == "2":
			mtype = "ME + NPI Exact Match"

		SE = raw_input("Enter Sale Executive Here: ")
		email = raw_input("Enter your email username here: ")
			
	#if user chooses Targeting then will be prompted the following Variables for input
	elif caseType == 'Targeting':
		print 'Targeting Inputs'
		suppApplied = config['suppressionApplied']
		listProduct = raw_input('DocAlert or Epoc_Quiz: ')
		# name = raw_input('Enter Target Client List Match Folder Here: ')
		date = raw_input('Enter TARGETS Folder Date Here: ')

		print 'Client List Inputs'
		print'---------------------------------------------------------------------------------------------------------------'
		print 'Only Choose Standard_Seg or Exact_Seg if you have a target that is meant to be broken out per segment Value'
		listMatchType = raw_input('What type of Target is this (Standard, Standard_Seg, Exact, Exact_Seg: ')
		if listMatchType == 'Standard' or listMatchType == 'Exact':
			targetNum = raw_input('Enter you single Target Number Here: ')
			segVariable = ''
			varValues = ''
		if listMatchType == 'Standard_Seg' or listMatchType == 'Exact_Seg':
			targetNum = raw_input('Enter Multiple Target Numbers Separated by a Space: ')
			segVariable = raw_input('Enter column name containing Segment to be Broken Out by value: ').lower()
			varValues = raw_input('Enter Unique Values of Segment Separated by a space (Remember to remove White Space from Values in Excel): ')


		print 'Client List DataSharing'
		print'-------------------------'
		dSharing = raw_input('Is there DataSharing Y or N: ')
		if dSharing == 'Y':
			addSeg = raw_input('Enter DataSharing Segments Here Here (Comma Separate Multiple Segments: ').lower()
			if segVariable != '':
				finalSeg = addSeg.split(', ')
				finalSeg.append(segVariable)
				for seg in finalSeg:    
					segmentList.append(str(seg).replace(' ', '_'))
					splitList = ", ".join(segmentList)
				
			if segVariable == '':
				finalSeg = addSeg.split(', ')
				for seg in finalSeg:    
					segmentList.append(str(seg).replace(' ', '_'))
					splitList = ", ".join(segmentList)
			keep_seg = raw_input('Do you want to keep the Segment Variable (Yes/No): ')
			# manu = raw_input('Enter Manufacture Name Here: ')

		if dSharing == 'N':
			dSharing = 'N'
			finalSeg = segVariable
			keep_seg = 'No'
			# manu = raw_input('Enter Manufacture Name Here: ')



		print 'USER INFO'
		print '-----------'
		manu = raw_input('Enter Manufacture Name Here: ')
		brand = raw_input('Enter Drug Name: ')
		#outFileFinal2 = """P:\\Epocrates Analytics\\TARGETS\\{date}\\{manu} {brand}\\Targeting Automation Code_OFFICIAL.sas""".format(date = date, manu = manu, brand = brand)
		yourIn = raw_input('Enter Your Initials Here: ')
		name = "T_{manu}_{brand}_{dt}_{initials}".format(manu=manu, brand=brand, dt=dt, initials=yourIn) 

		sDa = raw_input('Is there a Specialty Add On y/n: ')
		if sDa == 'y':
			sDa_only = raw_input('Are you Only Running a SDA or SDA + BDA (Y/N): ')
			SDA_Occ = raw_input('Enter Occupation(s) Here (Wrap in Quotations and use comma to separate): ')
			SDA_Spec = raw_input('Enter Specialty(s) Here (Wrap in Quotations and use comma to separate): ')
			SDA_Target = raw_input('Enter SDA Target Number: ')
		if sDa == 'n':
			sDa_only = 'N'
			SDA_Occ = '""'
			SDA_Spec = '""'
			SDA_Target = ''
			
		bDa = raw_input('Is there a Behavioral y/n: ')
		if bDa == 'y':
			bDa_only = raw_input('Are you only running a BDA (Y/N): ')
			deDup = raw_input('Dedup From SDA (Y/N): ')
			lookUpPeriod = raw_input('Enter Lookup Period Here: ')
			totalLookUps = raw_input('Enter # of look ups here: ')
			occupation = raw_input('Enter Occupation(s) Here (Wrap in Single Quotes and use comma to separate): ')
			specialty = raw_input('Enter Specialty(s) Here (Wrap in Quotations and use comma to separate): ')
			BDA_Target = raw_input('Enter Target Number Here: ')
			drugList = []
			try:
				while True:
					drugList.append(raw_input('Enter or Paste Drugs Here (Press Ctrl + Z After Entering/Pasting Then hit Enter): '))
			except EOFError:
				pass
			drugList = "\n".join(drugList)
		
		if bDa == 'n':
			bDa_only = 'N'
			deDup = 'N'
			lookUpPeriod = ''
			totalLookUps = ''
			occupation = '""'
			specialty = '""'
			BDA_Target = ''
			drugList = ''


	tableName = raw_input('ENTER TABLE NAME: ')


userhome = os.path.expanduser('~')
downloads = userhome + '\\Downloads\\'
newest = max(os.listdir(downloads), key=lambda f: os.path.getmtime("{}/{}".format(downloads, f)))
extension = os.path.splitext(os.path.join(downloads, newest))
justwork = downloads + listText

#MastDrug File
masterDrugs = 'P:\Epocrates Analytics\Drug Compare\Master Drug List\\drugs.csv'

#Set input File and Output File as well as establish default list of source attributes
outFile = """P:\\Epocrates Analytics\\List Match\\List Match Folder\\{folderName}{slashes}""".format(folderName = name, slashes = "\\")
outFileFinal = """P:\\Epocrates Analytics\\List Match\\List Match Folder\\{folderName}\\target.txt""".format(folderName = name)
outCode = """P:\\Epocrates Analytics\\List Match\\List Match Folder\\{folderName}""".format(folderName = name)
if caseType == 'Targeting':
	outFileFinal2 = """P:\\Epocrates Analytics\\TARGETS\\{date}\\{manu} {brand}\\target.txt""".format(date = date, manu = manu, brand = brand)
	outCode3 = """P:\\Epocrates Analytics\\TARGETS\\{date}\\{manu} {brand}""".format(date = date, slashes = "\\", manu = manu, brand = brand)
	rawOutfile = outCode3
	sdaBDAOnly = ""
	targetFolder = manu+" "+brand
	suppFileLocation = "P:\\Epocrates Analytics\\TARGETS\\{date}{slashes}{targetFolder}\\Supp".format(date = date, slashes = "\\", targetFolder=targetFolder)
else:
	suppFileLocation = "P:\\Epocrates Analytics\\List Match\\List Match Folder\\{folderName}\\Supp".format(folderName = name)
segmentListSingle = []

#Set Sas Code Variables
targetAuto = 'P:\\Epocrates Analytics\\Code_Library\\Standard_Codes\\Pre Sales\\DocAlert_Python_Reference\\Targeting Automation Code_OFFICIAL.sas'
basicMatch = 'P:\\Epocrates Analytics\\Code_Library\\Standard_Codes\\Pre Sales\\DocAlert_Python_Reference\\TGT_NPI_ME_3PT_20170515.sas'
dataSharing = 'P:\\Epocrates Analytics\\Code_Library\\Standard_Codes\\Pre Sales\\DocAlert_Python_Reference\\TGT_NPI_ME_3PT_DataSharing_20170515.sas'
autoCode = 'P:\\Epocrates Analytics\\Code_Library\\Standard_Codes\\Pre Sales\\DocAlert_Python_Reference\\Presales Automation.sas'
emailCode = 'P:\\Epocrates Analytics\\Code_Library\\Standard_Codes\\Pre Sales\\DocAlert_Python_Reference\\Presales Automation_Email_Final.sas'
supressionCode = "P:\\Epocrates Analytics\\Code_Library\\Standard_Codes\\Pre Sales\\DocAlert_Python_Reference\\Code Housing\\for_email\\Supp_Auto_DONOTEDIT\\Presales Automation_Email_Final_Suppressed.sas"

def createFolders():
	if caseType == 'listMatch':
		if os.path.exists(os.path.join("P:\\Epocrates Analytics\\List Match\\List Match Folder", name)):
			print colored('List Match Folder "'+name+'" Already Exists . . . Skipping Creation', 'yellow')
			print '------------------------------------------------------------------------------------'
			print''
		if not os.path.exists(os.path.join("P:\\Epocrates Analytics\\List Match\\List Match Folder", name)):
			os.chdir("P:\\Epocrates Analytics\\List Match\\List Match Folder\\")
			os.mkdir(name)

	if caseType == 'Targeting':
		targetFolder = manu+" "+brand
		if os.path.exists("P:\\Epocrates Analytics\\TARGETS\\{date}{slashes}{targetFolder}".format(date = date, slashes = "\\", targetFolder=targetFolder)):
			print colored('TARGETS Folder '+ manu + ' ' + brand + ' Already Exists . . . Skipping Creation', 'yellow')
		if not os.path.exists("P:\\Epocrates Analytics\\TARGETS\\{date}{slashes}{targetFolder}".format(date = date, slashes = "\\", targetFolder=targetFolder)):
			os.chdir("""P:\\Epocrates Analytics\\TARGETS\\{date}{slashes}""".format(date = date, slashes = "\\"))
			os.mkdir(targetFolder)

def checkDrugs():
	#read master drug file, strip white space and load into a list for comparison
	if caseType == 'listMatch' or caseType == 'Targeting':
		if bDa == 'y':
			with open(masterDrugs, 'rb') as myDrugs:
				reader = csv.DictReader(myDrugs)
				for row in reader:
					finalDrugs.append(row['drugs'].strip())
					
			for inputDrugs in drugList.replace(', ', '\n').split("\n"):
				if inputDrugs.strip() not in finalDrugs:
					unmatchedDrugs.append(inputDrugs)
					
			print colored('THESE DRUGS ARE SPELLED WRONG OR MISSING: ', 'yellow'), colored(unmatchedDrugs, 'yellow')
			print '----------------------------------------------------------------------------------------'
			for drug in unmatchedDrugs:
				print 'Possible Correct Spelling: ', colored(process.extract(drug, finalDrugs, limit=2), 'green'), colored(' - ', 'red'), colored(drug, 'red')
				print '----------------------------------------------------------------------------------------'
		if bDa == 'n':
			pass
	else:
		pass


def getMain():
	with open (downloads + 'csvFile.csv', 'rb') as inFile2, open(downloads + csvFileModTemp2, 'wb') as targetFile:
		r = csv.reader(inFile2)
		headers = r.next()
		foundMain = []
		mainCols = ['npi', 'me', 'fname', 'lname', 'zip']
		for index, col in enumerate(headers):
			cellVal = str(col).lower().replace('/', '_').replace('-', '_')
			#Regular Expression Rules. We can add new Rules as we need to this list to cover more common cases.
			if cellVal == 'npi' or cellVal == 'npi_id' or re.search('.+ npi .+', cellVal) or re.search('.+_npi_.+', cellVal) or re.search('^npi.+', cellVal) or re.search('.+ npi', cellVal) or re.search('.+ npi', cellVal):
				print cellVal, ': ',  colored('I found a NPI Number', 'green')
				headers[index] = 'npi'
				foundMain.append('npi')
			elif cellVal == 'me' or cellVal == 'me_id' or cellVal == 'me_' or cellVal == 'meded' or cellVal == 'menum' or re.search('.+ me .+', cellVal) or  re.search('.+_me_.+', cellVal) or re.search('^me .+', cellVal) or re.search('.+ me', cellVal) or re.search('.+ me', cellVal) or re.search('^me_.+', cellVal):
				print cellVal, ': ', colored('I found a ME Number', 'green')
				headers[index] = 'me'
				foundMain.append('me')
			elif cellVal == 'fname' or cellVal == 'firstname' or re.search('^first.+name', cellVal) or re.search('.+first.+name', cellVal) or re.search('.+fname', cellVal) or re.search('.+first', cellVal) or re.search('.+frst.+', cellVal):
				print cellVal, ': ', colored('I found a First Name', 'green')
				headers[index] = 'fname'
				foundMain.append('fname')
			elif re.search(r'^lname|^last.+name|.+last|.+last.+name', cellVal) or cellVal == 'lastname':
				print cellVal, ': ',  colored('I found a Last Name', 'green')
				headers[index] = 'lname'
				foundMain.append('lname')
			elif cellVal == 'full_name' or cellVal == 'fullname' or cellVal == 'prescriber_name' or re.search('^full.+name', cellVal) or re.search('.+full.+name', cellVal):
				print cellVal, ': ', colored('I found a Full Name', 'green')
				headers[index] = 'full_name'
				fullNameList.append('y')
			elif cellVal == 'zip_4' or cellVal == 'zip4' or cellVal == 'zip___4':
				headers[index] = 'whatever'
			elif cellVal == 'Group' or cellVal == 'group':
				headers[index] = '_group'
			elif cellVal == 'zip' or cellVal == 'Postal' or (re.search('^zip.+', cellVal) and (cellVal != 'zip_4' or cellVal != 'zip4')) or re.search('^postal.+', cellVal) or re.search('.+_zip', cellVal) or re.search('.+ zip', cellVal) or re.search('.+_postal', cellVal) or re.search('.+ zip', cellVal) or re.search('.+ postal', cellVal):
				print cellVal, ': ',  colored('I found a Zip/Postal Code', 'green')
				headers[index] = 'zip'
				foundMain.append('zip')
			elif cellVal.startswith(('0','1','2','3','4','5','6','7','8','9')):
				print 'I added an underscore to: ', cellVal
				headers[index] = '_' + headers[index]
		w = csv.writer(targetFile, lineterminator='\n')
		w.writerow(headers)
		for row in r:
			w.writerow(row)

		for col in mainCols:
			if col not in foundMain:
				print 'Did not find Main Column: ', colored(col, 'red')

		print '----------------------------------------------------------------------------------------------'
		print ''

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

		w = csv.writer(myOut, lineterminator='\n')
		w.writerow(visited)
		for row in reader:
			w.writerow(row)

def cmiCompasColumns():
	with open (downloads + csvFileMod, 'rb') as inFile2, open(downloads + csvFileModTemp, 'wb') as targetFile:
		r = csv.reader(inFile2)
		headers = r.next()
		for index, col in enumerate(headers):
			cellVal = str(col).lower().replace('/', '_').replace('-', '_')
			if cellVal == 'state':
				print cellVal, colored('I Found a State_Code', 'green')
				headers[index] = 'state_code'
			elif cellVal == 'address_1' or cellVal == 'addr1':
				print cellVal, colored('I Found a Address_1', 'green')
				headers[index] = 'address1'
			elif cellVal == 'client_id' or cellVal == 'client_id_1':
				print cellVal, colored('I Found an clientid', 'green')
				headers[index] = 'clientid'
			elif cellVal == 'compasid' or cellVal == 'compas_id' or cellVal == 'compas id':
				print cellVal, colored('I Found a CompasID', 'green')
				headers[index] = 'compasid'
			elif cellVal == 'specialty' or re.search('^specialty.+', cellVal) or re.search('.+specialty.+', cellVal):
				print cellVal, colored('I Found a Specialty', 'green')
				headers[index] = 'specialty'
		w = csv.writer(targetFile, lineterminator='\n')
		w.writerow(headers)
		for row in r:
			w.writerow(row)
	os.remove(os.path.join(downloads, csvFileMod))
	os.rename(os.path.join(downloads, csvFileModTemp), os.path.join(downloads, csvFileMod))


def postgresConn():

	# os.system("start cmd /C {}".format(os.path.join(desktop, 'EWOK\\pgsql\\pgLocal.bat')))
	# process = Popen(os.path.join(desktop, 'EWOK\\pgsql\\pgLocal.bat'), stdin=PIPE, stderr=PIPE, stdout=PIPE, shell=True)
	# time.sleep(3)


	# conn_string = "host='localhost' dbname='postgres' port='5432' user='postgres' password='{password}'".format(password=password)
	# conn = psycopg2.connect(conn_string)
	# cursor = conn.cursor()

	df = pandas.read_csv('{downloads}{csvFile}'.format(downloads=downloads, csvFile=csvFileMod))
	myHeaders = list(df)
	neededColumns = ['me', 'npi', 'fname', 'lname', 'zip']
	selectMain5 = "select REPLACE(npi, '.0', '') as npi, REPLACE(me, '.0', '') as me, fname, lname, REPLACE(zip, '.0', '') as zip"
	selectMain2 = "select REPLACE(npi, '.0', '') as npi, REPLACE(me, '.0', '') as me"
	selectFullName = """SELECT [REPLACE](npi, '.0', '') AS npi,
	   [REPLACE](me, '.0', '') AS me,
	   [replace](full_name, ltrim(full_name, [replace](full_name, ' ', '') ), '') AS fname,
	   [replace](full_name, rtrim(full_name, [replace](full_name, ' ', '') ), '') AS lname,
	   [REPLACE](zip, '.0', '') AS zip"""

	selectFuzzy = "Select fname, lname, [REPLACE](zip, '.0', '') AS zip"

	conn = sqlite3.connect(os.path.join(desktop,'Ewok', 'epocrates_tables.db'))
	conn.text_factory = str
	cur = conn.cursor()

	print colored("Connected to SQL LITE!\n", 'green')
	print colored('Checking if Table "{}" Exists. If so, were gonna drop dat table like 5th Period Chemistry. . .\n'.format(tableName), 'yellow')

	dropTable = """DROP TABLE IF EXISTS {}""".format(tableName)

	cur.execute(dropTable)
	conn.commit()


	df.to_sql('{}'.format(tableName), conn, if_exists='replace', index=False)

	for col in neededColumns:
		if col not in myHeaders:
			try:
				cur.executescript('ALTER TABLE {} ADD COLUMN {} text;'.format(tableName, col))
			except:
				print col, 'Already Exists . . . Skipping'
				continue

	conn.commit()

	if caseType == 'listMatch' and listMatchType == 'Standard_Seg' and foundFullName == 'n':

		export = """{select}, {seg} from {tableName};""".format(select=selectMain5, tableName=tableName, seg=splitList)
		pandas.read_sql_query(export, conn).to_csv(os.path.join(downloads, 'target.txt'), index=False, sep='\t')

	elif caseType == 'listMatch' and listMatchType == 'Standard_Seg' and foundFullName == 'y':

		export = """{select}, {seg} from {tableName};""".format(select=selectFullName, tableName=tableName, seg=splitList)
		pandas.read_sql_query(export, conn).to_csv(os.path.join(downloads, 'target.txt'), index=False, sep='\t')

	elif caseType == 'listMatch' and listMatchType =='Standard' and foundFullName == 'n':

		export = """{} from {};""".format(selectMain5, tableName)
		pandas.read_sql_query(export, conn).to_csv(os.path.join(downloads, 'target.txt'), index=False, sep='\t')

	elif caseType == 'listMatch' and listMatchType =='Standard' and foundFullName == 'y':

		export = """{select} from {tableName};""".format(select=selectFullName, tableName=tableName)
		pandas.read_sql_query(export, conn).to_csv(os.path.join(downloads, 'target.txt'), index=False, sep='\t')

	elif caseType == 'listMatch' and listMatchType == 'Exact_Seg':

		export = """{select}, {seg} from {tableName};""".format(select=selectMain2, tableName=tableName, seg=splitList)
		pandas.read_sql_query(export, conn).to_csv(os.path.join(downloads, 'target.txt'), index=False, sep='\t')

	elif caseType == 'listMatch' and listMatchType =='Exact':

		export = """{select} from {tableName};""".format(select=selectMain2, tableName=tableName)
		pandas.read_sql_query(export, conn).to_csv(os.path.join(downloads, 'target.txt'), index=False, sep='\t')

	elif caseType == 'listMatch' and listMatchType =='Fuzzy':

		export = """{} from {};""".format(selectFuzzy, tableName)
		pandas.read_sql_query(export, conn).to_csv(os.path.join(downloads, 'target.txt'), index=False, sep='\t')

	elif caseType == 'Targeting' and listMatchType =='Standard' and dSharing == 'Y' and foundFullName == 'n' and config['cmi_compass_client'] == 'N':

		export = """{select}, {seg} from {tableName};""".format(select=selectMain5, tableName=tableName, seg=splitList)
		pandas.read_sql_query(export, conn).to_csv(os.path.join(downloads, 'target.txt'), index=False, sep='\t')

	elif caseType == 'Targeting' and listMatchType =='Standard' and dSharing == 'Y' and foundFullName == 'y' and config['cmi_compass_client'] == 'N':

		export = """{select}, {seg} from {tableName};""".format(select=selectFullName, tableName=tableName, seg=splitList)
		pandas.read_sql_query(export, conn).to_csv(os.path.join(downloads, 'target.txt'), index=False, sep='\t') 

	elif caseType == 'Targeting' and listMatchType =='Exact' and dSharing == 'Y':

		export = """{select}, {seg} from {tableName};""".format(select=selectMain2, tableName=tableName, seg=splitList)
		pandas.read_sql_query(export, conn).to_csv(os.path.join(downloads, 'target.txt'), index=False, sep='\t')

	elif caseType == 'Targeting' and listMatchType =='Standard' and dSharing == 'N' and foundFullName == 'n':

		export = """{} from {};""".format(selectMain5, tableName)
		pandas.read_sql_query(export, conn).to_csv(os.path.join(downloads, 'target.txt'), index=False, sep='\t')

	elif caseType == 'Targeting' and listMatchType =='Standard' and dSharing == 'N' and foundFullName == 'n' and config['segmentListChecked'] == 'y':

		export = """{select} from {tableName};""".format(select=selectFullName, tableName=tableName)
		pandas.read_sql_query(export, conn).to_csv(os.path.join(downloads, 'target.txt'), index=False, sep='\t') 

	elif caseType == 'Targeting' and listMatchType =='Standard' and dSharing == 'N' and foundFullName == 'y':

		export = """{select}, {seg} from {tableName};""".format(select=selectFullName, tableName=tableName, seg=splitList)
		pandas.read_sql_query(export, conn).to_csv(os.path.join(downloads, 'target.txt'), index=False, sep='\t')

	elif caseType == 'Targeting' and listMatchType =='Exact' and dSharing == 'N' and config['segmentListChecked'] == 'n':

		export = """{select} from {tableName};""".format(select=selectMain2, tableName=tableName)
		pandas.read_sql_query(export, conn).to_csv(os.path.join(downloads, 'target.txt'), index=False, sep='\t')

	elif caseType == 'Targeting' and listMatchType =='Exact' and dSharing == 'N' and config['segmentListChecked'] == 'y':

		export = """{select}, {seg} from {tableName};""".format(select=selectMain2, tableName=tableName, seg=splitList)
		pandas.read_sql_query(export, conn).to_csv(os.path.join(downloads, 'target.txt'), index=False, sep='\t')		
		
	elif caseType == 'Targeting' and listMatchType =='Standard_Seg' and dSharing == 'Y' and foundFullName == 'n' and config['cmi_compass_client'] == 'N':

		export = """{select}, {seg} from {tableName};""".format(select=selectMain5, tableName=tableName, seg=splitList)
		pandas.read_sql_query(export, conn).to_csv(os.path.join(downloads, 'target.txt'), index=False, sep='\t')

	elif caseType == 'Targeting' and listMatchType =='Standard_Seg' and dSharing == 'Y' and foundFullName == 'y':

		export = """{select}, {seg} from {tableName};""".format(select=selectFullName, tableName=tableName, seg=splitList)
		pandas.read_sql_query(export, conn).to_csv(os.path.join(downloads, 'target.txt'), index=False, sep='\t')

	elif caseType == 'Targeting' and listMatchType =='Exact_Seg' and dSharing == 'Y':

		export = """{select}, {seg} from {tableName};""".format(select=selectMain2, tableName=tableName, seg=splitList)
		pandas.read_sql_query(export, conn).to_csv(os.path.join(downloads, 'target.txt'), index=False, sep='\t')

	elif caseType == 'Targeting' and listMatchType =='Standard_Seg' and dSharing == 'N' and foundFullName == 'n':

		export = """{select}, {seg} from {tableName};""".format(select=selectMain5, tableName=tableName, seg=splitList)
		pandas.read_sql_query(export, conn).to_csv(os.path.join(downloads, 'target.txt'), index=False, sep='\t')

	elif caseType == 'Targeting' and listMatchType =='Standard_Seg' and dSharing == 'N' and foundFullName == 'y':

		export = """{select}, {seg} from {tableName};""".format(select=selectFullName, tableName=tableName, seg=splitList)
		pandas.read_sql_query(export, conn).to_csv(os.path.join(downloads, 'target.txt'), index=False, sep='\t')

	elif caseType == 'Targeting' and listMatchType =='Exact_Seg' and dSharing == 'N':

		export = """{select}, {seg} from {tableName};""".format(select=selectMain2, tableName=tableName, seg=splitList)
		pandas.read_sql_query(export, conn).to_csv(os.path.join(downloads, 'target.txt'), index=False, sep='\t')

	elif caseType == 'Targeting' and config['cmi_compass_client'] == 'Y':
		for col in cmiList:
			if col not in myHeaders:
				try:
					cur.executescript('ALTER TABLE {} ADD COLUMN {} text;'.format(tableName, col))
				except:
					print col, 'Already Exists . . . Skipping'
					continue

		conn.commit()

		if listMatchType =='Standard' and config['cmi_compass_client'] == 'Y':
			export = """{select}, {seg} from {tableName};""".format(select=selectMain5, seg=cmiCompasSQL, tableName=tableName)
			pandas.read_sql_query(export, conn).to_csv(os.path.join(downloads, 'target.txt'), index=False, sep='\t')


		elif listMatchType =='Standard_Seg' and config['cmi_compass_client'] == 'Y':

			export = """{select}, {seg}, {segVar} from {tableName};""".format(select=selectMain5, seg=cmiCompasSQL, segVar=segVariable, tableName=tableName)
			pandas.read_sql_query(export, conn).to_csv(os.path.join(downloads, 'target.txt'), index=False, sep='\t')

	print ''
	print 'SQL Export Complete\n'

	print random.choice(randomList)+'\n'

def csv_from_excel():
	wb = xlrd.open_workbook(downloads + newest)
	sh = wb.sheet_by_index(0)
	your_csv_file = open (downloads + 'target.csv', 'wb')
	wr = unicodecsv.writer(your_csv_file, encoding='utf8', lineterminator='\n')

	for rownum in range(sh.nrows):
		wr.writerow(sh.row_values(rownum))

	your_csv_file.close()

def removeChar():
	inputFile = open(downloads + csvFile, 'r')
	outputFile = open(downloads + 'csvFile.csv', 'wb')
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


def get_Cols():
	with open(downloads + csvFileMod, 'rb') as f:
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

def copyTarget():
	if caseType == 'listMatch':
		copyfile(downloads + listText, outFileFinal)

	elif caseType == 'Targeting':
		copyfile(downloads + listText, outFileFinal2)

def removeFiles():
	os.remove(os.path.join(downloads, 'target.txt'))
	os.remove(os.path.join(downloads, 'target_mod.csv'))
	os.remove(os.path.join(downloads, 'target.csv'))
	# os.remove(os.path.join(downloads, 'csvFile.csv'))
	os.remove(os.path.join(downloads, 'csvFile1.csv'))
	os.remove(os.path.join(downloads, csvFileModTemp2))
	if caseType == 'Targeting' and sDa_only == 'N' and bDa_only == 'N':
		newest = str(config['loadedFile'])
		copyfile(os.path.join(downloads, newest), os.path.join(outCode3, newest))
	else:
		newest = str(config['loadedFile'])
		copyfile(os.path.join(downloads, newest), os.path.join(outCode, newest))

def fixSas():
	#fixes the listMatch Code
	#Straight List Match Only
	if suppApplied == 'N':
		if caseType == 'listMatch' and (listMatchType == 'Standard' or listMatchType == 'Exact' or listMatchType == 'Fuzzy'):
			copyfile(emailCode, os.path.join(outCode, 'Presales Automation_Email_Final.sas'))
			newInput = os.path.join(outCode, 'Presales Automation_Email_Final.sas')

			line_file = open(os.path.join(newInput),'r').readlines()
			new_file = open(os.path.join(newInput),'w')
			for line_in in line_file:
				#target_out = line_in.replace('/*DocORQuiz*/', listProduct).replace('/*listMatchType*/', listMatchType).replace('/*target_text_file*/', outCode).replace('/*Segments*/', '').replace('/*Segments2*/', '').replace('/*Brand*/', brand).replace('/*MY_INIT*/', yourIn).replace('/*Requester_Initials*/', reIn).replace('/*SDA_Occ*/', '').replace('/*SDA_Spec*/', '').replace('/*yesORno*/', '').replace('/*LookUpPeriod*/', '').replace('/*totalLoookUps*/', '').replace('/*BDA_Occ*/', '').replace('/*BDA_Spec*/', '').replace('/*drugList*/', '')
				target_out = line_in.replace('/*DocORQuiz*/', listProduct)
				target_out = target_out.replace('/*listMatchType*/', listMatchType)
				target_out = target_out.replace('/*target_text_file*/', outCode)
				target_out = target_out.replace('/*Segments*/', '')
				target_out = target_out.replace('/*Segments2*/', '')
				target_out = target_out.replace('/*Brand*/', brand)
				target_out = target_out.replace('/*MY_INIT*/', yourIn)
				target_out = target_out.replace('/*Requester_Initials*/', reIn)
				target_out = target_out.replace('/*SDA_Occ*/', SDA_Occ)
				target_out = target_out.replace('/*SDA_Spec*/', SDA_Spec)
				target_out = target_out.replace('/*therapyClass*/', therapyClass)
				target_out = target_out.replace('/*yesORno*/', deDup)
				target_out = target_out.replace('/*LookUpPeriod*/', lookUpPeriod)
				target_out = target_out.replace('/*totalLoookUps*/', totalLookUps)
				target_out = target_out.replace('/*BDA_Occ*/', occupation)
				target_out = target_out.replace('/*BDA_Spec*/', specialty)
				target_out = target_out.replace('/*drugList*/', ';')
				target_out = target_out.replace('/*caseno*/', caseno)
				target_out = target_out.replace('/*manu*/', manu)
				target_out = target_out.replace('/*mtype*/', mtype)
				target_out = target_out.replace('/*SE*/', SE)
				target_out = target_out.replace('/*username*/', email)
				target_out = target_out.replace('/*sdaocc2*/', SDA_Occ2)
				target_out = target_out.replace('/*bdaocc2*/', occupation2)
				target_out = target_out.replace('/*drugsnocomma*/', drugsnocomma)
				target_out = target_out.replace('/*dispPeriod*/', '')
				target_out = target_out.replace('/*suppApplied*/', 'No')
				target_out = target_out.replace('/*supp_text_file*/', '')
				target_out = target_out.replace('/*supp_Match_Type*/', '""')
				target_out = target_out.replace('/*bda_only*/', bDa_only)
				target_out = target_out.replace('/*sda_only*/', sDa_only)
				target_out = target_out.replace('/*createPivotTable*/', createPivotTable)
				target_out = target_out.replace('/*pivYes1*/', pivSeg1)
				target_out = target_out.replace('/*pivYes2*/', pivSeg2)
				target_out = target_out.replace('/*totalSDAS*/', finalSDATotal)
				target_out = target_out.replace('/*totalBDAS*/', finalBDATotal)
				new_file.write(target_out)
				line_file = new_file
				
			if bDa == 'y' and sDa == 'y':
				copyfile(emailCode, os.path.join(outCode, 'Presales Automation_Email_Final.sas'))
				newInput = os.path.join(outCode, 'Presales Automation_Email_Final.sas')

				line_file = open(os.path.join(newInput),'r').readlines()
				new_file = open(os.path.join(newInput),'w')
				for line_in in line_file:
					#target_out = line_in.replace('/*DocORQuiz*/', listProduct).replace('/*listMatchType*/', listMatchType).replace('/*target_text_file*/', outCode).replace('/*Segments*/', '').replace('/*Segments2*/', '').replace('/*Brand*/', brand).replace('/*MY_INIT*/', yourIn).replace('/*Requester_Initials*/', reIn).replace('/*SDA_Occ*/', SDA_Occ).replace('/*SDA_Spec*/', SDA_Spec).replace('/*yesORno*/', deDup).replace('/*LookUpPeriod*/', lookUpPeriod).replace('/*totalLoookUps*/', totalLookUps).replace('/*BDA_Occ*/', occupation).replace('/*BDA_Spec*/', specialty).replace('/*drugList*/', drugList)
					target_out = line_in.replace('/*DocORQuiz*/', listProduct)
					target_out = target_out.replace('/*listMatchType*/', listMatchType)
					target_out = target_out.replace('/*target_text_file*/', outCode)
					target_out = target_out.replace('/*Segments*/', '')
					target_out = target_out.replace('/*Segments2*/', '')
					target_out = target_out.replace('/*Brand*/', brand)
					target_out = target_out.replace('/*MY_INIT*/', yourIn)
					target_out = target_out.replace('/*Requester_Initials*/', reIn)
					target_out = target_out.replace('/*SDA_Occ*/', SDA_Occ)
					target_out = target_out.replace('/*SDA_Spec*/', SDA_Spec)
					target_out = target_out.replace('/*therapyClass*/', therapyClass)
					target_out = target_out.replace('/*yesORno*/', deDup)
					target_out = target_out.replace('/*LookUpPeriod*/', lookUpPeriod)
					target_out = target_out.replace('/*totalLoookUps*/', totalLookUps)
					target_out = target_out.replace('/*BDA_Occ*/', occupation)
					target_out = target_out.replace('/*BDA_Spec*/', specialty)
					target_out = target_out.replace('/*drugList*/', drugList)
					target_out = target_out.replace('/*caseno*/', caseno)
					target_out = target_out.replace('/*manu*/', manu)
					target_out = target_out.replace('/*mtype*/', mtype)
					target_out = target_out.replace('/*SE*/', SE)
					target_out = target_out.replace('/*username*/', email)
					target_out = target_out.replace('/*sdaocc2*/', SDA_Occ2)
					target_out = target_out.replace('/*bdaocc2*/', occupation2)
					target_out = target_out.replace('/*drugsnocomma*/', drugsnocomma)
					target_out = target_out.replace('/*dispPeriod*/', displayPeriod)
					target_out = target_out.replace('/*suppApplied*/', 'No')
					target_out = target_out.replace('/*supp_text_file*/', '')
					target_out = target_out.replace('/*supp_Match_Type*/', '""')
					target_out = target_out.replace('/*bda_only*/', bDa_only)
					target_out = target_out.replace('/*sda_only*/', sDa_only)
					target_out = target_out.replace('/*createPivotTable*/', createPivotTable)
					target_out = target_out.replace('/*pivYes1*/', pivSeg1)
					target_out = target_out.replace('/*pivYes2*/', pivSeg2)
					target_out = target_out.replace('/*totalSDAS*/', finalSDATotal)
					target_out = target_out.replace('/*totalBDAS*/', finalBDATotal)
					new_file.write(target_out)
					line_file = new_file
				
			else:
				if bDa == 'y':  
					copyfile(emailCode, os.path.join(outCode, 'Presales Automation_Email_Final.sas'))
					newInput = os.path.join(outCode, 'Presales Automation_Email_Final.sas')

					line_file = open(os.path.join(newInput),'r').readlines()
					new_file = open(os.path.join(newInput),'w')
					for line_in in line_file:
						#target_out = line_in.replace('/*DocORQuiz*/', listProduct).replace('/*listMatchType*/', listMatchType).replace('/*target_text_file*/', outCode).replace('/*Segments*/', '').replace('/*Segments2*/', '').replace('/*Brand*/', brand).replace('/*MY_INIT*/', yourIn).replace('/*Requester_Initials*/', reIn).replace('/*SDA_Occ*/', '').replace('/*SDA_Spec*/', '').replace('/*yesORno*/', deDup).replace('/*LookUpPeriod*/', lookUpPeriod).replace('/*totalLoookUps*/', totalLookUps).replace('/*BDA_Occ*/', occupation).replace('/*BDA_Spec*/', specialty).replace('/*drugList*/', drugList)
						target_out = line_in.replace('/*DocORQuiz*/', listProduct)
						target_out = target_out.replace('/*listMatchType*/', listMatchType)
						target_out = target_out.replace('/*target_text_file*/', outCode)
						target_out = target_out.replace('/*Segments*/', '')
						target_out = target_out.replace('/*Segments2*/', '')
						target_out = target_out.replace('/*Brand*/', brand)
						target_out = target_out.replace('/*MY_INIT*/', yourIn)
						target_out = target_out.replace('/*Requester_Initials*/', reIn)
						target_out = target_out.replace('/*SDA_Occ*/', SDA_Occ)
						target_out = target_out.replace('/*SDA_Spec*/', SDA_Spec)
						target_out = target_out.replace('/*therapyClass*/', therapyClass)
						target_out = target_out.replace('/*yesORno*/', deDup)
						target_out = target_out.replace('/*LookUpPeriod*/', lookUpPeriod)
						target_out = target_out.replace('/*totalLoookUps*/', totalLookUps)
						target_out = target_out.replace('/*BDA_Occ*/', occupation)
						target_out = target_out.replace('/*BDA_Spec*/', specialty)
						target_out = target_out.replace('/*drugList*/', drugList)                   
						target_out = target_out.replace('/*caseno*/', caseno)
						target_out = target_out.replace('/*manu*/', manu)
						target_out = target_out.replace('/*mtype*/', mtype)
						target_out = target_out.replace('/*SE*/', SE)
						target_out = target_out.replace('/*username*/', email)
						target_out = target_out.replace('/*sdaocc2*/', SDA_Occ2)
						target_out = target_out.replace('/*bdaocc2*/', occupation2)
						target_out = target_out.replace('/*drugsnocomma*/', drugsnocomma)
						target_out = target_out.replace('/*dispPeriod*/', displayPeriod)
						target_out = target_out.replace('/*suppApplied*/', 'No')
						target_out = target_out.replace('/*supp_text_file*/', '')
						target_out = target_out.replace('/*supp_Match_Type*/', '""')
						target_out = target_out.replace('/*bda_only*/', bDa_only)
						target_out = target_out.replace('/*sda_only*/', sDa_only)
						target_out = target_out.replace('/*createPivotTable*/', createPivotTable)
						target_out = target_out.replace('/*pivYes1*/', pivSeg1)
						target_out = target_out.replace('/*pivYes2*/', pivSeg2)
						target_out = target_out.replace('/*totalSDAS*/', finalSDATotal)
						target_out = target_out.replace('/*totalBDAS*/', finalBDATotal)
						new_file.write(target_out)
						line_file = new_file
						
				if sDa == 'y':  
					copyfile(emailCode, os.path.join(outCode, 'Presales Automation_Email_Final.sas'))
					newInput = os.path.join(outCode, 'Presales Automation_Email_Final.sas')

					line_file = open(os.path.join(newInput),'r').readlines()
					new_file = open(os.path.join(newInput),'w')
					for line_in in line_file:
						#target_out = line_in.replace('/*DocORQuiz*/', listProduct).replace('/*listMatchType*/', listMatchType).replace('/*target_text_file*/', outCode).replace('/*Segments*/', '').replace('/*Segments2*/', '').replace('/*Brand*/', brand).replace('/*MY_INIT*/', yourIn).replace('/*Requester_Initials*/', reIn).replace('/*SDA_Occ*/', SDA_Occ).replace('/*SDA_Spec*/', SDA_Spec).replace('/*yesORno*/', '').replace('/*LookUpPeriod*/', '').replace('/*totalLoookUps*/', '').replace('/*BDA_Occ*/', '').replace('/*BDA_Spec*/', '').replace('/*drugList*/', '')
						target_out = line_in.replace('/*DocORQuiz*/', listProduct)
						target_out = target_out.replace('/*listMatchType*/', listMatchType)
						target_out = target_out.replace('/*target_text_file*/', outCode)
						target_out = target_out.replace('/*Segments*/', '')
						target_out = target_out.replace('/*Segments2*/', '')
						target_out = target_out.replace('/*Brand*/', brand)
						target_out = target_out.replace('/*MY_INIT*/', yourIn)
						target_out = target_out.replace('/*Requester_Initials*/', reIn)
						target_out = target_out.replace('/*SDA_Occ*/', SDA_Occ)
						target_out = target_out.replace('/*SDA_Spec*/', SDA_Spec)
						target_out = target_out.replace('/*therapyClass*/', therapyClass)
						target_out = target_out.replace('/*yesORno*/', deDup)
						target_out = target_out.replace('/*LookUpPeriod*/', lookUpPeriod)
						target_out = target_out.replace('/*totalLoookUps*/', totalLookUps)
						target_out = target_out.replace('/*BDA_Occ*/', occupation)
						target_out = target_out.replace('/*BDA_Spec*/', specialty)
						target_out = target_out.replace('/*drugList*/', ';')
						target_out = target_out.replace('/*caseno*/', caseno)
						target_out = target_out.replace('/*manu*/', manu)
						target_out = target_out.replace('/*mtype*/', mtype)
						target_out = target_out.replace('/*SE*/', SE)
						target_out = target_out.replace('/*username*/', email)
						target_out = target_out.replace('/*sdaocc2*/', SDA_Occ2)
						target_out = target_out.replace('/*bdaocc2*/', occupation2)
						target_out = target_out.replace('/*drugsnocomma*/', drugsnocomma)
						target_out = target_out.replace('/*dispPeriod*/', '')
						target_out = target_out.replace('/*suppApplied*/', 'No')
						target_out = target_out.replace('/*supp_text_file*/', '')
						target_out = target_out.replace('/*supp_Match_Type*/', '""')
						target_out = target_out.replace('/*bda_only*/', bDa_only)
						target_out = target_out.replace('/*sda_only*/', sDa_only)
						target_out = target_out.replace('/*createPivotTable*/', createPivotTable)
						target_out = target_out.replace('/*pivYes1*/', pivSeg1)
						target_out = target_out.replace('/*pivYes2*/', pivSeg2)
						target_out = target_out.replace('/*totalSDAS*/', finalSDATotal)
						target_out = target_out.replace('/*totalBDAS*/', finalBDATotal)
						new_file.write(target_out)
						line_file = new_file

		elif caseType == 'listMatch' and listMatchType == 'None':
			copyfile(emailCode, os.path.join(outCode, 'Presales Automation_Email_Final.sas'))
			newInput = os.path.join(outCode, 'Presales Automation_Email_Final.sas')

			line_file = open(os.path.join(newInput),'r').readlines()
			new_file = open(os.path.join(newInput),'w')
			for line_in in line_file:
				#target_out = line_in.replace('/*DocORQuiz*/', listProduct).replace('/*listMatchType*/', listMatchType).replace('/*target_text_file*/', outCode).replace('/*Segments*/', '').replace('/*Segments2*/', '').replace('/*Brand*/', brand).replace('/*MY_INIT*/', yourIn).replace('/*Requester_Initials*/', reIn).replace('/*SDA_Occ*/', '').replace('/*SDA_Spec*/', '').replace('/*yesORno*/', '').replace('/*LookUpPeriod*/', '').replace('/*totalLoookUps*/', '').replace('/*BDA_Occ*/', '').replace('/*BDA_Spec*/', '').replace('/*drugList*/', '')
				target_out = line_in.replace('/*DocORQuiz*/', listProduct)
				target_out = target_out.replace('/*listMatchType*/', listMatchType)
				target_out = target_out.replace('/*target_text_file*/', outCode)
				target_out = target_out.replace('/*Segments*/', '')
				target_out = target_out.replace('/*Segments2*/', '')
				target_out = target_out.replace('/*Brand*/', brand)
				target_out = target_out.replace('/*MY_INIT*/', yourIn)
				target_out = target_out.replace('/*Requester_Initials*/', reIn)
				target_out = target_out.replace('/*SDA_Occ*/', SDA_Occ)
				target_out = target_out.replace('/*SDA_Spec*/', SDA_Spec)
				target_out = target_out.replace('/*therapyClass*/', therapyClass)
				target_out = target_out.replace('/*yesORno*/', deDup)
				target_out = target_out.replace('/*LookUpPeriod*/', lookUpPeriod)
				target_out = target_out.replace('/*totalLoookUps*/', totalLookUps)
				target_out = target_out.replace('/*BDA_Occ*/', occupation)
				target_out = target_out.replace('/*BDA_Spec*/', specialty)
				target_out = target_out.replace('/*drugList*/', ';')
				target_out = target_out.replace('/*caseno*/', caseno)
				target_out = target_out.replace('/*manu*/', manu)
				target_out = target_out.replace('/*mtype*/', mtype)
				target_out = target_out.replace('/*SE*/', SE)
				target_out = target_out.replace('/*username*/', email)
				target_out = target_out.replace('/*sdaocc2*/', SDA_Occ2)
				target_out = target_out.replace('/*bdaocc2*/', occupation2)
				target_out = target_out.replace('/*drugsnocomma*/', drugsnocomma)
				target_out = target_out.replace('/*dispPeriod*/', '')
				target_out = target_out.replace('/*suppApplied*/', 'No')
				target_out = target_out.replace('/*supp_text_file*/', '')
				target_out = target_out.replace('/*supp_Match_Type*/', '""')
				target_out = target_out.replace('/*bda_only*/', bDa_only)
				target_out = target_out.replace('/*sda_only*/', sDa_only)
				target_out = target_out.replace('/*createPivotTable*/', createPivotTable)
				target_out = target_out.replace('/*pivYes1*/', pivSeg1)
				target_out = target_out.replace('/*pivYes2*/', pivSeg2)
				target_out = target_out.replace('/*totalSDAS*/', finalSDATotal)
				target_out = target_out.replace('/*totalBDAS*/', finalBDATotal)
				new_file.write(target_out)
				line_file = new_file

		elif caseType == 'listMatch' and (listMatchType == 'Standard_Seg' or listMatchType == 'Exact_Seg'):
			segList = ', '.join(segmentList)
			segList2 = ' '.join(segmentList)
			copyfile(emailCode, os.path.join(outCode, 'Presales Automation_Email_Final.sas'))
			newInput = os.path.join(outCode, 'Presales Automation_Email_Final.sas')

			line_file = open(os.path.join(newInput),'r').readlines()
			new_file = open(os.path.join(newInput),'w')
			for line_in in line_file:
				target_out = line_in.replace('/*DocORQuiz*/', listProduct)
				target_out = target_out.replace('/*listMatchType*/', listMatchType)
				target_out = target_out.replace('/*target_text_file*/', outCode)
				target_out = target_out.replace('/*Segments*/', segList)
				target_out = target_out.replace('/*Segments2*/', segList2)
				target_out = target_out.replace('/*Brand*/', brand)
				target_out = target_out.replace('/*MY_INIT*/', yourIn)
				target_out = target_out.replace('/*Requester_Initials*/', reIn)
				target_out = target_out.replace('/*SDA_Occ*/', SDA_Occ)
				target_out = target_out.replace('/*SDA_Spec*/', SDA_Spec)
				target_out = target_out.replace('/*therapyClass*/', therapyClass)
				target_out = target_out.replace('/*yesORno*/', deDup)
				target_out = target_out.replace('/*LookUpPeriod*/', lookUpPeriod)
				target_out = target_out.replace('/*totalLoookUps*/', totalLookUps)
				target_out = target_out.replace('/*BDA_Occ*/', occupation)
				target_out = target_out.replace('/*BDA_Spec*/', specialty)
				target_out = target_out.replace('/*drugList*/', ';')
				target_out = target_out.replace('/*caseno*/', caseno)
				target_out = target_out.replace('/*manu*/', manu)
				target_out = target_out.replace('/*mtype*/', mtype)
				target_out = target_out.replace('/*SE*/', SE)
				target_out = target_out.replace('/*username*/', email)
				target_out = target_out.replace('/*sdaocc2*/', SDA_Occ2)
				target_out = target_out.replace('/*bdaocc2*/', occupation2)
				target_out = target_out.replace('/*drugsnocomma*/', drugsnocomma)
				target_out = target_out.replace('/*dispPeriod*/', '')
				target_out = target_out.replace('/*suppApplied*/', 'No')
				target_out = target_out.replace('/*supp_text_file*/', '')
				target_out = target_out.replace('/*supp_Match_Type*/', '""')
				target_out = target_out.replace('/*bda_only*/', bDa_only)
				target_out = target_out.replace('/*sda_only*/', sDa_only)
				target_out = target_out.replace('/*createPivotTable*/', createPivotTable)
				target_out = target_out.replace('/*pivYes1*/', pivSeg1)
				target_out = target_out.replace('/*pivYes2*/', pivSeg2)
				target_out = target_out.replace('/*totalSDAS*/', finalSDATotal)
				target_out = target_out.replace('/*totalBDAS*/', finalBDATotal)
				new_file.write(target_out)
				line_file = new_file
			
			if bDa == 'y' and sDa == 'y':
				copyfile(emailCode, os.path.join(outCode, 'Presales Automation_Email_Final.sas'))
				newInput = os.path.join(outCode, 'Presales Automation_Email_Final.sas')

				line_file = open(os.path.join(newInput),'r').readlines()
				new_file = open(os.path.join(newInput),'w')
				for line_in in line_file:
					target_out = line_in.replace('/*DocORQuiz*/', listProduct)
					target_out = target_out.replace('/*listMatchType*/', listMatchType)
					target_out = target_out.replace('/*target_text_file*/', outCode)
					target_out = target_out.replace('/*Segments*/', segList)
					target_out = target_out.replace('/*Segments2*/', segList2)
					target_out = target_out.replace('/*Brand*/', brand)
					target_out = target_out.replace('/*MY_INIT*/', yourIn)
					target_out = target_out.replace('/*Requester_Initials*/', reIn)
					target_out = target_out.replace('/*SDA_Occ*/', SDA_Occ)
					target_out = target_out.replace('/*SDA_Spec*/', SDA_Spec)
					target_out = target_out.replace('/*therapyClass*/', therapyClass)
					target_out = target_out.replace('/*yesORno*/', deDup)
					target_out = target_out.replace('/*LookUpPeriod*/', lookUpPeriod)
					target_out = target_out.replace('/*totalLoookUps*/', totalLookUps)
					target_out = target_out.replace('/*BDA_Occ*/', occupation)
					target_out = target_out.replace('/*BDA_Spec*/', specialty)
					target_out = target_out.replace('/*drugList*/', drugList)
					target_out = target_out.replace('/*caseno*/', caseno)
					target_out = target_out.replace('/*manu*/', manu)
					target_out = target_out.replace('/*mtype*/', mtype)
					target_out = target_out.replace('/*SE*/', SE)
					target_out = target_out.replace('/*username*/', email)
					target_out = target_out.replace('/*sdaocc2*/', SDA_Occ2)
					target_out = target_out.replace('/*bdaocc2*/', occupation2)
					target_out = target_out.replace('/*drugsnocomma*/', drugsnocomma)
					target_out = target_out.replace('/*dispPeriod*/', displayPeriod)
					target_out = target_out.replace('/*suppApplied*/', 'No')
					target_out = target_out.replace('/*supp_text_file*/', '')
					target_out = target_out.replace('/*supp_Match_Type*/', '""')
					target_out = target_out.replace('/*bda_only*/', bDa_only)
					target_out = target_out.replace('/*sda_only*/', sDa_only)
					target_out = target_out.replace('/*createPivotTable*/', createPivotTable)
					target_out = target_out.replace('/*pivYes1*/', pivSeg1)
					target_out = target_out.replace('/*pivYes2*/', pivSeg2)
					target_out = target_out.replace('/*totalSDAS*/', finalSDATotal)
					target_out = target_out.replace('/*totalBDAS*/', finalBDATotal)
					new_file.write(target_out)
					line_file = new_file
				
			else:
				if bDa == 'y':  
					copyfile(emailCode, os.path.join(outCode, 'Presales Automation_Email_Final.sas'))
					newInput = os.path.join(outCode, 'Presales Automation_Email_Final.sas')

					line_file = open(os.path.join(newInput),'r').readlines()
					new_file = open(os.path.join(newInput),'w')
					for line_in in line_file:
						target_out = line_in.replace('/*DocORQuiz*/', listProduct)
						target_out = target_out.replace('/*listMatchType*/', listMatchType)
						target_out = target_out.replace('/*target_text_file*/', outCode)
						target_out = target_out.replace('/*Segments*/', segList)
						target_out = target_out.replace('/*Segments2*/', segList2)
						target_out = target_out.replace('/*Brand*/', brand)
						target_out = target_out.replace('/*MY_INIT*/', yourIn)
						target_out = target_out.replace('/*Requester_Initials*/', reIn)
						target_out = target_out.replace('/*SDA_Occ*/', SDA_Occ)
						target_out = target_out.replace('/*SDA_Spec*/', SDA_Spec)
						target_out = target_out.replace('/*therapyClass*/', therapyClass)
						target_out = target_out.replace('/*yesORno*/', deDup)
						target_out = target_out.replace('/*LookUpPeriod*/', lookUpPeriod)
						target_out = target_out.replace('/*totalLoookUps*/', totalLookUps)
						target_out = target_out.replace('/*BDA_Occ*/', occupation)
						target_out = target_out.replace('/*BDA_Spec*/', specialty)
						target_out = target_out.replace('/*drugList*/', drugList)
						target_out = target_out.replace('/*caseno*/', caseno)
						target_out = target_out.replace('/*manu*/', manu)
						target_out = target_out.replace('/*mtype*/', mtype)
						target_out = target_out.replace('/*SE*/', SE)
						target_out = target_out.replace('/*username*/', email)
						target_out = target_out.replace('/*sdaocc2*/', SDA_Occ2)
						target_out = target_out.replace('/*bdaocc2*/', occupation2)
						target_out = target_out.replace('/*drugsnocomma*/', drugsnocomma)
						target_out = target_out.replace('/*dispPeriod*/', displayPeriod)
						target_out = target_out.replace('/*suppApplied*/', 'No')
						target_out = target_out.replace('/*supp_text_file*/', '')
						target_out = target_out.replace('/*supp_Match_Type*/', '""')
						target_out = target_out.replace('/*bda_only*/', bDa_only)
						target_out = target_out.replace('/*sda_only*/', sDa_only)
						target_out = target_out.replace('/*createPivotTable*/', createPivotTable)
						target_out = target_out.replace('/*pivYes1*/', pivSeg1)
						target_out = target_out.replace('/*pivYes2*/', pivSeg2)
						target_out = target_out.replace('/*totalSDAS*/', finalSDATotal)
						target_out = target_out.replace('/*totalBDAS*/', finalBDATotal)
						new_file.write(target_out)
						line_file = new_file
						
				if sDa == 'y':  
					copyfile(emailCode, os.path.join(outCode, 'Presales Automation_Email_Final.sas'))
					newInput = os.path.join(outCode, 'Presales Automation_Email_Final.sas')

					line_file = open(os.path.join(newInput),'r').readlines()
					new_file = open(os.path.join(newInput),'w')
					for line_in in line_file:
						target_out = line_in.replace('/*DocORQuiz*/', listProduct)
						target_out = target_out.replace('/*listMatchType*/', listMatchType)
						target_out = target_out.replace('/*target_text_file*/', outCode)
						target_out = target_out.replace('/*Segments*/', segList)
						target_out = target_out.replace('/*Segments2*/', segList2)
						target_out = target_out.replace('/*Brand*/', brand)
						target_out = target_out.replace('/*MY_INIT*/', yourIn)
						target_out = target_out.replace('/*Requester_Initials*/', reIn)
						target_out = target_out.replace('/*SDA_Occ*/', SDA_Occ)
						target_out = target_out.replace('/*SDA_Spec*/', SDA_Spec)
						target_out = target_out.replace('/*therapyClass*/', therapyClass)
						target_out = target_out.replace('/*yesORno*/', deDup)
						target_out = target_out.replace('/*LookUpPeriod*/', lookUpPeriod)
						target_out = target_out.replace('/*totalLoookUps*/', totalLookUps)
						target_out = target_out.replace('/*BDA_Occ*/', occupation)
						target_out = target_out.replace('/*BDA_Spec*/', specialty)
						target_out = target_out.replace('/*drugList*/', ';')
						target_out = target_out.replace('/*caseno*/', caseno)
						target_out = target_out.replace('/*manu*/', manu)
						target_out = target_out.replace('/*mtype*/', mtype)
						target_out = target_out.replace('/*SE*/', SE)
						target_out = target_out.replace('/*username*/', email)
						target_out = target_out.replace('/*sdaocc2*/', SDA_Occ2)
						target_out = target_out.replace('/*bdaocc2*/', occupation2)
						target_out = target_out.replace('/*drugsnocomma*/', drugsnocomma)
						target_out = target_out.replace('/*dispPeriod*/', '')
						target_out = target_out.replace('/*suppApplied*/', 'No')
						target_out = target_out.replace('/*supp_text_file*/', '')
						target_out = target_out.replace('/*supp_Match_Type*/', '""')
						target_out = target_out.replace('/*bda_only*/', bDa_only)
						target_out = target_out.replace('/*sda_only*/', sDa_only)
						target_out = target_out.replace('/*createPivotTable*/', createPivotTable)
						target_out = target_out.replace('/*pivYes1*/', pivSeg1)
						target_out = target_out.replace('/*pivYes2*/', pivSeg2)
						target_out = target_out.replace('/*totalSDAS*/', finalSDATotal)
						target_out = target_out.replace('/*totalBDAS*/', finalBDATotal)
						new_file.write(target_out)
						line_file = new_file

		elif caseType == 'Targeting':
			targetOut = """P:\\Epocrates Analytics\\TARGETS\\{date}\\{manu} {brand}""".format(date = date, manu = manu, brand = brand)
			outCode2 = """P:\\Epocrates Analytics\\TARGETS\\{date}\\{manu} {brand}{slashes}""".format(date = date, slashes = "\\", manu = manu, brand = brand)
			if dSharing == 'Y' and (listMatchType == 'Standard' or listMatchType == 'Standard_Seg' or listMatchType == 'Exact' or listMatchType == 'Exact_Seg'):
				segList = ', '.join(segmentList)
				copyfile(targetAuto, os.path.join(outCode2, 'Targeting Automation Code_OFFICIAL.sas'))
				newInput = os.path.join(outCode2, 'Targeting Automation Code_OFFICIAL.sas')

				line_file = open(os.path.join(newInput),'r').readlines()
				new_file = open(os.path.join(newInput),'w')
				# print 'Im running the write copy sas code that should copy datacap'
				for line_in in line_file:
					target_out = line_in.replace('/*DocORQuiz*/', listProduct)
					target_out = target_out.replace('/*listMatchType*/', listMatchType)
					target_out = target_out.replace('/*target_text_file*/', outCode3)
					target_out = target_out.replace('/*Date*/', date)
					target_out = target_out.replace('/*targetFoler*/', targetOut)
					target_out = target_out.replace('/*Brand*/', brand)
					target_out = target_out.replace('?MY_INIT?', yourIn)
					target_out = target_out.replace('/*TargetNum*/', targetNum)
					target_out = target_out.replace('/*segVar*/', segVariable)
					target_out = target_out.replace('/*segValues*/', varValues)
					target_out = target_out.replace('/*dataShareYorN*/', dSharing)
					target_out = target_out.replace('/*Segments*/', segList)
					target_out = target_out.replace('/*keep_seg*/', keep_seg)
					target_out = target_out.replace('/*Manu*/', manu)
					target_out = target_out.replace('/*SDAONLY*/', sDa_only)
					target_out = target_out.replace('/*SDA_Occ*/', SDA_Occ)
					target_out = target_out.replace('/*SDA_Spec*/', SDA_Spec)
					target_out = target_out.replace('/*SDATarget*/', SDA_Target)
					target_out = target_out.replace('/*BDAONLY*/', bDa_only)
					target_out = target_out.replace('/*yesORno*/', deDup)
					target_out = target_out.replace('/*LookUpPeriod*/', lookUpPeriod)
					target_out = target_out.replace('/*totalLoookUps*/', totalLookUps)
					target_out = target_out.replace('/*BDA_Occ*/', occupation)
					target_out = target_out.replace('/*BDA_Spec*/', specialty)
					target_out = target_out.replace('/*BDATarget*/', BDA_Target)
					target_out = target_out.replace('/*drugList*/', drugList)
					target_out = target_out.replace('/*suppApplied*/', 'No')
					target_out = target_out.replace('/*supp_text_file*/', '')
					target_out = target_out.replace('/*supp_Match_Type*/', '""')
					target_out = target_out.replace('/*dataCap*/', dataCap)
					target_out = target_out.replace('/*backFill*/', backFill)
					target_out = target_out.replace('/*needVals*/', neededValues)
					target_out = target_out.replace('/*backFillVals*/', backFillValue)
					target_out = target_out.replace('/*suppSDAOnly*/', suppSDAOnly)
					target_out = target_out.replace('/*suppBDAOnly*/', suppBDAOnly)
					new_file.write(target_out)
					line_file = new_file

			if dSharing == 'N' and (listMatchType == 'Standard_Seg' or listMatchType == 'Exact_Seg'):
				segList = finalSeg
				copyfile(targetAuto, os.path.join(outCode2, 'Targeting Automation Code_OFFICIAL.sas'))
				newInput = os.path.join(outCode2, 'Targeting Automation Code_OFFICIAL.sas')

				line_file = open(os.path.join(newInput),'r').readlines()
				new_file = open(os.path.join(newInput),'w')
				for line_in in line_file:
					target_out = line_in.replace('/*DocORQuiz*/', listProduct)
					target_out = target_out.replace('/*listMatchType*/', listMatchType)
					target_out = target_out.replace('/*target_text_file*/', outCode3)
					target_out = target_out.replace('/*Date*/', date)
					target_out = target_out.replace('/*targetFoler*/', targetOut)
					target_out = target_out.replace('/*Brand*/', brand)
					target_out = target_out.replace('?MY_INIT?', yourIn)
					target_out = target_out.replace('/*TargetNum*/', targetNum)
					target_out = target_out.replace('/*segVar*/', segVariable)
					target_out = target_out.replace('/*segValues*/', varValues)
					target_out = target_out.replace('/*dataShareYorN*/', dSharing)
					target_out = target_out.replace('/*Segments*/', segList)
					target_out = target_out.replace('/*keep_seg*/', keep_seg)
					target_out = target_out.replace('/*Manu*/', manu)
					target_out = target_out.replace('/*SDAONLY*/', sDa_only)
					target_out = target_out.replace('/*SDA_Occ*/', SDA_Occ)
					target_out = target_out.replace('/*SDA_Spec*/', SDA_Spec)
					target_out = target_out.replace('/*SDATarget*/', SDA_Target)
					target_out = target_out.replace('/*BDAONLY*/', bDa_only)
					target_out = target_out.replace('/*yesORno*/', deDup)
					target_out = target_out.replace('/*LookUpPeriod*/', lookUpPeriod)
					target_out = target_out.replace('/*totalLoookUps*/', totalLookUps)
					target_out = target_out.replace('/*BDA_Occ*/', occupation)
					target_out = target_out.replace('/*BDA_Spec*/', specialty)
					target_out = target_out.replace('/*BDATarget*/', BDA_Target)
					target_out = target_out.replace('/*drugList*/', drugList)
					target_out = target_out.replace('/*suppApplied*/', 'No')
					target_out = target_out.replace('/*supp_text_file*/', '')
					target_out = target_out.replace('/*supp_Match_Type*/', '""')
					target_out = target_out.replace('/*dataCap*/', dataCap)
					target_out = target_out.replace('/*backFill*/', backFill)
					target_out = target_out.replace('/*needVals*/', neededValues)
					target_out = target_out.replace('/*backFillVals*/', backFillValue)
					target_out = target_out.replace('/*suppSDAOnly*/', suppSDAOnly)
					target_out = target_out.replace('/*suppBDAOnly*/', suppBDAOnly)
					new_file.write(target_out)
					line_file = new_file

			if dSharing == 'N' and (listMatchType == 'Standard' or listMatchType == 'Exact') and (sDa_only == 'N' and bDa_only == 'N'):
				segList = str(finalSeg)
				# print 'copying sas code'
				copyfile(targetAuto, os.path.join(outCode2, 'Targeting Automation Code_OFFICIAL.sas'))
				newInput = os.path.join(outCode2, 'Targeting Automation Code_OFFICIAL.sas')

				line_file = open(os.path.join(newInput),'r').readlines()
				new_file = open(os.path.join(newInput),'w')
				for line_in in line_file:
					target_out = line_in.replace('/*DocORQuiz*/', listProduct)
					target_out = target_out.replace('/*listMatchType*/', listMatchType)
					target_out = target_out.replace('/*target_text_file*/', outCode3)
					target_out = target_out.replace('/*Date*/', date)
					target_out = target_out.replace('/*targetFoler*/', targetOut)
					target_out = target_out.replace('/*Brand*/', brand)
					target_out = target_out.replace('?MY_INIT?', yourIn)
					target_out = target_out.replace('/*TargetNum*/', targetNum)
					target_out = target_out.replace('/*segVar*/', segVariable)
					target_out = target_out.replace('/*segValues*/', str(config['varValues']))
					target_out = target_out.replace('/*dataShareYorN*/', dSharing)
					target_out = target_out.replace('/*Segments*/', segList)
					target_out = target_out.replace('/*keep_seg*/', keep_seg)
					target_out = target_out.replace('/*Manu*/', manu)
					target_out = target_out.replace('/*SDAONLY*/', sDa_only)
					target_out = target_out.replace('/*SDA_Occ*/', SDA_Occ)
					target_out = target_out.replace('/*SDA_Spec*/', SDA_Spec)
					target_out = target_out.replace('/*SDATarget*/', SDA_Target)
					target_out = target_out.replace('/*BDAONLY*/', bDa_only)
					target_out = target_out.replace('/*yesORno*/', deDup)
					target_out = target_out.replace('/*LookUpPeriod*/', lookUpPeriod)
					target_out = target_out.replace('/*totalLoookUps*/', totalLookUps)
					target_out = target_out.replace('/*BDA_Occ*/', occupation)
					target_out = target_out.replace('/*BDA_Spec*/', specialty)
					target_out = target_out.replace('/*BDATarget*/', BDA_Target)
					target_out = target_out.replace('/*drugList*/', drugList)
					target_out = target_out.replace('/*suppApplied*/', 'No')
					target_out = target_out.replace('/*supp_text_file*/', '')
					target_out = target_out.replace('/*supp_Match_Type*/', '""')
					target_out = target_out.replace('/*dataCap*/', dataCap)
					target_out = target_out.replace('/*backFill*/', backFill)
					target_out = target_out.replace('/*needVals*/', neededValues)
					target_out = target_out.replace('/*backFillVals*/', backFillValue)
					target_out = target_out.replace('/*suppSDAOnly*/', suppSDAOnly)
					target_out = target_out.replace('/*suppBDAOnly*/', suppBDAOnly)
					new_file.write(target_out)
					line_file = new_file

			if sDa_only == 'Y':
				# segList = ', '.join(segmentList)
				copyfile(targetAuto, os.path.join(outCode2, 'Targeting Automation Code_OFFICIAL.sas'))
				newInput = os.path.join(outCode2, 'Targeting Automation Code_OFFICIAL.sas')

				line_file = open(os.path.join(newInput),'r').readlines()
				new_file = open(os.path.join(newInput),'w')
				for line_in in line_file:
					target_out = line_in.replace('/*DocORQuiz*/', listProduct)
					target_out = target_out.replace('/*listMatchType*/', 'None')
					target_out = target_out.replace('/*target_text_file*/', '')
					target_out = target_out.replace('/*Date*/', date)
					target_out = target_out.replace('/*targetFoler*/', targetOut)
					target_out = target_out.replace('/*Brand*/', brand)
					target_out = target_out.replace('?MY_INIT?', yourIn)
					target_out = target_out.replace('/*TargetNum*/', '')
					target_out = target_out.replace('/*segVar*/', '')
					target_out = target_out.replace('/*segValues*/', '')
					target_out = target_out.replace('/*dataShareYorN*/', dSharing)
					target_out = target_out.replace('/*Segments*/', '')
					target_out = target_out.replace('/*keep_seg*/', keep_seg)
					target_out = target_out.replace('/*Manu*/', manu)
					target_out = target_out.replace('/*SDAONLY*/', sDa_only)
					target_out = target_out.replace('/*SDA_Occ*/', SDA_Occ)
					target_out = target_out.replace('/*SDA_Spec*/', SDA_Spec)
					target_out = target_out.replace('/*SDATarget*/', SDA_Target)
					target_out = target_out.replace('/*BDAONLY*/', bDa_only)
					target_out = target_out.replace('/*yesORno*/', deDup)
					target_out = target_out.replace('/*LookUpPeriod*/', lookUpPeriod)
					target_out = target_out.replace('/*totalLoookUps*/', totalLookUps)
					target_out = target_out.replace('/*BDA_Occ*/', occupation)
					target_out = target_out.replace('/*BDA_Spec*/', specialty)
					target_out = target_out.replace('/*BDATarget*/', BDA_Target)
					target_out = target_out.replace('/*drugList*/', drugList)
					target_out = target_out.replace('/*suppApplied*/', 'No')
					target_out = target_out.replace('/*supp_text_file*/', '')
					target_out = target_out.replace('/*supp_Match_Type*/', '""')
					target_out = target_out.replace('/*dataCap*/', dataCap)
					target_out = target_out.replace('/*backFill*/', backFill)
					target_out = target_out.replace('/*needVals*/', neededValues)
					target_out = target_out.replace('/*backFillVals*/', backFillValue)
					target_out = target_out.replace('/*suppSDAOnly*/', suppSDAOnly)
					target_out = target_out.replace('/*suppBDAOnly*/', suppBDAOnly)
					new_file.write(target_out)
					line_file = new_file

			if bDa_only == 'Y':
				# segList = ', '.join(segmentList)
				copyfile(targetAuto, os.path.join(outCode2, 'Targeting Automation Code_OFFICIAL.sas'))
				newInput = os.path.join(outCode2, 'Targeting Automation Code_OFFICIAL.sas')

				line_file = open(os.path.join(newInput),'r').readlines()
				new_file = open(os.path.join(newInput),'w')
				for line_in in line_file:
					target_out = line_in.replace('/*DocORQuiz*/', listProduct)
					target_out = target_out.replace('/*listMatchType*/', 'None')
					target_out = target_out.replace('/*target_text_file*/', '')
					target_out = target_out.replace('/*Date*/', date)
					target_out = target_out.replace('/*targetFoler*/', targetOut)
					target_out = target_out.replace('/*Brand*/', brand)
					target_out = target_out.replace('?MY_INIT?', yourIn)
					target_out = target_out.replace('/*TargetNum*/', '')
					target_out = target_out.replace('/*segVar*/', '')
					target_out = target_out.replace('/*segValues*/', '')
					target_out = target_out.replace('/*dataShareYorN*/', dSharing)
					target_out = target_out.replace('/*Segments*/', '')
					target_out = target_out.replace('/*keep_seg*/', keep_seg)
					target_out = target_out.replace('/*Manu*/', manu)
					target_out = target_out.replace('/*SDAONLY*/', sDa_only)
					target_out = target_out.replace('/*SDA_Occ*/', SDA_Occ)
					target_out = target_out.replace('/*SDA_Spec*/', SDA_Spec)
					target_out = target_out.replace('/*SDATarget*/', SDA_Target)
					target_out = target_out.replace('/*BDAONLY*/', bDa_only)
					target_out = target_out.replace('/*yesORno*/', deDup)
					target_out = target_out.replace('/*LookUpPeriod*/', lookUpPeriod)
					target_out = target_out.replace('/*totalLoookUps*/', totalLookUps)
					target_out = target_out.replace('/*BDA_Occ*/', occupation)
					target_out = target_out.replace('/*BDA_Spec*/', specialty)
					target_out = target_out.replace('/*BDATarget*/', BDA_Target)
					target_out = target_out.replace('/*drugList*/', drugList)
					target_out = target_out.replace('/*suppApplied*/', 'No')
					target_out = target_out.replace('/*supp_text_file*/', '')
					target_out = target_out.replace('/*supp_Match_Type*/', '""')
					target_out = target_out.replace('/*dataCap*/', dataCap)
					target_out = target_out.replace('/*backFill*/', backFill)
					target_out = target_out.replace('/*needVals*/', neededValues)
					target_out = target_out.replace('/*backFillVals*/', backFillValue)
					target_out = target_out.replace('/*suppSDAOnly*/', suppSDAOnly)
					target_out = target_out.replace('/*suppBDAOnly*/', suppBDAOnly)
					new_file.write(target_out)
					line_file = new_file

	if suppApplied == 'Y':
		if caseType == 'listMatch' and (listMatchType == 'Standard' or listMatchType == 'Exact'):
			copyfile(emailCode, os.path.join(outCode, 'Presales Automation_Email_Final.sas'))
			newInput = os.path.join(outCode, 'Presales Automation_Email_Final.sas')

			line_file = open(os.path.join(newInput),'r').readlines()
			new_file = open(os.path.join(newInput),'w')
			for line_in in line_file:
				#target_out = line_in.replace('/*DocORQuiz*/', listProduct).replace('/*listMatchType*/', listMatchType).replace('/*target_text_file*/', outCode).replace('/*Segments*/', '').replace('/*Segments2*/', '').replace('/*Brand*/', brand).replace('/*MY_INIT*/', yourIn).replace('/*Requester_Initials*/', reIn).replace('/*SDA_Occ*/', '').replace('/*SDA_Spec*/', '').replace('/*yesORno*/', '').replace('/*LookUpPeriod*/', '').replace('/*totalLoookUps*/', '').replace('/*BDA_Occ*/', '').replace('/*BDA_Spec*/', '').replace('/*drugList*/', '')
				target_out = line_in.replace('/*DocORQuiz*/', listProduct)
				target_out = target_out.replace('/*listMatchType*/', listMatchType)
				target_out = target_out.replace('/*target_text_file*/', outCode)
				target_out = target_out.replace('/*Segments*/', '')
				target_out = target_out.replace('/*Segments2*/', '')
				target_out = target_out.replace('/*Brand*/', brand)
				target_out = target_out.replace('/*MY_INIT*/', yourIn)
				target_out = target_out.replace('/*Requester_Initials*/', reIn)
				target_out = target_out.replace('/*SDA_Occ*/', SDA_Occ)
				target_out = target_out.replace('/*SDA_Spec*/', SDA_Spec)
				target_out = target_out.replace('/*yesORno*/', deDup)
				target_out = target_out.replace('/*LookUpPeriod*/', lookUpPeriod)
				target_out = target_out.replace('/*totalLoookUps*/', totalLookUps)
				target_out = target_out.replace('/*BDA_Occ*/', occupation)
				target_out = target_out.replace('/*BDA_Spec*/', specialty)
				target_out = target_out.replace('/*drugList*/', drugList)
				target_out = target_out.replace('/*caseno*/', caseno)
				target_out = target_out.replace('/*manu*/', manu)
				target_out = target_out.replace('/*mtype*/', mtype)
				target_out = target_out.replace('/*SE*/', SE)
				target_out = target_out.replace('/*username*/', email)
				target_out = target_out.replace('/*sdaocc2*/', SDA_Occ2)
				target_out = target_out.replace('/*bdaocc2*/', occupation2)
				target_out = target_out.replace('/*drugsnocomma*/', drugsnocomma)
				target_out = target_out.replace('/*dispPeriod*/', '')
				target_out = target_out.replace('/*suppApplied*/', 'Yes')
				target_out = target_out.replace('/*supp_text_file*/', suppFileLocation)
				target_out = target_out.replace('/*supp_Match_Type*/', str(config['suppMatchType']))
				target_out = target_out.replace('/*therapyClass*/', therapyClass)
				target_out = target_out.replace('/*bda_only*/', bDa_only)
				target_out = target_out.replace('/*sda_only*/', sDa_only)
				target_out = target_out.replace('/*createPivotTable*/', createPivotTable)
				target_out = target_out.replace('/*pivYes1*/', pivSeg1)
				target_out = target_out.replace('/*pivYes2*/', pivSeg2)
				target_out = target_out.replace('/*totalSDAS*/', finalSDATotal)
				target_out = target_out.replace('/*totalBDAS*/', finalBDATotal)
				new_file.write(target_out)
				line_file = new_file
				
			if bDa == 'y' and sDa == 'y':
				copyfile(emailCode, os.path.join(outCode, 'Presales Automation_Email_Final.sas'))
				newInput = os.path.join(outCode, 'Presales Automation_Email_Final.sas')

				line_file = open(os.path.join(newInput),'r').readlines()
				new_file = open(os.path.join(newInput),'w')
				for line_in in line_file:
					#target_out = line_in.replace('/*DocORQuiz*/', listProduct).replace('/*listMatchType*/', listMatchType).replace('/*target_text_file*/', outCode).replace('/*Segments*/', '').replace('/*Segments2*/', '').replace('/*Brand*/', brand).replace('/*MY_INIT*/', yourIn).replace('/*Requester_Initials*/', reIn).replace('/*SDA_Occ*/', '').replace('/*SDA_Spec*/', '').replace('/*yesORno*/', '').replace('/*LookUpPeriod*/', '').replace('/*totalLoookUps*/', '').replace('/*BDA_Occ*/', '').replace('/*BDA_Spec*/', '').replace('/*drugList*/', '')
					target_out = line_in.replace('/*DocORQuiz*/', listProduct)
					target_out = target_out.replace('/*listMatchType*/', listMatchType)
					target_out = target_out.replace('/*target_text_file*/', outCode)
					target_out = target_out.replace('/*Segments*/', '')
					target_out = target_out.replace('/*Segments2*/', '')
					target_out = target_out.replace('/*Brand*/', brand)
					target_out = target_out.replace('/*MY_INIT*/', yourIn)
					target_out = target_out.replace('/*Requester_Initials*/', reIn)
					target_out = target_out.replace('/*SDA_Occ*/', SDA_Occ)
					target_out = target_out.replace('/*SDA_Spec*/', SDA_Spec)
					target_out = target_out.replace('/*yesORno*/', deDup)
					target_out = target_out.replace('/*LookUpPeriod*/', lookUpPeriod)
					target_out = target_out.replace('/*totalLoookUps*/', totalLookUps)
					target_out = target_out.replace('/*BDA_Occ*/', occupation)
					target_out = target_out.replace('/*BDA_Spec*/', specialty)
					target_out = target_out.replace('/*drugList*/', drugList)
					target_out = target_out.replace('/*caseno*/', caseno)
					target_out = target_out.replace('/*manu*/', manu)
					target_out = target_out.replace('/*mtype*/', mtype)
					target_out = target_out.replace('/*SE*/', SE)
					target_out = target_out.replace('/*username*/', email)
					target_out = target_out.replace('/*sdaocc2*/', SDA_Occ2)
					target_out = target_out.replace('/*bdaocc2*/', occupation2)
					target_out = target_out.replace('/*drugsnocomma*/', drugsnocomma)
					target_out = target_out.replace('/*dispPeriod*/', displayPeriod)
					target_out = target_out.replace('/*suppApplied*/', 'Yes')
					target_out = target_out.replace('/*supp_text_file*/', suppFileLocation)
					target_out = target_out.replace('/*supp_Match_Type*/', str(config['suppMatchType']))
					target_out = target_out.replace('/*therapyClass*/', therapyClass)
					target_out = target_out.replace('/*bda_only*/', bDa_only)
					target_out = target_out.replace('/*sda_only*/', sDa_only)
					target_out = target_out.replace('/*createPivotTable*/', createPivotTable)
					target_out = target_out.replace('/*pivYes1*/', pivSeg1)
					target_out = target_out.replace('/*pivYes2*/', pivSeg2)
					target_out = target_out.replace('/*totalSDAS*/', finalSDATotal)
					target_out = target_out.replace('/*totalBDAS*/', finalBDATotal)
					new_file.write(target_out)
					line_file = new_file
				
			else:
				if bDa == 'y':  
					copyfile(emailCode, os.path.join(outCode, 'Presales Automation_Email_Final.sas'))
					newInput = os.path.join(outCode, 'Presales Automation_Email_Final.sas')

					line_file = open(os.path.join(newInput),'r').readlines()
					new_file = open(os.path.join(newInput),'w')
					for line_in in line_file:
						#target_out = line_in.replace('/*DocORQuiz*/', listProduct).replace('/*listMatchType*/', listMatchType).replace('/*target_text_file*/', outCode).replace('/*Segments*/', '').replace('/*Segments2*/', '').replace('/*Brand*/', brand).replace('/*MY_INIT*/', yourIn).replace('/*Requester_Initials*/', reIn).replace('/*SDA_Occ*/', '').replace('/*SDA_Spec*/', '').replace('/*yesORno*/', '').replace('/*LookUpPeriod*/', '').replace('/*totalLoookUps*/', '').replace('/*BDA_Occ*/', '').replace('/*BDA_Spec*/', '').replace('/*drugList*/', '')
						target_out = line_in.replace('/*DocORQuiz*/', listProduct)
						target_out = target_out.replace('/*listMatchType*/', listMatchType)
						target_out = target_out.replace('/*target_text_file*/', outCode)
						target_out = target_out.replace('/*Segments*/', '')
						target_out = target_out.replace('/*Segments2*/', '')
						target_out = target_out.replace('/*Brand*/', brand)
						target_out = target_out.replace('/*MY_INIT*/', yourIn)
						target_out = target_out.replace('/*Requester_Initials*/', reIn)
						target_out = target_out.replace('/*SDA_Occ*/', SDA_Occ)
						target_out = target_out.replace('/*SDA_Spec*/', SDA_Spec)
						target_out = target_out.replace('/*yesORno*/', deDup)
						target_out = target_out.replace('/*LookUpPeriod*/', lookUpPeriod)
						target_out = target_out.replace('/*totalLoookUps*/', totalLookUps)
						target_out = target_out.replace('/*BDA_Occ*/', occupation)
						target_out = target_out.replace('/*BDA_Spec*/', specialty)
						target_out = target_out.replace('/*drugList*/', drugList)                   
						target_out = target_out.replace('/*caseno*/', caseno)
						target_out = target_out.replace('/*manu*/', manu)
						target_out = target_out.replace('/*mtype*/', mtype)
						target_out = target_out.replace('/*SE*/', SE)
						target_out = target_out.replace('/*username*/', email)
						target_out = target_out.replace('/*sdaocc2*/', SDA_Occ2)
						target_out = target_out.replace('/*bdaocc2*/', occupation2)
						target_out = target_out.replace('/*drugsnocomma*/', drugsnocomma)
						target_out = target_out.replace('/*dispPeriod*/', displayPeriod)
						target_out = target_out.replace('/*suppApplied*/', 'Yes')
						target_out = target_out.replace('/*supp_text_file*/', suppFileLocation)
						target_out = target_out.replace('/*supp_Match_Type*/', str(config['suppMatchType']))
						target_out = target_out.replace('/*therapyClass*/', therapyClass)
						target_out = target_out.replace('/*bda_only*/', bDa_only)
						target_out = target_out.replace('/*sda_only*/', sDa_only)
						target_out = target_out.replace('/*createPivotTable*/', createPivotTable)
						target_out = target_out.replace('/*pivYes1*/', pivSeg1)
						target_out = target_out.replace('/*pivYes2*/', pivSeg2)
						target_out = target_out.replace('/*totalSDAS*/', finalSDATotal)
						target_out = target_out.replace('/*totalBDAS*/', finalBDATotal)
						new_file.write(target_out)
						line_file = new_file
						
				if sDa == 'y':  
					copyfile(emailCode, os.path.join(outCode, 'Presales Automation_Email_Final.sas'))
					newInput = os.path.join(outCode, 'Presales Automation_Email_Final.sas')

					line_file = open(os.path.join(newInput),'r').readlines()
					new_file = open(os.path.join(newInput),'w')
					for line_in in line_file:
						#target_out = line_in.replace('/*DocORQuiz*/', listProduct).replace('/*listMatchType*/', listMatchType).replace('/*target_text_file*/', outCode).replace('/*Segments*/', '').replace('/*Segments2*/', '').replace('/*Brand*/', brand).replace('/*MY_INIT*/', yourIn).replace('/*Requester_Initials*/', reIn).replace('/*SDA_Occ*/', '').replace('/*SDA_Spec*/', '').replace('/*yesORno*/', '').replace('/*LookUpPeriod*/', '').replace('/*totalLoookUps*/', '').replace('/*BDA_Occ*/', '').replace('/*BDA_Spec*/', '').replace('/*drugList*/', '')
						target_out = line_in.replace('/*DocORQuiz*/', listProduct)
						target_out = target_out.replace('/*listMatchType*/', listMatchType)
						target_out = target_out.replace('/*target_text_file*/', outCode)
						target_out = target_out.replace('/*Segments*/', '')
						target_out = target_out.replace('/*Segments2*/', '')
						target_out = target_out.replace('/*Brand*/', brand)
						target_out = target_out.replace('/*MY_INIT*/', yourIn)
						target_out = target_out.replace('/*Requester_Initials*/', reIn)
						target_out = target_out.replace('/*SDA_Occ*/', SDA_Occ)
						target_out = target_out.replace('/*SDA_Spec*/', SDA_Spec)
						target_out = target_out.replace('/*yesORno*/', deDup)
						target_out = target_out.replace('/*LookUpPeriod*/', lookUpPeriod)
						target_out = target_out.replace('/*totalLoookUps*/', totalLookUps)
						target_out = target_out.replace('/*BDA_Occ*/', occupation)
						target_out = target_out.replace('/*BDA_Spec*/', specialty)
						target_out = target_out.replace('/*drugList*/', drugList)
						target_out = target_out.replace('/*caseno*/', caseno)
						target_out = target_out.replace('/*manu*/', manu)
						target_out = target_out.replace('/*mtype*/', mtype)
						target_out = target_out.replace('/*SE*/', SE)
						target_out = target_out.replace('/*username*/', email)
						target_out = target_out.replace('/*sdaocc2*/', SDA_Occ2)
						target_out = target_out.replace('/*bdaocc2*/', occupation2)
						target_out = target_out.replace('/*drugsnocomma*/', drugsnocomma)
						target_out = target_out.replace('/*dispPeriod*/', '')
						target_out = target_out.replace('/*suppApplied*/', 'Yes')
						target_out = target_out.replace('/*supp_text_file*/', suppFileLocation)
						target_out = target_out.replace('/*supp_Match_Type*/', str(config['suppMatchType']))
						target_out = target_out.replace('/*therapyClass*/', therapyClass)
						target_out = target_out.replace('/*bda_only*/', bDa_only)
						target_out = target_out.replace('/*sda_only*/', sDa_only)
						target_out = target_out.replace('/*createPivotTable*/', createPivotTable)
						target_out = target_out.replace('/*pivYes1*/', pivSeg1)
						target_out = target_out.replace('/*pivYes2*/', pivSeg2)
						target_out = target_out.replace('/*totalSDAS*/', finalSDATotal)
						target_out = target_out.replace('/*totalBDAS*/', finalBDATotal)
						new_file.write(target_out)
						line_file = new_file

		elif caseType == 'listMatch' and (listMatchType == 'Standard_Seg' or listMatchType == 'Exact_Seg'):
			segList = ', '.join(segmentList)
			segList2 = ' '.join(segmentList)
			copyfile(emailCode, os.path.join(outCode, 'Presales Automation_Email_Final.sas'))
			newInput = os.path.join(outCode, 'Presales Automation_Email_Final.sas')

			line_file = open(os.path.join(newInput),'r').readlines()
			new_file = open(os.path.join(newInput),'w')
			for line_in in line_file:
				target_out = line_in.replace('/*DocORQuiz*/', listProduct)
				target_out = target_out.replace('/*listMatchType*/', listMatchType)
				target_out = target_out.replace('/*suppFile*/', suppFileLocation)
				target_out = target_out.replace('/*target_text_file*/', outCode)
				target_out = target_out.replace('/*Segments*/', segList)
				target_out = target_out.replace('/*Segments2*/', segList2)
				target_out = target_out.replace('/*Brand*/', brand)
				target_out = target_out.replace('/*MY_INIT*/', yourIn)
				target_out = target_out.replace('/*Requester_Initials*/', reIn)
				target_out = target_out.replace('/*SDA_Occ*/', SDA_Occ)
				target_out = target_out.replace('/*SDA_Spec*/', SDA_Spec)
				target_out = target_out.replace('/*yesORno*/', deDup)
				target_out = target_out.replace('/*LookUpPeriod*/', lookUpPeriod)
				target_out = target_out.replace('/*totalLoookUps*/', totalLookUps)
				target_out = target_out.replace('/*BDA_Occ*/', occupation)
				target_out = target_out.replace('/*BDA_Spec*/', specialty)
				target_out = target_out.replace('/*drugList*/', drugList)
				target_out = target_out.replace('/*caseno*/', caseno)
				target_out = target_out.replace('/*manu*/', manu)
				target_out = target_out.replace('/*mtype*/', mtype)
				target_out = target_out.replace('/*SE*/', SE)
				target_out = target_out.replace('/*username*/', email)
				target_out = target_out.replace('/*sdaocc2*/', SDA_Occ2)
				target_out = target_out.replace('/*bdaocc2*/', occupation2)
				target_out = target_out.replace('/*drugsnocomma*/', drugsnocomma)
				target_out = target_out.replace('/*dispPeriod*/', '')
				target_out = target_out.replace('/*suppApplied*/', 'Yes')
				target_out = target_out.replace('/*supp_text_file*/', suppFileLocation)
				target_out = target_out.replace('/*supp_Match_Type*/', str(config['suppMatchType']))
				target_out = target_out.replace('/*therapyClass*/', therapyClass)
				target_out = target_out.replace('/*bda_only*/', bDa_only)
				target_out = target_out.replace('/*sda_only*/', sDa_only)
				target_out = target_out.replace('/*createPivotTable*/', createPivotTable)
				target_out = target_out.replace('/*pivYes1*/', pivSeg1)
				target_out = target_out.replace('/*pivYes2*/', pivSeg2)
				target_out = target_out.replace('/*totalSDAS*/', finalSDATotal)
				target_out = target_out.replace('/*totalBDAS*/', finalBDATotal)
				new_file.write(target_out)
				line_file = new_file
			
			if bDa == 'y' and sDa == 'y':
				copyfile(emailCode, os.path.join(outCode, 'Presales Automation_Email_Final.sas'))
				newInput = os.path.join(outCode, 'Presales Automation_Email_Final.sas')

				line_file = open(os.path.join(newInput),'r').readlines()
				new_file = open(os.path.join(newInput),'w')
				for line_in in line_file:
					target_out = line_in.replace('/*DocORQuiz*/', listProduct)
					target_out = target_out.replace('/*listMatchType*/', listMatchType)
					target_out = target_out.replace('/*target_text_file*/', outCode)
					target_out = target_out.replace('/*Segments*/', segList)
					target_out = target_out.replace('/*Segments2*/', segList2)
					target_out = target_out.replace('/*Brand*/', brand)
					target_out = target_out.replace('/*MY_INIT*/', yourIn)
					target_out = target_out.replace('/*Requester_Initials*/', reIn)
					target_out = target_out.replace('/*SDA_Occ*/', SDA_Occ)
					target_out = target_out.replace('/*SDA_Spec*/', SDA_Spec)
					target_out = target_out.replace('/*yesORno*/', deDup)
					target_out = target_out.replace('/*LookUpPeriod*/', lookUpPeriod)
					target_out = target_out.replace('/*totalLoookUps*/', totalLookUps)
					target_out = target_out.replace('/*BDA_Occ*/', occupation)
					target_out = target_out.replace('/*BDA_Spec*/', specialty)
					target_out = target_out.replace('/*drugList*/', drugList)
					target_out = target_out.replace('/*caseno*/', caseno)
					target_out = target_out.replace('/*manu*/', manu)
					target_out = target_out.replace('/*mtype*/', mtype)
					target_out = target_out.replace('/*SE*/', SE)
					target_out = target_out.replace('/*username*/', email)
					target_out = target_out.replace('/*sdaocc2*/', SDA_Occ2)
					target_out = target_out.replace('/*bdaocc2*/', occupation2)
					target_out = target_out.replace('/*drugsnocomma*/', drugsnocomma)
					target_out = target_out.replace('/*dispPeriod*/', displayPeriod)
					target_out = target_out.replace('/*suppApplied*/', 'Yes')
					target_out = target_out.replace('/*supp_text_file*/', suppFileLocation)
					target_out = target_out.replace('/*supp_Match_Type*/', str(config['suppMatchType']))
					target_out = target_out.replace('/*therapyClass*/', therapyClass)
					target_out = target_out.replace('/*bda_only*/', bDa_only)
					target_out = target_out.replace('/*sda_only*/', sDa_only)
					target_out = target_out.replace('/*createPivotTable*/', createPivotTable)
					target_out = target_out.replace('/*pivYes1*/', pivSeg1)
					target_out = target_out.replace('/*pivYes2*/', pivSeg2)
					target_out = target_out.replace('/*totalSDAS*/', finalSDATotal)
					target_out = target_out.replace('/*totalBDAS*/', finalBDATotal)
					new_file.write(target_out)
					line_file = new_file
				
			else:
				if bDa == 'y':  
					copyfile(emailCode, os.path.join(outCode, 'Presales Automation_Email_Final.sas'))
					newInput = os.path.join(outCode, 'Presales Automation_Email_Final.sas')

					line_file = open(os.path.join(newInput),'r').readlines()
					new_file = open(os.path.join(newInput),'w')
					for line_in in line_file:
						target_out = line_in.replace('/*DocORQuiz*/', listProduct)
						target_out = target_out.replace('/*listMatchType*/', listMatchType)
						target_out = target_out.replace('/*target_text_file*/', outCode)
						target_out = target_out.replace('/*Segments*/', segList)
						target_out = target_out.replace('/*Segments2*/', segList2)
						target_out = target_out.replace('/*Brand*/', brand)
						target_out = target_out.replace('/*MY_INIT*/', yourIn)
						target_out = target_out.replace('/*Requester_Initials*/', reIn)
						target_out = target_out.replace('/*SDA_Occ*/', SDA_Occ)
						target_out = target_out.replace('/*SDA_Spec*/', SDA_Spec)
						target_out = target_out.replace('/*yesORno*/', deDup)
						target_out = target_out.replace('/*LookUpPeriod*/', lookUpPeriod)
						target_out = target_out.replace('/*totalLoookUps*/', totalLookUps)
						target_out = target_out.replace('/*BDA_Occ*/', occupation)
						target_out = target_out.replace('/*BDA_Spec*/', specialty)
						target_out = target_out.replace('/*drugList*/', drugList)
						target_out = target_out.replace('/*caseno*/', caseno)
						target_out = target_out.replace('/*manu*/', manu)
						target_out = target_out.replace('/*mtype*/', mtype)
						target_out = target_out.replace('/*SE*/', SE)
						target_out = target_out.replace('/*username*/', email)
						target_out = target_out.replace('/*sdaocc2*/', SDA_Occ2)
						target_out = target_out.replace('/*bdaocc2*/', occupation2)
						target_out = target_out.replace('/*drugsnocomma*/', drugsnocomma)
						target_out = target_out.replace('/*dispPeriod*/', displayPeriod)
						target_out = target_out.replace('/*suppApplied*/', 'Yes')
						target_out = target_out.replace('/*supp_text_file*/', suppFileLocation)
						target_out = target_out.replace('/*supp_Match_Type*/', str(config['suppMatchType']))
						target_out = target_out.replace('/*therapyClass*/', therapyClass)
						target_out = target_out.replace('/*bda_only*/', bDa_only)
						target_out = target_out.replace('/*sda_only*/', sDa_only)
						target_out = target_out.replace('/*createPivotTable*/', createPivotTable)
						target_out = target_out.replace('/*pivYes1*/', pivSeg1)
						target_out = target_out.replace('/*pivYes2*/', pivSeg2)
						target_out = target_out.replace('/*totalSDAS*/', finalSDATotal)
						target_out = target_out.replace('/*totalBDAS*/', finalBDATotal)
						new_file.write(target_out)
						line_file = new_file
						
				if sDa == 'y':  
					copyfile(emailCode, os.path.join(outCode, 'Presales Automation_Email_Final.sas'))
					newInput = os.path.join(outCode, 'Presales Automation_Email_Final.sas')

					line_file = open(os.path.join(newInput),'r').readlines()
					new_file = open(os.path.join(newInput),'w')
					for line_in in line_file:
						target_out = line_in.replace('/*DocORQuiz*/', listProduct)
						target_out = target_out.replace('/*listMatchType*/', listMatchType)
						target_out = target_out.replace('/*target_text_file*/', outCode)
						target_out = target_out.replace('/*Segments*/', segList)
						target_out = target_out.replace('/*Segments2*/', segList2)
						target_out = target_out.replace('/*Brand*/', brand)
						target_out = target_out.replace('/*MY_INIT*/', yourIn)
						target_out = target_out.replace('/*Requester_Initials*/', reIn)
						target_out = target_out.replace('/*SDA_Occ*/', SDA_Occ)
						target_out = target_out.replace('/*SDA_Spec*/', SDA_Spec)
						target_out = target_out.replace('/*yesORno*/', deDup)
						target_out = target_out.replace('/*LookUpPeriod*/', lookUpPeriod)
						target_out = target_out.replace('/*totalLoookUps*/', totalLookUps)
						target_out = target_out.replace('/*BDA_Occ*/', occupation)
						target_out = target_out.replace('/*BDA_Spec*/', specialty)
						target_out = target_out.replace('/*drugList*/', drugList)
						target_out = target_out.replace('/*caseno*/', caseno)
						target_out = target_out.replace('/*manu*/', manu)
						target_out = target_out.replace('/*mtype*/', mtype)
						target_out = target_out.replace('/*SE*/', SE)
						target_out = target_out.replace('/*username*/', email)
						target_out = target_out.replace('/*sdaocc2*/', SDA_Occ2)
						target_out = target_out.replace('/*bdaocc2*/', occupation2)
						target_out = target_out.replace('/*drugsnocomma*/', drugsnocomma)
						target_out = target_out.replace('/*dispPeriod*/', '')
						target_out = target_out.replace('/*suppApplied*/', 'Yes')
						target_out = target_out.replace('/*supp_text_file*/', suppFileLocation)
						target_out = target_out.replace('/*supp_Match_Type*/', str(config['suppMatchType']))
						target_out = target_out.replace('/*therapyClass*/', therapyClass)
						target_out = target_out.replace('/*bda_only*/', bDa_only)
						target_out = target_out.replace('/*sda_only*/', sDa_only)
						target_out = target_out.replace('/*createPivotTable*/', createPivotTable)
						target_out = target_out.replace('/*pivYes1*/', pivSeg1)
						target_out = target_out.replace('/*pivYes2*/', pivSeg2)
						target_out = target_out.replace('/*totalSDAS*/', finalSDATotal)
						target_out = target_out.replace('/*totalBDAS*/', finalBDATotal)
						new_file.write(target_out)
						line_file = new_file

		elif caseType == 'listMatch' and listMatchType == 'None':
			copyfile(emailCode, os.path.join(outCode, 'Presales Automation_Email_Final.sas'))
			newInput = os.path.join(outCode, 'Presales Automation_Email_Final.sas')

			line_file = open(os.path.join(newInput),'r').readlines()
			new_file = open(os.path.join(newInput),'w')
			for line_in in line_file:
				#target_out = line_in.replace('/*DocORQuiz*/', listProduct).replace('/*listMatchType*/', listMatchType).replace('/*target_text_file*/', outCode).replace('/*Segments*/', '').replace('/*Segments2*/', '').replace('/*Brand*/', brand).replace('/*MY_INIT*/', yourIn).replace('/*Requester_Initials*/', reIn).replace('/*SDA_Occ*/', '').replace('/*SDA_Spec*/', '').replace('/*yesORno*/', '').replace('/*LookUpPeriod*/', '').replace('/*totalLoookUps*/', '').replace('/*BDA_Occ*/', '').replace('/*BDA_Spec*/', '').replace('/*drugList*/', '')
				target_out = line_in.replace('/*DocORQuiz*/', listProduct)
				target_out = target_out.replace('/*listMatchType*/', listMatchType)
				target_out = target_out.replace('/*target_text_file*/', outCode)
				target_out = target_out.replace('/*Segments*/', '')
				target_out = target_out.replace('/*Segments2*/', '')
				target_out = target_out.replace('/*Brand*/', brand)
				target_out = target_out.replace('/*MY_INIT*/', yourIn)
				target_out = target_out.replace('/*Requester_Initials*/', reIn)
				target_out = target_out.replace('/*SDA_Occ*/', SDA_Occ)
				target_out = target_out.replace('/*SDA_Spec*/', SDA_Spec)
				target_out = target_out.replace('/*therapyClass*/', therapyClass)
				target_out = target_out.replace('/*yesORno*/', deDup)
				target_out = target_out.replace('/*LookUpPeriod*/', lookUpPeriod)
				target_out = target_out.replace('/*totalLoookUps*/', totalLookUps)
				target_out = target_out.replace('/*BDA_Occ*/', occupation)
				target_out = target_out.replace('/*BDA_Spec*/', specialty)
				target_out = target_out.replace('/*drugList*/', ';')
				target_out = target_out.replace('/*caseno*/', caseno)
				target_out = target_out.replace('/*manu*/', manu)
				target_out = target_out.replace('/*mtype*/', mtype)
				target_out = target_out.replace('/*SE*/', SE)
				target_out = target_out.replace('/*username*/', email)
				target_out = target_out.replace('/*sdaocc2*/', SDA_Occ2)
				target_out = target_out.replace('/*bdaocc2*/', occupation2)
				target_out = target_out.replace('/*drugsnocomma*/', drugsnocomma)
				target_out = target_out.replace('/*dispPeriod*/', '')
				target_out = target_out.replace('/*suppApplied*/', 'Yes')
				target_out = target_out.replace('/*supp_text_file*/', suppFileLocation)
				target_out = target_out.replace('/*supp_Match_Type*/', str(config['suppMatchType']))
				target_out = target_out.replace('/*therapyClass*/', therapyClass)
				target_out = target_out.replace('/*bda_only*/', bDa_only)
				target_out = target_out.replace('/*sda_only*/', sDa_only)
				target_out = target_out.replace('/*createPivotTable*/', createPivotTable)
				target_out = target_out.replace('/*pivYes1*/', pivSeg1)
				target_out = target_out.replace('/*pivYes2*/', pivSeg2)
				target_out = target_out.replace('/*totalSDAS*/', finalSDATotal)
				target_out = target_out.replace('/*totalBDAS*/', finalBDATotal)
				new_file.write(target_out)
				line_file = new_file



		elif caseType == 'Targeting': 
			targetOut = """P:\\Epocrates Analytics\\TARGETS\\{date}\\{manu} {brand}""".format(date = date, manu = manu, brand = brand)
			outCode2 = """P:\\Epocrates Analytics\\TARGETS\\{date}\\{manu} {brand}{slashes}""".format(date = date, slashes = "\\", manu = manu, brand = brand)
			if dSharing == 'Y' and (listMatchType == 'Standard' or listMatchType == 'Standard_Seg' or listMatchType == 'Exact' or listMatchType == 'Exact_Seg'):
				segList = ', '.join(segmentList)
				copyfile(targetAuto, os.path.join(outCode2, 'Targeting Automation Code_OFFICIAL.sas'))
				newInput = os.path.join(outCode2, 'Targeting Automation Code_OFFICIAL.sas')

				line_file = open(os.path.join(newInput),'r').readlines()
				new_file = open(os.path.join(newInput),'w')
				for line_in in line_file:
					target_out = line_in.replace('/*DocORQuiz*/', listProduct)
					target_out = target_out.replace('/*listMatchType*/', listMatchType)
					target_out = target_out.replace('/*target_text_file*/', outCode3)
					target_out = target_out.replace('/*Date*/', date)
					target_out = target_out.replace('/*targetFoler*/', targetOut)
					target_out = target_out.replace('/*Brand*/', brand)
					target_out = target_out.replace('?MY_INIT?', yourIn)
					target_out = target_out.replace('/*TargetNum*/', targetNum)
					target_out = target_out.replace('/*segVar*/', segVariable)
					target_out = target_out.replace('/*segValues*/', varValues)
					target_out = target_out.replace('/*dataShareYorN*/', dSharing)
					target_out = target_out.replace('/*Segments*/', segList)
					target_out = target_out.replace('/*keep_seg*/', keep_seg)
					target_out = target_out.replace('/*Manu*/', manu)
					target_out = target_out.replace('/*SDAONLY*/', sDa_only)
					target_out = target_out.replace('/*SDA_Occ*/', SDA_Occ)
					target_out = target_out.replace('/*SDA_Spec*/', SDA_Spec)
					target_out = target_out.replace('/*SDATarget*/', SDA_Target)
					target_out = target_out.replace('/*BDAONLY*/', bDa_only)
					target_out = target_out.replace('/*yesORno*/', deDup)
					target_out = target_out.replace('/*LookUpPeriod*/', lookUpPeriod)
					target_out = target_out.replace('/*totalLoookUps*/', totalLookUps)
					target_out = target_out.replace('/*BDA_Occ*/', occupation)
					target_out = target_out.replace('/*BDA_Spec*/', specialty)
					target_out = target_out.replace('/*BDATarget*/', BDA_Target)
					target_out = target_out.replace('/*drugList*/', drugList)
					target_out = target_out.replace('/*suppApplied*/', 'Yes')
					target_out = target_out.replace('/*supp_text_file*/', suppFileLocation)
					target_out = target_out.replace('/*supp_Match_Type*/', str(config['suppMatchType']))
					target_out = target_out.replace('/*dataCap*/', dataCap)
					target_out = target_out.replace('/*bda_only*/', bDa_only)
					target_out = target_out.replace('/*sda_only*/', sDa_only)
					target_out = target_out.replace('/*backFill*/', backFill)
					target_out = target_out.replace('/*needVals*/', neededValues)
					target_out = target_out.replace('/*backFillVals*/', backFillValue)
					target_out = target_out.replace('/*suppSDAOnly*/', suppSDAOnly)
					target_out = target_out.replace('/*suppBDAOnly*/', suppBDAOnly)
					new_file.write(target_out)
					line_file = new_file

			if dSharing == 'N' and (listMatchType == 'Standard_Seg' or listMatchType == 'Exact_Seg'):
				segList = finalSeg
				copyfile(targetAuto, os.path.join(outCode2, 'Targeting Automation Code_OFFICIAL.sas'))
				newInput = os.path.join(outCode2, 'Targeting Automation Code_OFFICIAL.sas')

				line_file = open(os.path.join(newInput),'r').readlines()
				new_file = open(os.path.join(newInput),'w')
				for line_in in line_file:
					target_out = line_in.replace('/*DocORQuiz*/', listProduct)
					target_out = target_out.replace('/*listMatchType*/', listMatchType)
					target_out = target_out.replace('/*target_text_file*/', outCode3)
					target_out = target_out.replace('/*Date*/', date)
					target_out = target_out.replace('/*targetFoler*/', targetOut)
					target_out = target_out.replace('/*Brand*/', brand)
					target_out = target_out.replace('?MY_INIT?', yourIn)
					target_out = target_out.replace('/*TargetNum*/', targetNum)
					target_out = target_out.replace('/*segVar*/', segVariable)
					target_out = target_out.replace('/*segValues*/', varValues)
					target_out = target_out.replace('/*dataShareYorN*/', dSharing)
					target_out = target_out.replace('/*Segments*/', segList)
					target_out = target_out.replace('/*keep_seg*/', keep_seg)
					target_out = target_out.replace('/*Manu*/', manu)
					target_out = target_out.replace('/*SDAONLY*/', sDa_only)
					target_out = target_out.replace('/*SDA_Occ*/', SDA_Occ)
					target_out = target_out.replace('/*SDA_Spec*/', SDA_Spec)
					target_out = target_out.replace('/*SDATarget*/', SDA_Target)
					target_out = target_out.replace('/*BDAONLY*/', bDa_only)
					target_out = target_out.replace('/*yesORno*/', deDup)
					target_out = target_out.replace('/*LookUpPeriod*/', lookUpPeriod)
					target_out = target_out.replace('/*totalLoookUps*/', totalLookUps)
					target_out = target_out.replace('/*BDA_Occ*/', occupation)
					target_out = target_out.replace('/*BDA_Spec*/', specialty)
					target_out = target_out.replace('/*BDATarget*/', BDA_Target)
					target_out = target_out.replace('/*drugList*/', drugList)
					target_out = target_out.replace('/*suppApplied*/', 'Yes')
					target_out = target_out.replace('/*supp_text_file*/', suppFileLocation)
					target_out = target_out.replace('/*supp_Match_Type*/', str(config['suppMatchType']))
					target_out = target_out.replace('/*dataCap*/', dataCap)
					target_out = target_out.replace('/*bda_only*/', bDa_only)
					target_out = target_out.replace('/*sda_only*/', sDa_only)
					target_out = target_out.replace('/*backFill*/', backFill)
					target_out = target_out.replace('/*needVals*/', neededValues)
					target_out = target_out.replace('/*backFillVals*/', backFillValue)
					target_out = target_out.replace('/*suppSDAOnly*/', suppSDAOnly)
					target_out = target_out.replace('/*suppBDAOnly*/', suppBDAOnly)
					new_file.write(target_out)
					line_file = new_file
			if dSharing == 'N' and (listMatchType == 'Standard' or listMatchType == 'Exact') and (sDa_only == 'N' and bDa_only == 'N'):
				segList = str(finalSeg)
				print 'copying sas code'
				copyfile(targetAuto, os.path.join(outCode2, 'Targeting Automation Code_OFFICIAL.sas'))
				newInput = os.path.join(outCode2, 'Targeting Automation Code_OFFICIAL.sas')

				line_file = open(os.path.join(newInput),'r').readlines()
				new_file = open(os.path.join(newInput),'w')
				for line_in in line_file:
					target_out = line_in.replace('/*DocORQuiz*/', listProduct)
					target_out = target_out.replace('/*listMatchType*/', listMatchType)
					target_out = target_out.replace('/*target_text_file*/', outCode3)
					target_out = target_out.replace('/*Date*/', date)
					target_out = target_out.replace('/*targetFoler*/', targetOut)
					target_out = target_out.replace('/*Brand*/', brand)
					target_out = target_out.replace('?MY_INIT?', yourIn)
					target_out = target_out.replace('/*TargetNum*/', targetNum)
					target_out = target_out.replace('/*segVar*/', segVariable)
					target_out = target_out.replace('/*segValues*/', varValues)
					target_out = target_out.replace('/*dataShareYorN*/', dSharing)
					target_out = target_out.replace('/*Segments*/', segList)
					target_out = target_out.replace('/*keep_seg*/', keep_seg)
					target_out = target_out.replace('/*Manu*/', manu)
					target_out = target_out.replace('/*SDAONLY*/', sDa_only)
					target_out = target_out.replace('/*SDA_Occ*/', SDA_Occ)
					target_out = target_out.replace('/*SDA_Spec*/', SDA_Spec)
					target_out = target_out.replace('/*SDATarget*/', SDA_Target)
					target_out = target_out.replace('/*BDAONLY*/', bDa_only)
					target_out = target_out.replace('/*yesORno*/', deDup)
					target_out = target_out.replace('/*LookUpPeriod*/', lookUpPeriod)
					target_out = target_out.replace('/*totalLoookUps*/', totalLookUps)
					target_out = target_out.replace('/*BDA_Occ*/', occupation)
					target_out = target_out.replace('/*BDA_Spec*/', specialty)
					target_out = target_out.replace('/*BDATarget*/', BDA_Target)
					target_out = target_out.replace('/*drugList*/', drugList)
					target_out = target_out.replace('/*suppApplied*/', 'Yes')
					target_out = target_out.replace('/*supp_text_file*/', suppFileLocation)
					target_out = target_out.replace('/*supp_Match_Type*/', str(config['suppMatchType']))
					target_out = target_out.replace('/*dataCap*/', dataCap)
					target_out = target_out.replace('/*bda_only*/', bDa_only)
					target_out = target_out.replace('/*sda_only*/', sDa_only)
					target_out = target_out.replace('/*backFill*/', backFill)
					target_out = target_out.replace('/*needVals*/', neededValues)
					target_out = target_out.replace('/*backFillVals*/', backFillValue)
					target_out = target_out.replace('/*suppSDAOnly*/', suppSDAOnly)
					target_out = target_out.replace('/*suppBDAOnly*/', suppBDAOnly)
					new_file.write(target_out)
					line_file = new_file

			if sDa_only == 'Y':
				# segList = ', '.join(segmentList)
				copyfile(targetAuto, os.path.join(outCode2, 'Targeting Automation Code_OFFICIAL.sas'))
				newInput = os.path.join(outCode2, 'Targeting Automation Code_OFFICIAL.sas')

				line_file = open(os.path.join(newInput),'r').readlines()
				new_file = open(os.path.join(newInput),'w')
				for line_in in line_file:
					target_out = line_in.replace('/*DocORQuiz*/', listProduct)
					target_out = target_out.replace('/*listMatchType*/', 'None')
					target_out = target_out.replace('/*target_text_file*/', '')
					target_out = target_out.replace('/*Date*/', date)
					target_out = target_out.replace('/*targetFoler*/', targetOut)
					target_out = target_out.replace('/*Brand*/', brand)
					target_out = target_out.replace('?MY_INIT?', yourIn)
					target_out = target_out.replace('/*TargetNum*/', '')
					target_out = target_out.replace('/*segVar*/', '')
					target_out = target_out.replace('/*segValues*/', '')
					target_out = target_out.replace('/*dataShareYorN*/', dSharing)
					target_out = target_out.replace('/*Segments*/', '')
					target_out = target_out.replace('/*keep_seg*/', keep_seg)
					target_out = target_out.replace('/*Manu*/', manu)
					target_out = target_out.replace('/*SDAONLY*/', sDa_only)
					target_out = target_out.replace('/*SDA_Occ*/', SDA_Occ)
					target_out = target_out.replace('/*SDA_Spec*/', SDA_Spec)
					target_out = target_out.replace('/*SDATarget*/', SDA_Target)
					target_out = target_out.replace('/*BDAONLY*/', bDa_only)
					target_out = target_out.replace('/*yesORno*/', deDup)
					target_out = target_out.replace('/*LookUpPeriod*/', lookUpPeriod)
					target_out = target_out.replace('/*totalLoookUps*/', totalLookUps)
					target_out = target_out.replace('/*BDA_Occ*/', occupation)
					target_out = target_out.replace('/*BDA_Spec*/', specialty)
					target_out = target_out.replace('/*BDATarget*/', BDA_Target)
					target_out = target_out.replace('/*drugList*/', drugList)
					target_out = target_out.replace('/*suppApplied*/', 'Yes')
					target_out = target_out.replace('/*supp_text_file*/', suppFileLocation)
					target_out = target_out.replace('/*supp_Match_Type*/', str(config['suppMatchType']))
					target_out = target_out.replace('/*dataCap*/', dataCap)
					target_out = target_out.replace('/*bda_only*/', bDa_only)
					target_out = target_out.replace('/*sda_only*/', sDa_only)
					target_out = target_out.replace('/*backFill*/', backFill)
					target_out = target_out.replace('/*needVals*/', neededValues)
					target_out = target_out.replace('/*backFillVals*/', backFillValue)
					target_out = target_out.replace('/*suppSDAOnly*/', suppSDAOnly)
					target_out = target_out.replace('/*suppBDAOnly*/', suppBDAOnly)
					new_file.write(target_out)
					line_file = new_file

			if bDa_only == 'Y':
				# segList = ', '.join(segmentList)
				copyfile(targetAuto, os.path.join(outCode2, 'Targeting Automation Code_OFFICIAL.sas'))
				newInput = os.path.join(outCode2, 'Targeting Automation Code_OFFICIAL.sas')

				line_file = open(os.path.join(newInput),'r').readlines()
				new_file = open(os.path.join(newInput),'w')
				for line_in in line_file:
					target_out = line_in.replace('/*DocORQuiz*/', listProduct)
					target_out = target_out.replace('/*listMatchType*/', 'None')
					target_out = target_out.replace('/*target_text_file*/', '')
					target_out = target_out.replace('/*Date*/', date)
					target_out = target_out.replace('/*targetFoler*/', targetOut)
					target_out = target_out.replace('/*Brand*/', brand)
					target_out = target_out.replace('?MY_INIT?', yourIn)
					target_out = target_out.replace('/*TargetNum*/', '')
					target_out = target_out.replace('/*segVar*/', '')
					target_out = target_out.replace('/*segValues*/', '')
					target_out = target_out.replace('/*dataShareYorN*/', dSharing)
					target_out = target_out.replace('/*Segments*/', '')
					target_out = target_out.replace('/*keep_seg*/', keep_seg)
					target_out = target_out.replace('/*Manu*/', manu)
					target_out = target_out.replace('/*SDAONLY*/', sDa_only)
					target_out = target_out.replace('/*SDA_Occ*/', SDA_Occ)
					target_out = target_out.replace('/*SDA_Spec*/', SDA_Spec)
					target_out = target_out.replace('/*SDATarget*/', SDA_Target)
					target_out = target_out.replace('/*BDAONLY*/', bDa_only)
					target_out = target_out.replace('/*yesORno*/', deDup)
					target_out = target_out.replace('/*LookUpPeriod*/', lookUpPeriod)
					target_out = target_out.replace('/*totalLoookUps*/', totalLookUps)
					target_out = target_out.replace('/*BDA_Occ*/', occupation)
					target_out = target_out.replace('/*BDA_Spec*/', specialty)
					target_out = target_out.replace('/*BDATarget*/', BDA_Target)
					target_out = target_out.replace('/*drugList*/', drugList)
					target_out = target_out.replace('/*suppApplied*/', 'Yes')
					target_out = target_out.replace('/*supp_text_file*/', suppFileLocation)
					target_out = target_out.replace('/*supp_Match_Type*/', str(config['suppMatchType']))
					target_out = target_out.replace('/*dataCap*/', dataCap)
					target_out = target_out.replace('/*bda_only*/', bDa_only)
					target_out = target_out.replace('/*sda_only*/', sDa_only)
					target_out = target_out.replace('/*backFill*/', backFill)
					target_out = target_out.replace('/*needVals*/', neededValues)
					target_out = target_out.replace('/*backFillVals*/', backFillValue)
					target_out = target_out.replace('/*suppSDAOnly*/', suppSDAOnly)
					target_out = target_out.replace('/*suppBDAOnly*/', suppBDAOnly)
					new_file.write(target_out)
					line_file = new_file

def buildSDAPreSalesMacro():
	totalIncludesBuilt = 1
	totalAdditionalSDAs = int(config['totalAdditionalSDAs'])
	sdaMacro = """%macro multipleSDAAdd_Ons;
%do i=1 %to 1;
	%if &totalSDAS. > 1 %then %do;\n"""
	macroEnd = """
	%end;
%end;
%mend;
%multipleSDAAdd_Ons;"""
	while totalIncludesBuilt <= totalAdditionalSDAs:
		includePath = '		%include "&filepath.\PS_SDA_plus_CL_Email_'+str(totalIncludesBuilt)+'.sas";'
		totalIncludesBuilt +=1
		includePath = '{}\n'.format(includePath)
		# totalIncludesBuilt +=1
		sdaMacro +=includePath
		# totalIncludesBuilt +=1
	finalSDAMacro = sdaMacro + macroEnd
	# print finalSDAMacro

	newInput = os.path.join(outCode, 'Presales Automation_Email_Final.sas')
	line_file = open(os.path.join(newInput),'r').readlines()
	new_file = open(os.path.join(newInput),'w')
	for line_in in line_file:
		target_out = line_in.replace('/*multiSDAMacro*/', finalSDAMacro)

		new_file.write(target_out)
		line_file = new_file	

def buildSDACodes():
	print 'Building Additional SDA Codes. . . '
	sdaCodeHousing = 'P:\\Epocrates Analytics\\Code_Library\\Standard_Codes\\Pre Sales\\DocAlert_Python_Reference\\CUSTOM\\Email Codes\\additionalSDA'
	sdaCode = 'PS_SDA_plus_CL_Email'
	suppApplied = str(config['suppressionApplied'])
	
	# if suppApplied == 'Y':
	# 	suppFolder = config['suppSASFile']
	# 	onlyFilesSupp = [f for f in listdir(suppFolder) if isfile(join(suppFolder, f))]

	# 	for files in sorted(set(onlyFilesSupp)):
	# 		file = files.split('.')[0]
	# 		if re.search('_supp_.+', file) or re.search('_matched_.+', file):
	# 			suppFile = file
	# else:
	# 	suppFolder = ''
	# 	suppFile = ''


	totalCodesBuilt = 1
	totalAdditionalSDAs = int(config['totalAdditionalSDAs'])
	while totalCodesBuilt <= totalAdditionalSDAs:
		# occupation = str(config['SDA_Occ'])
		occupation2 = str(config['additonalSdaOcc_'+str(totalCodesBuilt)]).replace('"', '')
		# specialty = str(config['SDA_Spec'])
		specialty2 = str(config['additonalSdaSpec_'+str(totalCodesBuilt)]).replace('"', '')


		copiedSDAFile = os.path.join(outCode, sdaCode+'_'+str(totalCodesBuilt)+'.sas')
		copyfile(os.path.join(sdaCodeHousing, sdaCode+'.sas'), copiedSDAFile)
		
		newInput = copiedSDAFile
		line_file = open(os.path.join(newInput),'r').readlines()
		new_file = open(os.path.join(newInput),'w')
		for line_in in line_file:
			# target_out = line_in.replace('/*folder*/', listMatchFolder)
			# target_out = target_out.replace('/*matchedFile*/', matchedFile)
			target_out = line_in.replace('/*suppApplied*/', suppApplied)
			# target_out = target_out.replace('/*suppFolder*/', suppFolder)
			# target_out = target_out.replace('/*suppFile*/', suppFile)
			target_out = target_out.replace('/*SDA_Occ*/', config['additonalSdaOcc_'+str(totalCodesBuilt)])
			target_out = target_out.replace('/*SDA_Occ_Disp*/', occupation2)
			target_out = target_out.replace('/*SDA_Spec*/', config['additonalSdaSpec_'+str(totalCodesBuilt)])
			target_out = target_out.replace('/*SDA_Spec_Disp*/', specialty2)			
			target_out = target_out.replace('/*username*/', email)
			target_out = target_out.replace('/*inc*/', '_'+str(totalCodesBuilt))
			new_file.write(target_out)
			line_file = new_file

		totalCodesBuilt +=1


def buildBDAPreSalesMacro():
	totalIncludesBuilt = 1
	totalAdditionalBDAs = int(config['totalAdditionalBDAs'])
	bdaMacro = """%macro multipleBDAAdd_Ons;
%do i=1 %to 1;
	%if &totalBDAS. > 1 %then %do;\n"""
	macroEnd = """
	%end;
%end;
%mend;
%multipleBDAAdd_Ons;"""
	while totalIncludesBuilt <= totalAdditionalBDAs:
		includePath = '		%include "&filepath.\PS_BDA_Mult_Lookup_plus_CL_Email_'+str(totalIncludesBuilt)+'.sas";'
		totalIncludesBuilt +=1
		includePath = '{}\n'.format(includePath)
		# totalIncludesBuilt +=1
		bdaMacro +=includePath
		# totalIncludesBuilt +=1
	finalBDAMacro = bdaMacro + macroEnd
	# print finalSDAMacro

	newInput = os.path.join(outCode, 'Presales Automation_Email_Final.sas')
	line_file = open(os.path.join(newInput),'r').readlines()
	new_file = open(os.path.join(newInput),'w')
	for line_in in line_file:
		target_out = line_in.replace('/*multiBDAMacro*/', finalBDAMacro)

		new_file.write(target_out)
		line_file = new_file

def buildBDACodes():
	print 'Building Additional BDA Codes. . . '
	bdaCodeHousing = 'P:\\Epocrates Analytics\\Code_Library\\Standard_Codes\\Pre Sales\\DocAlert_Python_Reference\\CUSTOM\\Email Codes\\additionalBDA'
	bdaCode = 'PS_BDA_Mult_Lookup_plus_CL_Email'
	suppApplied = str(config['suppressionApplied'])
	
	# if suppApplied == 'Y':
	# 	suppFolder = config['suppSASFile']
	# 	onlyFilesSupp = [f for f in listdir(suppFolder) if isfile(join(suppFolder, f))]

	# 	for files in sorted(set(onlyFilesSupp)):
	# 		file = files.split('.')[0]
	# 		if re.search('_supp_.+', file) or re.search('_matched_.+', file):
	# 			suppFile = file
	# else:
	# 	suppFolder = ''
	# 	suppFile = ''
	finalDrugs2 = []
	unmatchedDrugs2 = []
	totalCodesBuilt = 1
	totalAdditionalBDAs = int(config['totalAdditionalBDAs'])
	while totalCodesBuilt <= totalAdditionalBDAs:

		occupation2 = str(config['additonalBdaOcc_'+str(totalCodesBuilt)]).replace('"', '')
		specialty2 = str(config['additonalBdaSpec_'+str(totalCodesBuilt)]).replace('"', '')
		therapyClass = str(config['additonalBdatherapyChecked_'+str(totalCodesBuilt)])
		lookUpPeriod = str(config['additonalBdaLookUpPeriod_'+str(totalCodesBuilt)])
		totalLookUps = str(config['additonalBdaLookUps_'+str(totalCodesBuilt)])
		displayPeriod = str(int(lookUpPeriod)-1)
		dedupe = str(config['additonalBdaDedup_'+str(totalCodesBuilt)])

		drugList = str(config['additonalBdaDrugList_'+str(totalCodesBuilt)])
		drugsnocomma = str(config['additonalBdaDrugList_'+str(totalCodesBuilt)]).replace("\n", ", ").replace("'", '')

		print 'Checking Misspelled Drugs for Additional BDA '+str(totalCodesBuilt)
		# checkDrugs2()
		with open(masterDrugs, 'rb') as myDrugs:
			reader = csv.DictReader(myDrugs)
			for row in reader:
				finalDrugs2.append(row['drugs'].strip())
				
		for inputDrugs in drugList.replace(', ', '\n').split("\n"):
			if inputDrugs.strip() not in finalDrugs2:
				unmatchedDrugs2.append(inputDrugs)
						
		print colored('THESE DRUGS ARE SPELLED WRONG OR MISSING: ', 'yellow'), colored(unmatchedDrugs2, 'yellow')
		print '----------------------------------------------------------------------------------------'
		for drug in unmatchedDrugs2:
			print 'Possible Correct Spelling: ', colored(process.extract(drug, finalDrugs2, limit=2), 'green'), colored(' - ', 'red'), colored(drug, 'red')
			print '----------------------------------------------------------------------------------------'

		finalDrugs2 = []
		unmatchedDrugs2 = []

		copiedBDAFile = os.path.join(outCode, bdaCode+'_'+str(totalCodesBuilt)+'.sas')
		copyfile(os.path.join(bdaCodeHousing, bdaCode+'.sas'), copiedBDAFile)
		
		newInput = copiedBDAFile
		line_file = open(os.path.join(newInput),'r').readlines()
		new_file = open(os.path.join(newInput),'w')
		for line_in in line_file:
			target_out = line_in.replace('/*suppApplied*/', suppApplied)
			# target_out = target_out.replace('/*suppFolder*/', suppFolder)
			# target_out = target_out.replace('/*suppFile*/', suppFile)
			target_out = target_out.replace('/*therapyClass*/', therapyClass)
			target_out = target_out.replace('/*LookUpPeriod*/', lookUpPeriod)
			target_out = target_out.replace('/*totalLookUps*/', totalLookUps)
			target_out = target_out.replace('/*BDA_Occ*/', config['additonalBdaOcc_'+str(totalCodesBuilt)])
			target_out = target_out.replace('/*BDA_Occ_Disp*/', occupation2)
			target_out = target_out.replace('/*BDA_Spec*/', config['additonalBdaSpec_'+str(totalCodesBuilt)])
			target_out = target_out.replace('/*BDA_Spec_Disp*/', specialty2)
			target_out = target_out.replace('/*dispPeriod*/', displayPeriod)
			target_out = target_out.replace('/*drugList2*/', drugsnocomma)
			target_out = target_out.replace('/*drugList*/', drugList)				
			target_out = target_out.replace('/*username*/', email)
			target_out = target_out.replace('/*inc*/', '_'+str(totalCodesBuilt))
			target_out = target_out.replace('/*Yes_OR_No*/', dedupe)
			new_file.write(target_out)
			line_file = new_file

		totalCodesBuilt +=1

def checkDrugs2():
	finalDrugs2 = []
	unmatchedDrugs2 = []

	with open(masterDrugs, 'rb') as myDrugs:
		reader = csv.DictReader(myDrugs)
		for row in reader:
			finalDrugs2.append(row['drugs'].strip())
			
	for inputDrugs in drugList.replace(', ', '\n').split("\n"):
		if inputDrugs.strip() not in finalDrugs2:
			unmatchedDrugs2.append(inputDrugs)
					
			print colored('THESE DRUGS ARE SPELLED WRONG OR MISSING: ', 'yellow'), colored(unmatchedDrugs2, 'yellow')
			print '----------------------------------------------------------------------------------------'
			for drug in unmatchedDrugs2:
				print 'Possible Correct Spelling: ', colored(process.extract(drug, finalDrugs2, limit=2), 'green'), colored(' - ', 'red'), colored(drug, 'red')
				print '----------------------------------------------------------------------------------------'

	finalDrugs2 = []
	unmatchedDrugs2 = []	

if (caseType == 'listMatch' or caseType == 'Targeting') and listMatchType != 'None':
	get_cols_names()
	createFolders()
	getMain()
	if 'cmi_compass_client' in config:
		if config['cmi_compass_client'] == 'Y':
			cmiCompasColumns()
	postgresConn()
	checkDrugs()
	fixSas()
	copyTarget()
	removeFiles()
if int(config['totalAdditionalSDAs']) != 0:
	buildSDACodes()
	buildSDAPreSalesMacro()
if int(config['totalAdditionalBDAs']) != 0:
	buildBDACodes()
	buildBDAPreSalesMacro()
if (caseType == 'listMatch' or caseType == 'Targeting') and listMatchType == 'None':
	createFolders()
	if bDa_only == 'Y':
		checkDrugs()
	fixSas()

print ''
print colored('P', 'cyan')+colored('R', 'red')+colored('O', 'green')+colored('G', 'yellow')+colored('R', 'blue')+colored('A', 'magenta')+colored('M', 'cyan')+' '+colored('C', 'magenta')+colored('O', 'red')+colored('M', 'green')+colored('P', 'blue')+colored('L', 'red')+colored('E', 'cyan')+colored('T', 'yellow')+colored('E', 'white')
# print 'PROGRAM COMPLETE!'