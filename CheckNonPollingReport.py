import openpyxl
import os
import glob
import re

list_site_id_with_case = list()
#list_case_id_with_case = list()

list_site_id_daily_poll = list() 
#list_case_id_daily_poll = list() 

list_sites_missing = list()

patternSiteID = '\d\d\d\d\d\d'
patternCaseNumber = '\d\d\d\d\d\d\d'



def _get_current_site_ID_with_open_case():

	os.chdir(os.curdir)

	for file in glob.glob('*.log'): #get all log files, save directory //glob.glob returns an array  
		fileName = file

	with open(fileName, "r") as currentSites: #Open text file containing all prior days openend cases
		print("Opened LOG FILE")
		line = currentSites.readline()
		

		for line in currentSites:
			matchObject = re.match(patternSiteID, line) #Find the BR store #
			if matchObject:
				list_site_id_with_case.append(int(matchObject.group(0)))
				matchObject = None
	print("Parsed Pre-Existing Sites listed in file {} \n".format(fileName))


def _get_updated_site_ID_from_daily_report(): #Open Daily Excel File and pull BR Store #
	
	for file in glob.glob('*.xlsx'): #Get excel file and save file location
		fileName = file

	currentWorkbook = openpyxl.load_workbook(fileName) #Open excel doc
	sheetList = currentWorkbook.get_sheet_names() #get list of sheets
	currentSheet = currentWorkbook.get_sheet_by_name(sheetList[0]) #Set active sheet to first in list

	columnA = currentSheet['A'] #Return all values in active sheet's column A
	
	for column in columnA:

		columnValue = column.value
				
		if columnValue is not None:
			try:
				columnValue = int(columnValue) #Excel Doc is reading in with trailing zeros. This removes those zeros.
	
				if columnValue is not None:
					matchObject = re.match(patternSiteID, str(columnValue)) #Find the BR store number, convert to string for regular expression scanning
					
					if matchObject:
						list_site_id_daily_poll.append(columnValue) #Match found, save it.
						#print("Found Match: {}".format(columnValue)
	
			except ValueError:
				print("Excel: '{}'  Failed to convert Type {}".format(columnValue, type(columnValue)))
		
		

	print("Parsed Updates Sites listed in Excel file {} \n".format(fileName))
	
def _compare_current_and_daily_lists(): #Compare lists to find BR Store number missing from list. These can be closed.
	print(len(list_site_id_with_case))
	print(len(list_site_id_daily_poll))
	if(len(list_site_id_with_case) >= len(list_site_id_daily_poll)):
		#check new daily excel list against pre-existing list
		for i in range(len(list_site_id_with_case), 0, -1):
			#print("Type1: {}       Type2: {}".format(type(list_site_id_with_case[i-1]), type(list_site_id_daily_poll[i-2])))
			#print("Type1: {}              Type2: {}".format(list_site_id_with_case[i-1], list_site_id_daily_poll[i-2]))
			if list_site_id_with_case[i-1] not in list_site_id_daily_poll:
				list_sites_missing.append(list_site_id_with_case[i-1])
	else:
		#check pre-existing list against daily excel
		for i in range(len(list_site_id_daily_poll), 0, -1):
			#print("Type1: {}     Type2: {}".format(type(list_site_id_with_case[i-(len(list_site_id_daily_poll)-1)]), type(list_site_id_daily_poll[i-1])))
			#print("Type1: {}            Type2: {}".format(list_site_id_with_case[i-(len(list_site_id_daily_poll)-1)], list_site_id_daily_poll[i-1]))
			if int(list_site_id_daily_poll[i-1]) not in list_site_id_with_case:
				list_sites_missing.append(list_site_id_daily_poll[i-1])


if __name__ == "__main__":
	_get_current_site_ID_with_open_case()
	_get_updated_site_ID_from_daily_report()
	_compare_current_and_daily_lists()
	print(list_sites_missing)

