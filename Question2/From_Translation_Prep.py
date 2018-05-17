import openpyxl
import os

#Load the path that the translated file is located in
sourcePathStem = 'translation/translated/'
#Loads the path that you want the finished file to be put in
targetPathStem = 'translation/for-delivery/'

docs = os.listdir(sourcePathStem)

for doc in docs:
	sourcePath = sourcePathStem + doc
	targetPath = targetPathStem + doc
	#Opens up workbook in stated source path
	wb = openpyxl.load_workbook(sourcePath)
	#Collects all sheet names in the workbook
	ws = wb.get_sheet_names()
	for count in ws:
	    sheet = wb.sheetnames(count)
	    #This list represents the columns in the source file that need to be excluded from translation (This file needs to be updated to reflect new columns if source changes)
	    listA = ['A', 'B', 'D', 'G', 'H']
	    lengthOfList = len(listA)
	    i = 0
	    #Program runs as many times as there are columns that are hidden
	    while lengthOfList != 0:
	        #UnHides columns one at a time taking data from the list as input
	        sheet.column_dimensions[listA[i]].hidden = False
	        #After columns are unhidden their default width is set to 0.00. This resizes the widtch to 13.71 which is the column with on the source doc
	        sheet.column_dimensions[listA[i]].width = 13.71
	        i += 1
	        lengthOfList -= 1
	#Saves the finished translated file with original formatting into target path location
	wb.save(targetPath)
