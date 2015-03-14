#Programmer: Brent E. Chambers (q0m)
#Date: 2/21/2014
#Description: This program is a part of the DEFTRN training module.  


import xlrd
import sets

class PenInstance:
	tArray = []
	sArray = []
	dArray = []
	CurrentModeData = (tArray, sArray, dArray)
	book = xlrd.open_workbook(".\ForensicsResource.xls")
	secMode = book.sheet_names() #sheet array
	sheetCount = len(book.sheet_names())
	
	def listDisciplines(self):
		secMode = self.book.sheet_names()
		#for item in secMode:
		#	print item
			
	def loadCurrentMode(self, mode):
		sheet = self.book.sheet_by_name(mode) #set the sheet to the supplied mode
		return sheet
		
	def populateData(self, sheet):
		techniqueIndex = sheet.col_values(0)		#creates the first column data
		syntaxIndex = sheet.col_values(1)			#creates the second column data
		descriptionIndex = sheet.col_values(2)		#creates the third column daata
		self.CurrentModeData = [self.tArray, self.sArray, self.dArray]
		#print CurrentModeData
		for item in techniqueIndex:
			self.CurrentModeData[0].append(item)
		for item in syntaxIndex:
			self.CurrentModeData[1].append(item)
		for item in descriptionIndex:
			self.CurrentModeData[2].append(item)
				
	def createMasterList(self):
		testy = []
		for item in self.secMode:
			testy.append(item)
		for item in testy:
			sheet = self.loadCurrentMode(item)
			self.populateData(sheet)
			masterList = self.CurrentModeData
		return masterList

	def dump_techniques(self):	
		masterList = self.createMasterList()
		all_technique_items = []
		for item in masterList[0]:
			all_technique_items.append(item)
		all_technique_items = sets.Set(all_technique_items)
		return all_technique_items
	
	def dump_syntax(self):
		masterList = self.createMasterList()
		all_syntax_items = []
		for item in masterList[1]:
			all_syntax_items.append(item)
		all_syntax_items = sets.Set(all_syntax_items)
		return all_syntax_items
	
	def dump_descriptions(self):
		masterList = self.createMasterList()
		all_description_items = []
		for item in masterList[2]:
			all_description_items.append(item)
		all_description_items = sets.Set(all_description_items)
		return all_description_items
		
	def newsearch(self, query):

		masterResults = []
		searchResults = []
		masterList = self.createMasterList()
		techniqueSearchResults = []
		techniqueSearchResults = [t for t in self.dump_techniques() if query in t]
		techniqueSearchResults = sets.Set(techniqueSearchResults)
		#First searches the technique Array for the search query
		for technique in techniqueSearchResults:
			technique_ref_Number = masterList[0].index(technique)
			searchResults.append([technique, 
								masterList[1][technique_ref_Number],
								masterList[2][technique_ref_Number]])
				
		#Then we check the syntax, looking for the tool mayhaps?
		syntaxSearchResults = []
		syntaxSearchResults = [s for s in self.dump_syntax() if query in s]
		syntaxSearchResults = sets.Set(syntaxSearchResults)
		for syntax in syntaxSearchResults:
			syntax_ref_Number = masterList[1].index(syntax)
			searchResults.append([masterList[0][syntax_ref_Number],
								syntax,
								masterList[2][syntax_ref_Number]])
		#Then look throught he description for stuff and roll with it -- ya heard?
		descrptionSearchResults = []
		descriptionSearchResults = [d for d in self.dump_descriptions() if query in d]
		descriptionSearchResults = sets.Set(descriptionSearchResults)
		for description in descriptionSearchResults:
			description_ref_Number = masterList[2].index(description)
			searchResults.append([masterList[0][description_ref_Number],
								masterList[1][description_ref_Number],
								description])
		
		#print "\n\nDEFTRN Returned", len(searchResults), "results.\n\n"
		return searchResults
		
		
