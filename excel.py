# from xlutils.copy import copy
import xlrd
import xlwt

class ReadObject(object):
	rows = 0
	cols = 0
	lineNumber = 0
	book = 0
	sheet = 0
	def __init__(self, bookName, sheetName, lineNumber):
		self.book = xlrd.open_workbook(bookName,  on_demand = True)
		if(isinstance(sheetName, int)):
			self.sheet = self.book.sheet_by_index(sheetName)
		else:
			self.sheet = self.book.sheet_by_name(sheetName)
		self.rows = self.sheet.nrows
		self.cols = self.sheet.ncols
		self.lineNumber = lineNumber

class WriteObject(object):
	rows = 0
	cols = 0
	lineNumber = 0
	book = 0
	sheet = 0
	def __init__(self, worksheet):
		self.book = xlwt.Workbook()
		self.sheet = self.book.add_sheet(worksheet)
		# self.rows = self.sheet.nrows
		# self.cols = self.sheet.ncols
		# self.lineNumber = lineNumber

# Saves s as name
# s is a WriteObj
# (ReadObjs can't be saved)
def save(s, name):
	s.book.save(name)

def readCell(readObj, col):
	return((readObj.sheet.cell(readObj.lineNumber, col)).value)

def writeCell(writeObj, col, val):
	return(writeObj.sheet.write(writeObj.lineNumber, \
			col, val))

def write_row(readObj, writeObj):
	read_sheet = readObj.sheet
	write_sheet = writeObj.sheet
	for j in range(0, readObj.cols):
		writeCell(writeObj, j, readCell(readObj, j))
	readObj.lineNumber += 1
	writeObj.lineNumber += 1

def ordToNum(c):
	return(ord(c)-65)


# col is column letter in Upper case character
# rem is a bool. True => remove the following from column col
# False implies keep these...
def filter(readObj, writeObj, col, rem, *args):
	length = len(args)
	j = 0
	boolIn = False
	for i in range(readObj.lineNumber, readObj.rows):
		boolIn = False
		rowIValue = readCell(readObj, ordToNum(col))
		for j in range(length):
			if(args[j] in rowIValue):
				boolIn = True
				break
		if(not rem):
			if(boolIn):
				write_row(readObj, writeObj)
			else:
				readObj.lineNumber +=1
		if(rem):
			if(boolIn):
				readObj.lineNumber +=1
			else:
				write_row(readObj, writeObj)


def changeLine(Obj, lineNumber):
	Obj.lineNumber = lineNumber


### CHANGE ###
def group(readObj, writeObj, strings):
	for x in strings:
		changeLine(readObj, 1)
		for i in range(readObj.lineNumber, readObj.rows):
			if x == readCell(readObj, ordToNum('I')):
				print "True for " + x
				write_row(readObj, writeObj)
			else:
				readObj.lineNumber += 1


def main():
	readObj = ReadObject('monitor_iscala_psasia_05_17_2016.xls', 0, 0)
	writeObj = WriteObject('Sheet1')
	write_row(readObj, writeObj)
	filter(readObj, writeObj, 'H', True, "Invalid object name", "FIRST in N type MEM not found")
	save(writeObj, 'New.xls')
	# for i in range(0,7):
	# 	readObj = ReadObject('monitor_iscala_psasia_05_17_2016.xls', i, 0)
	# 	print("DOing " + str(i) + " now")
	# 	writeObj = WriteObject(strings[i])

	# 	#Writing the first row...
	# 	write_row(readObj, writeObj)
	# 	filter(readObj, writeObj)
	# # 	# st = [] 
	# # 	# st = filter(readObj, writeObj)

	# 	save(writeObj, str(strings[i]) + '.xls')
	# # 	# newR = ReadObject('New.xls', 'Sheet1', 0)
	# # 	# newW = WriteObject('Sheet1')
	# # 	# write_row(newR, newW)
	# 	# # group(newR, newW, st)
	# 	# save(newW, 'Grouped.xls')

if __name__== "__main__":
	main()
