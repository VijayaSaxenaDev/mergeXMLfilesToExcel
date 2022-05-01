from xml.etree import cElementTree as ET
import xml.dom.minidom as minidom
from tkinter import filedialog
import xlsxwriter
import os
import mimetypes
import datetime
from io import BytesIO
row = 0
col = 0
listDir =[]
xmlBool = "false"
def xmlToExcel(f, writer, content):
	global row
	global col
	try:
		tree = minidom.parse(f)
		itemlist = tree.getElementsByTagName('Event') 
		for s in itemlist :
			col = 0
			if (s.attributes["Type"].value == "video"):
				row += 1
				for item in content :
					try:
						worksheet.write(row, col, s.attributes[item].value)
						col += 1
					except:
						continue
	except Exception as e:
		print("Error occured in " , f, " : ", e)

def processFile(directory, filename, writer, content):
	global xmlBool
	f = os.path.join(directory, filename)
	#print(f,os.path.isfile(f), mimetypes.guess_type(filename)[0])
	# checking if it is a file
	if (os.path.isfile(f) and mimetypes.guess_type(filename)[0] is not None and (mimetypes.guess_type(filename)[0] == "text/xml")):
		xmlToExcel(f, writer, content)
		print(f,os.path.isfile(f))				
		xmlBool = "true"

def loopD(directory, writer, content):
# iterate over files in
# that directory
	global xmlBool
	global listDir
	i=0
	for root, dirs, files in os.walk(directory):
		for filename in files:
			processFile(directory, filename, writer, content)
		for dir in dirs:
			for root2, dirs2, files2 in os.walk(os.path.join(root, dir)):
				for filename2 in files2:
					processFile(os.path.join(root, dir), filename2, writer, content)
try:
	directory = filedialog.askdirectory()
	workbook = xlsxwriter.Workbook(directory +'/VideoPlayFrequency_new.xlsx')
	worksheet = workbook.add_worksheet()
	content = ["Title","Type","Start","Transition","TransitionTime","EventDuration","InPoint","MediaDuration","KeepDuration", "Looping","Path"]
	for item in content :
		# write operation perform
		worksheet.write(row, col, item)
		col += 1
	loopD(directory, worksheet, content)
	if (xmlBool == "true"):
		print("Writing to file...", directory +'/VideoPlayFrequency_new.xlsx')
	else:
		print("No xml is present in the folder")
	key_pressed = input('Press ENTER to continue: ')
	workbook.close()
except Exception as e:
	print("Error occured: ", e)
	key_pressed = input('Press ENTER to continue: ')