import pandas as pd
import glob
import pathlib
from pathlib import Path
import csv

text_files_folder = "TextFiles"
input_excel = "input.xlsx"
input_excel_sheetname = "Sheet1"
input_excel_key_column = "Key"
text_file_delimeter = "="
output_file_name = "output.csv"

def getTextFilenamesFromFolder(folder_name):
	files = pathlib.Path(folder_name).glob("*.txt")
	filenames = []
	for file in files:
		filenames.append(file.stem)
	return filenames

text_files = getTextFilenamesFromFolder(text_files_folder)

print("Text files in folder are:", text_files)

def readKeysFromInputExcel(filename,sheet_name,key_col_name):
	data_frame = pd.read_excel(filename,sheet_name,index_col=None)
	keys = data_frame[key_col_name].tolist()
	return keys

keys = readKeysFromInputExcel(input_excel,input_excel_sheetname,input_excel_key_column)

print("Keys in input file:", keys)

def combineKeysAndGenerateOutput(output_file_name):

	file_data = {}	
	for file in text_files:
		data = []
		fileobject = open(text_files_folder+"/"+file+".txt",'r',encoding="utf-8") 
		lines = fileobject.readlines()
		for line in lines:
			try:
				splitted_line = line.split(text_file_delimeter)
				key = splitted_line[0].strip()
				value = splitted_line[1].strip()
				data.append((key,value))
			except:
				pass
		file_data[file] = data

	#print(file_data)

	rows = []

	for key in keys:
		row = [key]
		for file in text_files:
			file_entries = file_data[file]
			found=False
			for fkey,fvalue in file_entries:
				if(fkey==key):
					row.append(fvalue)
					found = True
					break
			if(not found):
				row.append("")
		rows.append(row)

	#print(rows)

	header = ["Key"]+text_files
	final_rows = [header] + rows
	f = open(output_file_name, 'w',newline='',encoding="utf-8")
	with f:
	    writer = csv.writer(f)
	    writer.writerows(final_rows)

	print("Output Generated :)")


combineKeysAndGenerateOutput(output_file_name)