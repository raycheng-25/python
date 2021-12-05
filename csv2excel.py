def xlsxreport(csvfile, xlsxpath, xlsdestfpath):
	import csv, openpyxl, shutil
	import pandas as pd
	from openpyxl import load_workbook

	# read and convert the csv content to pandas dataframe
	with open(csvfile) as f:
		readcsv = csv.reader(f)
		df0 = pd.DataFrame(data=readcsv)

	# convert the dataframe to numpy, convert the string to float
	numpy_data = df0.values
	for n in range(len(numpy_data)):
		for i in range(len(numpy_data[n])):
			try:
				numpy_data[n][i] = float(numpy_data[n][i])#
			except:
				pass
	# read the workbook then delete the sheet that contains old data
	book = load_workbook(xlsxpath)
	print(type(book))
	del book['Data']
	# write the data to excel sheet
	with pd.ExcelWriter(xlsxpath, engine='openpyxl', strings_to_numbers=True) as writer:
		writer.book = book
		df = pd.DataFrame(data = numpy_data)
		df0.to_excel(writer, sheet_name='Data', header=False, index=False)
	# copy to the excel to specific location
	shutil.copyfile(xlsxpath, xlsdestfpath)
	print(f'{xlsdestfpath} is created')

if __name__ =='__main__':
	import shutil
	csvfile = '/Users/raymond.cheng/My Drive/01-Test Dev/Test Data/raw.csv'
	xlsxpath = '/Users/raymond.cheng/My Drive/01-Test Dev/Test Data/report.xlsx'
	xlsdestfpath = '/Users/raymond.cheng/My Drive/01-Test Dev/Test Data/new_report.xlsx'
	'''
	For define the output path
	destDir11 = ''
	fname = ''
	xlsdestfpath = os.path.join(destDir11, fname) + ".xlsx"
	'''
	shutil.copy(xlsxpath1, xlsxpath)
	xlsxreport(csvfile, xlsxpath, xlsdestfpath)