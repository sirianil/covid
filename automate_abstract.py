import xlsxwriter
import xlrd

datePattern = {
	'day' : '20',
	'month' : '09',
	'year' : '2020',
	'day_suffix' : 'th', #st/nd/rd/th 
	'month_words' : 'Sep'
}

input_file_names = {
	'testing' : 'Testing ' + datePattern.get('day') + '.' + datePattern.get('month') + '.' + datePattern.get('year') + '.xlsx', #Ex: 'Testing 20.09.2020' 
	'testing-team' : 'Testing team-P6-Pharma-Data-' + datePattern.get('day') + "-" + datePattern.get('month')
	'hfw' : 'HFW-' + datePattern.get('day') + datePattern.get('month') + datePattern.get('year') + '.xlsx', #'HFW-20092020'
	'bmc' : 'BMC-' + datePattern.get('day') + datePattern.get('month') + datePattern.get('year') + '.xlsx', #'BMC-20092020'
	'apthamitra' : 'Apthamithra Covid Testing Outreach 19' + datePattern.get('day_suffix') + ' ' + datePattern.get('month_words') + ' Phase 4 BLR.xlsx'
}

rows_to_skip = set(['(blank)', 'PHC name', 'Grand Total'])

output_file_name = 'xyz_012' #TODO: what's the correct output format?

def main_processing():
	output_sheet_counter = 499

	def write_to_file(abstract_worksheet, output_new_worksheet):
		print("Inside callingFunction")
		for row in range(abstract_worksheet.nrows):
			if abstract_worksheet.cell_value(row, 0) not in rows_to_skip:
				for col in range(0,2):
					output_value = abstract_worksheet.cell_value(row, col)
					output_new_worksheet.write(row, col, output_value)

	for file_type, file_name in input_file_names.items():
		output_sheet_counter = output_sheet_counter + 1
		print("<--Processing type: ", file_type  + "-->")
		
		input_file = xlrd.open_workbook(file_name)
		print("<--"  + file_name + " opened successfully-->")	

		''''if file_type is 'testing':
			output_new_worksheet = output_file.add_worksheet(output_file_name + str(output_sheet_counter))
			print("<--Additional output sheet added successfully-->")

			abstract_worksheet = input_file.sheet_by_name('abstract')
			write_to_file(abstract_worksheet, output_new_worksheet)

		elif file_type is 'bmc' or 'hfw':'''
		input_sheets = input_file.sheet_names()
		print("Sheets in file are ")
		print(sheets)

		for sheet in input_sheets:
			sheet = str(sheet)
			if sheet.startswith("abs"):
				output_new_worksheet = output_file.add_worksheet(output_file_name + str(output_sheet_counter))
				print("<--Additional output sheet added successfully-->")

				abstract_worksheet = input_file.sheet_by_name(sheet)
				write_to_file(abstract_worksheet, output_new_worksheet)
				output_sheet_counter = output_sheet_counter + 1

		print("<--Done processing file-->")

	output_file.close()

print ("<--Beginning automated script to generate abstract sheets-->")
output_file = xlsxwriter.Workbook('Output_Abstract.xlsx')

main_processing()




		
