import xlrd
import csv

def csv_from_excel(xlsx_name, sheet_name, txt_name):
	wb = xlrd.open_workbook(xlsx_name)
	sh = wb.sheet_by_name(sheet_name)
	csv_file = open(txt_name, 'w', encoding='utf-16le', newline='')
	csv_file.write('\ufeff')
	wr = csv.writer(csv_file, quoting=csv.QUOTE_NONE, delimiter="\t")

	col_num = sh.ncols

	for rownum in range(sh.nrows):
		row_list = sh.row_values(rownum, 0, sh.ncols)
		for i in range(sh.ncols):
			item_end = str(row_list[i])[-2:]
			if str(".0") in item_end:
				row_list[i] = str(row_list[i])[:-2]
			row_list[i] = str(row_list[i])
			
		wr.writerow(row_list)
	csv_file.close()

def read_list():
	file = open('.DataTableList.txt', 'r')
	csv_data = csv.reader(file, delimiter='\t')
	c_list = list(csv_data)
	for i in range(len(c_list)):
		csv_from_excel(c_list[i][0], c_list[i][1], c_list[i][2])
		
read_list()