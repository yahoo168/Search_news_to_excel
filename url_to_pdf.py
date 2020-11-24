from openpyxl import *
import pdfkit
import os
import time
import eventlet

def check_file_exist(filename, directory_path):
	filename_list = os.listdir(directory_path)
	if filename in filename_list:
		print("檔案已存在： ", filename)
		return True
	else:
		return False

def to_pdf(pdf_directory_path, title_url_tuple):
	eventlet.monkey_patch()
	for item in title_url_tuple:
		title = item[0]
		url = item[1]
		
		if check_file_exist(title+'.pdf', pdf_directory_path):
			continue
		
		file_path = os.path.join(pdf_directory_path,title+'.pdf')
		
		try:
			with eventlet.Timeout(5,False):
				pdfkit.from_url(url, file_path)
			
		except Exception as e:
			print(e)
			continue
			
		print("檔案已存檔", title)
	
excel_directory_path = os.path.join(".", "news_excel_file")
excel_filename_list = os.listdir(excel_directory_path)
pdf_directory_path = os.path.join(".", "news_pdf_file")

# print(excel_filename_list)
for excel_filename in excel_filename_list:
	excel_filename_path = os.path.join(excel_directory_path, excel_filename)
	
	try:
		wb = load_workbook(filename = excel_filename_path)
		ws = wb["news"]
	except Exception as e:
		print(e)
		continue

	title_list = [i.value for i in list(ws.columns)[1][1:]]
	url_list = [i.value for i in list(ws.columns)[3][1:]]
	title_url_tuple = tuple(zip(title_list,url_list))
	industry_name = excel_filename.split('.')[0]
	indusrty_pdf_directory_path = os.path.join(pdf_directory_path, industry_name)
	
	if not os.path.isdir(indusrty_pdf_directory_path):
		os.mkdir(indusrty_pdf_directory_path)
	
	to_pdf(indusrty_pdf_directory_path, title_url_tuple)
	