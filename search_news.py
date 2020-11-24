from GoogleNews import GoogleNews
from openpyxl import load_workbook
import datetime
import sys
import os

#得到新聞的原始資訊
def get_news_result_list(keyword, language="cn", SEARCH_PAGE_NUM = 1):
	googlenews = GoogleNews(language)
	print("\n開始搜尋「{keyword}」".format(keyword=keyword))
	googlenews.search(keyword)
	result = []
	for i in range(SEARCH_PAGE_NUM):
		googlenews.getpage(i)
		result += googlenews.result()
		googlenews.clear()
		#顯示處理進度
		sys.stdout.write('\r'+str(round((i+1) / SEARCH_PAGE_NUM, 2)*100) +"%") 

	return result

#篩出日期、標題等重要資訊
def get_info_from_news_result(result):
	news_list = []
	length = len(result)
	for index, news in enumerate(result, 1):
		news_item = []
		try:
			news_item.append(adjust_date(news["date"]))
			news_item.append(news["title"])
			news_item.append(adjust_link(news["link"]))
			news_item.append(news["link"])
			news_item.append(news["desc"])
			#顯示處理進度
			sys.stdout.write('\r'+str(round(index / length, 2)*100) +"%") 
			
		except Exception as e:
			print(e)

		news_list.append(news_item)
	return news_list

# 去除重複的新聞（依標題判斷）
def delete_overlap_news(news_list):
	existing_news_title = []
	set_news_list = []

	for news in news_list:
		if news[1] in existing_news_title:
			continue
		else:
			existing_news_title.append(news[1])
			set_news_list.append(news)

	return set_news_list

def adjust_date(date):
	try:
		# 由字串長度判斷為 「"2020年7月18日"」形式
		if len(date) >= 9:
			adjusted_date = ''.join(["/" if idx in ["年","月"] else idx for idx in date[0:-1]]) 
		
		# 由字串長度判斷為 "XX 小時前" 形式，大致計算日期
		else:
			splitted_date = date.split(' ')
			num = int(splitted_date[0])
			unit = splitted_date[1]
			day = 0
			
			if unit == "小時前":
				day = num / 24

			elif unit == "天前":
				day = num

			elif unit == "週前":
				day = num*7

			elif unit == "個月前":
				day = num*30

			elif unit == "月前":
				day = num*30

			adjusted_date = (datetime.datetime.now()-datetime.timedelta(days=day)).strftime("%Y/%m/%d")

		return adjusted_date
	except:
		return date

#將連結文字加上超連結
def adjust_link(link):
	#Excel限制文字長度不得超過252
	if len(link) < 252:
		adjusted_link  = "=HYPERLINK(\"{0}\",\"{1}\")".format(link,"連結")
		
	else:
		adjusted_link = link

	return adjusted_link

#將新聞寫入Excel檔
def write_news_result_to_excel(news_list, template_filename, save_filename):
	wb = load_workbook(filename = template_filename)
	sheet_news = wb["news"]
	for news_item in news_list:
	    sheet_news.append(news_item)

	wb.save(save_filename)

def search_news(template_filename, save_filename, key_word_list, SEARCH_PAGE_NUM):
	result_list = []
	
	for keyword in key_word_list:
		result_list += get_news_result_list(keyword, language="cn", SEARCH_PAGE_NUM=SEARCH_PAGE_NUM)

	news_list = get_info_from_news_result(result_list)
	news_list = delete_overlap_news(news_list)
	write_news_result_to_excel(news_list, template_filename, save_filename)
	return news_list

def read_config(filename):
	config_list=[]
	with open(filename, "r") as f:
		lines = f.readlines()
		_ ,page_num = lines[0].split(':')
		print(page_num)
		for line in lines[1:]:
			line = line.replace('\n', '')
			topic, keyword_list, save_or_not = line.split(':')
			config_list.append((topic, keyword_list, save_or_not))
	return page_num, config_list

template_filename = "news_excel_template.xlsx"
directory_path = "news_excel_file"
config_filename = "config.txt"
page_num, config_list = read_config(config_filename)

for config_item in config_list:
	topic = config_item[0]
	keyword_list = config_item[1].split(',')
	save_or_not = config_item[2]
	print(topic, keyword_list, save_or_not, '\n')
	save_filename = os.path.join(".", directory_path, topic+".xlsx")
	if save_or_not == 'O':
		search_news(template_filename, save_filename, keyword_list, SEARCH_PAGE_NUM=int(page_num))
