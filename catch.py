import requests
from bs4 import BeautifulSoup
from urllib.parse import quote
import cpca
import csv
import os
import xlwt


def csv_2_xls():
    csvfiles = os.listdir('.')
    csvfiles = filter(lambda x: x.endswith('csv'), csvfiles)
    for csvfile in list(csvfiles):
        finename = csvfile.split('.')[0]
        if not os.path.exists('excel'):
            os.mkdir('excel')
        xlsfile = 'excel/' + finename + '.xls'
        with open(csvfile, 'r') as f:
            reader = csv.reader(f)
            workbook = xlwt.Workbook()
            sheet = workbook.add_sheet('sheet1')  # 创建一个sheet表格
            i = 0
            for line in reader:
                j = 0
                for v in line:
                    sheet.write(i, j, v)
                    j += 1
                i += 1
            workbook.save(xlsfile)  # 保存Excel
        print(f'转换完成: {csvfile} -> {xlsfile}')

def search_data(key_word):
	headers = {
		'Cookie': "ssuid=3295451614; _ga=GA1.2.379723253.1654941791; jsid=SEO-BAIDU-ALL-SY-000001; TYCID=5e145c803b7111ee9012697c2ff98ee5; HWWAFSESID=f26f06c784d6c4bf3fd; HWWAFSESTIME=1692508043749; csrfToken=Wx0yqUR5rOvy0BqKkFzc1SKU; bannerFlag=true; Hm_lvt_e92c8d65d92d534b0fc290df538b4758=1692106873,1692508070; refresh_page=0; _gid=GA1.2.1265700635.1692509886; searchV6CompanyResultCid=2356710914; searchV6CompanyResultCid.sig=ceb_VMG1BLT8TPD_RSqliMwtt8SOw2xzD2nJuXp4e-o; RTYCID=fc4aecfecb64442a9d587a6b76a37fc9; bdHomeCount=4; searchV6CompanyResultName=%E6%AE%B5%E6%99%93%E9%B9%83; searchV6CompanyResultName.sig=F3UqiTd8_kAEFGFv2_IRiYo2vYZf_rxkV7JmN2suRE4; tyc-user-info={%22state%22:%220%22%2C%22vipManager%22:%220%22%2C%22mobile%22:%2213111880088%22}; tyc-user-info-save-time=1692515655319; auth_token=eyJhbGciOiJIUzUxMiJ9.eyJzdWIiOiIxMzExMTg4MDA4OCIsImlhdCI6MTY5MjUxNTYzNCwiZXhwIjoxNjk1MTA3NjM0fQ.tLUk06FxWu3mAJePr7D5HH0OzETn2ml91oOOOwJo_sG1NneUasAryGgfwMBMEp35ir0fn_qC4qzXbbd2EInL2w; tyc-user-phone=%255B%252213111880088%2522%255D; sensorsdata2015jssdkcross=%7B%22distinct_id%22%3A%22314175433%22%2C%22first_id%22%3A%22178f823b2e190a-0e5100bf1bbb15-3f356b-2073600-178f823b2e2ede%22%2C%22props%22%3A%7B%22%24latest_traffic_source_type%22%3A%22%E7%9B%B4%E6%8E%A5%E6%B5%81%E9%87%8F%22%2C%22%24latest_search_keyword%22%3A%22%E6%9C%AA%E5%8F%96%E5%88%B0%E5%80%BC_%E7%9B%B4%E6%8E%A5%E6%89%93%E5%BC%80%22%2C%22%24latest_referrer%22%3A%22%22%7D%2C%22%24device_id%22%3A%22178f823b2e190a-0e5100bf1bbb15-3f356b-2073600-178f823b2e2ede%22%2C%22identities%22%3A%22eyIkaWRlbnRpdHlfY29va2llX2lkIjoiMTg5Zjk2ZGQ2NGE1OWUtMDlhOGI3YmFlM2ZhMTk4LTI2MDMxYzUxLTMxNzAzMDQtMTg5Zjk2ZGQ2NGIxNjk5IiwiJGlkZW50aXR5X2xvZ2luX2lkIjoiMzE0MTc1NDMzIn0%3D%22%2C%22history_login_id%22%3A%7B%22name%22%3A%22%24identity_login_id%22%2C%22value%22%3A%22314175433%22%7D%7D; searchSessionId=1692515643.31394186; Hm_lpvt_e92c8d65d92d534b0fc290df538b4758=1692515833; _gat_gtag_UA_123487620_1=1",
		'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/115.0.0.0 Safari/537.36'
	}#需要根据用户内容自己修改
	url = 'https://www.tianyancha.com/search?key={}'.format(quote(key_word))
	html = requests.get(url,headers = headers,timeout=3)
	if html.status_code != 200:
		print("*" * 30)
		print("网络连接错误！！！请检查Cookie，User-Agent是否设置正确！！！,或者被网页拦截请打开网页重新认证")
		print("*" * 30)
		return [key_word,None,None,None,None]
	soup = BeautifulSoup(html.text,'html.parser')
	if len(soup.find_all("a",class_ = "index_alink__zcia5 link-click")) == 0:
		print(f"并未查到相关公司或个人的资料：该或个人为{key_word}，若连续未查找到请检查重新更换cookie，可能被服务器检测到爬虫行为而导致拒绝访问")
		return [key_word,None,None,None,None]
	info = soup.find_all("a",class_="index_alink__zcia5 link-click")[0]#查找首个搜索结果
	company_name = info.text #公司全称
	company_url = info['href']
	# print(company_name,company_url)


	html_detail = requests.get(company_url,headers = headers)
	soup_detail = BeautifulSoup(html_detail.text,'html.parser')
	table_infos = soup_detail.find("table",class_ = "index_tableBox__ZadJW")#查找表格内容

	table_list = table_infos.find_all("tr")
	if len(table_list)== 11:
		registered_capital = table_list[2].find("div").text#公司注册资本
		if registered_capital == '-':
			registered_capital = '未查到注册成本'
		company_site = table_list[-2].find("span", class_="index_copy-text__ri7W6").text#公司地址
		company_locate_list = cpca.transform([company_site]).values.tolist()[0]
		if company_locate_list[0] == None or company_locate_list[1] == None:
			company_locate = None
		else:
			company_locate = company_locate_list[0]+company_locate_list[1]

		out_put_data = [company_name,registered_capital,company_locate,company_site,company_url]#信息汇总
	elif len(table_list) == 5:
		registered_capital = table_list[0].contents[3].text # 公司注册资本
		company_site = table_list[-2].contents[1].text  # 公司地址
		company_locate_list = cpca.transform([company_site]).values.tolist()[0]
		company_locate = company_locate_list[0] + company_locate_list[1]
		out_put_data = [company_name, registered_capital, company_locate, company_site, company_url]  # 信息汇总
	else:
		print(f"该公司或者学院名称不符合查找表格标准，需要用户手动查找添加，查找的关键词为{key_word}")
		return [key_word,None,None,None,company_url]
	print(out_put_data)
	return out_put_data



if __name__ == "__main__":
	with open("company.txt",encoding="utf-8") as f:
		key_words = f.read().splitlines()#将关键词存在company的txt的文件当中
	f.close()
	csv_path = "output.csv"
	if os.path.getsize(csv_path):
		header_list = ["公司名称", "注册资本", "公司属地", "公司地址", "天眼查网址"]
		with open(csv_path,"w",newline="") as f:
			writer = csv.writer(f)
			writer.writerow(header_list)
		f.close()
	for key_word in key_words:
			line_data =  search_data(key_word)
			with open(csv_path, "a+", newline="") as f:
				writer = csv.writer(f)
				writer.writerow(line_data)
			f.close()
	csv_2_xls()