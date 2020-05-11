'''
file: CoInfoSpider.py
encoding: utf-8
data: 2020.04.23
name: LSY
email: 394883561@qq.com
introduction: 爬取企业工商注册信息
url: https://gongshang.mingluji.com/anhui/riqi
'''

import requests
import re
import time
import random
import xlwt
from traceback import print_exc


class CompanyInfo_Spide():
    
    # 构建请求头
    def __init__(self):
        
        self.headers = {
                        "Cookie": "has_js=1; Hm_lvt_f733651f7f7c9cfc0c1c62ebc1f6388e=1587622287; __utmc=152261551; __utmz=152261551.1587622287.1.1.utmcsr=(direct)|utmccn=(direct)|utmcmd=(none); __utma=152261551.1014135744.1587622287.1587637372.1587647755.5; PHPSESSID=1kaqkqeem52qqceq1iksoit604; __utmt=1; __utmb=152261551.5.10.1587647755; Hm_lpvt_f733651f7f7c9cfc0c1c62ebc1f6388e=1587648934",
                        "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/78.0.3904.97 Safari/537.36"
                       }


    # 获取响应
    def get_respone(self, url):
        
        res = requests.get(url, headers=self.headers, timeout=10)
        res.encoding = 'utf-8'
        self.html = res.text


    # 获取所有注册日期
    def get_all_data(self): 

        self.all_data = []  # 存所有获取到的注册日期
        
        # 此处可将1改为其他数值，表示要爬取的页数
        for page in range(1):
            page_url = 'https://gongshang.mingluji.com/anhui/riqi?page=%s'%page
            
            try:
                self.get_respone(page_url)
                
                # 正则匹配注册日期
                regex = r'"field-content"><a href="(.+?)"'
                data_regex = re.findall(regex, self.html)
                
                print("页面：%s 爬取成功"%page_url)
                self.all_data += data_regex
            
            except Exception :
                print(">>> 页面：%s 爬取失败"%page_url)
                print("失败原因如下：")
                print_exc()
            
            finally:
                # 延时
                sec = random.uniform(1, 2)
                time.sleep(sec)


    # 获取所有企业的详情页链接
    def get_all_company(self):

        self.all_company_url = []  # 存所有公司的详情页链接

        for data in self.all_data:
            data_url = "https://gongshang.mingluji.com" + data
            
            try:
                self.get_respone(data_url)
                
                # 正则匹配企业详情页链接
                regex = r'(https://gongshang.mingluji.com/anhui/name/.+?)">'
                result = re.findall(regex, self.html)
                
                print("页面：%s 爬取成功"%data_url)
                self.all_company_url += result
            
            except Exception :
                print(">>> 页面：%s 爬取失败"%data_url)
                print("失败原因如下：")
                print_exc()
            
            finally:
                # 延时
                sec = random.uniform(1, 2)
                time.sleep(sec)


    # 获取所有企业的注册信息
    def get_company_info(self):

        self.all_company_info = []  # 存所有企业的注册信息
        count = 0  # 记录当前爬取数
        
        for company_url in self.all_company_url:
            try:
                self.get_respone(company_url)
                
                # 正则匹配公司名称
                regex_company = r'''<span class='field-label'>企业名称.+?<span itemprop='name'>(.+?)</span></span>'''
                company = re.findall(regex_company, self.html)[0]
            
            except Exception :
                print(">>> 页面：%s 爬取失败"%company_url)
                print("失败原因如下：")
                print_exc()
            
            else:
                try:
                    # 正则匹配工商编码
                    regex_number = r'''<span itemprop='identifier'>.+?>(.+?)</a></span>'''
                    number = re.findall(regex_number, self.html)[0]
                except Exception :
                    number = ''
                try:
                    # 正则匹配地址
                    regex_add = r'''<span itemprop='address'>(.+?)</span></span>'''
                    add = re.findall(regex_add, self.html)[0]
                except Exception :
                    add = ''
                try:
                    # 正则匹配注册日期
                    regex_data = r'''<span itemprop='foundingDate'>.+?>(.+?)</a></span>'''
                    data = re.findall(regex_data, self.html)[0]
                except Exception :
                    data = ''
                try:
                    # 正则匹配法定代表人
                    regex_name = r'''<span itemprop='founder'>.+?>(.+?)</a></span>'''
                    name = re.findall(regex_name, self.html)[0]
                except Exception :
                    name = ''
                try:
                    # 正则匹配经营范围
                    regex_range = r'''<span itemprop='makesOffer'>(.+?)</span></span>'''
                    Range = re.findall(regex_range, self.html)[0]
                except Exception :
                    Range = ''
                try:
                    # 正则匹配专题
                    regex_item = r'''<span class='field-label'>专题.+?<span itemprop='name'></span><a href=".+?">(.+?)</a>'''
                    item = re.findall(regex_item, self.html)[0]
                except Exception :
                    item = ''
                try:
                    # 正则匹配地区
                    regex_area = r"<span class='field-label'>地区.+?<span itemprop='foundingLocation'><a href=.+?>(.+?)</a></span>"
                    area = re.findall(regex_area, self.html)[0]
                except Exception :
                    area = ''
                try:
                    # 正则匹配市县
                    regex_city = r"<span class='field-label'>县市.+?<span itemprop='foundingLocation'><a href=.+?>(.+?)</a></span>"
                    city = re.findall(regex_city, self.html)[0]
                except Exception :
                    city = ''
                try:
                    # 正则匹配企业类型
                    regex_Type = r"<span class='field-label'>企业类型.+?itemprop='name'>(.+?)</span></span>"
                    Type = re.findall(regex_Type, self.html)[0]
                except Exception :
                    Type = ''
                try:
                    # 正则匹配企业状态
                    regex_status = r"<span itemprop='serverStatus'>(.+?)</span></span>"
                    status = re.findall(regex_status, self.html)[0]
                except Exception :
                    status = ''

                count += 1  # 计当前爬取数
                print("%d  页面：%s 爬取成功"%(count, company_url))
                # 存所有公司信息
                self.all_company_info.append((company, number, add, data, name, Range, item, area, city, Type, status))
            
            # 延时
            sec = random.uniform(1, 2)
            time.sleep(sec)


    # 爬取到的信息写入文件
    def write_file(self):

        # 建excel表
        workbook = xlwt.Workbook(encoding='utf-8')
        worksheet = workbook.add_sheet("企业注册信息")
        
        # 写入标题行
        title = ['企业名称', '统一社会信用代码/工商注册号', '注册地址', '注册日期',
                 '法定代表人', '经营范围', '专题', '地区', '市县', '公司类型', '公司现状']
        for index,col_name in enumerate(title):
            worksheet.write(0, index, label=col_name)

        line = 1  # 当前写入的行数
        # 逐行写入爬取到的企业信息
        for company_info in self.all_company_info:
            for index,info in enumerate(company_info):
                worksheet.write(line, index, label=info)
            line += 1

        # 存文件
        workbook.save('企业注册信息.xls')


    # 主循环
    def main(self):

        print("开始爬取所有注册日期...")
        self.get_all_data()
        print("所有注册日期页面爬取完成！")

        print("\n开始爬取所有企业链接...")
        self.get_all_company()
        print("所有企业链接爬取完成！，共成功获取到 %d 条企业链接"%len(self.all_company_url))

        print("\n开始爬取所有企业信息...预计耗时 %d 秒"%(len(self.all_company_url) * 1.5))
        self.get_company_info()
        print("所有企业信息爬取完成！")

        print("\n开始将爬取的所有企业信息写入文件...")
        self.write_file()
        print("所有企业信息写入完成！")


if __name__=='__main__':

    CompanyInfo = CompanyInfo_Spide()
    CompanyInfo.main()