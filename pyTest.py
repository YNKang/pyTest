# encoding:utf-8
import datetime

import xlrd
import xlwt
from xlutils.copy import copy
from bs4 import BeautifulSoup
import requests
import time

# 爬取“天气网”天气预报
from pandas import json

#数据请求
def requestsData(url):  # city为字符串，year为列表，month为列表
    #url = 'http://www.nmc.cn/f/rest/province'
    headers = {
        'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/60.0.3112.113 Safari/537.36'}           # 设置头文件信息
    response = requests.get(url, headers=headers).content    # 提交requests get 请求
    #soup = BeautifulSoup(response, "html.parser")       # 用Beautifulsoup 进行解析
    jsonData = json.loads(response)
    return jsonData

#写文件
def list_to_excel(weather_result):
    path = 'C:\Users\Administrator\Desktop\shuju11.xls'
    workbook = xlwt.Workbook(encoding='utf-8')
    sheet = workbook.add_sheet('weather_report',cell_overwrite_ok=False)
    title = ['城市', '天气','温度', '相对湿度','风向风速', '降水', '体感温度']
    for i in range(len(title)):
        sheet.write(0, i, title[i],set_color(0x00,True))
    row, col = 1, 0
    sheet.write(row, col, weather_result['station']['city'])
    sheet.write(row, col + 1, weather_result['weather']['info'])
    sheet.write(row, col + 2, bytes(weather_result['weather']['temperature'])+'C')
    sheet.write(row, col + 3, bytes(weather_result['weather']['humidity'])+'%')
    sheet.write(row, col + 4, weather_result['wind']['direct']+weather_result['wind']['power'])
    sheet.write(row, col + 5, bytes(weather_result['weather']['rain'])+'mm')
    sheet.write(row, col + 6, bytes(weather_result['weather']['feelst'])+'C')

    xlrd.open_workbook('path','')
    newWord = copy(workbook);
    newWord.get_sheet(0)

    workbook.save(path)

#样式设置
def set_color(color,bold):
    style=xlwt.XFStyle()
    font=xlwt.Font()
    font.colour_index=color
    font.bold = bold
    style.font=font
    return style
#主方法
if __name__ == '__main__':
    #provinces = ['ABJ','ATJ','AHE','ASX','ANM','ALN','AJL','AHL','ASH','AJS','AZJ','AAH','AFJ','AJX','ASD','AHA','AHB','AHN','AGD','AGX','AHI','ACQ','ASC','AGZ','AYN','AXZ','ASN','AGS','AQH','ANX','AXJ','AXG','AAM','ATW']
    provinces = ['ASD']
    citys = ['54511','54517','53698','53772']
    tt=[]
    now = datetime.datetime.now()
    for city in citys:
        # time.sleep(2)
        url = 'http://www.nmc.cn/f/rest/real/'+city+'?_='+bytes(now.microsecond)
        data = requestsData(url)
        # tt.append(data)
        list_to_excel(data)

