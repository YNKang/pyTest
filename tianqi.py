# encoding:utf-8
import datetime
import os

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
    nowTime = datetime.datetime.now()
    city = weather_result['real']['station']['city']
    path = r'D:\tqsj\\'+str(nowTime.year)+'\\' +city
    fileName = path + '\\' + 'tianqi_'+city+'.xls'
    folder = os.path.exists(path)
    if not folder:
        os.makedirs(path)
    if os.path.exists(fileName):
        oldWd = xlrd.open_workbook(fileName)
        newWd = copy(oldWd)
        sheet = newWd.get_sheet(0)
        row = len(sheet.rows)
        sheet.write(row, 0, weather_result['real']['station']['city'])
        sheet.write(row, 1,bytes(nowTime.year)+'_'+bytes(nowTime.month)+'_'+bytes(nowTime.day))
        sheet.write(row, 2,bytes(nowTime.hour))
        sheet.write(row, 3, weather_result['real']['weather']['info'])
        sheet.write(row, 4, bytes(weather_result['real']['weather']['temperature'])+'C')
        sheet.write(row, 5, bytes(weather_result['real']['weather']['humidity'])+'%')
        sheet.write(row, 6, bytes(weather_result['aqi']['aqi']))
        sheet.write(row, 7, weather_result['aqi']['text'])
        sheet.write(row, 8, weather_result['real']['wind']['direct']+weather_result['real']['wind']['power'])
        sheet.write(row, 9, bytes(weather_result['real']['weather']['rain'])+'mm')
        sheet.write(row, 10, bytes(weather_result['real']['weather']['feelst'])+'C')
        newWd.save(fileName)
    else:
        workbook = xlwt.Workbook(encoding='utf-8')
        sheet = workbook.add_sheet('weather_report',cell_overwrite_ok=False)
        title = ['城市', '日期','小时','天气','温度', '相对湿度','AQI','空气质量','风向风速', '降水', '体感温度']
        for i in range(len(title)):
            sheet.write(0, i, title[i],set_color(0x00,True))
        row = 1
        sheet.write(row, 0, weather_result['real']['station']['city'])
        sheet.write(row, 1,bytes(nowTime.year)+'_'+bytes(nowTime.month)+'_'+bytes(nowTime.day))
        sheet.write(row, 2,bytes(nowTime.hour))
        sheet.write(row, 3, weather_result['real']['weather']['info'])
        sheet.write(row, 4, bytes(weather_result['real']['weather']['temperature'])+'C')
        sheet.write(row, 5, bytes(weather_result['real']['weather']['humidity'])+'%')
        sheet.write(row, 6, bytes(weather_result['aqi']['aqi']))
        sheet.write(row, 7, weather_result['aqi']['text'])
        sheet.write(row, 8, weather_result['real']['wind']['direct']+weather_result['real']['wind']['power'])
        sheet.write(row, 9, bytes(weather_result['real']['weather']['rain'])+'mm')
        sheet.write(row, 10, bytes(weather_result['real']['weather']['feelst'])+'C')
        workbook.save(fileName)

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
    nowTime = datetime.datetime.now()
    citys = ['54511','54517','53698','53772','53463','54342','54161','50953','58367','58238','58457','58321','58847','58606','54823','57083','57494','57679','59287','59431','59758','57516','56294','56294','56778','55591','57036','52889','52866','53614','51463']
    while True:
        print "Start : %s" % time.ctime()
        now = datetime.datetime.now()
        for city in citys:
            dict={}
            url = 'http://www.nmc.cn/f/rest/real/'+city+'?_='+bytes(now.microsecond)
            data = requestsData(url)
            dict['real'] = data;
            url = 'http://www.nmc.cn/f/rest/aqi/'+city+'?_='+bytes(now.microsecond)
            data2 = requestsData(url)
            dict['aqi'] = data2;
            list_to_excel(dict)
        print "end : %s" % time.ctime()
        time.sleep(3600)


