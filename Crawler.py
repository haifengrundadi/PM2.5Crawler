# -*- encoding:utf-8 -*-
__author__ = 'juanecho'
import base64
import json
import xlwt
import urllib
import urllib2
from xlwt import Workbook

"""
# 气象数据
# 1：编辑重发post请求（starttime改一下）得到相应并保存为（data.txt)
# 2；需要经过两次base64解码
# 3：解析json数据 得到结果保存为result.txt
"""


def chageValuesToBase64(values):
    """
    :param values: 普通的字典
    :return: 值为base64编码的字典
    """
    dict = {}
    for (k, v) in values.items():
        dict[k] = base64.encodestring(v).strip("\n").strip("=")
    return dict


def sendRequestAndGetSourceFile(url, values, cityName, JSONFileName):
    """
    编辑http请求 解码得到JSON文件
    :param url: 获取数据的网址
    :param values: 值为base64编码的字典（post的请求值)
    :param cityName: 想要获取城市的名称（要写正确)
    :param JSONFileName: 最后所需数据的json文件
    :return:
    """
    values = chageValuesToBase64(values)
    data = urllib.urlencode(values)
    req = urllib2.Request(url, data)
    response = urllib2.urlopen(req)
    the_page = response.read()
    data = base64.decodestring(base64.decodestring(the_page)).strip(cityName)
    with open(JSONFileName, 'w')as f_result:
        f_result.write(data)


def getResultFromData(JSONFileName, sheet):
    """
    分析JSON
    :param JSONFileName: 所需数据的json文件
    :param sheet: 处理后的表格
    :return:
    """
    style = xlwt.easyxf('font: bold 1, color black;')
    with open(JSONFileName, 'r')as f_source:
        dict = json.load(f_source)
        content = dict["rows"]
        m = 0
        for row in content:
            r = m
            c = 0
            m += 1
            for (k, v) in row.items():
                sheet.write(r, c, k, style)
                c += 1
                sheet.write(r, c, v)
                c += 1


if __name__ == "__main__":
    urlPM = 'http://www.aqistudy.cn/api/getdata_citydetail.php'
    urlWeather = 'http://www.aqistudy.cn/api/getdata_cityweather.php'
    values = {'city': '上海',  'type': 'HOUR', 'startTime': '2014-01-01 00:00:00', 'endTime': '2015-08-28 09:00:00'}
    cityname = '上海'
    sendRequestAndGetSourceFile(urlPM, values, cityname, "data.dat")
    sendRequestAndGetSourceFile(urlWeather, values, cityname, "dataweather.dat")
    book = Workbook(encoding="utf-8")
    sheetPM = book.add_sheet(u"PM2.5")
    sheetWeather = book.add_sheet(u"weather")
    getResultFromData("data.dat", sheetPM)
    getResultFromData("dataweather.dat", sheetWeather)
    book.save("resultPM.xls")

