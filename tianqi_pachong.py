# 爬取常熟天气数据
import requests
import re
import datetime
import csv


# 第一步
# 根据时间获取网页
def getPageByTime(timeStr):
    url = 'https://www.tianqishi.com/changshu/' + timeStr + '.html'
    resp = requests.get(url,verify=False)
    return extractDataByText(resp.text)


# 第二步
# 根据获取的text提取数据
def extractDataByText(text):
    # 使用正则表达式截取字符串
    pattern = '天气简报</b></p>\n<p>.*?</p>'
    p = re.compile(pattern)
    match = re.search(p, text)
    retult = match.group(0).replace("天气简报</b></p>\n<p>", "").replace("</p>", "")
    return retult


# 第三步
# 将查询到的数据写入到excel
def writeDataToExcel(path, data):
    with open(path, 'w', newline='') as csv_file:
        csv_writer = csv.writer(csv_file)
        for list in data:
            csv_writer.writerow(list)


# 开始
def start():
    # 结果集合
    resultList = []
    # 遍历年份
    for year in range(2021, 2022):
        # 遍历月份
        for month in range(8, 9):
            # 遍历日期
            for day in range(1, 31):
                # 判断日期是否出错，否则跳过循环
                monthStr = '{}'.format(month)
                dayStr = '{}'.format(day)
                if (month < 10):
                    monthStr = '0{}'.format(month)
                if (day < 10):
                    dayStr = '0{}'.format(day)
                date = '{}{}{}'.format(year, monthStr, dayStr)
                try:
                    datetime.date(year, month, day)
                    # 日期合法继续下一步
                    try:
                        retult = getPageByTime(date)
                        resultList.append([date, retult])
                        print(date + ' Success')
                    except Exception as err:
                        print(err)
                        print(date + ' error')
                except:
                    pass
    # 将结果写入excel中
    path = r'E:\Zph\0704常熟日报\2天气信息\out21.csv'
    writeDataToExcel(path, resultList)


start()