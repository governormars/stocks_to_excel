# _*_ coding:utf-8 _*_

import requests, re, json, time, os
import heapq
from bs4 import BeautifulSoup
import xlwt


def get_data():
    stock_code = raw_input('plz in put stock code：')
    base_url = "http://f10.eastmoney.com/f10_v2/FinanceAnalysis.aspx?code="
    url = base_url + stock_code
    # 请求数据
    orihtml = requests.get(url).content
    # 创建 beautifulsoup 对象
    soup = BeautifulSoup(orihtml, 'lxml')
    # 采集每一个股票的信息
    count = 1
    list = []
    if None == soup.find('div', attrs={'class': 'cnt'}):
        fuck = 1
    else:
        for a in soup.find('div', attrs={'class': 'cnt'}).find_all('a', attrs={'target': "_blank"}):
            list.append(a.get_text().strip(' '))
        name = list[0].replace("\n", "").replace("\r", "").strip(' ')
        list = [name+stock_code]
        for a in soup.find_all(attrs={'class': 'tips-data-Right'}):
            num = a.get_text()
            if count%9==2:
                list.append(num)
                if len(list) == 34:
                    break
            count = count + 1
    return list


def main():
    # stock_list = get_data()
    # top_10 = heapq.nlargest(10, result, key=lambda r: float(r['data'][7].strip('%')))
    f = xlwt.Workbook()  # 创建工作簿
    '''
    创建第一个sheet:
      sheet1
    '''
    sheet1 = f.add_sheet(u'sheet1', cell_overwrite_ok=True)  # 创建sheet
    row0 = [u'股票(代码)', u'基本每股收益(元)', u'扣非每股收益(元)', u'稀释每股收益(元)', u'每股净资产(元)', u'每股公积金(元)', u'每股未分配利润(元)',\
            u'每股经营现金流(元)', u'营业总收入(元)', u'毛利润(元)', u'归属净利润(元)', u'扣非净利润(元)', u'营业总收入同比增长(%)',\
            u'归属净利润同比增长(%)', u'扣非净利润同比增长(%)', u'营业总收入滚动环比增长(%)', u'归属净利润滚动环比增长(%)', \
            u'扣非净利润滚动环比增长(%)', u'加权净资产收益率(%)', u'摊薄净资产收益率(%)', u'摊薄总资产收益率(%)', u'毛利率(%)',\
            u'净利率(%)', u'实际税率(%)', u'预收款/营业收入', u'销售现金流/营业收入', u'经营现金流/营业收入', u'总资产周转率(次)',\
            u'应收账款周转天数(天)', u'存货周转天数(天)', u'资产负债率(%)', u'流动负债/总负债(%)', u'流动比率', u'速动比率']

    # 生成第一行
    for i in range(0, len(row0)):
        sheet1.write(0, i, row0[i])
    num = 1
    while(1):
        stock_list = get_data()
        if len(stock_list) == 0:
            if 'n' == raw_input("NOT FIND!\nadd next code?：(y),(n):"):
                break
        else:
            for t in range(0, len(stock_list)):
                sheet1.write(num, t, stock_list[t])
            # 保存文件
            if 'n' == raw_input("add next code?：(y),(n):"):
                break
            num = num + 1
    f.save('demo.xls')


if __name__ == '__main__':
    main()


