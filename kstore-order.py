# _*_ encoding:utf-8 _*_

import urllib
import urllib2
import cookielib
import os
import sys
import re
import time
from bs4 import BeautifulSoup
from openpyxl import Workbook

reload(sys)
sys.setdefaultencoding('utf-8')

"""
帮人写的一个爬Kstore会员订单的爬虫
"""

login_url = 'http://kstore.qianmi.com/checklogin.htm'
order_url = 'http://kstore.qianmi.com/myorder.htm?pageNo='

cookie = cookielib.CookieJar()
opener = urllib2.build_opener(urllib2.HTTPCookieProcessor(cookie))

order_code_pattern = re.compile(r'\d{16}')
order_amount_pattern = re.compile(r'\d+\.\d{2}')

login_data = {
    'username': 'qianmi',
    'password': 'qianmi520',
    'url': 'index.html',
    'type': '0'
}

def login():
    request = urllib2.Request(login_url, urllib.urlencode(login_data))
    response = opener.open(request)
    print response.read()


def crawl_order():
    page_no = 0
    order_list = []
    while True:
        page_no += 1
        request = urllib2.Request(order_url+str(page_no))
        response = opener.open(request)
        html = response.read()
        soup = BeautifulSoup(html)
        order_list_div = soup.find('div', {'class': 'new_order_list'})
        order_body_list = order_list_div.find_all('tbody')
        for order in order_body_list:
            order_code = get_order_code(order)
            print '订单 %s ====================' % order_code
            order_goods = get_order_goods(order)
            customer_name = get_customer(order)
            order_amount = get_order_amount(order)
            print '=========================\n'
            order_list.append([order_code, order_goods, customer_name, order_amount])
        if not has_next_page(order_list_div):
            break
    return order_list


def has_next_page(order_content):
    non_next_btn = order_content.find('div', {'class': 'paging_area'}).find('a', {'class': 'next_null'})
    if non_next_btn:
        return False
    else:
        return True


def write_excel(order_list):
    wb = Workbook()
    ws = wb.active
    ws = wb.create_sheet()
    ws.title = 'customer orders'
    ws.append(['订单号', '订单商品', '会员名', '订单金额'])
    for order in order_list:
        ws.append(order)
    save_path = 'order_list.' + str(time.time()) + '.xlsx'
    print save_path
    wb.save(save_path)


def get_order_amount(order):
    order_content = order.find('tr', {'class': 'order-bd'})
    amount_text = order_content.find_all('td')[3].get_text()
    amount_match = order_amount_pattern.search(amount_text)
    if amount_match:
        print amount_match.group()
        return amount_match.group()
    else:
        return 00.00


def get_customer(order):
    order_content = order.find('tr', {'class': 'order-bd'})
    customer = order_content.find_all('td')[1].get_text()
    name = ' '.join(customer.split())
    print name
    return name


def get_order_goods(order):
    order_content = order.find('tr', {'class': 'order-bd'})
    order_goods_descs = order_content.find('td', {'class': 'baobei'})\
    .find_all('div', {'class': 'desc'})
    order_goods = []
    for goods_desc in order_goods_descs:
        goods_title = goods_desc.find('a', {'class': 'name'}).get_text()
        print goods_title
        order_goods.append(goods_title)
    return '\n'.join(order_goods)


def get_order_code(order):
    order_header = order.find('tr', {'class': 'order-hd'})
    order_code_text = order_header.find('td', {'class': 'first'}).find('span').get_text()
    code_match = order_code_pattern.search(order_code_text)
    if code_match:
        return code_match.group()
    else:
        return ''


def test_re():
    text = '胜多负少12345678'
    pattern = re.compile(r'\d+')
    match = pattern.search(text)
    if match:
        print match.group()
    else:
        print '不匹配'


def test_write_excel():
    wb = Workbook()
    ws = wb.active             #默认创建第一个表，默认名字为sheet
    ws1 = wb.create_sheet()    #创建第二个表
    ws1.title = "New Title"    #为第二个表设置名字
    ws2 = wb.get_sheet_by_name("New Title")                #通过名字获取表，和第二个表示一个表
    wb.save('your_name.xlsx') #保存

if __name__ == '__main__':
    login()
    # crawl_order()
    write_excel(crawl_order())
    # test_re()
    # test_write_excel()
