# -*- coding: utf-8 -*-
from __future__ import unicode_literals
import time
import os
import urlparse
import urllib
import urllib2
import string
import logging
from pyExcelerator import *
import xlrd
from bs4 import BeautifulSoup
import json


INPUT_FILE_PATH = u'input.xls'
JU_ITEM_HEADER = [
        u'标题',
        u'描述',
        u'日期',
        u'现价',
        u'折扣',
        u'原价',
        u'销量',
        u'已售/预售',
        u'品牌',
        u'类型',
        u'图片链接',
        u'商品链接',
]


def get_file_logger(name, format='%(asctime)s %(levelname)s\n%(message)s\n', level=logging.DEBUG):
    log = logging.getLogger(name)
    if not log.handlers:
        handler = logging.FileHandler('%s/%s.log' % ('./logs', name))
        formatter = logging.Formatter(format)
        handler.setFormatter(formatter)
        log.addHandler(handler)
        log.setLevel(level)
    return log

def check_url(url):
    """
    如果 url头部没有加上http://则补充上，然后进行正则匹配，看url是否合法。
    """
    if not isinstance(url, unicode):
        return ''
    url_fields = list(urlparse.urlsplit(url))
    if url_fields[0] != 'http':
        url = 'http:' + url
    return url

class DealExcel(object):
    """
    将信息读出或存入excel文件中
    """

    def __init__(self, file_name, sheet_name):
        self.file_name = file_name
        self.sheet_name = sheet_name
        self.log = get_file_logger('excel_handler')

    def read_column_excel(self, column_num=1, start_row_num=1):
        """
        从某一行开始，把对应的某一列的值读取出来
        """
        try:
            file_info = xlrd.open_workbook(self.file_name)
            sheet_info = file_info.sheet_by_name(self.sheet_name)
            sheet_rows = sheet_info.nrows
            sheet_columns = sheet_info.ncols

            column_values = []
            if start_row_num > sheet_rows:
                start_row_num = 1
            elif column_num > sheet_columns:
                column_num = 0
            for i in range(start_row_num-1, sheet_rows):
                column_values.append(sheet_info.cell_value(i, column_num-1))
        except Exception, e:
            self.log.error(e)
            return []
        return column_values

    def excel_insert(self, excel_file_name, values, header, row_num=0):
        """
        将输入存入, 返回数据记录总数
        """
        w = Workbook()
        ws = w.add_sheet(u'聚划算报表')
        len_values = len(values)
        if row_num == 0:
            for index, item in enumerate(header):
                #print row_num,index,item
                try:
                    ws.write(row_num, index, item)
                except Exception, e:
                    self.log.error(e + str(item))
                    continue
            row_num = row_num + 1
        for index, value in enumerate(values):
            try:
                tmp = index + row_num
                ws.write(tmp, 0, value['name'])
                ws.write(tmp, 1, value['desc'])
                ws.write(tmp, 2, value['date_time'])
                ws.write(tmp, 3, value['price'])
                ws.write(tmp, 4, value['discount'])
                ws.write(tmp, 5, value['orig_price'])
                ws.write(tmp, 6, value['sold_num'])
                ws.write(tmp, 7, value['str_people'])
                ws.write(tmp, 8, value['brand_name'])
                ws.write(tmp, 9, value['item_type'])
                ws.write(tmp, 10, value['img_src'])
                ws.write(tmp, 11, value['src_detail'])
            except Exception, e:
                self.log.error(e)
                continue

        w.save(excel_file_name)
        return row_num

class GetPageData(object):
    '''
    输入url抓取网页信息，并保存unicode格式的文本信息
    '''
    log = get_file_logger('get_page')

    def __init__(self, url_list, brand_names=[]):
        self.urls = []
        self.brand_names = []
        for i in url_list:
            self.urls.append(check_url(i))
        for i in brand_names:
            self.brand_names.append(i)

        user_agent = 'Mozilla/6.0 (compatible; MSIE 5.5; Windows NT)'
        self.headers = {'User-Agent': user_agent}

    def get_page(self, url='', page_title=u'', decode_str='gbk'):
        """
        将url对应的界面数据以及名称打包成结果返回
        """
        try:
            page = ''
            if not url.strip():
                raise Exception('url is None')
            req = urllib2.Request(url, headers=self.headers)
            response = urllib2.urlopen(req)
            html = response.read()
            page = html.decode(decode_str)
        except Exception, e:
            self.log.error(e)
            return None

        result = {
            'data': page,
            'title': page_title,
        }
        return result

    def get_pages(self, page_title=u'', decode_str='gbk'):
        """
        针对ajax返回结果集
        """
        data = []
        for i in self.urls:
            try:
                if not i.strip():
                    raise Exception('i url is None')
                req = urllib2.Request(i, headers=self.headers)
                response = urllib2.urlopen(req)
                html = response.read()
                # page = html.decode(decode_str)
                data.append(html)
            except Exception, e:
                self.log.error(e)
                continue
        return data

    @classmethod
    def get_images(cls, url='', item_name='Unknown'):
        try:
            if not url.strip():
                raise Exception('url is None')
            dateline = time.strftime('%Y-%m-%d-%H-%M-%S', time.localtime(time.time()))
            path = './images/'
            if not os.path.isdir(path):
                os.makedirs(path)
            item_name = item_name.replace('/', '-')
            item_name = item_name.replace('*', '×')
            imgname = path+item_name+dateline+'.jpg'
            data = urllib.urlopen(url).read()
            f = file(imgname,'wb')
            f.write(data)
            f.close()
            time.sleep(0.2)
            return True
        except Exception, e:
            cls.log.error(e)
            return False

def get_image_input(url, brand_name):
    is_get_image = False
    print u"===》是否抓取的图片？（是：yes或y，否：任意其他输入） "
    flag = raw_input()
    if flag.strip().lower() in ['yes', 'y']:
        is_get_image = True
    return is_get_image

def say_hi_to_start(url, brand_name):
    print u"===》品牌名称：%s" % brand_name
    print u"===》url地址：%s" % url
    print u"===》开始抓取，制作报表......"

def check_page(data_page):
    if not isinstance(data_page, unicode):
        return ''
    else:
        return data_page

class GetJuFloor(object):
    """获取unicode中的floor数据"""
    def __init__(self, page, brand_name):
        self.unicode_page = check_page(page)
        self.log = get_file_logger('ju_report')
        self.brand_name = brand_name

    def get_floors(self):
        try:
            if not self.unicode_page.strip():
                raise Exception(u'unicode_page为空')
            small_floors = {
                'urls' : [],
                'data' : [],
            }
            floors = {
                'big': '',
                'small': small_floors,
                'brand_name': self.brand_name,
            }
            soup_page = BeautifulSoup(self.unicode_page)
            i = 1
            while True:
                floor = soup_page.find(id="floor%i" % i)
                if floor:
                    data_url = floor.attrs.get('data-url', '')
                    if data_url:
                        floors['small']['urls'].append(unicode(data_url.encode('utf-8')))
                        floors['small']['data'].append(floor)
                    else:
                        floors['big'] = floor
                    i = i + 1
                else:
                    break
        except Exception, e:
            self.log.error(e)
            return None
        return floors

class GetJuItem(object):
    """
    输入: 一个unicode的html页面或部分页面
    处理聚划算每个楼层的数据
    输出: 每个商品组成的item_list
    """

    def __init__(self, floor_data, brand_name):
        self.log = get_file_logger('ju_report')
        self.floor_data = floor_data
        self.brand_name = brand_name

    def get_big_items(self):
        try:
            if not self.floor_data:
                raise Exception(u'big floor_data为空')
        except Exception, e:
            self.log.error(e)
            return None

        result = []
        soup_li = self.floor_data.find_all('li')
        for i in soup_li:
            if not i.get('class'):
                pass
            else:
                try:
                    row = {
                        'name': '',
                        'desc': '',
                        'date_time': '',
                        'price': 0.0,
                        'discount': 0.0,
                        'orig_price': 0.0,
                        'sold_num': 0,
                        'str_people': u'',
                        'brand_name': '',
                        'item_type': u'热款',
                        'img_src': '',
                        'src_detail': '',
                    }
                    row['name'] = i.h3.get('title', '')
                    row['desc'] = i.find('ul', 'desc').text
                    row['date_time'] = time.strftime('%Y-%m-%d-%H:%M:%S',time.localtime(time.time()))
                    soup_price = i.find('span', 'price')
                    price = soup_price.find('span', attrs={'class': 'yen'}).text
                    try:
                        cent = soup_price.find('span', attrs={'class': 'cent'}).text
                        row['price'] = float(price + cent)
                    except Exception, e:
                        row['price'] = int(price)
                    try:
                        row['orig_price'] = string.atof(string.atof(soup_price.find('del', 'oriPrice').string.encode('gbk', 'ignore')))
                    except Exception, e:
                        row['orig_price'] = row['price']
                    row['discount'] = round(float(row['price'])/float(row['orig_price']), 2)
                    row['sold_num'] = string.atoi(i.find('div', 'soldcount').em.next)
                    row['str_people'] = u'件已付款'
                    row['brand_name'] = self.brand_name
                    row['item_type'] = u'热款'
                    row['img_src'] = i.img.attrs.get('data-ks-lazyload', '')
                    row['src_detail'] = 'http://detail.tmall.com/item.htm?id=%s' % i.a.attrs.get('href', '').split('item_id=')[1]
                except Exception, e:
                    self.log.error(e)
                    continue
                result.append(row)
        return result

    def get_small_items(self):
        result = []
        floor_dict = json.loads(self.floor_data)
        for i in floor_dict['itemList']:
            try:
                row = {
                    'name': i['name']['title'],
                    'desc': i['name']['longName'],
                    'date_time': time.strftime('%Y-%m-%d-%H:%M:%S', time.localtime(time.time())),
                    'price': float(i['price']['actPrice']),
                    'discount': float(i['price']['discount'])/10,
                    'orig_price': float(i['price']['origPrice']),
                    'sold_num': i['remind']['soldCount'],
                    'str_people': u'件已付款',
                    'brand_name': self.brand_name,
                    'item_type': u'普通',
                    'img_src': i['baseinfo']['picUrl'],
                    'src_detail': 'http://detail.tmall.com/item.htm?id=%s' % i['baseinfo']['itemId'],
                }
            except Exception, e:
                self.log.error(e)
                continue
            result.append(row)
        return result



def main():

    excel_handler = DealExcel(INPUT_FILE_PATH, 'Sheet1')
    ju_brands = excel_handler.read_column_excel(1, 2)
    ju_urls = excel_handler.read_column_excel(2, 2)
    ju_pages = GetPageData(ju_urls, ju_brands)
    result = []
    for i, j in zip(ju_urls, ju_brands):
        page = ju_pages.get_page(i, j)
        row_big_item = []
        row_small_item = []
        say_hi_to_start(i, j)
        result = []
        if not page:
            print u"===》抓取失败，请检查网路是否正常，按任意键退出......"
            failed = raw_input()
            return False
        else:
            floor = GetJuFloor(page['data'], page['title']).get_floors()
            time_start = time.strftime('%Y-%m-%d-%H-%M-%S',time.localtime(time.time()))
            row_big_item = GetJuItem(floor.get('big'), floor['brand_name']).get_big_items()
            result.extend(row_big_item)
            small_pages = GetPageData(floor['small'].get('urls'), floor['brand_name']).get_pages()
            for i in small_pages:
                row_small_item.extend(GetJuItem(i, j).get_small_items())
            result.extend(row_small_item)
            excel_handler.excel_insert(j+time_start+'.xls', result, JU_ITEM_HEADER)
            print u"%s 报表制作完成" % (j+time_start+'.xls')
            if get_image_input(i, j):
                for item in result:
                    GetPageData.get_images(check_url(item['img_src']), item['name'])
    print u"===》运行结束，按任意键退出......"
    success = raw_input()
    return True

if __name__ in "__main__":
    main()
