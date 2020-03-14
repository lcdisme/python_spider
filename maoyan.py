import requests
import re
import xlwt
import os
import time


#构造猫眼的url
def build_url_spider(page):
    url_init = 'http://maoyan.com/board/4?offset='
    page = [i*10 for i in range(page)]
    for i in page:
        yield url_init + str(i)


#爬取单个网页，使用requests库
def get_html_spider(url):
    try:
        headers = {
            'User-Agent': 'Mozilla/5.0 (Macintosh; Intel Mac OS X 10_13_3) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/65.0.3325.162 Safari/537.36'
        }
        r = requests.get(url,headers = headers)
        if r.status_code == 200:
            return r.text
        else:
            print ('网络访问失败')
    except:
        print ('get url 环节发生错误')


#解析单个网页，使用最原始的正则表达式匹配，返回电影信息列表
def parse_html(html):
    pattern = re.compile(r'<dd>.*?board-index.*?>(\d+)</i>'
                         +r'.*?<p class="name"><a href=".*?" title=".*?" data-act="boarditem-click" data-val="{movieId:.*?}">(.*?)</a></p>'
                         +r'.*?<p class="star">(.*?)</p>'
                         +r'.*?<p class="releasetime">(.*?)</p>'
                          r'.*?<p class="score"><i class="integer">(.*?)</i>',re.S)
    items = re.findall(pattern,html)
    for item in items:
        rank = item[0]
        name = item[1]
        actor = item[2].strip()
        on_time = item[3]
        score = item[4].strip('.')
        yield [rank,name,actor,on_time,score]


#将汇总的电影信息列表写入excel
def write_to_excel(movie_info):
    workbook = xlwt.Workbook (encoding='utf-8')
    worksheet = workbook.add_sheet ('My Worksheet')
    for i in range(len(movie_info)):
        for j in range(len(movie_info[1])):
            worksheet.write (i,j, label=movie_info[i][j])
    if os.path.exists('MaoYan_Movies.xls'):
        os.remove('MaoYan_Movies.xls')
    workbook.save ('MaoYan_Movies.xls')

#主程序框架
def main(page):
    movie_info = []
    for url in build_url_spider (page):
        html = get_html_spider(url)
        movie_info_single = parse_html (html)
        movie_info += (list(movie_info_single))
        time.sleep(0.5)
    write_to_excel (movie_info)


if __name__ == '__main__':
    main(10)
