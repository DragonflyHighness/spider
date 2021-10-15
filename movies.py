import os
import requests
import pandas as pd
from lxml import etree
import xlwt
import time
import random


def GetMovieList(link):
    res = requests.get(link, headers=headers, allow_redirects=False)
    res.encoding = 'utf-8'
    movies = []
    stars = []
    comments = []
    dates = []
    # print(res.text)
    selector = etree.HTML(res.text)
    temp = selector.xpath('//div[@class="item"]/div[@class="info"]/ul')
    pictrues = selector.xpath('//div[@class="item"]/div[@class="pic"]/a/img/@src')
    pic_title = selector.xpath('//div[@class="item"]/div[@class="pic"]/a/img/@alt')
    # print(pictrues)
    for element in temp:
        name = element.xpath('./li[@class="title"]/a/em/text()')
        comment = element.xpath('./li[4]/span[@class="comment"]/text()')
        star = element.xpath('./li[3]/span[1]/@class')
        date = element.xpath('./li[3]/span[@class="date"]/text()')
        if not star:
            star = ['No star']
        if not comment:
            comment = ['No comment.']
        movies.append(name[0])
        stars.append(star[0])
        comments.append(comment[0])
        dates.append(date[0])
    # print(movies)
    # print(comments)
    # print(stars)
    # print(dates)
    results = {'movie': movies, 'date': dates, 'comment': comments, 'rating': stars, 'picture': pictrues}
    return results


def exportexcel(result, filepath):
    dirs = os.listdir(filepath)  # 目录下文件
    order = ['Name', 'date', 'stars', 'comments']
    pd1 = pd.DataFrame.from_dict(result)
    if 'douban.xlsx' in dirs:
        df = pd.read_excel(os.path.join(filepath, 'douban.xlsx'))
        df = df.append(pd1, ignore_index=True)
        df.to_excel(os.path.join(filepath, 'douban.xlsx'), encoding='utf-8', index=False)
    else:
        pd1.to_excel(os.path.join(filepath, 'douban.xlsx'), encoding='utf-8', index=False)


if __name__ == '__main__':
    start_link = 'https://movie.douban.com/people/158707318/collect?start='  # 自行修改你的域名
    headers = {
        "Host": "movie.douban.com",
        "Referer": "https://www.douban.com/people/158707318/",
        'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; WOW64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/63.0.3239.132 '
                      'Safari/537.36 ',
        'Cookie': 'll="108288"; bid=sC04JLLd66Q; push_doumail_num=0; __utmv=30149280.15870; douban-fav-remind=1; ct=y; _vwo_uuid_v2=DC6410694325D8D34E871A55D20EC2398|a4947ebb95b86ee4720ccdc7694bfbf0; _vwo_uuid_v2=DC6410694325D8D34E871A55D20EC2398|a4947ebb95b86ee4720ccdc7694bfbf0; dbcl2="158707318:Vzg1SibqReQ"; gr_user_id=344cc8d4-628c-430a-8eba-d5a93d45bc88; __utmz=30149280.1629704489.24.4.utmcsr=baidu|utmccn=(organic)|utmcmd=organic; ck=D3aW; ap_v=0,6.0; push_noty_num=0; __utma=30149280.2126853780.1628495426.1630640734.1630905452.36; __utmc=30149280; __utma=223695111.2114745503.1628670511.1630640734.1630905458.5; __utmb=223695111.0.10.1630905458; __utmc=223695111; __utmz=223695111.1630905458.5.4.utmcsr=douban.com|utmccn=(referral)|utmcmd=referral|utmcct=/; _pk_ses.100001.4cf6=*; _pk_ref.100001.4cf6=["","",1630905458,"https://www.douban.com/"]; __utmt=1; __utmb=30149280.4.10.1630905452; _pk_id.100001.4cf6=e836fd6c50db952c.1628670511.5.1630906696.1630642404.'
    }  # 在Referer里修改你的域名
    depth = 54  # 你的页数
    lists = {'MOVIE': [], 'DATE': [], 'COMMENT': [], 'RATING': []}
    print("Start : %s" % time.ctime())
    for i in range(depth):
        link = start_link + str(15 * i)
        time.sleep(random.random() * 3)
        MovieInfo = GetMovieList(link)
        # print(MovieInfo)
        lists['MOVIE'].extend(MovieInfo['movie'])
        lists['DATE'].extend(MovieInfo['date'])
        lists['COMMENT'].extend(MovieInfo['comment'])
        lists['RATING'].extend(MovieInfo['rating'])
        for j in range(len(MovieInfo['picture'])):
            pic_url = MovieInfo['picture'][j]
    print(lists)
    exportexcel(lists, 'E:\Study\spider')
    print("End : %s" % time.ctime())
