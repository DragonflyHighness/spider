import os
import requests
import pandas as pd
from lxml import etree
import xlwt
import time
import random
import json

proxy = "proxy.lfk.qianxin-inc.cn"
proxies = {"http": "http://" + proxy,
           "https": "https://" + proxy}


def GetMovieList(link):
    res = requests.get(link, headers=headers, allow_redirects=True)
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
        "Referer": "https://movie.douban.com/people/158707318",
        'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; WOW64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/63.0.3239.132 '
                      'Safari/537.36 ',
        'Cookie': 'll="108288"; bid=sC04JLLd66Q; push_doumail_num=0; __utmv=30149280.15870; douban-fav-remind=1; _vwo_uuid_v2=DC6410694325D8D34E871A55D20EC2398|a4947ebb95b86ee4720ccdc7694bfbf0; gr_user_id=344cc8d4-628c-430a-8eba-d5a93d45bc88; _ga=GA1.2.2126853780.1628495426; push_noty_num=0; dbcl2="158707318:5LgU+CYX6l8"; ck=nQjZ; __utmc=30149280; __utmz=30149280.1639991661.115.16.utmcsr=accounts.douban.com|utmccn=(referral)|utmcmd=referral|utmcct=/; _pk_ref.100001.8cb4=["","",1640079427,"https://accounts.douban.com/"]; _pk_id.100001.8cb4=68985021adf8057f.1628495425.110.1640079427.1639992082.; _pk_ses.100001.8cb4=*; ap_v=0,6.0; __utma=30149280.2126853780.1628495426.1639991661.1640079427.116; __utmt=1; __utmb=30149280.2.10.1640079427'
    }
    depth = 54  # 你的页数
    lists = {'MOVIE': [], 'DATE': [], 'COMMENT': [], 'RATING': []}
    print("Start : %s" % time.ctime())
    # for i in range(depth):
    for i in range(1):
        link = start_link + str(15 * i)
        time.sleep(random.random() * 6)
        MovieInfo = GetMovieList(link)
        # print(MovieInfo)
        lists['MOVIE'].extend(MovieInfo['movie'])
        lists['DATE'].extend(MovieInfo['date'])
        lists['COMMENT'].extend(MovieInfo['comment'])
        lists['RATING'].extend(MovieInfo['rating'])
        # print(lists)
        for j in range(len(MovieInfo['picture'])):
            pic_url = MovieInfo['picture'][j]
    js = json.dumps(lists, ensure_ascii=False)
    # print(lists)
    # print(js)
    # exportexcel(lists, 'E:\Study\spider')
    print("End : %s" % time.ctime())
