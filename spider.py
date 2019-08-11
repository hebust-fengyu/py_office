import requests
import re
import time
import win32com
from win32com.client import Dispatch
import os




headers = {
    "Accept": "text/html,application/xhtml+xml,application/xml;q=0.9,image/webp,image/apng,*/*;q=0.8,application/signed-exchange;v=b3",
    "Accept-Encoding": "gzip, deflate, br",
    "Accept-Language": "zh-CN,zh;q=0.9,en;q=0.8",
    "Cache-Control": "max-age=0",
    "Connection": "keep-alive",
    "Cookie": "CURRENT_FNVAL=16; buvid3=61973FF1-1865-46F2-85F1-467EDB8E5452110258infoc; LIVE_BUVID=AUTO5515555107116764; UM_distinctid=16a2be281e0472-04282432a5351e-39395704-100200-16a2be281e15a9; fts=1555682345; rpdid=|(kmR~|Jluu|0J'ullYuuJm)Y; _uuid=9AF6BA14-719B-C99F-973B-A2817D068CD485522infoc; DedeUserID=417476917; DedeUserID__ckMd5=5f2fc44f8e7a260c; SESSDATA=0593dea5%2C1562164122%2Cd7741961; bili_jct=173bffbaf83dcdb6a7a45f9a86bccff2; finger=b3372c5f; stardustvideo=1; bp_t_offset_417476917=264345070884874854; _dfcaptcha=a5b006ee5fbdf9ce631f787139fa88e0; CURRENT_QUALITY=80; sid=la7ha6l2",
    "Host": "search.bilibili.com",
    "Referer": "https://www.bilibili.com/video/av34952152?from=search&seid=16008096283420799714",
    "Upgrade-Insecure-Requests": "1",
    "User-Agent": "Mozilla/5.0 (Windows NT 6.1; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/74.0.3729.169 Safari/537.36",
    }

def get_html(url):
    return requests.get(url, headers = headers)

def get_title_link(html_text):
    string_pattern = '''<a title="(.*?)" href="//(.*?)" target="_blank" class="title">.*?</a>'''
    return re.findall(string_pattern, html_text)

def main():
    app = Dispatch("excel.application")
    #excel = app.WorkBooks.Open(os.path.join(os.getcwd(), "bibi_acm"))
    app.visible = True
    app.WorkBooks.Add()
    app.WorkSheets.Add().name = "acm"
    sheet = app.WorkSheets("acm")
    row_count = 1
    for i in range(1, 44):
        url = "https://search.bilibili.com/all?keyword=acm&from_source=nav_search&page={}".format(i)
        html = get_html(url)
        html.encoding = html.apparent_encoding
        #print(html.apparent_encoding)
        res = get_title_link(html.text)
        for title in res:
            #f.write(title[0] + "\t" + title[1] + "\n")
            sheet.Cells(row_count, 1).value = title[0]
            string = '''=hyperlink(concatenate("https://","{}"))'''.format(title[1])
            print(string)
            sheet.Cells(row_count, 2).Value = string
            row_count += 1
        #print(type(res))
        print(i)
        time.sleep(0.11)
        

    print(dir(html))
    print(html)
    sheet.saveas(os.path.join(os.getcwd(),"bilibili"))
    del app
    
    #print(html.text)
    



if __name__ == "__main__":
    main()