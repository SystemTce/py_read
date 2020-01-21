# ！ /usr/bin/python # 第一行是特殊注释行，称之为组织行，用来告诉我们GUN/Linux系统应该使用哪个解释器来执行该程序
# -*- coding: utf-8 -*-
# FileName: new.py
# Author: arry$tce
# Date:2019-05-29 22:01
# Python http 客户端，编写爬虫和测试服务器响应的库
import requests
import re
import urllib.request


# 掌握多门开发语言的技巧：掌握 语言特性
# java
# public static void main(String[] args){
#  System.out.println("太棒了")
# }
# Python
# if __name__=='__main__':
# print("太棒了")
# 输出 0-9
# for i in range(10):
# print(i)

def spiderPic(html, keyword):
    print('正在查找：'+keyword + '对应的图像，正在下载中，请稍等...')
    x = 0
    # for addr in re.findall('"objURL":"(.*?)"', html, re.S):
    
    for addr in re.findall('source srcset="(.*?)"', html, re.S):
        print(addr)
        imgres = requests.get(addr)
        with open("E:/bigdata/zhuajian/{}.jpg".format(x), "wb")as f:
            f.write(imgres.content)
            x += 1
            print("第",x,"张")


word = input('请输入关键字')
# result = requests.get('http://image.baidu.com/search/index?tn=baiduimage&fm=result&ie=utf-8&word='+word)
# result = requests.get('https://dribbble.com/search?q='+word)
result = requests.get('https://dribbble.com/search?q=%E6%8A%93')

# print(result.text)
spiderPic(result.text, word)
