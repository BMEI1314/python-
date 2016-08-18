# -*- coding: cp936 -*-
import urllib.request
import re
import os
import time
import xlwt
def movie(movieTag):
  tagUrl=urllib.request.urlopen(url)
  tagUrl_read = tagUrl.read().decode('utf-8')
  return tagUrl_read
def subject(tagUrl_read):
#������ʽƥ���Ӱ�����֣����ӣ�������������
  nameURL= re.findall(r'\s+title="(.+)"',tagUrl_read)
  scoreURL = re.findall(r'<span\s+class="rating_nums">([0-9.]+)<\/span>',tagUrl_read)
  evaluateURL = re.findall(r'<span\s+class="pl">\((\w+)������\)<\/span>',tagUrl_read)
  movieLists = list(zip(nameURL,scoreURL,evaluateURL))
  newlist.extend(movieLists)
  return newlist
def find_imgs(url):
    html = movie(url)
    img_addrs = []
    a = html.find('img src=')
    while a != -1:
        b = html.find('.jpg', a, a+255)
        if b != -1:
            string=html[a+9:b+4].replace("thumb","photo")
            img_addrs.append(string)  
        else:
            b = a + 9
        a = html.find('img src=', b)
    return img_addrs
def save_imgs(img_addrs,no): 
    for each in img_addrs:
        filename=str(no)
        with open(filename, 'wb') as f: 
         response = urllib.request.urlopen(each)
         img = response.read()
         f.write(img)
         no+=1
#��quote�������⣨���ģ��ַ�
movie_type = urllib.request.quote(input('�������Ӱ����(����顢ϲ�硢����)��'))

page_end=int(input('��������������ʱ��ҳ�룺'))
file1=urllib.request.unquote(movie_type)
num_end=page_end*20

num=0
page_num=1
count=1
no=1
newlist=[]
os.mkdir(file1)
os.chdir(file1)

while num<num_end:

  url=r'http://movie.douban.com/tag/%s?start=%d'%(movie_type,num)

  movie_url = movie(url)
  subject_url=subject(movie_url)
  img_addrs = find_imgs(movie_url)
  save_imgs(img_addrs,no)
  num=page_num*20
  page_num+=1

else:

#ʹ��sorted�������б�������У�reverse����ΪTrueʱ����Ĭ�ϻ�FalseʱΪ���� key=lambda�����Ǻ����������ԭ����

  movieLIST = sorted(newlist, key=lambda movieList : movieList[1],reverse = True)
  f = open(file1+'.txt','w')
  file=xlwt.Workbook()
  table=file.add_sheet('data')
  table.write(0,0,'name')
  table.write(0,1,'score')
  table.write(0,2,'evaluate')
  for i in range(len(movieLIST)):
     m=i+1
     table.write(m,0,movieLIST[i][0])
     table.write(m,1,movieLIST[i][1])
     table.write(m,2,movieLIST[i][2])
  table = file.add_sheet('sheet_name',cell_overwrite_ok=True)
  file.save(str(file1)+'.xls')     # �����ļ�

   # ���⣬ʹ��style
  style = xlwt.XFStyle()    # ��ʼ����ʽ
  font = xlwt.Font()        # Ϊ��ʽ��������
  font.name = 'Times New Roman'
  font.bold = True
  style.font = font         #Ϊ��ʽ��������
  table.write(0, 0, 'some bold Times text', style) # ʹ����ʽ
  for movie in movieLIST:
    k=str(count)+":"+str(movie)
    f.write(k+"\n")
    print(movie)
    count+=1
  f.close()
time.sleep(3)
print('����')
