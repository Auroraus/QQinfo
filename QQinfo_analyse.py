# -*- coding: utf-8 -*-
#@author:zf
import xlrd
import pylab
import matplotlib.pyplot as plt
import re

friend='某某'
zh=re.compile(u'[\u4e00-\u9fa5]+')
plt.rcParams['font.sans-serif']=['SimHei'] #用来正常显示中文标签
plt.rcParams['axes.unicode_minus']=False #用来正常显示负号
data=xlrd.open_workbook(r'C:\test\好友说说(zhong).xls')
table=data.sheets()[0]
col= table.col_values(2)
text=''
for i in col:
    aaa=i[1:-1]
    text=text+aaa
positive=['理想','奋斗','梦想','早起','相信','勇敢','信','美','担当','阳光','感恩','谢','乐','笑','嘿','哈哈','珍惜','操心','感激','暖','满足'\
          ,'怀念','支持','辛苦','实现','艰苦','美好','健康','激动']
negtive=['悲','伤','痛','累','泪','哭','苦','呜','滚','妈卖批','TMD','mmp','贱','无聊','智障','崩溃','睡不着','生气','解气','蠢','辜负'\
          ,'怨','恨','丑','仇','臭','死','杀','慌','谎','废','懒','孤独','毁','虚伪','压力','忘记','呵','绝望','离','倒霉','惨'\
          ,'烦','恼','啊啊','恐','祸','病'] 
study=['学','单词','自习','学习','学霸','考','录取','同学','学生','学校','老师','打卡','英语','大学生','单片机','比赛','协会']
eat=['吃','饭','菜','外卖','餐','面','米','麻辣','土豆','盐','水果','烧烤','肉','摊','小吃','火锅','酒店','食'\
          ,'瓜','果']
play=['跑步','健身','旅游','游','骑','走路','玩','登','开车','照','相机','山']
music=['音乐','演唱会','琴','哼','唱','表演','晚会','网易云','酷狗','听','歌','曲']
dog=['单身狗','一万点','情人劫','电灯泡','单身','脱单']
love=['亲爱的','爱情','恋','喜欢','浪漫','幸福','么么哒','感情'] 
pos=0 
neg=0
stu=0
ea=0
pla=0
mus=0
do=0
lov=0

pos1=0 
neg1=0
stu1=0
ea1=0
pla1=0
mus1=0
do1=0
lov1=0
for m in col:
    for n in positive:
        if n in m:
            pos1=pos1+1
for m in col:
    for n in negtive:
        if n in m:
            neg1+=1   
for m in col:
    for n in study:
        if n in m:
            stu1+=1
for m in col:
    for n in eat:
        if n in m:
            ea1+=1
for m in col:
    for n in play:
        if n in m:
            pla1+=1
for m in col:
    for n in music:
        if n in m:
            mus1+=1
for m in col:
    for n in dog:
        if n in m:
            do1+=1
for m in col:
    for n in love:
        if n in m:
            lov1+=1 
            
for m in col:
    for n in positive:
        if n in m:
            pos=pos+1
            
            break
for m in col:
    for n in negtive:
        if n in m:
            print(m,n)
            neg+=1
            break
for m in col:
    for n in study:
        if n in m:
            stu+=1
            break
for m in col:
    for n in eat:
        if n in m:
            ea+=1
            break
for m in col:
    for n in play:
        if n in m:
            pla+=1
            break
for m in col:
    for n in music:
        if n in m:
            mus+=1
            break
for m in col:
    for n in dog:
        if n in m:
            do+=1
            break
for m in col:
    for n in love:
        if n in m:
            lov+=1
            break
def bar(word,word_fre,xlabel,ylabel,title,weizhi):
    pylab.figure(figsize=(9,9))
    #pylab.pie(single_word_fre,labels=single_word,labeldistance = 1.1,autopct = '%3.1f%%',shadow = False,startangle = 90,pctdistance = 0.6)
    pylab.bar(range(len(word)),word_fre,color='green',tick_label=word)
    pylab.xlabel(xlabel,fontproperties='SimHei',fontsize=20)
    pylab.ylabel(ylabel,fontproperties='SimHei',fontsize=20)
    pylab.title(title,fontproperties='SimHei',fontsize=20)
    plt.xticks(rotation=90)
    pylab.savefig(weizhi)
ran=[pos,neg,stu,ea,pla,mus,do,lov] 
rano=['积极生活态度','消极生活态度','学习相关','吃货一枚','爱运动，爱游玩','音乐','单身狗','秀恩爱的']   
bar(rano,ran,'','出现频率','相关说说分布',"F:\\说说1.png")    

rano=[pos1,neg1,stu1,ea1,pla1,mus1,do1,lov1] 
ranoo=['积极生活态度','消极生活态度','学习相关','吃货一枚','爱运动，爱游玩','音乐','单身狗','秀恩爱的']   
bar(ranoo,rano,'','出现频率','相关词汇分布',"F:\\说说2.png")    
# 年
timelist=[]
time_fre=[]
time_value=[]
time=table.col_values(0)
for ti in time[1:]:
    timelist.append(int(ti[:4]))
timm={}
key_list=[]
value_list=[]
for a in timelist:
    if a in timm:
        timm[a]+=1
    else:
        timm[a]=1
for key,value in timm.items():
        key_list.append(key)
        value_list.append(value)
time_value=sorted(key_list)
for nn in range(len(time_value)):
    time_fre.append(timm[time_value[nn]])
  #月
mlist=[]
m_fre=[]
m_value=[]
for ti in time[1:]:
    if '月' in ti[5:7]:
        mlist.append(int(ti[5:6]))
    else:
        mlist.append(int(ti[5:7]))
mmm={}
key_list=[]
value_list=[]
for a in mlist:
    if a in mmm:
        mmm[a]+=1
    else:
        mmm[a]=1
for key,value in mmm.items():
        key_list.append(key)
        value_list.append(value)
m_value=sorted(key_list)
for nn in range(len(m_value)):
    m_fre.append(mmm[m_value[nn]])       

avr_text=int(table.nrows)/len(time_fre)
def word_fre(art,num):
    key_list=[]
    value_list=[]
    word=[]
    word_num=[]
    word_fre=[]
    artical=art
    string={}
    number=int(num)
    for ii in range(len(artical)-number):
        word.append(artical[ii:ii+number])

    for aa in (word):
        if aa in string:
            string[aa]+=1
        else:
            string[aa]=1
    for value,key in string.items():
        key_list.append(key)
        value_list.append(value)
    key=sorted(key_list,reverse=True)
    for n in range(100):
        for k,v in string.items():
            if (v==key[n]) and (k not in word_fre) and zh.search(k[0]) and('，'or'.'or'。') and zh.search(k[-1]):
                word_fre.append(k)
                word_num.append(key[n])
            else:
                pass
    return word_fre,word_num

def bar(word,word_fre,xlabel,ylabel,title,weizhi):
    pylab.figure(figsize=(9,9))
    #pylab.pie(single_word_fre,labels=single_word,labeldistance = 1.1,autopct = '%3.1f%%',shadow = False,startangle = 90,pctdistance = 0.6)
    pylab.bar(range(len(word)),word_fre,color='green',tick_label=word)
    pylab.xlabel(ylabel,fontproperties='SimHei',fontsize=15)
    pylab.ylabel(ylabel,fontproperties='SimHei',fontsize=15)
    pylab.title(title,fontproperties='SimHei',fontsize=20)
    plt.xticks(rotation=90)
    pylab.savefig(weizhi)
    pylab.legend()   #是plot函数中的label标签生效，下同，不再赘述
def pie(word,word_fre,title,weizhi):
    pylab.figure(figsize=(9,9))
    pylab.pie(word_fre,labels=word,labeldistance = 1.1,autopct = '%3.1f%%',shadow = False,startangle = 90,pctdistance = 0.6)
    pylab.title(title,fontproperties='SimHei',fontsize=20)
    pylab.savefig(weizhi)
    pylab.legend()   #是plot函数中的label标签生效，下同，不再赘述
print('平均每年发表说说条数：',int(avr_text))


bar(time_value,time_fre,'','年份',friend+'说说随年份的变化情况','C:\\test'+friend+'的说说bar10.png')
bar(m_value,m_fre,'','月份',friend+'说说随月份的变化情况','C:\\test'+friend+'的说说bar11.png')
pie(time_value,time_fre,friend+'说说随年份的变化情况','C:\\test'+friend+'的说说pie12.png',)
pie(m_value,m_fre,friend+'说说随月份的变化情况','C:\\test'+friend+'的说说pie12.png',)

word1,word_fre1=word_fre(text,1)
bar(word1,word_fre1,'','出现频率',friend+'说说中出现频率最高的一字词','C:\\test'+friend+'的说说bar2.png')
pie(word1,word_fre1,friend+'说说中出现频率最高的一字词','C:\\test'+friend+'的说说pie.png',)

word2,word_fre2=word_fre(text,2)
bar(word2,word_fre2,'','出现频率',friend+'说说中出现频率最高的两字词','C:\\test'+friend+'的说说bar3.png')
pie(word2,word_fre2,friend+'说说中出现频率最高的两字词','C:\\test'+friend+'的说说pie4.png')

word4,word_fre4=word_fre(text,4)
bar(word4,word_fre4,'','出现频率',friend+'说说中出现频率最高的四字词','C:\\test'+friend+'的说说bar7.png')
pie(word4,word_fre4,friend+'说说中出现频率最高的四字词','C:\\test'+friend+'的说说pie8.png')


   
