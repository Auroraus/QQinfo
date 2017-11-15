# -*- coding: utf-8 -*-
"""
Created on Wed Nov 15 18:12:42 2017

@author: zf

email:1379875051@qq.com 或zf083415@gmail.com


"""
import tkinter
import xlrd
import Praise


data=xlrd.open_workbook(r'QQ.xls')
table=data.sheets()[0]
col= table.col_values(2)
box=[]
for i in col:
    if "qq.com" in i and "zone" not in i:
        box.append(i[:-7])

r=tkinter.Tk() #tkinter root初始化 
r.title('QQ说说下载器  制作者:张凡')#界面标题栏
r.geometry()#界面大小（自适应）

mu=tkinter.Menu(r)
fi=tkinter.Menu(mu,tearoff=False)
fi.add_command(label='开发者：张凡',command='callback')
fi.add_command(label='版本号：1.0.0 ',command='callback')
fi.add_command(label='软件介绍：用这款软件让你下载目标好友的全部说说！',command='callback')
fi.add_command(label='退出',command=r.destroy)
mu.add_cascade(label='关于',menu=fi)
r.config(menu=mu)

tkinter.Label(r,text='注意：电脑一定要有火狐浏览器才可以使用！！！').pack()
tkinter.Label(r,text='警告：严禁非法使用，侵犯个人隐私！！').pack()

tkinter.Label(r,text='你的QQ号').pack()
input1=tkinter.StringVar()
xen=tkinter.Entry(r,textvariable=input1,width=25)
input1.set('1379875051')
xen.pack()

tkinter.Label(r,text='你的QQ密码').pack()
input2=tkinter.StringVar()
xen1=tkinter.Entry(r,textvariable=input2,width=25)
input2.set('')
xen1.pack()


def mohu():
    QQ =input1.get() #input(u"输入QQ号：")
    password =input2.get() #input(u"输入QQ密码：")
    for number in range(len(box)):
        praise_spider = Praise.Praise(QQ,password,box[number])
        praise_spider.login_qzone()
    tkinter.Label(r,text='下载完成，信息已经保存！').pack()#第一个标签

tkinter.Label(r,text='                                                       ').pack()#第一个标签

tkinter.Button(r,text=('开始'),command=mohu,width=10,height=1,bg='blue').pack()
tkinter.Label(r,text='                                                       ').pack()#第一个标签

r.mainloop()


