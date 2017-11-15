# -*- coding: utf-8 -*-
"""
Created on Wed Nov 15 18:12:42 2017

@author: zf

email:1379875051@qq.com 或zf083415@gmail.com


"""

from selenium import webdriver
import time
import xlwt
import xlrd
import traceback
from xlutils.copy import copy


work=xlwt.Workbook()
sheet =work.add_sheet("说说")
style = xlwt.easyxf('font: bold 1, color red;')
sheet.row(0).height = 256 *3
sheet.col(0).width = 256 *20
sheet.col(1).width = 256 *20
sheet.col(2).width = 256 *100
sheet.col(3).width = 256 *20

sheet.write(0,0,"日期",style)                      
sheet.write(0,0+1,"名字",style)
sheet.write(0,1+1,"说说",style)
sheet.write(0,3,"QQ号",style)

weizhi="好友说说.xls"
work.save(weizhi)

class Praise():
    def __init__(self,QQ,password,friendQQ):    
        self.QQ=QQ
        self.n = 0
        self.logs=""
        self.praised_list = []       
        self.data=xlrd.open_workbook(r'F:\QQ.xls')
        self.table=self.data.sheets()[0]
        self.col= self.table.col_values(2)
        self.password=password
        self.friendQQ=friendQQ
        self.QQ=QQ
        self.QQ_box=[]
        self.old = xlrd.open_workbook('好友说说.xls',formatting_info=True)
        self.new = copy(self.old)
        self.news = self.new.get_sheet(0)
        self.sheet_data=self.old.sheet_by_index(0)
        self.col0= self.sheet_data.col_values(0)
        for co in self.col:
            if "qq.com" in co and "zone" not in co:
                self.QQ_box.append(co[:-7])
        if self.friendQQ:
            self.num=len(self.col0)
        if self.friendQQ:
            self.url="https://user.qzone.qq.com/"+self.friendQQ+"/311"
    def run(self):
        if self.friendQQ:
            
            self.browser.refresh()
            time.sleep(5)
            print("正在浏览：  ",self.friendQQ,"的空间")
            self.praise_someone()
            try:
                    
                    self.new.save(r'C:\\好友说说.xls')
                    print('文件保存在'+'C:\\好友说说.xls')
                    self.browser.quit()    
            except:
                traceback.print_exc()
                self.browser.quit()
        else:
                self.browser.quit()
        

    def login_qzone(self):
      try:
        self.browser = webdriver.Chrome(executable_path=r'C:\chromedriver\chromedriver.exe')
        #self.browser.maximize_window()
        self.browser.get(self.url)
        self.browser.switch_to.frame("login_frame")
        self.browser.find_element_by_id("switcher_plogin").click()
        self.browser.find_element_by_id("u").clear()
        self.browser.find_element_by_id("u").send_keys(self.QQ)
        self.browser.find_element_by_id("p").clear()
        self.browser.find_element_by_id("p").send_keys(self.password)
        self.browser.find_element_by_id("login_button").click()
        time.sleep(5)
        print ("登录成功")
        self.run()
        # 解决FireFox的登录成功后，直接访问新页面出现can't access dead object错误的方法链接：
        # http://stackoverflow.com/questions/16396767/firefox-bug-with-selenium-cant-access-dead-object
        # 通过下面这句解决，可能时因为上面switch_to到了login_frame，所以现在它是dead object
        self.browser.switch_to.default_content()
      except:
          self.browser.quit()
          

    def praise_someone(self):
        try:
            self.browser.switch_to.frame("app_canvas_frame")  # 个人主页才有  mod_pagenav_more
            try:
                s=self.browser.find_element_by_id("pager_last_0").text
                if s!=None:
                  
                  self.test=self.browser.find_elements_by_class_name("mod_pagenav_more")
                  self.pagenumbers=int(self.browser.find_element_by_css_selector("[id='pager_last_0']").text)
                  self.browser.switch_to.default_content()
                  if int(self.pagenumbers)>80:
                      self.pagenumbers=80
                  for page in range(1,self.pagenumbers+1):
                       self.browser.switch_to.frame("app_canvas_frame")  # 个人主页才有
                       self.browser.find_element_by_css_selector("input.textinput").send_keys(str(page))
                       self.browser.find_element_by_id("pager_gobtn_"+str(page-1)).click()
                       self.browser.switch_to.default_content()
                       time.sleep(2)
                       self.browser.switch_to.frame("app_canvas_frame")  # 个人主页才有
                       self.log_head = self.browser.find_element_by_id("msgList")
                       self.log_list=self.log_head.find_elements_by_class_name("feed")
                       self.start_praising()
                       self.browser.switch_to.default_content()
                       for i in range(50):
                           self.browser.execute_script("window.scrollBy(0,500);")
                       time.sleep(2)
                
            except:
                try:
                    self.pagenumbers=4
                    self.browser.switch_to.default_content()                      
                    for page in range(1,self.pagenumbers+1): 
                           self.browser.switch_to.frame("app_canvas_frame")  # 个人主页才有
                           self.browser.find_element_by_css_selector("input.textinput").send_keys(str(page))
                           self.browser.find_element_by_id("pager_gobtn_"+str(page-1)).click()
                           self.browser.switch_to.default_content()
                           time.sleep(2)
                           self.browser.switch_to.frame("app_canvas_frame")  # 个人主页才有
                           self.log_head = self.browser.find_element_by_id("msgList")
                           self.log_list=self.log_head.find_elements_by_class_name("feed")
                           self.start_praising()
                           self.browser.switch_to.default_content()
                           for i in range(50):
                               self.browser.execute_script("window.scrollBy(0,500);")
                           time.sleep(2)
                       
                except:
                    self.browser.switch_to.default_content()
                    self.browser.switch_to.frame("app_canvas_frame")  # 个人主页才有
                    self.log_head = self.browser.find_element_by_id("msgList")
                    self.log_list=self.log_head.find_elements_by_class_name("feed")
                    self.start_praising()
        except:
             print("无权限访问")
             self.browser.quit()
        
    def start_praising(self):
        for log in self.log_list:
            try:
                print('sfsgffsds')
                self.date=log.find_element_by_xpath("./div[3]/div[4]/div/span/a").get_attribute("title")
                # 说说内容
                self.name="【"+str(log.find_element_by_xpath("./div[3]/div[2]/a").text)+"】"
                
                self.text="【"+str(log.find_element_by_xpath("./div[3]/div[2]/pre").text)+"】"
                
                if not(self.text):
                       self.text=="【图片说说/转发无评论】" 
                if self.date =='' and self.name=='' and self.text=='':
                    break
                self.news.write(self.num,0,self.date)
                self.news.write(self.num,1,self.name)
                self.news.write(self.num,2,self.text)
                self.news.write(self.num,3,self.friendQQ)
                print(self.date)
                print(self.name)
                print(self.text)
                self.num+=1
            except:
                print(traceback.print_exc())
                continue
