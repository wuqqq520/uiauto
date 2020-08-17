# -*- coding: utf-8 -*-
import json
import time
from appium import webdriver
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.by import By
from appium.webdriver.common.touch_action import TouchAction
import os
import xlrd
from xlutils.copy import copy
import xlwt

class devices_test():
    def __init__(self,device,appPackage,appActivity,server,appfile=None):
        """Constructor"""
        print(device,appPackage,appActivity,server)
        desired_caps = {}
        desired_caps['platformName'] = 'Android'
        desired_caps['platformVersion'] = '4.4'
        desired_caps['deviceName'] = device
        desired_caps['appPackage'] = appPackage
        if appfile != None:
            desired_caps['app'] = appfile
        desired_caps['appActivity'] = appActivity
        desired_caps['unicodeKeyboard'] = True    
        desired_caps['resetKeyboard'] = True        
        self.driver = webdriver.Remote(server, desired_caps)
        #self.driver.implicitly_wait(8) 
        self.Wait=3
        print(self.driver.current_activity)
        Popup=[('xpath','//android.widget.TextView[@resource-id="com.ellabookhome:id/confirm" and text="确定"'),('id','com.ellabookhome:id/update_close'),('id','com.ellabookhome:id/iv_market_close'),('id','com.ellabookhome:id/ivTaskNoticeClose')]
        for pop in Popup:
            che=eval('self.'+pop[0])(pop[1],2)
            if che!=None:
                che.click()
        print(self.driver.current_activity)
    
    def id(self,key,timing=None):
        if timing == None:
            timing=self.Wait
        try:
            wait = WebDriverWait(self.driver,timing,0.5).until(EC.presence_of_element_located((By.ID,key)),message='lllllllllllllll') 
            return wait
        except:
            return None
            
    def className(self,key,timing=None):
        if timing == None:
            timing=self.Wait        
        try:
            wait = WebDriverWait(self.driver,timing,0.5).until(EC.presence_of_element_located((By.CLASS_NAME,key))) 
            return wait
        except:
            return None 
            
    def name(self,key,timing=None):
        if timing == None:
            timing=self.Wait        
        try:
            wait = WebDriverWait(self.driver,timing,0.5).until(EC.presence_of_element_located((By.NAME,key))) 
            return wait
        except:
            return None 

    def xpath(self,key,timing=None):
        if timing == None:
            timing=self.Wait        
        try:
            wait = WebDriverWait(self.driver,timing,0.5).until(EC.presence_of_element_located((By.XPATH,key))) 
            return wait
        except:
            return None
            
    def tap(self,x,y):
        #self.driver.tap([(141,97)(903,196)],500)  
        TouchAction(self.driver).press(x=x, y=y).release().perform()
        
    def seipe(self,x,y,x2,y2,durat=500):
        time.sleep(2)
        self.driver.swipe(x,y,x2,y2,durat)
    def keyevent(self,keycode):
        self.driver.keyevent(keycode)
    def check(self,typ,content):
        comm=[]
        if typ=='Activity':
            for i in content:
                if i[1]!=self.driver.current_activity:
                    comm.append (i[1]+':Not found')  
        elif typ=='element':            
            for i in content:
                cck=eval('self.'+i[0])(i[1])
                if cck==None:
                    comm.append(i[2]+':('+i[0]+')'+i[1]+':Not found')
        return comm

    def homing(self):
        while(True):
            if self.driver.current_activity != 'com.ellahome.home.HomeActivity':
                self.driver.keyevent(4)  
            else:
                break
    def reset(self):
        self.driver.reset()
        time.sleep(5)
        self.popup(3)
    def popup(self,wait=3):
        Popup=[('xpath','//android.widget.TextView[@resource-id="com.ellabookhome:id/confirm" and text="确定"'),('id','com.ellabookhome:id/update_close'),('id','com.ellabookhome:id/iv_market_close'),('id','com.ellabookhome:id/ivTaskNoticeClose')]
        for pop in Popup:
            che=eval('self.'+pop[0])(pop[1],wait)
            if che!=None:
                che.click()        
            

            
        
            
def setp(case):
    step=case.split('->') 
    Notes=[]
    Result='Pass'
    for carryout in step:
        print(carryout)
        n=carryout.index(')')+1
        en=carryout[0:n]
        obje=carryout[n:]
        Event=['check']
        if en[0:en.index('(')] in Event:
            if en[0:en.index('(')]=='check':
                p=obje.split(',')
                content=[]
                for i in p:
                    if i in value:
                        content.append((value[i]['type'],value[i]['value'],i))   
                    else:
                        Notes.append(i+':Not collected')
                Result_step=devices.check(en[en.index('(')+1:en.index(')')], content)
                if len(Result_step)>0:
                    for r in Result_step:
                        Notes.append(r)
                    Result='Fail'
                         
        elif len(obje)>0:
            if obje in value:
                element=eval('devices.'+value[obje]['type'])(value[obje]['value'])    
                if element !=None:
                    print(element)
                    exec('element.'+en)
                else:
                    Notes.append(obje+':('+value[obje]['type']+')'+value[obje]['value']+',Not found') 
                    Result='Fail'
                    break
            else:
                Notes.append(obje+':Not collected')
                if Result !='Fail':
                    Result='Block'
                break
        else:
            exec('devices.'+en) 
            
        print(time.strftime('%Y-%m-%d %H:%M:%S',time.localtime(time.time())),devices.driver.current_activity)
    return (Result,Notes)
    
def readtext(filename):
    f = open(filename,"r")
    str =f.read().replace('\n','')
    f.close()
    data=json.loads(str)
    return data



value=readtext('data.txt')  
devices=devices_test('G2W7N16115000817', 'com.ellabookhome','com.ellahome.start.splash.SplashActivity', 'http://127.0.0.1:4723/wd/hub')

filenmae='D:\\Python\\test\\test.xls'

workbook = xlrd.open_workbook(filenmae)        # 打开工作簿
sheets = workbook.sheet_names()                # 获取工作簿中的所有表格
worksheet = workbook.sheet_by_name(sheets[0])  # 获取工作簿中所有表格中的的第一个表格
rows_old = worksheet.nrows                     # 获取表格中已存在的数据的行数
new_workbook = copy(workbook)                  # 将xlrd对象拷贝转化为xlwt对象
new_worksheet = new_workbook.get_sheet(0)      # 获取转化后工作簿中的第一个表格
for i in range(2,rows_old):
    reus=''
    Resu=setp(worksheet.cell(i,2).value)
    new_worksheet.write(i, 2, Resu[0])        # 追加写入数据，注意是从i+rows_old行开始写入    
    for n in Resu[1]:
        if len(reus)>1:
            reus=reus+'\r\n'+n
        else:
            reus=n
    print(type(reus),'reus:',reus)
    new_worksheet.write(i, 4, reus)
new_workbook.save(filenmae)                    # 保存工作簿


devices.driver.quit()

#devices.id('com.ellabookhome:id/image')
