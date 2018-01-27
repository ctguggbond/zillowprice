import os
import requests
import re
import json
from bs4 import BeautifulSoup
import time
import wx
from threading import Thread
from wx.lib.pubsub import pub 
import xlrd
import xlwt

class PriceThread(Thread):
    def __init__(self,mainFram,address):
        Thread.__init__(self)
        self.address = address
        self.mainFram = mainFram
        self.headers = {
            "Host": "www.zillow.com",
            "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64; rv:57.0) Gecko/20100101 Firefox/57.0",
            "Accept": "text/html,application/xhtml+xml,application/xml;q=0.9,*/*;q=0.8",
            "Accept-Language": "zh-CN,zh;q=0.8,zh-TW;q=0.7,zh-HK;q=0.5,en-US;q=0.3,en;q=0.2",
            "Accept-Encoding": "gzip, deflate, br",
            "Connection": "keep-alive",
            "Upgrade-Insecure-Requests": "1",
            "Pragma": "no-cache",
            "Cache-Control": "no-cache",
        }
        self.session = requests.session()
        self.setDaemon(True)
        self.start()

    def run(self):
        f = xlwt.Workbook()
            #读取文件
        workbook = xlrd.open_workbook(self.address)
        sheet = workbook.sheet_by_index(0)
        
        urls = sheet.col_values(1)

        #写文件
        sheet1 = f.add_sheet('sheet1',cell_overwrite_ok=True) 
        sheet2 = f.add_sheet('sheet2',cell_overwrite_ok=True)

        row0 = ['url','date','event','price','$/sqft or agent' ,'source']
        for i in range(0,len(row0)):
            sheet1.write(0,i,row0[i])
        nextLine = 1
        errorline = 0
        for url in urls:
            #self.control.AppendText('正在处理:' + url + '\n')
            wx.CallAfter(self.sendData,'正在处理:' + url + '\n')
            #获取对应url信息
            try:
                resultRows = self.getPriceInfo(url.strip())
            except Exception as e:
                wx.CallAfter(self.sendData,url + ' : 处理失败\n')
                print(repr(e))
                sheet2.write(errorline,0,url)
                errorline += 1
                continue
            sheet1.write_merge(nextLine,nextLine+len(resultRows)-1,0,0,url)
            for i in range(0,len(resultRows)):
                for j in range(0,len(resultRows[i])):
                    sheet1.write(nextLine+i,j+1,resultRows[i][j])                    

            nextLine += len(resultRows)
        wx.CallAfter(self.sendData,'处理完成!\n')
 
        savepath = self.address #默认保存到与源文件同目录
        dlg2 = wx.FileDialog(self.mainFram,message=u"选择目标保存文件路径",  
                    defaultDir=os.getcwd(),  
                    defaultFile="",  
                    wildcard="*.xls",  
                    style=wx.FD_SAVE)  
        if dlg2.ShowModal() == wx.ID_OK:
            savepath = os.path.join(dlg2.GetDirectory(),dlg2.GetFilename())
        try:
            f.save(savepath)
            wx.CallAfter(self.sendData,'保存成功!  其中处理失败url在sheet2')
        except:
            wx.CallAfter(self.sendData,'保存失败!')
    #发送数据给主线程
    def sendData(self,message):
        pub.sendMessage("update",msg=message)

    #获取价格信息
    def getPriceInfo(self,url): 
        resp = self.session.post(url,headers=self.headers)
        soup = BeautifulSoup(resp.text,"lxml")
        table = soup.find_all('script',text=re.compile(r'.*hdp-price-history.*'))
        #print(resp.text)
        priceAndtaxUrl = re.findall(r"ajaxURL:\"(.+?)\"",table[0].string)

        datas = self.session.post("https://www.zillow.com" +priceAndtaxUrl[1] ,headers=self.headers)
        jdata = json.loads(datas.text)

        soup = BeautifulSoup(jdata.get('html'),"lxml")
        pricetable = soup.find('tbody')
        result = []
        for tr in pricetable.children:
           if len(tr) >= 5:
                row = [tr.contents[0].text,tr.contents[1].text,tr.contents[2].text,tr.contents[3].text,
                tr.contents[4].text]
            #历史数据不可见时只有一列
           elif len(tr) == 1:
                row = [tr.contents[0].text]
           result.append(row)
        return result
