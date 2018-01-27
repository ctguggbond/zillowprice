# -*- coding: utf-8 -*-
import os
import wx
from priceThread import PriceThread
from threading import Thread  
from wx.lib.pubsub import pub 

class myframe(wx.Frame):
    
    def __init__(self,parent, title):
        self.address = ""

        wx.Frame.__init__(self, parent, title=title,size=(600,500))
        self.control = wx.TextCtrl(self, -1,u"请选择要处理的文件\n", style=wx.TE_MULTILINE,)
        self.Show(True)
        filemenu = wx.Menu()
        menu_open = filemenu.Append(wx.ID_OPEN, U"打开文件", " ")
        menu_exit = filemenu.Append(wx.ID_EXIT, "Exit", "Termanate the program")

        menuBar = wx.MenuBar()
        menuBar.Append(filemenu, u"选项")
        self.SetMenuBar(menuBar)
        self.Show(True)

        #设置监听器
        self.Bind(wx.EVT_MENU, self.on_open, menu_open)
        self.Bind(wx.EVT_MENU, self.on_exit, menu_exit)
        #设置线程消息回调函数
        pub.subscribe(self.updateDisplay, "update")  
    
    def updateDisplay(self,msg):
        self.control.AppendText(msg)
    
    #退出程序
    def on_exit(self,e):
        self.Close(True)

    #打开文件
    def on_open(self,e):
        self.dirname = ''
        dlg = wx.FileDialog(self,"选择源文件", self.dirname, "",
        "xls file(*.xlsx) |*.xls?| all file(*.*) |*.*",wx.FD_OPEN)
        if dlg.ShowModal() == wx.ID_OK:
            self.filename = dlg.GetFilename()
            self.dirname = dlg.GetDirectory()
            self.address = os.path.join(self.dirname,self.filename)
            #调用线程处理
            PriceThread(self,self.address)
        dlg.Destroy()


if __name__ == '__main__':
    app = wx.App(False)
    frame = myframe(None, u'zillow price')
    app.MainLoop()