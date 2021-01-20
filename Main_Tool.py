from tkinter import *
import tkinter as tk
import hashlib
import time
import os
import sys
import threading

from AddCustomer import AddCustomer
# import PyQt5.sip#用于打包使用其他情况注释掉

class MY_GUI():
    def __init__(self,MYWindow):
        self.mywindow  = MYWindow

    def set_init_window(self):
        w = 500
        h = 550
        ws = self.mywindow.winfo_screenwidth()
        hs = self.mywindow.winfo_screenheight()

        self.mywindow.geometry('%dx%d+%d+%d' % (w, h, ws-w-100,hs-h-100))
        self.mywindow.title("医客通-添加客户小工具V1.0")
        self.mywindow.iconphoto(False,tk.PhotoImage(file=os.getcwd() + "\\Picture\\windowicon.png"))
        
        self.type_default = IntVar()
        self.label_type = Label(self.mywindow, text = '工具类型:')
        self.label_type.place(x = 30, y = 30)
        self.radiobutton_type_one = Radiobutton(self.mywindow, text='添加新客户', value=1, variable=self.type_default)
        self.radiobutton_type_one.place(x=150,y=30)
        self.radiobutton_type_group = Radiobutton(self.mywindow, text='添加群客户', value=2, variable=self.type_default)
        self.radiobutton_type_group.place(x=250,y=30)
        self.radiobutton_type_group = Radiobutton(self.mywindow, text='频率测试', value=3, variable=self.type_default)
        self.radiobutton_type_group.place(x=350,y=30)
        self.type_default.set(1)

        self.word_default = StringVar()
        self.label_word = Label(self.mywindow, text = '添加话术:')
        self.label_word.place(x = 30, y = 60 )
        self.entry_word = Entry(self.mywindow,width = 40,textvariable = self.word_default)
        self.entry_word.place(x=150,y=60)

        self.time1_default = IntVar()
        self.time2_default = IntVar()
        self.label_time = Label(self.mywindow, text = '秒(每次):')
        self.label_time.place(x = 30, y = 100 )
        self.entry_time = Entry(self.mywindow,width = 5,textvariable = self.time1_default)
        self.entry_time.place(x=150,y=100)
        self.label_time = Label(self.mywindow, text = '~')
        self.label_time.place(x = 175, y = 100 )
        self.entry_time = Entry(self.mywindow,width = 4,textvariable = self.time2_default)
        self.entry_time.place(x=190,y=100)
        self.label_time = Label(self.mywindow, text = '添加单个客户时间间隔不得小于11秒!')
        self.label_time.place(x = 225, y = 100 )

        self.button_ensure = Button(self.mywindow,text ="开始执行",command = self.StartWork)
        self.button_ensure.place(x=30,y=140)

        self.label_log = Label(self.mywindow, text = '执行日志')
        self.label_log.place(x = 30, y = 180 )
        self.text_log = Text(self.mywindow,width = 60)
        self.text_log.place(x = 30, y = 210)

        self.label_alert = Label(self.mywindow, text = '执行过程中，请不要移动鼠标或者操作键盘！')
        self.label_alert.place(x = 100,y=135)
        self.label_alert = Label(self.mywindow, text = '执行过程中，如想中断程序，请将鼠标移至屏幕左上角！')
        self.label_alert.place(x = 100,y=155)

    def StartWork(self):
        configdata = {}
        configdata['type'] = self.type_default.get()
        configdata['word'] = self.word_default.get()
        configdata['time1'] = self.time1_default.get()
        configdata['time2'] = self.time2_default.get()

        if len(configdata['word']) == 0:
            self.text_log.insert(1.0,time.strftime("%Y-%m-%d %H:%M:%S\n", time.localtime())+"请填写添加话术!\n")
            return 
        if configdata['time1'] == 0 or configdata['time1'] == 0:
            self.text_log.insert(1.0,time.strftime("%Y-%m-%d %H:%M:%S\n", time.localtime())+"添加单个客户时间间隔不得为0!\n")
            return 
        if configdata['time1'] <= 10:
            self.text_log.insert(1.0,time.strftime("%Y-%m-%d %H:%M:%S\n", time.localtime())+"添加单个客户时间间隔不得小于11秒!\n")
            return
        if configdata['time2'] < configdata['time1']:
            self.text_log.insert(1.0,time.strftime("%Y-%m-%d %H:%M:%S\n", time.localtime())+"最大时间不得小于最小时间!\n")
            return
        self.button_ensure['text'] ="执行中"
        self.button_ensure['state'] = DISABLED
        AddCustomer().ChooseTypeToWork(configdata,self.text_log)
        self.button_ensure['text'] ="开始执行"
        self.button_ensure['state'] = NORMAL
            

if __name__ =='__main__': 
    my_window  = Tk()
    ZMJ_PORTAL = MY_GUI(my_window)
     # 设置根窗口默认属性
    ZMJ_PORTAL.set_init_window()
    my_window.mainloop()