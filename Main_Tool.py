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
        h = 500
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
        self.button_ensure = Button(self.mywindow,text ="开始执行",command = self.StartWork)
        self.button_ensure.place(x=30,y=100)

        self.label_log = Label(self.mywindow, text = '执行日志')
        self.label_log.place(x = 30, y = 140 )
        self.text_log = Text(self.mywindow,width = 60)
        self.text_log.place(x = 30, y = 170)

        self.label_alert = Label(self.mywindow, text = '执行过程中，请不要移动鼠标或者操作键盘！')
        self.label_alert.place(x = 100,y=90)
        self.label_alert = Label(self.mywindow, text = '执行过程中，如想中断程序，请将鼠标移至屏幕左上角！')
        self.label_alert.place(x = 100,y=110)

    def StartWork(self):
        configdata = {}
        configdata['type'] = self.type_default.get()
        configdata['word'] = self.word_default.get()

        if len(configdata['word']) == 0:
            self.text_log.insert(1.0,time.strftime("%Y-%m-%d %H:%M:%S\n", time.localtime())+"请填写添加话术!\n")
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