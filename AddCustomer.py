import pyautogui
import time
import random
import os
import pyperclip
from ReadExcl import ReadExcl
from ReadTxt import ReadTxt

from tkinter import *

class AddCustomer():
    def __init__(self):
        pyautogui.PAUSE = 1 # 调用在执行动作后暂停的秒数，只能在执行一些pyautogui动作后才能使用，建议用#time.sleep
        pyautogui.FAILSAFE = True # 启用自动防故障功能，左上角的坐标为（0，0），将鼠标移到屏幕的左上角，来抛出failSafeException异常

        self.exclhandle = ReadExcl("Phone")
        self.txthandle = ReadTxt("Log")
        
    def ChooseTypeToWork(self,configdata,loghandle):
        self.configdata = configdata
        self.loghandle = loghandle
        self.WriteLog('程序执行-开始')

        if self.configdata['type'] ==1:
            self.AddNewCustomer()
        elif self.configdata['type'] ==2:
            self.AddGroupCustomer()
        self.WriteLog('程序执行-结束')
        return
    
    def WriteLog(self,logdata):
        #命令窗口展示
        print(logdata)
        #日志文件展示
        self.txthandle.write_txt(logdata)
        #客户端展示
        self.loghandle.insert(1.0,time.strftime("%Y-%m-%d %H:%M:%S\n", time.localtime())+logdata + "\n")
        self.loghandle.update()
        
    #添加群客户--通过滑动添加
    def AddGroupCustomer(self):
        # try:
            MemX,MemY = 0,0,
            WehcatY,WehcatY = 0,0
            MemberPointArray =[]
            OldMemberPointArray = []
            
            AddContactX,AddContactY = 0,0
            ClearX,ClearY = 0,0
            SendX,SendY = 0,0
            SliderX,SliderY = 0,0
            ContrastX,ContrastY = 0,0
            ContrastXX,ContrastYY =0,0

            pyperclip.copy(self.configdata['word'])
            path = os.getcwd()

            while True:
                #获取所有当页成员的坐标
                MemberPointArray1 = self.GetAllPointOfPicture("Member1")
                MemberPointArray2 = self.GetAllPointOfPicture("Member2")
                #冒泡排序下
                MemberPointArray = self.BubbleSort(MemberPointArray1,MemberPointArray2,ContrastYY)
                if MemberPointArray == OldMemberPointArray:
                    break
                if len(MemberPointArray) ==0:
                    self.WriteLog('成员图片没有再屏幕出现,程序中断')
                    return 
                else:
                    for memberpoint in MemberPointArray:
                        #鼠标移动到成员上
                        pyautogui.click(memberpoint['pointx'],memberpoint['pointy'])

                        #点击添加按钮
                        AddContactX,AddContactY = self.GetOnePointOfPicture("AddContact")
                        if AddContactX != 0 or AddContactY != 0:
                            pyautogui.click(AddContactX,AddContactY)

                            #添加文案
                            ClearX,ClearY = self.GetOnePointOfPicture("Clear")
                            if ClearX != 0 or ClearY != 0:
                                pyautogui.click(ClearX,ClearY)
                            else:
                                break

                            #粘贴文案
                            pyautogui.click(ClearX-20,ClearY)
                            pyautogui.hotkey('ctrl', 'v')

                            #点击发送按钮
                            SendX,SendY = self.GetOnePointOfPicture("Send")
                            if SendX != 0 or SendY != 0:  
                                pyautogui.click(SendX,SendY)
                            else:
                                break
                        else:
                            self.WriteLog("该客户以添加为好友！")
                            continue
                        
                        if memberpoint == MemberPointArray[-1]:
                            #最后一个截取图片
                            pyautogui.screenshot(os.getcwd()+'\\Picture\\contrast.png', region=(memberpoint['pointx']-10,memberpoint['pointy']-10,150,20))
                OldMemberPointArray = MemberPointArray
                SliderX,SliderY = self.GetOnePointOfPicture("Slider")
                if SliderX != 0 or SliderY != 0:
                    pyautogui.moveTo(SliderX,SliderY)
                    ContrastX,ContrastY = self.GetOnePointOfPicture("contrast")
                    while ContrastX !=0 and ContrastY !=0:
                        if  ContrastYY != ContrastY:
                            SliderY = SliderY + 3.5
                            pyautogui.dragTo(SliderX,SliderY,1, button='left')
                            ContrastXX = ContrastX
                            ContrastYY = ContrastY
                            ContrastX,ContrastY = self.GetOnePointOfPicture("contrast")
                        else:
                            ContrastYY = ContrastY
                            break
                else:
                    break

        # except Exception as ex_results:
            # print(str(ex_results))
            # self.txthandle.write_txt(str(ex_results))
            # self.WriteLog('鼠标移至屏幕左上角坐标(0,0),程序中断！')

    #添加新客户
    def AddNewCustomer(self):
        try:
            AddCustomerX,AddCustomerY = 0,0
            SearchX,SearchY = 0,0
            AddPointArray=[]
            ClearX,ClearY =0,0
            SendX,SendY = 0,0

            phonedata = self.exclhandle.get_xls_next()
            pyperclip.copy(self.configdata['word'])

            path = os.getcwd()
            i=0
            while i <len(phonedata): 
                #鼠标点击添加按钮
                AddCustomerX,AddCustomerY= self.GetOnePointOfPicture("AddCustomer",number=i+1,identify=phonedata[i]["phone"])
                if AddCustomerX != 0 or AddCustomerY != 0:
                    pyautogui.click(AddCustomerX,AddCustomerY)
                else:
                    break

                #输入手机号
                pyautogui.typewrite(phonedata[i]["phone"])

                #鼠标点击搜索
                SearchX,SearchY = self.GetOnePointOfPicture("Search",number=i+1,identify=phonedata[i]["phone"])
                if SearchX != 0 or SearchY != 0:
                    pyautogui.click(SearchX,SearchY)
                else:
                    break

                pyautogui.press('enter')

                #鼠标点击好友添加
                AddPointArray = self.GetAllPointOfPicture("Add")
                if len(AddPointArray) ==0:
                    self.WriteLog('编号：'+str(i+1)+',识别号：'+phonedata[i]["phone"]+"不存在或者以添加为好友！")
                    i = i+1
                    continue
                else:
                    #点击第一个添加，通常情况下第一个为微信，第二个为企业微信
                    pyautogui.click(AddPointArray[0]['pointx'],AddPointArray[0]['pointy'])

                #添加文案，系统有记忆功能，只对前2个粘贴
                if i <2:
                    ClearX,ClearY = self.GetOnePointOfPicture("Clear",number=i+1,identify=phonedata[i]["phone"])
                    if ClearX != 0 or ClearY != 0:
                        pyautogui.click(ClearX,ClearY)
                    else:
                        break
                
                    #粘贴文案
                    pyautogui.click(ClearX-20,ClearY)
                    pyautogui.hotkey('ctrl', 'v')

                #点击发送按钮
                SendX,SendY = self.GetOnePointOfPicture("Send",number=i+1,identify=phonedata[i]["phone"])
                if SendX != 0 or SendY != 0:  
                    pyautogui.click(SendX,SendY)
                    self.WriteLog('编号'+str(i+1)+"成功,识别号:"+ phonedata[i]["phone"])
                    self.exclhandle.write_excl(phonedata[i])
                else:
                    break
                time.sleep(random.randint(self.configdata['time1']-8,self.configdata['time2']-8))
                    
                i = i+1
        except Exception as ex_results:
            print(str(ex_results))
            self.txthandle.write_txt(str(ex_results))
            self.WriteLog('鼠标移至屏幕左上角坐标(0,0),程序中断！')           

    def GetOnePointOfPicture(self,PictureName,OTx = 0,OTy =0,X=0,Y=0,number=0,identify=''):
        if X==0 and Y == 0:
            #获取当前目录
            path = os.getcwd()
            x = 0 
            y = 0
            # 如果截图没找到，pyautogui.locateOnScreen()函数返回None
            PicturePath = path + "\\Picture\\" + PictureName+".png"
            num = 1
            while num <=2:
                a = pyautogui.locateOnScreen(PicturePath,confidence=0.8)
                if a is not None:
                    break
                if identify!='' and number!=0:
                    self.WriteLog("编号："+str(number)+",识别号："+identify+"第"+str(num)+"次定位"+PictureName+"失败！重启定位！")
                else:
                    self.WriteLog("第"+str(num)+"次定位"+PictureName+"失败！重启定位！")
                num = num + 1

            if a is not None:
                x, y = pyautogui.center(a) # 获得文件图片在现在的屏幕上面的中心坐标
            else:
                if identify!='' and number!=0:
                    self.WriteLog("编号："+str(number)+",识别号："+identify+","+PictureName.split('\\')[-1].strip()+"未在屏幕中出现")
                else:
                    self.WriteLog(PictureName.split('\\')[-1].strip()+"未在屏幕中出现")
            return x+OTx,y+OTy
        else:
            return X,Y

    def GetAllPointOfPicture(self,PictureName,OTx = 0,OTy =0,AddPointArray=[]):
        if len(AddPointArray) == 0:
            path = os.getcwd()
            point = []
            PointArray = []
            x = 0 
            y = 0
            PicturePath = path + "\\Picture\\" + PictureName+".png"
            num = 1
            while num <=2:
                point = list(pyautogui.locateAllOnScreen(PicturePath,confidence=0.9))
                if len(point) !=0:
                    break
                self.WriteLog("第"+str(num)+"次定位"+PictureName+"失败！重启定位！")
                num = num +1

            if len(point) !=0:
                for pos in point:
                    pointxy = {}
                    x,y = pyautogui.center(pos)
                    pointxy["pointx"] = x+OTx
                    pointxy["pointy"] = y+OTy
                    PointArray.append(pointxy)
            return PointArray
        else:
            return AddPointArray

    def BubbleSort(self,MemberPointArray1,MemberPointArray2,ContrastYY):
        arr = MemberPointArray1+MemberPointArray2
        arr1 = []

        for i in range(len(arr)):
            for j in range(0,len(arr)-i-1):
                if arr[j]["pointy"] > arr[j+1]['pointy']:
                    arr[j], arr[j+1] = arr[j+1], arr[j]
        for point in arr:
            if point["pointy"] > ContrastYY:
                arr1.append(point)
        return arr1



    


