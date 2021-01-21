import sys
import os
import time
class ReadTxt:
    def __init__(self,filename):
        proDir = os.getcwd()#获取当前目录
        self.txtPath = os.path.join(proDir, "{0}.txt".format(filename))
        self.caseList = []

    def write_txt(self,logdata):
        try:
            with open(self.txtPath,"a") as self.file:
                self.file.write(time.strftime("%Y-%m-%d %H:%M:%S\n", time.localtime())+logdata+"\n")
        except Exception as ex_results:
            print("程序终止,抓了一个异常：",ex_results,)
            os._exit(0)
        finally:
            self.close_txt()  
    
    def close_txt(self):
        self.file.close()
