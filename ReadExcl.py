import xlrd
import xlwt
from xlutils.copy import copy
import os
from datetime import datetime
from xlrd import xldate_as_tuple

class ReadExcl:
    def __init__(self,exclname):
        self.exclname = exclname
        proDir = os.getcwd()#获取当前目录
        self.ExclPath = os.path.join(proDir, "{0}.xls".format(exclname))

    def read_excl(self):
        try:
            self.readfile = xlrd.open_workbook(self.ExclPath,'w',formatting_info= True)
        except Exception as ex_results:
            print("程序终止,抓了一个异常：",ex_results,)
            os._exit(0)

    def write_excl(self,phonedata):
        self.read_excl()
        try:
            self.newfile = copy(self.readfile)
            newsheet = self.newfile.get_sheet(self.exclname)
            newsheet.write(int(phonedata["id"]),2,"success")
            self.save_excl()
        except Exception as ex_results:
            print("程序终止,抓了一个异常：",ex_results,)
            os._exit(0)

    def save_excl(self):
        self.newfile.save(self.ExclPath)

    #遍历sheet中的用例
    def get_xls_next(self):
        self.read_excl()
        sheet = self.readfile.sheet_by_name(self.exclname)

        merged = self.merge_cell(sheet)  ###合并单元格

        row = sheet.row_values(0)
        rowNum  = sheet.nrows#获取行
        colNum = sheet.ncols #获取列
        
        cls = []
        curRowNo = 1
        while self.hasNext(rowNum,curRowNo):
            s = {}  
            col = sheet.row_values(curRowNo) 
            i = colNum  
            for x in range(i):

                if merged.get((curRowNo,x)):###合并单元格
                    col[x] = sheet.cell_value(*merged.get((curRowNo,x))) ###合并单元格

                s[row[x]] = self.conversion_cell(sheet,curRowNo,x,col[x])
            if s['results'] == "failure":
                cls.append(s)  
            curRowNo += 1
        return cls

    def hasNext(self,rownum,curRowNo):  
        if rownum == 0 or rownum <= curRowNo :  
            return False  
        else:  
            return True  

    #将读取excl整形的float，转换成int
    def conversion_cell(self,sheet,curRowNo,curColNo,cell):
        #判断python读取的返回类型  0 --empty,1 --string, 2 --number(都是浮点), 3 --date, 4 --boolean, 5 --error  
        if sheet.cell(curRowNo,curColNo).ctype == 2:
            no = str(int(cell))
             
        elif sheet.cell(curRowNo,curColNo).ctype == 3:
            # 转成datetime对象
            date = datetime(*xldate_as_tuple(cell, 0))
            no = date.strftime('%Y-%m-%d')
        else:
            no = cell
        return no


    #拆分合并单元格，合并单元格的第一个单元格的的映射到其他被合并的单元格中
    def merge_cell(self,sheet):
        mc ={}
        if sheet.merged_cells:
            for item in sheet.merged_cells:
                for row in range(item[0],item[1]):
                    for col in range(item[2],item[3]):
                        mc.update({(row,col):(item[0],item[2])})
        return mc

if __name__ =='__main__': 
    aa = ReadExcl("Phone")
    data = aa.get_xls_next()
    i = 1
    while i<20000:
        aa.write_excl({"id":str(i)})
        print(str(i))
        i = i+1




    