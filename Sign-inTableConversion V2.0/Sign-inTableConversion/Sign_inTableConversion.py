import xlrd as xr   #导入模块
import xlwt as xw
import re,os
from xlutils.copy import copy
#签到表格转换
#制作者：梁梓熙

class Conversion():
    """签到表格转换"""
    def __init__(self,sourceName: str,targetName: str):
        """初始化程序"""
        self.wrong = 0 #记录错误学生
        self.marked = []  #记录学生

        self.source = xr.open_workbook(sourceName)
        self.targetSource = xr.open_workbook(targetName,formatting_info=True)  #打开
        self.target = copy(self.targetSource)  #复制一个以供使用

    def targetSheet(self):
        """获取目标表格"""
        return self.target.get_sheet(0)


    def sourceSheet(self):
        """获取资源表格"""
        return self.source.sheet_by_index(0)


    def center(self):
        """使一个文本加框并居中"""
        style = xw.XFStyle() 

        borders = xw.Borders()   #修改边框
        borders.left = xw.Borders.THIN
        borders.right = xw.Borders.THIN
        borders.top = xw.Borders.THIN
        borders.bottom = xw.Borders.THIN
        style.borders = borders
 
        alignment = xw.Alignment()   #居中
        alignment.horz = xw.Alignment.HORZ_CENTER 
        style.alignment = alignment 
        
        return style


    def readStudentName(self,row: int,col: int):
        """读取第row行col列的学生名字"""
        string = self.sourceSheet().cell_value(row,col)  
        if string == "": return None

        target = re.sub("[A-Za-z0-9\!\%\[\]\,\。\ ]","", string)  #将所有非中文字符都删掉
        if self._finePos(target) == None:  #找不到学生的位置则写入错误信息
            self.targetSheet().write(self.wrong + 4 ,\
                10,self.sourceSheet().cell_value(row,col),self.center())
            self.targetSheet().write(self.wrong + 4 ,\
                11,c.readTime(row,col),self.center())
            self.wrong += 1 #更新指针
            return None

        return target


    def writeStudent(self,name: str,mark: str):
        """记录一个学生签到"""
        pos = self._finePos(name)  #找到学生的位置列表
        self.targetSheet().write(pos[0],pos[1] + 1,mark,self.center())


    def writeStudentTime(self,name: str,time: str):
        """记录一个学生在线时间"""
        pos = self._finePos(name)  #找到学生的位置列表

        if self.isMark(name):   #出现重名情况
            self.targetSheet().write(self.wrong + 4 ,10,name,self.center())
            self.targetSheet().write(self.wrong + 4 ,11,time,self.center())
            self.wrong += 1 #更新指针

        self.targetSheet().write(pos[0],pos[1] + 2,time,self.center())


    def _finePos(self,name: str):
        """寻找一个学生的位置，返回列表，没有则返回None"""
        for i in range(self.targetSource.sheet_by_index(0).nrows):

            if self.targetSource.sheet_by_index(0).cell_value(i,2) in name \
                and self.targetSource.sheet_by_index(0).cell_value(i,2) != "":
                return [i,2]

            if self.targetSource.sheet_by_index(0).cell_value(i,7) in name \
                and self.targetSource.sheet_by_index(0).cell_value(i,7) != "":
                return [i,7]
        return None  


    def readTime(self,row: int,col: int) -> str:
        """返回一个学生时间"""
        return self.sourceSheet().cell_value(row,col+4)  


    def isMark(self,name: str) -> bool:
        """返回一个学生是否录入"""
        return self._finePos(name) in self.marked


    def mark(self,name: str):
        """标记一个学生"""
        self.marked.append(self._finePos(name))   #标记学生已签到


def getTargetName(sources: str) -> str:
    """获得对应的目标模板"""
    for root, dirs, files in os.walk('templates'):   #在模板文件中查找
        if len(files) == 0:
            return None
        for file in files:
            #判断后缀为xls的
            if os.path.splitext(file)[1] == '.xls':
                if makeSure(sources,file):
                    return file

    for root, dirs, files in os.walk('templates'):   #找不到时，默认使用第一个
        return files[0]


def makeSure(sources: str,file: str)-> bool:
    """确定学生在此模板中"""
    temp = xr.open_workbook(sources) #先打开资源文件

    for i in range(5,10):
        string = temp.sheet_by_index(0).cell_value(i,3)  #读取里面几名学生的名字
        for j in range(len(nums)):   #得到班级的序号
            if nums[j] not in string:  #不在这个班
                continue
            elif nums[j] in file:
                return True
    return False


def inputStudent(name:str,conversion: Conversion,mark: str):
    """完成输入"""
    c.writeStudent(name,mark)
    c.writeStudentTime(name,conversion.readTime(i,3))
    c.mark(name)


nums = '一二三四五六七八九十卓创'  #关键词搜索列表
mark = '√'       #签到成功的标记

for root, dirs, files in os.walk('sources'):

    if len(files) == 0: 
        print("无资源文件!")
        break

    for file in files:
        #判断后缀为xlsx的
        if os.path.splitext(file)[1] == '.xlsx' :
            targetName = getTargetName('sources/' + file)  #得到对应的目标模板

            if targetName == None:
                print("无模板文件!")
                break

            c = Conversion('sources/' + file,'templates/' + targetName)

            for i in range(5,c.sourceSheet().nrows):
                name = c.readStudentName(i,3)
                if name != None:
                    inputStudent(name,c,mark)

            if c.wrong > 0:
                c.targetSheet().col(10).width = 256 * 30  #用于储存出错同学

            os.remove('sources/' + file) #删掉文件
            print(file + "转换完成!\n")
            c.target.save('targets/' + file)   #保存文件
