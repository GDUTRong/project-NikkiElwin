import xlrd as xr   #导入模块
import xlwt as xw
import re,os,shutil,collections
from xlutils.copy import copy
#签到表格转换
#制作者：梁梓熙

class Conversion():
    """签到表格转换"""
    def __init__(self,sourceName: str,targetName: str):
        """初始化程序"""

        self.source = xr.open_workbook(sourceName,logfile=open(os.devnull, 'w'))
        self.targetSource = xr.open_workbook(targetName,formatting_info=True,logfile=open(os.devnull, 'w'))  #打开
        self.target = copy(self.targetSource)  #复制一个以供使用

        self.sameName = self._findSameName() #记录重名
        self.wrong = 0 #记录错误学生
        self.marked = {}  #记录学生时间

    def _findSameName(self):
        """寻找有没有重名的人"""
        temp,ans = [],[]
        for i in range(4,self.targetSource.sheet_by_index(0).nrows):
            val1 = self.targetSource.sheet_by_index(0).cell_value(i,2)
            val2 = self.targetSource.sheet_by_index(0).cell_value(i,7)
            if not val1 =='':
                temp.append(val1)
            if not val2 =='':
                temp.append(val2)

        if temp != []: #确实有重名的
            for key,value in collections.Counter(temp).items():
                if value > 1:
                    ans.append(key)
        return ans

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

        if self._finePos(string) == None:  #找不到学生的位置则写入错误信息
            self.targetSheet().write(self.wrong + 4 ,\
                10,self.sourceSheet().cell_value(row,col),self.center())
            self.targetSheet().write(self.wrong + 4 ,\
                11,c.readTime(row,col),self.center())
            self.wrong += 1 #更新指针
            return None

        return string


    def writeStudent(self,name: str,mark: str):
        """记录一个学生签到"""
        pos = self._finePos(name)  #找到学生的位置列表
        self.targetSheet().write(pos[0],pos[1] + 1,mark,self.center())


    def writeStudentTime(self,name: str,time: str):
        """记录一个学生在线时间"""
        pos = self._finePos(name)  #找到学生的位置列表

        if self.isMark(name):   #出现重名情况
            newTime = self.marked[str(pos[0]) + ',' + str(pos[1])] + ',' + time
            self.targetSheet().write(pos[0],pos[1] + 2,newTime,self.center())
            self.mark(name,newTime)
            return

        self.targetSheet().write(pos[0],pos[1] + 2,time,self.center())
        self.mark(name,time)

    def _sameName(self,name) ->bool:
        """寻找有没有重名的人"""
        if self.sameName == []: return False
        else:
            for val in self.sameName:
                if val in name:
                    return True
            return False

    def _finePosSame(self,name):
        """通过序号信息找到同名人的位置"""
        num = name[:2] #把学生的序号提取出来
        if not num.isdigit():
            return None  #找不到序号，当没有这个人

        for i in range(4,self.targetSource.sheet_by_index(0).nrows):
            if str(int(num)) == self.targetSource.sheet_by_index(0).cell_value(i,0) and self.targetSource.sheet_by_index(0).cell_value(i,2) in name:
                return [i,2]

            if str(int(num)) == self.targetSource.sheet_by_index(0).cell_value(i,5) and self.targetSource.sheet_by_index(0).cell_value(i,7) in name:
                return [i,7]
        return None

    def _finePos(self,name: str):
        """寻找一个学生的位置，返回列表，没有则返回None"""
        for i in range(4,self.targetSource.sheet_by_index(0).nrows):

            if self.targetSource.sheet_by_index(0).cell_value(i,2) in name and self.targetSource.sheet_by_index(0).cell_value(i,2) != "":
                if not self._sameName(name):
                    return [i,2]
                else:
                    return self._finePosSame(name)

            if self.targetSource.sheet_by_index(0).cell_value(i,7) in name and self.targetSource.sheet_by_index(0).cell_value(i,7) != "":
                if not self._sameName(name):
                    return [i,7]
                else:
                    return self._finePosSame(name)

        return None  


    def readTime(self,row: int,col: int) -> str:
        """返回一个学生时间"""
        return self.sourceSheet().cell_value(row,col+4)  


    def isMark(self,name: str) -> bool:
        """返回一个学生是否录入"""
        temp = self._finePos(name)   #标记学生已签到
        return str(temp[0]) + ',' + str(temp[1]) in self.marked.keys()


    def mark(self,name: str,time: str):
        """标记一个学生"""
        temp = self._finePos(name)   #找到学生位置
        self.marked[str(temp[0]) + ',' + str(temp[1])] = time

def getFromExcel(sources:str):
    """从课程中得到对应的目标模板"""
    temp = xr.open_workbook(sources)
    string = temp.sheet_by_index(0).cell_value(5,2)
    for root, dirs, files in os.walk('templates'):   #在模板文件中查找
        if len(files) == 0:
            return None

        ans = []
        for i in range(len(files)):
            #判断后缀为xls的
            if os.path.splitext(files[i])[1] == '.xls':
                mes = os.path.splitext(files[i])[0].split(' ')  #把模板名分割
                if mes[0] in string and mes[-1] in string:
                    ans.append(files[i]) #匹配的情况都放进去
                    
        if not ans:  #没有匹配情况
            return None
        elif len(ans) == 1:  #唯一匹配，提高速度
            return ans[0]
        else:
            return askTeacher(sources,ans)  #询问老师

def getTargetName(sources: str):
    """获得对应的目标模板"""

    for root, dirs, files in os.walk('templates'):   #在模板文件中查找
        if len(files) == 0:
            return None

        ans = []
        for i in range(len(files)):
            #判断后缀为xls的
            if os.path.splitext(files[i])[1] == '.xls':
                mes = os.path.splitext(files[i])[0].split(' ')  #把模板名分割
                if mes[0] in sources and mes[-1] in sources:
                    ans.append(files[i]) #匹配的情况都放进去
                    
        if not ans:  #没有匹配情况
            return None
        elif len(ans) == 1:  #唯一匹配，提高速度
            return ans[0]
        else:
            return askTeacher(sources,ans)  #询问老师

def askTeacher(sources:str,ans:list):
    """询问老师"""
    print(sources + "检测到多个匹配模板情况！")
    for i in range(len(ans)):
        print(str(i+1) + '.' + ans[i])

    while True:
        num=None
        try:
            num=int(input("请输入选定模板的下标："))
        except:
            pass
        if num in range(len(ans)+1):
            return ans[num-1]
        else:
            print("输入错误!请检查您的输入情况！")

def inputStudent(name:str,conversion: Conversion,mark: str,i):
    """完成输入"""
    c.writeStudent(name,mark)
    time = conversion.readTime(i,3)
    c.writeStudentTime(name,time)
  

mark = '√'       #签到成功的标记

for root, dirs, files in os.walk('sources'):

    if len(files) == 0: 
        break

    for file in files:
        #判断后缀为xlsx的
        if os.path.splitext(file)[1] == '.xlsx' :

            targetName = getTargetName('sources/' + file)  #得到对应的目标模板
            if targetName == None:
                targetName = getFromExcel('sources/' + file)  #从Excel中得到对应的模板
                if targetName == None:
                    print("找不到模板文件！\n已结束录入文件" + file)
                    continue

            shutil.copyfile('sources/' + file,'achieve/' + file)  #复制文件
            c = Conversion('sources/' + file,'templates/' + targetName)

            for i in range(5,c.sourceSheet().nrows):
                name = c.readStudentName(i,3)
                if name != None:
                    inputStudent(name,c,mark,i)

            if c.wrong > 0:
                c.targetSheet().col(10).width = 256 * 30  #用于储存出错同学

            os.remove('sources/' + file) #删掉文件
            print(file + "转换完成!\n")
            finalTargetName = targetName.replace('签到空表','')[:-4]
            c.target.save('targets/' + os.path.splitext(file)[0][:19] + finalTargetName + '.xls')   #保存文件

