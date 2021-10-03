import xlsxwriter
import xlrd
from xlrd import xldate_as_tuple

TestNum = 223 # 要处理的数据组数
Save_FilePath = 'E:/test_07.xls' # 要保存的地址及名称 每次都要换名字
Save_SheetName = 'sheet1'# 要保存的工作表
Data_path = "E:/cycle.xlsx" # 要打开的文件
Data_sheetname = "cycle224"# 要打开的工作表

'''  
xlrd中单元格的数据类型
数字一律按浮点型输出，日期输出成一串小数，布尔型输出0或1
成我们想要的数据类型
0 empty,1 string, 2 number, 3 date, 4 boolean, 5 error
'''
class ExcelData():
    # 初始化方法
    def __init__(self, data_path, sheetname):
        #定义一个属性接收文件路径
        self.data_path = data_path
        # 定义一个属性接收工作表名称
        self.sheetname = sheetname
        # 使用xlrd模块打开excel表读取数据
        self.data = xlrd.open_workbook(self.data_path)
        # 根据工作表的名称获取工作表中的内容（方式①）
        self.table = self.data.sheet_by_name(self.sheetname)
        # 根据工作表的索引获取工作表的内容（方式②）
        # self.table = self.data.sheet_by_name(0)
        # 获取第一行所有内容,如果括号中1就是第二行，这点跟列表索引类似
        self.keys = self.table.row_values(0)
        # 获取工作表的有效行数
        self.rowNum = self.table.nrows
        # 获取工作表的有效列数
        self.colNum = self.table.ncols

    # 定义一个读取excel表的方法
    def readExcel(self):
        # 定义一个空列表
        k1 = 0
        k2 = TestNum
        workbook = xlsxwriter.Workbook(Save_FilePath)
        worksheet = workbook.add_worksheet(Save_SheetName)

        for i in range(1, self.rowNum):
            c_type5 = self.table.cell(i, 0).ctype
            if c_type5 !=2:
                if k1>10:
                    break
                continue
            else:
                k1 = k1+1  #有效行数
        i = 0
        for m in range(0,2 * k2,2):
            for k in range(k1):
                    while 1:
                        c_cell1 = self.table.cell_value(i, 0)
                        c_cell3 = self.table.cell_value(i, 1)
                        c_type1 = self.table.cell(i, 0).ctype
                        if c_type1 != 2:
                            i=i+1
                            continue
                        else:
                            worksheet.write(k, m, c_cell1)
                            worksheet.write(k, m+1, c_cell3)
                            i = i + 1

                            break

        workbook.close()

if __name__ == "__main__":
    get_data = ExcelData(Data_path, Data_sheetname)
    get_data.readExcel()
