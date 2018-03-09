import xlrd
from xlwt import *
from xlutils.copy import copy
from Account import Account
from ExcelUtils import dateFormat

font = Font()
font.name = '宋体'

titleFont = Font()
titleFont.name = '宋体'
titleFont.bold = True

alignment = Alignment()
alignment.horz = Alignment.HORZ_CENTER
alignment.vert = Alignment.VERT_CENTER

borders = Borders()
borders.bottom = Borders.THIN
borders.left = Borders.THIN
borders.right = Borders.THIN

amountStyle = XFStyle()
amountStyle.font = font
amountStyle.num_format_str = '#,##0.00'
amountStyle.alignment = alignment
amountStyle.borders = borders

stringStyle = XFStyle()
stringStyle.font = font
stringStyle.alignment = alignment
stringStyle.borders = borders


titleStyle = XFStyle()
titleStyle.font = titleFont


def processData(itemPath, templatePath):
    try:
        #明细表
        book = xlrd.open_workbook(itemPath)
        #模板表
        wb = xlrd.open_workbook(templatePath, formatting_info=True, on_demand=True)
        newBook = copy(wb)
        newSheet = newBook.get_sheet(0)
        newSheet.header_str = ''.encode('utf-8')
        newSheet.footer_str = ''.encode('utf-8')
        
        #定义表单数量
        sheetCount = book.nsheets
        print('表单数量:%d' % sheetCount)
        
        #定义表单名称
        for sheetNum in range(0, sheetCount):
            sheetName = book.sheet_names()[sheetNum]
            print('表单名称:%s' % sheetName)
        #定义账户个数
        accountCount = 0
        #定义收入项目个数
        projectCount = 0
        #定义账户类型所在的行数
        accountTypeRow = 3
        #遍历excel表格，查找账户个数
        sheet = book.sheet_by_name('函证明细表')
        
        #sheet录入数据总行数行数
        rowTotalCount = sheet.nrows
        dataRowTotalCount = rowTotalCount - 5
        print('sheet录入数据总行数行数：%d' % dataRowTotalCount)
        if(dataRowTotalCount < 1):
            raise Exception('函证明细表中无录入数据，请查看数据')
        for column in range(len(sheet.row_values(accountTypeRow))):
            string = sheet.row_values(accountTypeRow)[column]
            if '账户' in string:
                accountCount = accountCount + 1
            if '收入项目' in string:
                projectCount = projectCount + 1
        #打印账户个数
        print('账户个数为：%d' % accountCount)
        print('收入项目个数为：%d' % projectCount)
        
        #计算录入数据总列数 编号+单位名称+账户*5+收入项目*5+是否回函+其他事项
        columnNum = 2 + accountCount * 5 + projectCount * 5 + 2
        print('录入数据总列数为:%d' % columnNum)
        
        print('process data start.')
        
        #校验统一函证截止日期
        generalEndDate = sheet.row_values(1)[1]
        value1 = generalEndDate
        print('统一函证截止日期：%s' % generalEndDate)
        
        for rowNum in range(5, rowTotalCount):
            #一行数据测试
            #定义一个空列表，填装账户信息
            list = []
            print('process rowNum:%d start.' % rowNum)
            newSheetNo = sheet.row_values(rowNum)[0]
            newSheetName = sheet.row_values(rowNum)[1]
            print('函证模板sheet编号：%s,单位名称：%s' % (newSheetNo, newSheetName))
            n = 0
            for column in range(2, accountCount * 5 + 2):
                num = column - 1
                if (num - n * 5 == 1) and generalEndDate == '':
                    value1 = sheet.row_values(rowNum)[column]
                elif num - n * 5 == 2:
                    value2 = sheet.row_values(rowNum)[column]
                elif num - n * 5 == 3:
                    value3 = sheet.row_values(rowNum)[column]
                elif num - n * 5 == 4:
                    value4 = sheet.row_values(rowNum)[column]
                else:
                    value5 = sheet.row_values(rowNum)[column]
                if num - n * 5 == 5:
                    n = n + 1
                    account = Account(value1, value2, value3, value4, value5)
                    list.append(account)
            i = 9
            contextTitle = '致：' + newSheetName
            newSheet.write(2, 0, contextTitle, titleStyle)
            for account in list:
                date = dateFormat(account.date)
                newSheet.write(i, 0, date, stringStyle)
                newSheet.write(i, 1, account.bussinessContext, stringStyle)
                newSheet.write(i, 2, account.debit, amountStyle)
                newSheet.write(i, 3, account.credit, amountStyle)
                newSheet.write(i, 4, account.item, stringStyle)
                i = i + 1
            newBook.save(r"../result/" + newSheetNo + "-" + newSheetName + ".xls")
            print('process rowNum:%d end.' % rowNum)
        print('process data end.')
    except Exception as e:
        raise e


if __name__ == "__main__":
    try:
        itemPath = "../template/函证明细表.xls"
        templatePath = "../template/询证函模版.xls"
        processData(itemPath, templatePath)
    except Exception as e:
        raise e

    input("\n\nPress the enter key to exit.")
