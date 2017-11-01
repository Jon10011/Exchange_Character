# -*- coding: utf-8 -*-
import sys
reload(sys)
sys.setdefaultencoding('utf-8')
import xlrd
import xlwt
import HTMLParser
#import openpyxl

class OperExcel():
    def rExcel(self,infile,outfile):
        rfile = xlrd.open_workbook(infile)
        table = rfile.sheet_by_index(0)
        #table = rfile.sheets()[0]
        #table = rfile.sheet_by_name(u'Sheet1')
        table.row_values(0)
        table.col_values(0)
        nrows = table.nrows
        #ncols = table.ncols
        wfile = open(outfile,'w')
        print nrows

        for i in range(nrows):
            html_parser = HTMLParser.HTMLParser()
            a = str(table.cell(i,3).value)
            try:
                if "&" in a:
                    a = html_parser.unescape(a).encode("utf-8 ")+"\n"
                else:
                    a = a.encode("utf-8 ")+""+"\n"
            except TypeError:
                print '++++++++++++++++++++++++'
            # print i
            # print table.cell(i,3)
            # print "--------------------"
            wfile.write(a)
            #wfile.close()

    def wExcel(self,infile,outfile):
        rfile = open(infile,'r')

        buf = rfile.read().split('\n')
        rfile.close()
        #设置excel编码以及单元格格式
        w = xlwt.Workbook(encoding = 'utf-8')
        sheet = w.add_sheet('Sheet1')
        for i in range(len(buf)):
            #print buf[i]
            sheet.write(i,0,buf[i].encode("utf-8"))
            w.save(outfile)

    # def cExcel(self, infile, outfile):
    #     wb1 = openpyxl.load_workbook(infile)
    #     wb2 = openpyxl.load_workbook(outfile)
    #     sheets1 = wb1.get_sheet_names()
    #     sheets2 = wb2.get_sheet_names()
    #     sheet1 = wb1.get_sheet_by_name(sheets1[0])
    #     sheet2 = wb2.get_sheet_by_name(sheets2[0])
    #     max_row = sheet1.max_row
    #     max_column = sheet1.max_column
    #     for m in range(1, max_row + 1):
    #         for n in range(97, 97 + max_column):  # chr(97)='a'
    #             n = chr(n)
    #             i = '%s%d' % (n, m)
    #             cell1 = sheet1[i].value
    #
    #             html_parser = HTMLParser.HTMLParser()
    #
    #             if n == 3 :
    #                 print '+++++++++++++++++++++++++++'
    #                 print cell1
    #                 if "&" in cell1:
    #                     cell1 = html_parser.unescape(sheet1[i].value)
    #                 else:
    #                     cell1 = sheet1[i].value
    #             else:
    #                 cell1 = sheet1[i].value
    #             sheet2[i].value = cell1
    #             wb2.save(outfile)
    #             wb1.close()
    #             wb2.close()


if __name__ == '__main__':
   t = OperExcel()
   t.rExcel('C:\Users\Jon\Desktop\Atest1020\Autodesk_Data_Wizard_ExcelConnector_1.0.xlsm','C:\Users\Jon\Desktop\Atest1020\A20171020.txt')
   t.wExcel('C:\Users\Jon\Desktop\Atest1020\A20171020.txt','C:\Users\Jon\Desktop\Atest1020\Character_Exchange.csv')
   # t.cExcel('C:\Users\Jon\Desktop\Atest1020\Autodesk_Data_Wizard_ExcelConnector_1.0.xlsm','C:\Users\Jon\Desktop\Atest1020\Autodesk_Data.xlsx')