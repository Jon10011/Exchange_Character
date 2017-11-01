# -*- coding: utf-8 -*-
import sys
reload(sys)
sys.setdefaultencoding('utf-8')
import xlrd
import xlwt
import HTMLParser


class OperExcel():
    def rExcel(self,inEfile,outfile):
        rfile = xlrd.open_workbook(inEfile)
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
            if "&" in table.cell(i,3).value:
                a = html_parser.unescape(table.cell(i,3).value).encode("utf-8 ")+"\n"
            else:
                a = table.cell(i,3).value.encode("utf-8 ")+"\n"
            print "--------------------"
            print i
            print table.cell(i,3)
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

if __name__ == '__main__':
   t = OperExcel()
   t.rExcel('C:\Users\Jon\Desktop\Atest1020\Autodesk_Data_Wizard_ExcelConnector_1.0.xlsm','C:\Users\Jon\Desktop\Atest1020\A20171020.txt')
   t.wExcel('C:\Users\Jon\Desktop\Atest1020\A20171020.txt','C:\Users\Jon\Desktop\Atest1020\Character_Exchange.csv')