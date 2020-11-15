# -*- coding: utf-8 -*-  

import re  
import urllib  
import urllib2  
import cookielib
import  xdrlib ,sys
import xlrd
from xlutils.copy import copy
import xlwt
import os

__version = "1.0"
__date__ = "2013/9/14"
__doc__ = "Grap stock data from cnki.net,Powered by iniwaper@gmail.com"

'''
本程序主要实现自动从磁盘读取股票数据，然后从知网采集所需数据后，存入磁盘

本程序实现采用Python编写，是一种开发效率高，非常适合用于科研领域

软件思想基于面向对象编程，将任务划分成两个类，分别用于处理数据抓取以及文件处理。
其中，GrapData类，主要实现参数配置，网页请求，网页解析，数据获取等；
    ProcessExcelDate类主要实现excel文件的创建，读写，保存等

另外，main函数主要负责整个任务的执行

软件执行步骤：1，程序开始
            2，读取excel文件数据
            3，设置参数后，依次执行数据采集
            4，将结果写入相应excel文件
            5，程序结束
            
本软件实现过程中的难点：1，Url分析，请求参数分析，Html页面分析等
                    2，需要执行两次Http请求才能得到正确的数据，此点是关键
                    3，断点续传的支持，由于网络原因以及源网站的屏蔽，导致的中断，已经完成的工作需要重来
                       支持断点续传是本软件的亮点
                            
'''


print __doc__
print "processing....pleast wait...."

########################## html operation ##################################

class GrapData:
    HandlerUrl = 'http://epub.cnki.net/KNS/request/SearchHandler.ashx?'
    DataUrl = "http://epub.cnki.net/kns/brief/brief.aspx?"
    
    def __init__(self):
        self.handlerParam = {
                                'action':'',
                                'NaviCode':'*',
                                'ua':'1.21',
                                'PageName':'ASP.brief_result_aspx',
                                'DbPrefix':'CCND',
                                'DbCatalog':'中国重要报纸全文数据库',
                                'ConfigFile':'CCND.xml',
                                'db_opt':'中国重要报纸全文数据库',
                                'db_value':'中国重要报纸全文数据库',
                                'magazine_value1':'中国证券报+上海证券报+证券时报+证券日报',#set by input
                                'magazine_special1':'%',
                                'publishdate_from':'2010-01-01',#set by input
                                'publishdate_to':'2011-12-31',#set by input
                                'au_1_sel':'AU',
                                'au_1_special1':'=',
                                'txt_1_sel':'FT',
                                'txt_1_value1':'000012',#read from excel
                                'txt_1_value2':'南  玻Ａ',#read from excel
                                'txt_1_relation':'#CNKI_OR',
                                'txt_1_special1':'=',
                                'his':'0'
                            }
        self.dataParam = {
                                   'pagename':'ASP.brief_result_aspx',
                                    'dbPrefix':'CCND',
                                    'dbCatalog':'中国重要报纸全文数据库',
                                    'ConfigFile':'CCND.xml',
                                    'research':'off',
                                    'keyValue':'000012',
                                    'S':'1',
                                }
        
        self.cj = cookielib.CookieJar()
        self.op = urllib2.build_opener(urllib2.HTTPCookieProcessor(self.cj))

        urllib2.install_opener(self.op)
    def setMagazine(self,value):
        '''
        parm value exmple:'中国证券报+上海证券报+证券时报+证券日报'
        '''
        self.handlerParam['magazine_value1'] = value
    def setPublishdateFrom(self,value):
        '''
        parm value exmple:'2010-01-01'
        '''
        self.handlerParam['publishdate_from'] = value
    def setPublishdateTo(self,value):
        '''
        parm value exmple:'2010-01-01'
        '''
        self.handlerParam['publishdate_to'] = value
    def setStockCode(self,value):
        '''
        parm value exmple:'000014
        '''
        self.handlerParam['txt_1_value1'] = value

        #self.dataParam
        self.dataParam['keyValue'] = value
        
    def setStockName(self,value):
        '''
        parm value exmple:'沙河股份'
        '''
        self.handlerParam['txt_1_value2'] = value
    def doReq(self,url,values):
        data = urllib.urlencode(values)
        req = urllib2.Request(url,data)
        response = self.op.open(req) 
        html = response.read()
        return html
    def doDoubleReq(self):
        self.doReq(self.HandlerUrl,self.handlerParam)
        html = self.doReq(self.DataUrl,self.dataParam)
        return html
    def parseHtml(self,regular):
        html = self.doDoubleReq()
        html = "".join(html.split())
        html = "".join(html.split('&nbsp;'))
        html = "".join(html.split(','))
        v = re.findall(regular,html, re.S)
        return v
    def getTotalArticleNum(self):
        regular = r"找到(\d*)条结果"
        ret = self.parseHtml(regular)
        if len(ret) == 0:
            num = 0
        else:
            num = ret[0]
        return num

    
#################################### excel operation ###########################################
class ProcessExcelDate:
    
    def __init__(self,path,name):
        self.data = self.openExcel(path)
        self.table = self.data.sheet_by_name(name)
        self.write = copy(self.data)
        self.writeTable = self.write.get_sheet(0)
    def openExcel(self,path= ''):
        try:
            return xlrd.open_workbook(path)
        except Exception,e:
            return None
            
    #
    def readExcelTableByname(self,colnameindex=0):

        nrows = self.table.nrows #行数 
        colnames =  self.table.row_values(colnameindex) #某一行数据 
        data =[]
        for rownum in range(1,nrows):
             row = self.table.row_values(rownum)
             if row:
                 app = {}
                 for i in range(len(colnames)):
                    app[colnames[i]] = row[i]
                 data.append(app)
        return data
    def writeExcelTableByname(self,cell_row,cell_col,num):
       self.writeTable.write(cell_row, cell_col, num)
    def saveExcel(self,path):
        self.write.save(path)
def getRPath(filename):
    return os.path.join(os.path.dirname(os.path.abspath(sys.argv[0])), filename)
def readExcelTable(path):
        data = xlrd.open_workbook(path)
        table = data.sheet_by_name(u'Sheet1')
        nrows = table.nrows 
        colnames =  table.row_values(0)
        data =[]
        for rownum in range(1,nrows):
             row = table.row_values(rownum)
             if row:
                 app = {}
                 for i in range(len(colnames)):
                    app[colnames[i]] = row[i]
                 data.append(app)
        return data
################################ main ###########################################
def main():
    processExcelDate = ProcessExcelDate(getRPath("stockData.xlsx"),u'Sheet1')
    tables = processExcelDate.readExcelTableByname()

    #for break tx
    if(os.path.exists(getRPath("stockData.xls"))):
        breakPointTables = readExcelTable(getRPath("stockData.xls"))

    else:
        breakPointTables = None

    ###
    for i in range(0,len(tables)):
        row = tables[i]
        #saved not process
        if breakPointTables:
            print str(breakPointTables[i]['Smedia'])
            if(str(breakPointTables[i]['Smedia'])!= ''):
                processExcelDate.writeExcelTableByname(i+1,6,breakPointTables[i]['Smedia'])
                continue
        
        Stock = row['stock']
        Company = "".join(row['company'].split())
        
        PublishdateFrom = row['year'] + "-01-01"
        PublishdateTo = row['year'] + "-12-31"
        SearchMagazine = row['magazine']
        
        #process all data
        grap = GrapData()
        grap.setMagazine(SearchMagazine.encode('utf8'))
        grap.setPublishdateFrom(PublishdateFrom)
        grap.setPublishdateTo(PublishdateTo)
        grap.setStockCode(Stock)
        grap.setStockName(Company.encode('utf8'))

        processExcelDate.writeExcelTableByname(i+1,6,grap.getTotalArticleNum())
        print 'current:',i,row['Smedia']
        processExcelDate.saveExcel(getRPath("stockData.xls"))

##########################################
if __name__=="__main__":
    main()





