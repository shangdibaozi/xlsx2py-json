#!/usr/bin/env python
# -*- coding:utf-8 -*-

#######################################################
#  用于批量删除excel的指定行                          #
#  适用于所有office，前提需要安装pywin32和office软件  #
#######################################################

import os
import sys
import time
import glob
import shutil
import string
import os.path
import traceback
import ConfigParser
import win32com.client

SPATH = ""  # 需处理的excel文件目录
DPATH = ""  # 处理后的excel存放目录

SKIP_FILE_LIST = []  # 需要跳过的文件列表
MAX_SHEET_INDEX = 1  # 每个excel文件的前几个表需要处理
DELETE_ROW_LIST = []  # 需要删除的行号


def dealPath(pathname=''):
    '''deal with windows file path'''
    if pathname:
        pathname = pathname.strip()
    if pathname:
        pathname = r'%s' % pathname
        pathname = string.replace(pathname, r'/', '\\')
        pathname = os.path.abspath(pathname)
        if pathname.find(":\\") == -1:
            pathname = os.path.join(os.getcwd(), pathname)
    return pathname


class EasyExcel(object):
    '''class of easy to deal with excel'''

    def __init__(self):
        '''initial excel application'''
        self.m_filename = ''
        self.m_exists = False
        # 也可以用Dispatch，前者开启新进程，后者会复用进程中的excel进程
        self.m_excel = win32com.client.DispatchEx('Excel.Application')
        self.m_excel.DisplayAlerts = False  # 覆盖同名文件时不弹出确认框

    def open(self, filename=''):
        '''open excel file'''
        if getattr(self, 'm_book', False):
            self.m_book.Close()
        self.m_filename = dealPath(filename) or ''
        self.m_exists = os.path.isfile(self.m_filename)
        if not self.m_filename or not self.m_exists:
            self.m_book = self.m_excel.Workbooks.Add()
        else:
            self.m_book = self.m_excel.Workbooks.Open(self.m_filename)

    def reset(self):
        '''reset'''
        self.m_excel = None
        self.m_book = None
        self.m_filename = ''

    def save(self, newfile=''):
        '''save the excel content'''
        assert type(newfile) is str, 'filename must be type string'
        newfile = dealPath(newfile) or self.m_filename
        if not newfile or (self.m_exists and newfile == self.m_filename):
            self.m_book.Save()
            return
        pathname = os.path.dirname(newfile)
        if not os.path.isdir(pathname):
            os.makedirs(pathname)
        self.m_filename = newfile
        self.m_book.SaveAs(newfile)

    def close(self):
        '''close the application'''
        self.m_book.Close(SaveChanges=1)
        self.m_excel.Quit()
        time.sleep(2)
        self.reset()

    def addSheet(self, sheetname=None):
        '''add new sheet, the name of sheet can be modify,but the workbook can't '''
        sht = self.m_book.Worksheets.Add()
        sht.Name = sheetname if sheetname else sht.Name
        return sht

    def getSheet(self, sheet=1):
        '''get the sheet object by the sheet index'''
        assert sheet > 0, 'the sheet index must bigger then 0'
        return self.m_book.Worksheets(sheet)

    def getSheetByName(self, name):
        '''get the sheet object by the sheet name'''
        for i in xrange(1, self.getSheetCount()+1):
            sheet = self.getSheet(i)
            if name == sheet.Name:
                return sheet
        return None

    def getCell(self, sheet=1, row=1, col=1):
        '''get the cell object'''
        assert row > 0 and col > 0, 'the row and column index must bigger then 0'
        return self.getSheet(sheet).Cells(row, col)

    def getRow(self, sheet=1, row=1):
        '''get the row object'''
        assert row > 0, 'the row index must bigger then 0'
        return self.getSheet(sheet).Rows(row)

    def getCol(self, sheet, col):
        '''get the column object'''
        assert col > 0, 'the column index must bigger then 0'
        return self.getSheet(sheet).Columns(col)

    def getRange(self, sheet, row1, col1, row2, col2):
        '''get the range object'''
        sht = self.getSheet(sheet)
        return sht.Range(self.getCell(sheet, row1, col1), self.getCell(sheet, row2, col2))

    def getCellValue(self, sheet, row, col):
        '''Get value of one cell'''
        return self.getCell(sheet, row, col).Value

    def setCellValue(self, sheet, row, col, value):
        '''set value of one cell'''
        self.getCell(sheet, row, col).Value = value

    def getRowValue(self, sheet, row):
        '''get the row values'''
        return self.getRow(sheet, row).Value

    def setRowValue(self, sheet, row, values):
        '''set the row values'''
        self.getRow(sheet, row).Value = values

    def getColValue(self, sheet, col):
        '''get the row values'''
        return self.getCol(sheet, col).Value

    def setColValue(self, sheet, col, values):
        '''set the row values'''
        self.getCol(sheet, col).Value = values

    def getRangeValue(self, sheet, row1, col1, row2, col2):
        '''return a tuples of tuple)'''
        return self.getRange(sheet, row1, col1, row2, col2).Value

    def setRangeValue(self, sheet, row1, col1, data):
        '''set the range values'''
        row2 = row1 + len(data) - 1
        col2 = col1 + len(data[0]) - 1
        range = self.getRange(sheet, row1, col1, row2, col2)
        range.Clear()
        range.Value = data

    def getSheetCount(self):
        '''get the number of sheet'''
        return self.m_book.Worksheets.Count

    def getMaxRow(self, sheet):
        '''get the max row number, not the count of used row number'''
        return self.getSheet(sheet).Rows.Count

    def getMaxCol(self, sheet):
        '''get the max col number, not the count of used col number'''
        return self.getSheet(sheet).Columns.Count

    def clearCell(self, sheet, row, col):
        '''clear the content of the cell'''
        self.getCell(sheet, row, col).Clear()

    def deleteCell(self, sheet, row, col):
        '''delete the cell'''
        self.getCell(sheet, row, col).Delete()

    def clearRow(self, sheet, row):
        '''clear the content of the row'''
        self.getRow(sheet, row).Clear()

    def deleteRow(self, sheet, row):
        '''delete the row'''
        self.getRow(sheet, row).Delete()

    def clearCol(self, sheet, col):
        '''clear the col'''
        self.getCol(sheet, col).Clear()

    def deleteCol(self, sheet, col):
        '''delete the col'''
        self.getCol(sheet, col).Delete()

    def clearSheet(self, sheet):
        '''clear the hole sheet'''
        self.getSheet(sheet).Clear()

    def deleteSheet(self, sheet):
        '''delete the hole sheet'''
        self.getSheet(sheet).Delete()

    def deleteRows(self, sheet, fromRow, count=1):
        '''delete count rows of the sheet'''
        maxRow = self.getMaxRow(sheet)
        maxCol = self.getMaxCol(sheet)
        endRow = fromRow+count-1
        if fromRow > maxRow or endRow < 1:
            return
        self.getRange(sheet, fromRow, 1, endRow, maxCol).Delete()

    def deleteCols(self, sheet, fromCol, count=1):
        '''delete count cols of the sheet'''
        maxRow = self.getMaxRow(sheet)
        maxCol = self.getMaxCol(sheet)
        endCol = fromCol + count - 1
        if fromCol > maxCol or endCol < 1:
            return
        self.getRange(sheet, 1, fromCol, maxRow, endCol).Delete()


def echo(msg):
    '''echo message'''
    print msg


def dealSingle(excel, sfile, dfile):
    '''deal with single excel file'''
    echo("deal with %s" % sfile)
    basefile = os.path.basename(sfile)
    excel.open(sfile)
    sheetcount = excel.getSheetCount()
    if not (basefile in SKIP_FILE_LIST or file in SKIP_FILE_LIST):
        for sheet in range(1, sheetcount+1):
            if sheet > MAX_SHEET_INDEX:
                continue
            reduce = 0
            for row in DELETE_ROW_LIST:
                excel.deleteRow(sheet, row-reduce)
                reduce += 1
            #excel.deleteRows(sheet, 2, 2)
    excel.save(dfile)


def dealExcel(spath, dpath):
    '''deal with excel files'''
    start = time.time()
    # check source path exists or not
    spath = dealPath(spath)
    if not os.path.isdir(spath):
        echo("No this directory :%s" % spath)
        return
    # check destination path exists or not
    dpath = dealPath(dpath)
    if not os.path.isdir(dpath):
        os.makedirs(dpath)
    shutil.rmtree(dpath)
    # list the excel file
    filelist = glob.glob(os.path.join(spath, '*.xlsx'))
    if not filelist:
        echo('The path of %s has no excel file' % spath)
        return
    # deal with excel file
    excel = EasyExcel()
    for file in filelist:
        basefile = os.path.basename(file)
        destfile = os.path.join(dpath, basefile)
        dealSingle(excel, file, destfile)
    echo('Use time:%s' % (time.time()-start))
    excel.close()


def loadConfig(configfile='./config.ini'):
    '''parse config file'''
    global SPATH
    global DPATH
    global SKIP_FILE_LIST
    global MAX_SHEET_INDEX
    global DELETE_ROW_LIST

    file = dealPath(configfile)
    if not os.path.isfile(file):
        echo('Can not find the config.ini')
        return False
    parser = ConfigParser.ConfigParser()
    parser.read(file)
    SPATH = parser.get('pathconfig', 'spath').strip()
    DPATH = parser.get('pathconfig', 'dpath').strip()
    filelist = parser.get('otherconfig', 'filelist').strip()
    index = parser.get('otherconfig', 'maxindex').strip()
    rowlist = parser.get('otherconfig', 'deleterows').strip()
    if filelist:
        SKIP_FILE_LIST = filelist.split(";")
    if rowlist:
        DELETE_ROW_LIST = map(int, rowlist.split(";"))
    MAX_SHEET_INDEX = int(index) if index else MAX_SHEET_INDEX


def main():
    '''main function'''
    loadConfig()
    if SPATH and DPATH and MAX_SHEET_INDEX:
        dealExcel(SPATH, DPATH)
    raw_input("Please press any key to exit!")


if __name__ == "__main__":
    main()

"""
config.ini文件如下：
[pathconfig]
#;spath表示需要处理的excel文件目录
spath=./tests
#;dpath表示处理后的excel文件目录
dpath=./dest

[otherconfig]
#;filelist表示不需要做特殊处理的excel文件列表,以英文分号分隔
filelist=
#;maxindex表示需要处理每个excel文件的前几张表
maxindex=1
#;deleterows表示需要删除的阿拉伯数字行号，用英文分号分隔
deleterows=2;3
"""