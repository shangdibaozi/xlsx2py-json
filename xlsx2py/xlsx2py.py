import sys
import re
import os
import signal
import time
import codecs
import json
import copy
import tqdm


from ExcelTool import ExcelTool
import functions
import xlsxtool
import xlsxError

import config

SYS_CODE = sys.getdefaultencoding()


def siginit(sigNum, sigHandler):
    print("byebye")
    sys.exit(1)


signal.signal(signal.SIGINT, siginit)  # Ctrl-c处理


def hasFunc(funcName):
    return hasattr(functions, funcName)


def getFunc(funcName):
    return getattr(functions, funcName)


g_dctDatas = {}
g_fdatas = {}


class xlsx2py(object):
    """
    将excel数据导出为py文件 使用过程需要进行编码转换

    targets = 'py|json|lua'
    """

    def __init__(self, infile, outfile, targets):
        sys.excepthook = xlsxError.except_hook  # traceback处理,希望输出中文
        self.infile = os.path.abspath(infile)  # 暂存excel文件名
        self.outfile = os.path.abspath(outfile)  # data文件名
        self.targets = targets

    def __initXlsx(self):
        self.xbook = ExcelTool(self.infile)

        while not self.xbook.getWorkbook(forcedClose=True):
            xlsxtool.exportMenu(config.EXPORT_INFO_RTEXCEL, OCallback=self.resetXlsx)

    def resetXlsx(self):
        """
        输入O(other)的回调
        关闭已打开的excel，然后重新打开
        """
        self.xbook.getWorkbook(forcedClose=True)

    def __initInfo(self):
        self.__exportSheetIndex = []  # 存储可导表的索引
        self.headerDict = {}  # 导出表第一行转为字典
        self.mapDict = {}  # 代对表生成的字典(第一行是代对表说明忽略)

    def run(self):
        """
        带有$的列数据需要代对表,首先生成代对字典
        """
        self.__initXlsx()  # 初始excel相关
        self.__initInfo()  # 初始导表相关
        self.sth4Nth()  # 进入下一个阶段
        self.constructMapDict()  # 生成代对字典
        self.__onRun()

    def __onRun(self):
        self.writeLines = 0  # 记录已写入的excel的行数
        self.parseDefineLine()  # 分析文件

# 寻找代对表和标记导入的表
    def sth4Nth(self):
        """
        something for nothing, 代对表和导入表需要有
        """
        for index in range(0, self.xbook.getSheetCount()):
            sheetName = self.xbook.getSheetNameByIndex(index)
            if sheetName == config.EXPORT_MAP_SHEET:
                self.__onFindMapSheet(index)
            if sheetName.startswith(config.EXPORT_PREFIX_CHAR):
                self.__onFindExportSheet(index)
        self.onSth4Nth()

    def onSth4Nth(self):
        """
        """
        if not hasattr(self, 'mapIndex'):
            self.xlsxClear(config.EXPORT_ERROR_NOMAP)

        if len(self.__exportSheetIndex) == 0:
            xlsxError.error_input(config.EXPORT_ERROR_NOSHEET)

        return

    def __onFindMapSheet(self, mapIndex):
        self.mapIndex = mapIndex
        return

    def __onFindExportSheet(self, Eindex):
        """
        完毕
        """
        self.__exportSheetIndex.append(Eindex)

    def constructMapDict(self):
        """
        生成代对字典， 代对表只有一个
        """
        mapDict = {}
        sheet = self.xbook.getSheetByIndex(self.mapIndex)
        if not sheet:
            return

        for col in range(0, self.xbook.getRowCount(self.mapIndex)):
            colValues = self.xbook.getColValues(sheet, col)
            if colValues:
                for v in [e for e in colValues[1:] if e and isinstance(e, str) and e.strip()]:
                    print(v)
                    mapStr = v.replace('：', ":")  # 中文"："和":"
                    try:
                        k, v = mapStr.split(":")
                        k = str.strip(k)
                        v = str.strip(v)
                        mapDict[k] = v
                    except Exception as errstr:
                        print("waring：需要检查代对表 第%d列, err=%s" % (col, errstr))
        self.__onConstruct(mapDict)
        return

    def __onConstruct(self, mapDict):
        """
        代对字典生成完毕
        """
        self.mapDict = mapDict
        return

# 文件头检测
    def parseDefineLine(self):
        self.__checkDefine()  # 检查定义是否正确
        self.__checkData()  # 检查数据是否符合规则

    def __reCheck(self, head):
        pattern = "(\w+)(\[.*])(\[\w+\])"
        reGroups = re.compile(pattern).match(head)

        if not reGroups:
            return ()
        return reGroups.groups()

    def __checkDefine(self):
        """
        第一行的个元素是否符合定义格式"name[signs][func]"以及key是否符合规定
        """
        for index in self.__exportSheetIndex:
            print("检测表[%s]文件头(第一行)是否正确" % self.xbook.getSheetNameByIndex(index))
            self.sheetKeys = []
            headList = self.xbook.getRowValues(
                self.xbook.getSheetByIndex(index), config.EXPORT_DEFINE_ROW - 1)
            enName = []  # 检查命名重复临时变量

            self.headerDict[index] = {}
            for c, head in enumerate(headList):
                if head is None or head.strip() == '':  # 导出表的第一行None, 则这一列将被忽略
                    self.__onCheckSheetHeader(self.headerDict[index], c, None)
                    continue

                reTuple = self.__reCheck(head)

                if len(reTuple) == 3:  # 定义被分拆为三部分:name, signs, func, signs可以是空
                    name, signs, funcName = reTuple[0], reTuple[1][1:-1], reTuple[2][1:-1]
                    for s in signs:  # 符号定义是否在规则之内
                        if s not in config.EXPORT_ALL_SIGNS:
                            self.xlsxClear(config.EXPORT_ERROR_NOSIGN,
                                           (config.EXPORT_DEFINE_ROW, c + 1))

                    if config.EXPORT_SIGN_GTH in signs:  # 是否为key
                        self.sheetKeys.append(c)

                    if len(self.sheetKeys) > config.EXPORT_KEY_NUMS:  # key是否超过规定的个数
                        self.xlsxClear(config.EXPORT_ERROR_NUMKEY,
                                       (config.EXPORT_DEFINE_ROW, c + 1))

                    if name not in enName:  # name不能重复
                        enName.append(name)
                    else:
                        self.xlsxClear(config.EXPORT_ERROR_REPEAT,
                                       (self.xbook.getSheetNameByIndex(index).encode(config.FILE_CODE), config.EXPORT_DEFINE_ROW, c + 1))

                    if not hasFunc(funcName):  # funcName是否存在
                        self.xlsxClear(config.EXPORT_ERROR_NOFUNC,
                                       (xlsxtool.toGBK(funcName), c + 1))

                else:
                    self.xlsxClear(config.EXPORT_ERROR_HEADER, (self.xbook.getSheetNameByIndex(
                        index).encode(config.FILE_CODE), config.EXPORT_DEFINE_ROW, c + 1))

                self.__onCheckSheetHeader(
                    self.headerDict[index], c, (name, signs, funcName))  # 定义一行经常使用存起来了

            self.__onCheckDefine()

        return

    def __onCheckSheetHeader(self, DataDict, col, headerInfo):
        DataDict[col] = headerInfo

    def __onCheckDefine(self):
        if len(self.sheetKeys) != config.EXPORT_KEY_NUMS:  # key也不能少
            self.xlsxClear(config.EXPORT_ERROR_NOKEY, ("需要%d而只有%d" % (config.EXPORT_KEY_NUMS, len(self.sheetKeys))))

        print("文件头检测正确", time.ctime(time.time()))

    def sheetIndex2Data(self):
        self.sheet2Data = {}
        for index in self.__exportSheetIndex:
            SheetName = self.xbook.getSheetNameByIndex(index)
            sheetName = SheetName[SheetName.find(config.EXPORT_PREFIX_CHAR) + 1:]
            if sheetName in self.mapDict:
                dataName = self.mapDict[sheetName]
                if dataName in self.sheet2Data:
                    self.sheet2Data[dataName].append(index)
                else:
                    self.sheet2Data[dataName] = [index]

    def __checkData(self):
        """
        列数据是否符合命名规范, 生成所需字典
        """
        self.sheetIndex2Data()
        self.dctDatas = g_dctDatas
        self.hasExportedSheet = []

        for dataName, indexList in self.sheet2Data.items():
            print('开始处理表：%s' % dataName)
            self.curIndexMax = len(indexList)
            self.curProIndex = []
            for index in indexList:
                sheet = self.xbook.getSheetByIndex(index)
                self.curProIndex.append(index)

                rows = self.xbook.getRowCount(index)
                cols = self.xbook.getColCount(index)
                if dataName not in self.dctDatas:
                    self.dctDatas[dataName] = {}
                self.dctData = self.dctDatas[dataName]

                # for row in range(3, rows + 1):
                for row in tqdm.tqdm(range(3, rows + 1), ncols=50):
                    rowval = self.xbook.getRowValues(sheet, row - 1)
                    childDict = {}
                    for col in range(1, cols + 1):
                        val = rowval[col - 1]
                        if val is not None:
                            val = (str(rowval[col - 1]),)
                        else:
                            val = ("",)
                            
                        if self.headerDict[index][col - 1] is None:
                            continue

                        name, sign, funcName = self.headerDict[index][col - 1]
                        if '$' in sign and len(val[0]) > 0:
                            self.needReplace({'v': val[0], "pos": (row, col)})
                            if ',' in val[0]:
                                nv = val[0].strip()
                                vs = nv.split(',')
                                v = ''
                                for item in vs:
                                    v += (self.mapDict[xlsxtool.GTOUC(
                                        xlsxtool.val2Str(item))] + ',')
                                v = v[:-1]  # 去掉最后的','
                            else:
                                # mapDict:key是unicode.key都要转成unicode
                                v = self.mapDict[xlsxtool.GTOUC(
                                    xlsxtool.val2Str(val[0]))]
                        else:
                            v = val[0]
                        if config.EXPORT_SIGN_DOT in sign and v is None:
                            self.xlsxClear(config.EXPORT_ERROR_NOTNULL, (col, row))

                        sv = v

                        func = getFunc(funcName)

                        try:
                            v = func(self.mapDict, self.dctData, childDict, sv)
                        except Exception as errstr:
                            self.xlsxClear(config.EXPORT_ERROR_FUNC, (errstr, funcName, sv, row, col))

                        for ss in sign.replace('$', ''):
                            if len(sv) == 0 and ss == '!':
                                continue
                            config.EXPORT_SIGN[ss](self, {'tableName': dataName, "v": v, "pos": (row, col)})

                        childDict[name] = v

                    self.dctData[self.tempKeys[-1]] = copy.deepcopy(childDict)

            # self.writeHead()

            overFunc = self.mapDict.get('overFunc')
            if overFunc is not None:
                func = getFunc(overFunc)
                self.dctData = func(self.mapDict, self.dctDatas, self.dctData, dataName)
                self.dctDatas[dataName] = self.dctData

            g_dctDatas.update(self.dctDatas)
            self.__onCheckSheet()

        self.writeBody()

    def __onCheckSheet(self):
        if hasattr(self, "tempKeys"):
            del self.tempKeys
        return

    # 符号字典的相关设置EXPORT_SIGN
    def isNotEmpty(self, cellData):
        if cellData['v'] is None:
            self.xlsxClear(config.EXPORT_ERROR_NOTNULL, (cellData['pos'], ))

    def needReplace(self, cellData):
        """宏替代"""
        v = cellData["v"].strip()

        if isinstance(v, float):  # 防止数字报错(1:string) mapDict 是unicode字符串
            v = str(int(v))

        vs = None
        if ',' in v:
            vs = v.split(',')
        else:
            vs = [v]

        for v in vs:
            if v not in self.mapDict:  # 检测而不替换
                self.xlsxClear(config.EXPORT_ERROR_NOTMAP, (cellData['pos'], v))

    def isKey(self, cellData):
        if not hasattr(self, "tempKeys"):
            self.tempKeys = []

        if cellData['v'] not in self.tempKeys:
            self.tempKeys.append(cellData['v'])
        else:
            self.xlsxClear(config.EXPORT_ERROR_REPKEY, (cellData['tableName'], cellData['pos'], (self.tempKeys.index(cellData['v']) + 3, cellData['pos'][1]), cellData['v']))


    def writeHead(self):
        print("开始写入文件:", time.ctime(time.time()))
        try:
            SheetName = self.xbook.getSheetNameByIndex(self.curProIndex[-1])
        except Exception as err:
            print("获取表的名字出错", err)

        sheetName = SheetName[SheetName.find(config.EXPORT_PREFIX_CHAR) + 1:]
        print('表：%s' % sheetName)
        if sheetName in self.mapDict:
            # dataName = self.mapDict[sheetName]
            self.hasExportedSheet.append(self.curProIndex[-1])
        else:
            self.xlsxClear(2, (sheetName.encode(config.FILE_CODE),))

    def writeBody(self):
        for dataName, datas in g_dctDatas.items():
            
            if 'py' in self.targets:
                # 创建目录
                pyPath = os.path.join(self.outfile, 'py')
                if os.path.exists(pyPath) is False:
                    os.makedirs(pyPath)
                # 写py
                pyFilePath = os.path.join(pyPath, 'd_%s.py' % dataName)
                pyHandle = codecs.open(pyFilePath, 'w+', 'utf-8')
                pyHandle.write('datas = %s' % str(datas))
                pyHandle.close()

            if 'json' in self.targets:
                # json.dumps在默认情况下，对于非ascii字符生成的是相对应的字符编码，而非原始字符，只需要ensure_ascii = False
                # sort_keys：是否按照字典排序（a-z）输出，True代表是，False代表否。
                # indent=4：设置缩进格数，一般由于Linux的习惯，这里会设置为4。
                # separators：设置分隔符，在dic = {'a': 1, 'b': 2, 'c': 3}这行代码里可以看到冒号和逗号后面都带了个空格，这也是因为Python的默认格式也是如此，
                # 如果不想后面带有空格输出，那就可以设置成separators=(',', ':')，如果想保持原样，可以写成separators=(', ', ': ')。
                jsonStr = json.dumps(datas, ensure_ascii=False, sort_keys=False, indent=4, separators=(',', ': '))

                jsonPath = os.path.join(self.outfile, 'json')
                if os.path.exists(jsonPath) is False:
                    os.makedirs(jsonPath)
                # 写json
                jsonFilePath = os.path.join(jsonPath, 'd_%s.json' % dataName)
                jsonhandle = codecs.open(jsonFilePath, "w+", 'utf-8')
                jsonhandle.write("%s" % jsonStr)
                jsonhandle.close()
        self.xlsxbyebye()
        print("写完了time:", time.ctime(time.time()))

    def xlsxClose(self):
        """
        关闭文档
        """
        if hasattr(self, "fileHandler"):
            self.fileHandler.close()

        self.xbook.close()
        return

    def xlsxClear(self, errno=0, msg=''):
        """
        程序异常退出清理打开的Excel
        """
        self.xlsxClose()
        if errno > 0:
            raise xlsxError.xe(errno, msg)
        else:
            sys.exit(1)

    def xlsxbyebye(self):
        """
        正常退出
        """
        self.xlsxClose()
        return


config.EXPORT_SIGN['.'] = xlsx2py.isNotEmpty
config.EXPORT_SIGN['$'] = xlsx2py.needReplace
config.EXPORT_SIGN['!'] = xlsx2py.isKey


def main():
    """
    使用方法：
    python3 xlsx2py excelName.xls(x) 输出目录
    """
    try:
        outfile = sys.argv[1]
    except Exception as err:
        print(main.__doc__, err)
        return

    infile = sys.argv[2]
    print("开始导表:[%s]" % (infile))
    if os.path.isfile(infile):
        targets = sys.argv[3:]
        a = xlsx2py(infile, outfile, targets)
        xlsxtool.exportMenu(config.EXPORT_INFO_OK)
        a.run()
    else:
        xlsxError.error_input(config.EXPORT_ERROR_NOEXISTFILE, (infile,))
    print('-------------------------------THE END------------------------------------------------')

    sys.exit()


if __name__ == '__main__':
    main()
    infile = r'E:\github\xlsx2py-json\dist-sample\xlsx\stall.xlsx'
    outfilePath = r'E:\github\xlsx2py-json\dist-sample\datas'
    targets = 'json|py'
    # infile = r'E:\ComblockEngine\2\Games\Config1\xlsx\stall.xlsx'
    # outfilePath = r'E:\ComblockEngine\2\Games\Config1\pydatas'
    if os.path.isfile(infile):
        a = xlsx2py(infile, outfilePath, targets)
        xlsxtool.exportMenu(config.EXPORT_INFO_OK)
        a.run()
