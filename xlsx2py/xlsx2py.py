import sys
import re
import os
import signal
import time
import copy
import tqdm

from ExcelTool import ExcelTool
import functions
import xlsxtool
import xlsxError
import ExportType

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
        self.tempKeys = []

    def __initXlsx(self):
        self.xbook = ExcelTool(self.infile)
        self.xbook.getWorkbook()

    def __initInfo(self):
        self.__exportSheetName = []  # 存储可导表的索引
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
        self.parseDefineLine()  # 分析文件

    # 寻找代对表和标记导入的表
    def sth4Nth(self):
        """
        something for nothing, 代对表和导入表需要有
        """

        # 获得所有需要导出的表
        for sheetName in self.xbook.getSheetNames():
            if sheetName.startswith(config.EXPORT_PREFIX_CHAR):
                self.__exportSheetName.append(sheetName)
        
        # 检查是否有代对表
        if not self.xbook.getSheetBySheetName(config.EXPORT_MAP_SHEET):
            self.xlsxClear(config.EXPORT_ERROR_NOMAP)

        # 检查导出表数量
        if len(self.__exportSheetName) == 0:
            xlsxError.error_input(config.EXPORT_ERROR_NOSHEET)

    def constructMapDict(self):
        """
        生成代对字典， 代对表只有一个
        """
        mapDict = self.mapDict
        sheet = self.xbook.getSheetBySheetName(config.EXPORT_MAP_SHEET)
        for col in range(0, self.xbook.getRowCount(config.EXPORT_MAP_SHEET)):
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
                        mapDict[v] = f'@{k}'
                    except Exception as errstr:
                        print("waring：需要检查代对表 第%d列, err=%s" % (col, errstr))

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
        for sheetName in self.__exportSheetName:
            print("检测表[%s]文件头(第一行)是否正确" % sheetName)
            self.sheetKeys = []
            headList = self.xbook.getRowValues(
                self.xbook.getSheetBySheetName(sheetName), config.EXPORT_DEFINE_ROW - 1)
            enName = []  # 检查命名重复临时变量

            self.headerDict[sheetName] = {}
            for c, head in enumerate(headList):
                if head is None or head.strip() == '':  # 导出表的第一行None, 则这一列将被忽略
                    self.headerDict[sheetName][c] = None
                    continue

                reTuple = self.__reCheck(head)

                if len(reTuple) == 3:  # 定义被分拆为三部分:name, signs, func, signs可以是空
                    name, signs, funcName = reTuple[0], reTuple[1][1:-1], reTuple[2][1:-1]
                    for s in signs:  # 符号定义是否在规则之内
                        if s not in config.EXPORT_ALL_SIGNS:
                            self.xlsxClear(config.EXPORT_ERROR_NOSIGN, (config.EXPORT_DEFINE_ROW, c + 1))

                    if config.EXPORT_SIGN_GTH in signs:  # 是否为key
                        self.sheetKeys.append(c)

                    if len(self.sheetKeys) > config.EXPORT_KEY_NUMS:  # key是否超过规定的个数
                        self.xlsxClear(config.EXPORT_ERROR_NUMKEY, (config.EXPORT_DEFINE_ROW, c + 1))

                    if name not in enName:  # name不能重复
                        enName.append(name)
                    else:
                        self.xlsxClear(config.EXPORT_ERROR_REPEAT, (sheetName.encode(config.FILE_CODE), config.EXPORT_DEFINE_ROW, c + 1))

                    if not hasFunc(funcName):  # funcName是否存在
                        self.xlsxClear(config.EXPORT_ERROR_NOFUNC, (xlsxtool.toGBK(funcName), c + 1))

                else:
                    self.xlsxClear(config.EXPORT_ERROR_HEADER, (sheetName.encode(config.FILE_CODE), config.EXPORT_DEFINE_ROW, c + 1))

                self.headerDict[sheetName][c] = (name, signs, funcName)

            self.__onCheckDefine()

    def __onCheckDefine(self):
        if len(self.sheetKeys) != config.EXPORT_KEY_NUMS:  # key也不能少
            self.xlsxClear(config.EXPORT_ERROR_NOKEY, ("需要%d而只有%d" % (config.EXPORT_KEY_NUMS, len(self.sheetKeys))))

        print("文件头检测正确", time.ctime(time.time()))

    def sheetName2Data(self):
        self.sheet2Data = {}
        for sheetName in self.__exportSheetName:
            exportSheetName = sheetName[1:]  # 截取表名：@tbname -> tbname
            if exportSheetName in self.mapDict:
                dataName = self.mapDict[exportSheetName]  # 拿到要导出的表名：sheetName可能为中文，需要代对表将表名映射出去
                if dataName in self.sheet2Data:
                    self.sheet2Data[dataName].append(sheetName)
                else:
                    self.sheet2Data[dataName] = [sheetName]

    def __checkData(self):
        """
        列数据是否符合命名规范, 生成所需字典
        """
        self.sheetName2Data()
        self.dctDatas = g_dctDatas
        self.hasExportedSheet = []

        for dataName, sheetNameLst in self.sheet2Data.items():
            print('开始处理表：%s' % dataName)
            self.curProIndex = []
            for sheetName in sheetNameLst:
                sheet = self.xbook.getSheetBySheetName(sheetName)
                self.curProIndex.append(sheetName)

                rows = self.xbook.getRowCount(sheetName)
                cols = self.xbook.getColCount(sheetName)
                if dataName not in self.dctDatas:
                    self.dctDatas[dataName] = {}
                self.dctData = self.dctDatas[dataName]


                # for row in range(3, rows + 1):
                for row in tqdm.tqdm(range(3, rows + 1), ncols=50):
                    keyName: str = None
                    rowval = self.xbook.getRowValues(sheet, row - 1)
                    childDict = {}
                    for col in range(1, cols + 1):
                        val = rowval[col - 1]
                        if val is not None:
                            val = (str(rowval[col - 1]),)
                        else:
                            val = ("",)
                            
                        if self.headerDict[sheetName][col - 1] is None:
                            continue

                        name, sign, funcName = self.headerDict[sheetName][col - 1]
                        if '$' in sign and len(val[0]) > 0:
                            self.needReplace({'v': val[0], "pos": (row, col)})
                            if ',' in val[0]:
                                nv = val[0].strip()
                                vs = nv.split(',')
                                v = ''
                                for item in vs:
                                    v += (self.mapDict[xlsxtool.GTOUC(xlsxtool.val2Str(item))] + ',')
                                v = v[:-1]  # 去掉最后的','
                            else:
                                # mapDict:key是unicode.key都要转成unicode
                                v = self.mapDict[xlsxtool.GTOUC(xlsxtool.val2Str(val[0]))]
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
                            if ss == '!':
                                keyName = name

                            config.EXPORT_SIGN[ss](self, {'tableName': dataName, "v": v, "pos": (row, col)})

                        childDict[name] = v

                    if keyName is not None:
                        self.dctData[childDict[keyName]] = copy.deepcopy(childDict)

            overFunc = self.mapDict.get('overFunc')
            if overFunc is not None:
                func = getFunc(overFunc)
                self.dctData = func(self.mapDict, self.dctDatas, self.dctData, dataName)
                self.dctDatas[dataName] = self.dctData

            self.tempKeys.clear()
            g_dctDatas.update(self.dctDatas)

        self.writeBody()

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

    def checkKey(self, cellData):
        """
        检测是否有重复的键值
        """
        if cellData['v'] not in self.tempKeys:
            self.tempKeys.append(cellData['v'])
        else:
            self.xlsxClear(config.EXPORT_ERROR_REPKEY, (cellData['tableName'], cellData['pos'], (self.tempKeys.index(cellData['v']) + 3, cellData['pos'][1]), cellData['v']))

    def writeBody(self):
        print('writeBody %s' % self.targets)
        for exportTableName, datas in g_dctDatas.items():
            if 'py' in self.targets:
                print('export py')
                ExportType.toPy(self.outfile, exportTableName, datas)

            if 'json' in self.targets:
                print('导出json配置')
                ExportType.toJson(self.outfile, exportTableName, datas)

            if 'lua' in self.targets:
                ExportType.toLua(self.outfile, exportTableName, datas)

        if 'C#' in self.targets or 'c#' in self.targets:
            # 将一个excel上所有表声明文件放在一个C#文件内
            fileName, _ = os.path.splitext(os.path.basename(self.infile))
            exportTypes = {}
            for exportTableName, _ in g_dctDatas.items():
                exportTypes[exportTableName] = self.headerDict[self.mapDict[exportTableName]]
            ExportType.generateCSharpTypeFile(fileName, self.outfile, exportTypes)

        self.xlsxbyebye()
        print("写完了time:", time.ctime(time.time()))

    def xlsxClose(self):
        """
        关闭文档
        """
        if hasattr(self, "fileHandler"):
            self.fileHandler.close()

        self.xbook.close()

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


config.EXPORT_SIGN['.'] = xlsx2py.isNotEmpty
config.EXPORT_SIGN['$'] = xlsx2py.needReplace
config.EXPORT_SIGN['!'] = xlsx2py.checkKey


def main():
    """
    使用方法：
    set datas=datas/
    set excel=xlsx/stall.xlsx
    set targets=json py
    echo on
    xlsx2py.exe %datas% %excel% %targets%
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
    infile = r'E:\xWorld\trunk\Config\xlsx\VoxelEditor.xlsx'
    outfilePath = r'E:\github\xlsx2py-json\dist-sample\datas'
    targets = ['json', 'py', 'C#']
    # infile = r'E:\ComblockEngine\2\Games\Config1\xlsx\stall.xlsx'
    # outfilePath = r'E:\ComblockEngine\2\Games\Config1\pydatas'
    if os.path.isfile(infile):
        a = xlsx2py(infile, outfilePath, targets)
        xlsxtool.exportMenu(config.EXPORT_INFO_OK)
        a.run()
