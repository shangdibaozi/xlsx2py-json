import os
import codecs
import json

import xlsxError
import config
import functions

def checkSubPath(outfile, subFolderName):
    path = os.path.join(outfile, subFolderName)
    if os.path.exists(path) is False:
        os.makedirs(path)
    return path

def toPy(outfile, dataName, datas):
    # 创建目录
    pyPath = checkSubPath(outfile, 'py')
    """
    dict = {value}
    value = key:value
    key = int|float|string
    value = int|float|string|[value]|dict
    """

    # 写py
    strList = []
    strList.append('datas = {\n')
    spaces4 = ' ' * 4
    spaces8 = ' ' * 8
    for k in datas:
        key = k
        if isinstance(k, int):
            key = str(k)
        elif isinstance(k, str):
            key = '"%s"' % key
        else:
            raise xlsxError.xe(config.EXPORT_ERROR_KEY_FLOAT, dataName)

        strList.append('%s%s: {\n' % (spaces4, key))

        for e in datas[k]:
            key1 = e
            if isinstance(e, int):
                key1 = str(e)
            elif isinstance(e, str):
                key1 = '"%s"' % key1
            else:
                raise xlsxError.xe(config.EXPORT_ERROR_KEY_FLOAT, dataName)

            value = datas[k][e]
            if isinstance(value, int) or isinstance(value, float) or isinstance(value, tuple) or isinstance(value, list):
                value = str(value)
            elif isinstance(value, str):
                value = '"%s"' % value
            else:
                raise xlsxError.xe(config.EXPORT_ERROR_NOFUNC, dataName)

            strList.append('%s%s: %s' % (spaces8, key1, value))
            strList.append(',\n')

        strList.pop()
        strList.append('\n%s}' % spaces4)
        strList.append(',\n')
    
    strList.pop()
    strList.append('\n}\n')
    pyFilePath = os.path.join(pyPath, 'd_%s.py' % dataName)
    pyHandle = codecs.open(pyFilePath, 'w+', 'utf-8')
    dataStr = ''.join(strList)
    pyHandle.write(dataStr)
    pyHandle.close()

def toJson(outfile, dataName, datas):
    # json.dumps在默认情况下，对于非ascii字符生成的是相对应的字符编码，而非原始字符，只需要ensure_ascii = False
    # sort_keys：是否按照字典排序（a-z）输出，True代表是，False代表否。
    # indent=4：设置缩进格数，一般由于Linux的习惯，这里会设置为4。
    # separators：设置分隔符，在dic = {'a': 1, 'b': 2, 'c': 3}这行代码里可以看到冒号和逗号后面都带了个空格，这也是因为Python的默认格式也是如此，
    # 如果不想后面带有空格输出，那就可以设置成separators=(',', ':')，如果想保持原样，可以写成separators=(', ', ': ')。
    jsonStr = json.dumps(datas, ensure_ascii=False, sort_keys=False, indent=4, separators=(',', ': '))

    jsonPath = checkSubPath(outfile, 'json')
    # 写json
    jsonFilePath = os.path.join(jsonPath, 'd_%s.json' % dataName)
    jsonhandle = codecs.open(jsonFilePath, "w+", 'utf-8')
    jsonhandle.write("%s" % jsonStr)
    jsonhandle.close()


def toLua(outfile, dataName, datas):
    pass


def generateCSharpTypeFile(fileName, outfile, headDict):
    cSharpPath = checkSubPath(outfile, 'C#')

    strList = []
    for tbName in headDict:
        properties = headDict[tbName]
        oneClass = [f'public class Tbl_{tbName}\n']
        oneClass.append('{\n')
        for propertyInfo in properties.values():
            if propertyInfo is not None:
                name, _, funcName = propertyInfo
                oneClass.append(f'    public {functions.functionType2CSharpType[funcName]} {name};\n')
        oneClass.append('}\n\n')
        strList.append(''.join(oneClass))

    cSharpFilePath = os.path.join(cSharpPath, 'd_%s.cs' % fileName)
    fileHandler = codecs.open(cSharpFilePath, 'w+', 'utf-8')
    dataStr = ''.join(strList)
    fileHandler.write(dataStr)
    fileHandler.close()
