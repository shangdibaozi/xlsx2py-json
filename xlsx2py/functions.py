
def funcPos2D(mapDict, dctData, chilidDict, data):
    """
    返回int数据
    """
    if data is None or (type(data) == str and len(data) == 0):
        return ()

    data = str(data)

    return (int(data.split(",")[0]), 0, int(data.split(",")[1]))

def funcInt(mapDict, dctData, chilidDict, data):
    """
    返回int数据
    """
    if len(data) == 0:
        return 0

    # 16进制
    if '0x' in data or '0X' in data:
        return int(data, 16)

    # 科学计数法
    if 'e' in data or 'E' in data:
        data = eval(data)
    
    return int(data)

def funcLong(mapDict, dctData, chilidDict, data):
    return funcInt(mapDict, dctData, chilidDict, data)

def funcUInt(mapDict, dctData, chilidDict, data):
    """
    返回int数据
    """
    if len(data) == 0:
        return 0

    # 16进制
    if '0x' in data or '0X' in data:
        return int(data, 16)

    # 科学计数法
    if 'e' in data or 'E' in data:
        data = eval(data)  # eval返回的是float类型
    
    value = int(data)

    if value < 0:
        raise Exception(f'数值为负：{value}')

    return value


def funcFloat(mapDict, dctData, chilidDict, data):
    """
    返回float数据，保留2位小数
    """
    if data is None or (type(data) == str and len(data) == 0):
        return 0.0

    return round(float(data), 2)

def funcStr(mapDict, dctData, chilidDict, data):
    """
    返回字符串数据
    """
    if data is None:
        return ""

    if type(data) == str:
        return data
    else:
        data = str(data)
        data = data.encode('utf8')
        return str(data)

def funcEval(mapDict, dctData, chilidDict, data):
    """
    返回eval数据
    """
    if data is None or (type(data) == str and len(data) == 0):
        return ""
    return eval(data)

def funcTupleInt(mapDict, dctData, chilidDict, data):
    """
    返回tuple数据
    """
    if data is None or (type(data) == str and len(data) == 0):
        return ()

    data = str(data)

    return tuple([int(e) for e in data.split(",") if len(e) > 0])

def funcTupleLong(mapDict, dctData, chilidDict, data):
    return funcTupleInt(mapDict, dctData, chilidDict, data)

def funcTupleUInt(mapDict, dctData, chilidDict, data):
    """
    返回tuple数据
    """
    if data is None or (type(data) == str and len(data) == 0):
        return ()

    arr = []
    for e in str(data).split(','):
        if len(e) > 0:
            val = int(e)
            if val < 0:
                raise Exception(f'数值为负：{val}')
            arr.append(val)

    return tuple(arr)


def funcTupleFloat(mapDict, dctData, chilidDict, data):
    """
    返回tuple数据
    """
    if data is None or (type(data) == str and len(data) == 0):
        return ()

    data = str(data)

    return tuple([float(e) for e in data.split(",") if len(e) > 0])
    
def funcDict(mapDict, dctData, chilidDict, data):
    """
    返回dict数据
    "xx:1'2'3;fff:2'3'4"
    """
    if data is None or (type(data) == str and len(data) == 0):
        return ''
    
    data = str(data)
    dict1 = {}
    for item in data.split(';'):
        if item != '':
            e = item.split(':')
            if len(e) == 1:
                dict1[int(e[0])] = ()
            elif len(e) == 2:
                dict1[int(e[0])] = tuple([index for index in e[1].split('`') if index != ''])

    return dict1

def funcTupleStr(mapDict, dctData, chilidDict, data):
    """
    返回tuple数据
    """
    if data is None or (type(data) == str and len(data) == 0):
        return ()

    data = str(data)
    return tuple([e for e in data.split(",") if len(e) > 0])

def funcTupleEval(mapDict, dctData, chilidDict, data):
    """
    返回tuple数据
    """
    if data is None or (type(data) == str and len(data) == 0):
        return ()

    data = str(data)
    return tuple([eval(e) for e in data.split(",") if len(e) > 0])

def funcTupleEvalMD(mapDict, dctData, chilidDict, data):
    """
    返回tuple数据 使用代对表
    """
    if data is None or (type(data) == str and len(data) == 0):
        return ()
    
    data = str(data)
    try:
        return tuple([eval(mapDict[e.decode("gb2312")]) for e in data.split(",") if len(e) > 0])
    except Exception as errstr:
        print("函数中发生错误:%s" % errstr)
        return ()
    
def funcTupleEval1(mapDict, dctData, chilidDict, data):
    """
    返回tuple数据
    1'100/2'100/3'54
    """
    if data is None or (type(data) == str and len(data) == 0):
        return ()

    data = str(data)
    ret = []
    for e in data.split("/"):
        try:
            i, v = e.split("'")
        except Exception as errstr:
            print("函数中发生错误:%s" % errstr)
            continue
        ret.append((eval(i), eval(v)))
    return tuple(ret)
    
def funcBool(mapDict, dctData, chilidDict, data):
    """
    返回布尔值
    """
    if data is None or (type(data) == str and len(data) == 0):
        return False
    return float(data) > 0.0001
    # return int(data) > 0 # 不知道为什么，excel里面的0读取出来的时候是'0.0'

def funcNotBool(mapDict, dctData, chilidDict, data):
    """
    返回取反的布尔值
    """
    return not funcBool(mapDict, dctData, chilidDict, data)

def funcNull(mapDict, dctData, chilidDict, data):
    """
    什么也不做 直接返回
    """
    return data

def funcZipFloat(mapDict, dctData, chilidDict, data):
    """
    返回float数据
    """
    if data is None or (type(data) == str and len(data) == 0):
        return 0

    return int(float(data) * 10000)

def funcUNZipFloat(mapDict, dctData, chilidDict, data):
    """
    返回float数据
    """
    if data is None or (type(data) == str and len(data) == 0):
        return 0.0

    return int(data) / 10000.0
    
def funcFlags(mapDict, dctData, chilidDict, data):
    """
    返回标记组合数据
    比如： 想在excel上配置标记组合
    近程攻击:0x00000001
    远程攻击:0x00000002
    暴击:0x00000004
    用此函数可以输出多个标记组成一个uint32的数字
    """
    val = 0
    for x in data.split(","):
        if len(x) > 0:
            val |= int(mapDict[x])

    return val


functionType2PyType = {
    'funcBool': 'bool',
    'funcFloat': 'float',
    'funcInt': 'int',
    'funcStr': 'str',
    'funcTupleInt': 'List[int]',
    'funcTupleStr': 'List[str]',
    'funcTupleFloat': 'List[float]'
}


functionType2CSharpType = {
    'funcBool': 'bool',
    'funcFloat': 'float',
    'funcInt': 'int',
    'funcUInt': 'uint',
    'funcTupleInt': 'int[]',
    'funcTupleUInt': 'uint[]',
    'funcTupleFloat': 'float[]',
    'funcTupleLong': 'long[]',
    'funcStr': 'string',
    'funcLong': 'long'
}
