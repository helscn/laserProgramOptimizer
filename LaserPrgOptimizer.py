#!/usr/bin/python
# -*-coding:utf-8-*-

from __future__ import print_function
import sys
import os
import re
import platform
if sys.version_info[0] == 3:
    # 导入Python3的Tkinter模块
    from tkinter import Tk
    from tkinter.messagebox import showinfo, showerror
    from tkinter.filedialog import askopenfilename
elif sys.version_info[0] == 2:
    # 针对Python2将默认编码改为utf8
    import sys
    reload(sys)
    sys.setdefaultencoding('utf8')
    # 导入Python2的Tkinter模块
    from Tkinter import Tk
    from tkMessageBox import showinfo, showerror
    from tkFileDialog import askopenfilename
else:
    sys.exit(1)

BLOCK_SIZE = 46                 # 指定转换的扫描区块大小及路径优化时的区块间距大小，单位mm

if 'Windows' in platform.platform():
    LINE_BREAK = '\n'           # 写入文件时的换行符，Windows系统下设为\n，Linux系统下设为\r\n
else:
    LINE_BREAK = '\r\n'

# 隐藏Tk主界面窗口
root = Tk()
root.withdraw()
root.update()


class GridData(object):
    """
    二维栅格化矩阵数据类对象，可以根据数据的X,Y坐标将其对应保存在矩阵中等间距分隔的Cell单元格数组中
    矩阵通过三级嵌套数组保存数据，一级索引为矩阵Cell列X坐标，二级索引为Cell行Y坐标，三级索引为保存在Cell单元格中的数据编号
    矩阵单元格的间距由创建对象时进行指定，而宽度和高度会根据数据的坐标位置自动扩展
    扩展后的矩阵的基准Cell(行、列坐标均为零)会平移至所有数据中最小X,最小Y坐标所在的Cell，以保证所有Cell索引编号不小于零
    """

    def __init__(self, pitchX, pitchY):
        """
        初始化GridData对象.
        @param pitchX:Grid对象中每个Cell区块在X方向上的间隔大小.
        @param pitchY:Grid对象中每个Cell区块在Y方向上的间隔大小.
        @return:None
        """
        self.__pitchX = abs(pitchX)
        self.__pitchY = abs(pitchY)
        self.__originX = None
        self.__originY = None
        self.__width = 0
        self.__height = 0
        self.__gridData = [[[]]]

    @property
    def pitchX(self):
        """
        只读属性，返回当前Grid对象在X轴方向划分Cell的间隔
        @return: Grid在X轴方向的Cell间距
        @rtype: float or int
        """
        return self.__pitchX

    @property
    def pitchY(self):
        """
        只读属性，返回当前Grid对象在Y轴方向划分Cell的间隔
        @return: Grid在Y轴方向的Cell间距
        @rtype: float or int
        """
        return self.__pitchY

    @property
    def width(self):
        """
        只读属性，返回Grid在X方向的宽度，其值为X方向最大索引编号+1
        @return: Grid在X方向的Cell个数
        @rtype: int
        """
        return self.__width

    @property
    def height(self):
        """
        只读属性，返回Grid在Y方向的高度，其值为Y方向最大索引编号+1
        @return: Grid在Y方向的Cell个数
        @rtype: int
        """
        return self.__height

    def posParser(self, data):
        """
        数据中的X,Y坐标值的解析器函数，可以重写此函数方法以便正确识别数据坐标
        默认识别包含x/y的属性或元素的对象,或者为list/tuple前二个值(第0个值为x，第1个值为y).
        @param data: 需要解析坐标的数据对象
        @return: 返回数据中(x,y)坐标的元组
        @rtype: tuple(x,y)
        @raise ValueError:无法解析x,y坐标
        """
        x = None
        y = None
        if hasattr(data, 'x'):
            x = data.x
        elif isinstance(data, (list, tuple)):
            x = data[0]
        elif hasattr(data, '__getitem__'):
            x = data['x']

        if hasattr(data, 'y'):
            y = data.y
        elif isinstance(data, (list, tuple)):
            y = data[1]
        elif hasattr(data, '__getitem__'):
            y = data['y']

        if x is None or y is None:
            raise ValueError('Can not parse the posX or posY in data')
        else:
            return (x, y)

    def addItem(self, data):
        """
        根据数据的坐标位置添加至对应的Grid表Cell单元中
        @param data: 需要增加的数据，其中数据保存时的x,y坐标由创建对象时的posParser解析器对data解析
        @return: None
        """
        x, y = self.posParser(data)

        # 第一次添加数据时初始化保存数据的数组
        if self.__originX is None:
            self.__originX = x
            self.__originY = y
            self.__width = 1
            self.__height = 1

        # X坐标大于当前Grid上界时扩充X正方向边界
        while int(round((x-self.__originX)/self.__pitchX))+1 > self.__width:
            column = []
            for i in range(self.__height):
                column.append([])
            self.__gridData.append(column)
            self.__width += 1

        # X坐标小于当前Grid下界时扩充X负方向边界
        while int(round((x-self.__originX)/self.__pitchX)) < 0:
            column = []
            for i in range(self.__height):
                column.append([])
            self.__gridData.insert(0, column)
            self.__width += 1
            self.__originX -= self.__pitchX

        # Y坐标大于当前Grid上界时扩充Y正方向边界
        while int(round((y-self.__originY)/self.__pitchY))+1 > self.__height:
            for column in self.__gridData:
                column.append([])
            self.__height += 1

        # Y坐标小于当前Grid下界时扩充X负方向边界
        while int(round((y-self.__originY)/self.__pitchY)) < 0:
            for column in self.__gridData:
                column.insert(0, [])
            self.__height += 1
            self.__originY -= self.__pitchY

        posX = int(round((x-self.__originX)/self.__pitchX))
        posY = int(round((y-self.__originY)/self.__pitchY))
        self.__gridData[posX][posY].append(data)

    def getIndex(self, data):
        """
        根据数据的X，Y坐标计算保存在Grid中的索引位置
        @param data:需要查询Grid中索引位置的数据,其中数据的x,y坐标由posParser解析器解析
        @return:返回待查询数据在Grid中的索引元组 tuple(column,row)，超出Grid界限时返回 tuple(None,None)
        @rtype:tuple(column,row)
        @raise IndexError:待查询的数据索引超过了当前Grid的界限范围
        """
        x, y = self.posParser(data)
        column = int(round((x-self.__originX)/self.__pitchX))
        row = int(round((y-self.__originY)/self.__pitchY))
        if 0 <= column < self.__width and 0 <= row < self.height:
            return (column, row)
        else:
            raise IndexError('The data index out of grid range')

    def getItems(self, column, row):
        """
        根据(column,row)索引返回对应Cell区间中的所有元素
        @param column:要查询的Cell在x方向的列索引，索引可以为负值
        @type column:int
        @param row:要查询的Cell在y方向的行索引，索引可以为负值
        @type row:int
        @return:返回包含指定Cell中全部元素的list对象
        @rtype:list
        @raise IndexError:待查询的数据索引超过了当前Grid的界限范围
        """
        return self.__gridData[column][row]

    def getItem(self, column, row, index):
        """
        根据(column,row)索引返回对应Cell区间中index指定的元素
        @param column:要查询的Cell在x方向的列索引，索引可以为负值
        @type column:int
        @param row:要查询的Cell在y方向的行索引，索引可以为负值
        @type row:int
        @param index:需要查询Cell中的第几个元素，可以为负值
        @type index:int
        @return: 返回(column,row)的Cell中index指定的元素
        @rtype: object
        @raise IndexError:待查询的数据索引超过了当前Grid的界限范围
        """
        return self.__gridData[column][row][index]

    def popItem(self, column, row, index=-1):
        """
        根据(column,row)索引取出指定Cell中index指定的元素
        @param column: 要查询的Cell在x方向的列索引，索引可以为负值
        @type column: int
        @param row: 要查询的Cell在y方向的行索引，索引可以为负值
        @type row: int
        @param index: 需要pop出Cell中的第几个元素，可以为负值，默认为-1
        @type index: int
        @return: 返回(column,row)的Cell中index指定的元素，同时将该元素从Grid中删除
        @rtype: object
        @raise IndexError:待查询的数据索引超过了当前Grid的界限范围
        """
        return self.__gridData[column][row].pop(index)

    def countItems(self, column=None, row=None):
        """
        根据(column,row)索引返回对应Cell区间中的元素个数
        @param column:要查询的Cell在x方向的列索引，索引可以为负值，为None时查询整行的元素个数
        @type column:int
        @param row:要查询的Cell在y方向的行索引，索引可以为负值，为None时查询整列的元素个数
        @type row:int
        @return:返回指定Cell中所有元素的个数，column和row均为None时返回所有元素个数
        @rtype:int
        @raise IndexError:待查询的数据索引超过了当前Grid的界限范围
        """
        count = 0
        if column is None and row is None:
            for column in range(self.__width):
                for row in range(self.__height):
                    count += len(self.__gridData[column][row])
        elif column is None:
            for c in range(self.__width):
                count += len(self.__gridData[c][row])
        elif row is None:
            for r in range(self.__height):
                count += len(self.__gridData[column][r])
        else:
            count = len(self.__gridData[column][row])
        return count

    def delAllItems(self):
        """
        删除所有保存的数据
        @return: None
        """
        self.__originX = None
        self.__originY = None
        self.__width = 0
        self.__height = 0
        self.__gridData = [[[]]]

    def optimizeGrid(self):
        """
        删除Grid外围边缘Cell中不包含任何元素的空行、空列，以减小数据占用空间，注意优化后数据的索引号可能发生变更
        当Grid中没有任何元素时至少会保留一个空的Cell
        @return:None
        """
        while self.countItems(column=0) == 0 and self.__width > 1:
            # 删除左侧空列
            self.__gridData.pop(0)
            self.__originX += self.__pitchX
            self.__width -= 1
        while self.countItems(column=self.__width-1) == 0 and self.__width > 1:
            # 删除右侧空列
            self.__gridData.pop(-1)
            self.__width -= 1
        while self.countItems(row=0) == 0 and self.__height > 1:
            # 删除下方空行
            for i in range(self.__width):
                self.__gridData[i].pop(0)
            self.__originY += self.__pitchY
            self.__height -= 1
        while self.countItems(row=self.__height-1) == 0 and self.__height > 1:
            # 删除上方空行
            for i in range(self.__width):
                self.__gridData[i].pop(-1)
            self.__height -= 1


def parseNumber(s, length=3, lead_zero=False):
    """
    根据Excellon指令中的文本坐标数据识别转换为实际坐标值
    @param s: 需要识别的文本坐标数据
    @type s: str
    @param length: 数据位数长度，当lead_zero为True时代表整数位数，当lead_zero为False(即trail_zero)时代表小数位数
    @type length: int
    @param lead_zero: 设置数据是否为前导零格式，值为Fasle时表示数据为后补零格式，此时length指定了数据的小数位数
    @type lead_zero: bool
    @return: 根据文本数据识别出的实际数值
    @rtype: float
    """

    # 删除文本头尾的空格
    s = re.sub(r'(^\s+|\s+$)', '', s)
    # 如果文本中含有小数点直接转换为浮点数
    if '.' in s:
        return float(s)
    # 判断数据的正负号
    if s[0] == '-':
        sign = -1
        s = s[1:]
    else:
        sign = 1
    if lead_zero:
        # 识别前导零格式数据，其中length为整数位数
        if len(s) < length:
            # 当数据位数小于整数位数时根据缺少的位置进行进位
            v = sign * float(s) * pow(10, length-len(s))
        else:
            # 数据位数大于等于整数位时在整数位后插入小数点
            v = sign * float(s[0:length]+'.'+s[length:])
    else:
        # 将识别后导零格式数据，其中length为小数位数
        v = sign * float(s)/pow(10, length)
    return v


def parseBlockXY(s):
    """
    解析镭射机加工区块的x,y坐标
    @param s: 三镭射机加工区块指令，如 N1G1X10Y300
    @type s: str
    @return: 获取三菱机区块指令中的区块坐标，其中坐标已经转换为象限1(-90度旋转)
    @rtype: tuple(x,y)
    """
    # 识别三菱机区块指令的正则表达式
    regBlock = re.compile(r'N\d+G1X(-?\d+)Y(-?\d+)')
    result = regBlock.match(s)
    if result:
        BlockX, BlockY = result.groups()
        # 将三菱机区块指令中的坐标转换为象限1(-90度旋转)，并返回转换后的坐标元组
        return (-1*parseNumber(BlockY), parseNumber(BlockX))


def optimizeBlockOrder(grid, clockwise=True, mirrorX=True):
    """
    将GridData中保存的三菱机区块指令按照从外向内加工路径依次输出的生成器
    @param grid: 保存三菱机区块数据的GridData对象
    @param clockwise: 输出区块时是否按顺时针顺序进行输出，值为True时按顺时针输出，值为False时按逆时针输出
    @param mirrorX: 反面加工钻带的镜像方式，值为True时按照X轴镜像从右下角开始逆时针输出，值为False时按照Y镜从左上角开始逆时针输出，只有当clockwise为False时生效
    @return :该函数为生成器函数，每次yield一个区块，直至取出所有区块后清空grid中的数据
    @rtype: None
    """
    while grid.countItems() > 0:
        minX = 0
        minY = 0
        maxX = grid.width-1
        maxY = grid.height-1
        if maxX == 0 or maxY == 0:
            # 判断矩阵只有单行或单列时的特殊情况，只需循环一次行或列
            if clockwise:
                # 正面钻带加工顺序
                for y in range(maxY+1):
                    for x in range(maxX+1):
                        if grid.countItems(x, y) > 0:
                            yield grid.popItem(x, y, 0)
            else:
                # 反面钻带加工顺序
                if mirrorX:
                    # 反面使用X轴镜像时
                    for y in range(maxY+1):
                        for x in range(maxX, -1, -1):
                            if grid.countItems(x, y) > 0:
                                yield grid.popItem(x, y, 0)
                else:
                    # 反面使用Y轴镜像时
                    for y in range(maxY,-1,-1):
                        for x in range(maxX+1):
                            if grid.countItems(x, y) > 0:
                                yield grid.popItem(x, y, 0)
        else:
            # 当矩阵行和列均大于1时，需要循环一圈yield区块
            if clockwise:
                # 正面钻带加工顺序为左下角起始的顺时针方向
                for y in range(minY, maxY):
                    # 从左下到左上
                    if grid.countItems(minX, y) > 0:
                        yield grid.popItem(minX, y, 0)
                for x in range(minX, maxX):
                    # 从左上到右上
                    if grid.countItems(x, maxY) > 0:
                        yield grid.popItem(x, maxY, 0)
                for y in range(maxY, minY, -1):
                    # 从右上到右下
                    if grid.countItems(maxX, y) > 0:
                        yield grid.popItem(maxX, y, 0)
                for x in range(maxX, minX, -1):
                    # 从右下到左下
                    if grid.countItems(x, minY) > 0:
                        yield grid.popItem(x, minY, 0)
            else:
                # 反面钻带加工顺序为右下角起始的逆时针方向
                if mirrorX:
                    # 反面使用X轴镜像时
                    for y in range(minY, maxY):
                        # 从右下到右上
                        if grid.countItems(maxX, y) > 0:
                            yield grid.popItem(maxX, y, 0)
                    for x in range(maxX, minX, -1):
                        # 从右上到左上
                        if grid.countItems(x, maxY) > 0:
                            yield grid.popItem(x, maxY, 0)
                    for y in range(maxY, minY, -1):
                        # 从左上到左下
                        if grid.countItems(minX, y) > 0:
                            yield grid.popItem(minX, y, 0)
                    for x in range(minX, maxX):
                        # 从左下到右下
                        if grid.countItems(x, minY) > 0:
                            yield grid.popItem(x, minY, 0)
                else:
                    # 反面使用Y轴镜像时
                    for y in range(maxY, minY, -1):
                        # 从左上到左下
                        if grid.countItems(minX, y) > 0:
                            yield grid.popItem(minX, y, 0)
                    for x in range(minX, maxX):
                        # 从左下到右下
                        if grid.countItems(x, minY) > 0:
                            yield grid.popItem(x, minY, 0)
                    for y in range(minY, maxY):
                        # 从右下到右上
                        if grid.countItems(maxX, y) > 0:
                            yield grid.popItem(maxX, y, 0)
                    for x in range(maxX, minX, -1):
                        # 从右上到左上
                        if grid.countItems(x, maxY) > 0:
                            yield grid.popItem(x, maxY, 0)
        # 清除外围取出数据后的空Cell，只有矩阵外围四边全部为空Cell时才会删除外围空Cell
        if grid.countItems(column=0) == 0 and grid.countItems(column=maxX) == 0 and grid.countItems(row=0) == 0 and grid.countItems(row=maxY) == 0:
            grid.optimizeGrid()
    grid.delAllItems()


def getGlvFileIndex(N, glvFiles):
    """
    根据当前区块编号返回区块所在的.glv文件索引
    @param N: 需要查询所在GLV文件中的区块编号
    @type N: int
    @param glvFiles: 保存每个GLV文件起始区块编号的列表，列表的索引编号代表GLV的文件编号，值代表该GLV文件中第一个区块编号
    @type glvFiles: list
    @return: 返回区块N所在的GLV文件编号，值为0表示在第一个glv文件，值为1表示在第二个glv文件，依此类推
    @rtype: int
    """
    for i in range(len(glvFiles)-1, -1, -1):
        if N >= glvFiles[i]:
            return i


def outputBlock(f, gridData, curGlvIndex, glvFiles, isTopSide,mirrorX):
    """将GridData中的区块按优化后路径保存到文件中"""
    reBlockN = re.compile(r'N(\d+)G1X-?\d+Y-?\d+')
    for block in optimizeBlockOrder(grid, isTopSide,mirrorX):
        N = int(reBlockN.match(block).groups()[0])
        # 获取当前区块所在的GLV文件编号
        glvFileIndex = getGlvFileIndex(N, glvFiles)
        # 当处理的区块不在当前GLV文件编号中时，增加M90x指令切换GLV数据文件
        if curGlvIndex != glvFileIndex:
            curGlvIndex = glvFileIndex
            f.write('M9'+str(curGlvIndex).zfill(2)+LINE_BREAK)
        f.write(block+LINE_BREAK)
        f.write('M300'+LINE_BREAK)
    return curGlvIndex


def formatNum(val, deci=1):
    """对数字进行四舍五入，对整型数字返回整型值"""
    val = float(val)
    if round(val, 0) == round(val, deci):
        return int(round(val, 0))
    else:
        return round(val, deci)


def checkPrgSide(filePath):
    """
        从钻带文件名中判断钻带是否为正面钻带:
            正面钻带返回True
            反面钻带返回False
            无法判断面次时退出程序
    """
    side = re.search(r'lsr(\d\d)(\d\d)', filePath)
    if side:
        side = side.groups()
    else:
        showerror(title='错误', message='无法判断加工程序面次!')
        sys.exit(1)
    if side[0] < side[1]:
        return True
    else:
        return False


def checkSourcePrg(filePath):
    """检查原始镭射钻带是否满足要求，并返回钻带中的板厚及孔径"""

    if not os.path.isfile(name+ext):
        raise ValueError('钻带程序未找到: '+filePath)

    prg = []
    coreThick = None
    viaSize = None
    mirrorX = True
    with open(filePath) as f:
        for line in f:
            line = line.strip()
            if line == '%':
                break
            elif re.search(r'T(\d\d)C(\d+\.?\d*)', line):
                # 获取程式头中的刀具直径，当T2-T20中的刀具直径不相等时抛出异常
                t, c = re.search(r'T(\d\d)C(\d+\.?\d*)', line).groups()
                t = int(t)                        # 刀具编号
                c = formatNum(float(c)/0.0254)    # 刀具直径
                if 2 <= t <= 20:
                    if viaSize is None:
                        viaSize = c
                    elif viaSize != c:
                        raise ValueError('原始钻带中有设计多种不同孔径，请确认孔径设计是否有异常！')
            elif re.search(r'M47,Core_thick:(\d+\.?\d*)mil', line):
                # 获取程式头备注中的板厚大小
                h, = re.search(r'M47,Core_thick:(\d+\.?\d*)mil', line).groups()
                h = formatNum(h)
                if coreThick is None:
                    coreThick = h
            elif re.search(r'M47,\*\*Scale X:(\d\.?\d*)\*\*Y:(\d\.?\d*)', line):
                # 获取程式头备注中的涨缩系数信息
                x, y = re.search(
                    r'M47,\*\*Scale X:(\d\.?\d*)\*\*Y:(\d\.?\d*)', line).groups()
                x = formatNum(x, 6)
                y = formatNum(x, 6)
                if x == 1 or y == 1:
                    raise ValueError('原始钻带中的X或Y轴涨缩系数为0，请确认钻带是否有拉伸系数！')
            prg.append(line)

    # 检查反面钻带是否有添加 VER 指令
    isTopSide = checkPrgSide(filePath)
    
    if not isTopSide:
        if 'VER,4' in prg:
            mirrorX=True
        elif 'VER,7' in prg:
            mirrorX=False
        else:
            raise ValueError('反面钻带程式头没有添加 VER 指令进行镜像！')

    if coreThick is None:
        raise ValueError('无法获取原始钻带中的板厚数据！')
    if viaSize is None:
        raise ValueError('无法获取原始钻带中的孔径数据！')
    return (coreThick, viaSize, mirrorX)


def checkMitsuPrg(filePath, blockSize):
    """检查三菱镭射钻带是否满足要求"""

    if not os.path.isfile(name+ext):
        raise ValueError('钻带程序未找到: '+filePath)

    # 读取程序文件并按行保存到列表中
    prg = []
    with open(filePath) as f:
        for line in f:
            prg.append(line.strip())

    # 判断是否为三菱机加工钻带
    if prg[0] != '%':
        raise ValueError('加工钻带无法识别，请确认是否为三菱机加工钻带!')

    # 判断是否使用回形加工转换，即使没有使用回形加工方法转换也能正常优化路径，此判断逻辑非必需
    if '(BEST DIVISION:SP1_DIV)' not in prg:
        raise ValueError('请使用SP1_DIV回形加工方法转换钻带!')

    # 判断扫描区域大小设置是否正确
    if '(Area:X={0}.000,Y={0}.000)'.format(str(blockSize)) not in prg:
        raise ValueError('请使用{0}mm*{0}mm扫描区域大小转换钻带!'.format(str(BLOCK_SIZE)))

    # 判断是否打开90 ANGLE进行转换
    if '(90 ANGLE:OFF)' in prg:
        raise ValueError('请打开90 ANGLE设置进行钻带转换!')

    # 判断是否关闭X-Mirror进行转换
    if '(X MIRROR:ON)' in prg:
        raise ValueError('请关闭X Mirror设置进行钻带转换!')

    # 判断是否关闭Y-Mirror进行转换
    if '(Y MIRROR:ON)' in prg:
        raise ValueError('请关闭Y Mirror设置进行钻带转换!')

    # 判断镭射钻带是否已经优化过加工路径
    if '(Drilling Path Optimized)' in prg:
        raise ValueError('加工钻带已经优化过加工路径!')


if __name__ == '__main__':
    argvCount = len(sys.argv)
    if argvCount > 1:
        # 从第1个参数获取钻带文件名
        filePath = sys.argv[1]
    else:
        # 从文件选择对话框获取文件名
        filePath = askopenfilename()
        if not filePath:
            sys.exit(1)

    # 解析文件名及扩展名
    name, ext = os.path.splitext(filePath)
    # 忽略原文件扩展名，读取扩展名为.prg的程序文件
    ext = '.prg'

    # 检查原始钻带并获取原始钻带中的板厚、孔径信息
    try:
        coreThick, viaSize, mirrorX = checkSourcePrg(name)
    except ValueError as err:
        showerror(title='错误', message=err.args[0])
        sys.exit(1)
    except IOError as err:
        showerror(title='错误', message='无法读取原始钻带文件！')
        sys.exit(1)

    # 检查三菱机转换后钻带是否符合要求
    try:
        checkMitsuPrg(name+ext, BLOCK_SIZE)
    except ValueError as err:
        showerror(title='错误', message=err.args[0])
        sys.exit(1)
    except IOError as err:
        showerror(title='错误', message='无法读取待转换的.prg钻带文件！')
        sys.exit(1)

    # 读取钻带内容并按行保存至prg的列表中，删除每行头尾的空字符和换行符
    prg = []
    with open(name+ext) as f:
        for line in f:
            prg.append(line.strip())

    # 在钻带程式头添加参数
    cond = r"ldd8um-{0}mil-core-{1}mil".format(
        coreThick, viaSize).replace(".", "'")
    isTopSide = checkPrgSide(name)
    if isTopSide:
        # 正面钻带添加参数指令
        cond = r"M100(1st-{0})".format(cond)
    else:
        # 反面钻带添加参数指令
        cond = r"M100(2nd-{0})".format(cond)
    prg.insert(1, cond)

    # 钻带末尾添加已执行路径优化的备注
    prg.insert(-1, '(Drilling Path Optimized)')

    curTool = 0                                     # 当前区块的刀具编号
    curBlock = 0                                    # 当前的区块编号
    curGlvIndex = 0                                 # 当前区块所在的GLV文件编号
    if 'M900' in prg:
        curGlvIndex = -1                            # 当钻带中有GLV文件切换指令时才增加切换指令，单个GLV文件不添加M90x指令
    
    glvFiles = [1]                                  # 保存每个GLV文件中起始区块编号的列表，钻带读取过程中根据M90x指令自动识别更新此列表
    grid = GridData(BLOCK_SIZE, BLOCK_SIZE)         # 按照指定的大小划分每个回形加工路径之间的间距
    grid.posParser = parseBlockXY                   # 指定区块坐标的解析器
    regBlock = re.compile(r'N(\d+)G1X-?\d+Y-?\d+')  # 识别区块指令的正则表达式
    regTool = re.compile(r'M1(0[1-9]|[1-4]\d|50)')  # 识别刀具切换指令的正则表达式
    regGlvIndex = re.compile(r'M9(0\d)')            # 识别GLV数据文件切换指令M90x的正则表达式，通常该指令后第一个区块即为该GLV文件的起始区块编号
    flagIndex = -1                                  # 当识别到GLV数据文件切换指令时，此标识符会设为指令中对应的GLV文件编号，用于后续自动更新glvFiles列表

    # 重写镭射机加工程序至临时文件中
    with open(name+'.tmp', 'w') as f:
        for line in prg:
            if regTool.match(line):
                # 识别刀具切换指令
                toolNum = int(regTool.match(line).groups()[0])
                # 将T03-T20合并为T02
                if 2 < toolNum < 21:
                    toolNum = 2
                if toolNum != curTool:
                    # 如果当前刀具编号与之前编号不一致时，将前一把刀中保存的区块按优化后路径写入的临时文件中
                    if grid.countItems() > 0:
                        curGlvIndex = outputBlock(
                            f, grid, curGlvIndex, glvFiles, isTopSide,mirrorX)
                    # 将新刀具的切换指令写入到文件中，并更新当前刀具编号
                    curTool = toolNum
                    f.write('M1'+str(curTool).zfill(2)+LINE_BREAK)
            elif regBlock.match(line):
                # 识别区块指令
                curBlock = int(regBlock.match(line).groups()[0])
                grid.addItem(line)
                # 如果前一个指令为Glv切换指令，则更新glvFiles列表，其中列表索引为GLV文件的编号，值为该GLV文件的起始区块编号
                if flagIndex > -1:
                    # 当前区块编号小于列表中保存的起始区块编号时才更新列表
                    if curBlock < glvFiles[flagIndex]:
                        glvFiles[flagIndex] = curBlock
                    flagIndex = -1
            elif line == 'M300':
                # 识别区块加工执行指令，输出优化后区块路径时可以自动添加，故删除原M300指令
                pass
            elif regGlvIndex.match(line):
                # 识别Glv数据文件切换指令
                glvIndex = int(regGlvIndex.match(line).groups()[0])
                # 当GLV文件编号大于glvFiles列表中的索引上限时增加一个临时区块编号至列表中
                while glvIndex >= len(glvFiles):
                    glvFiles.append(999999)
                flagIndex = glvIndex
            else:
                # 识别到未知程序指令时先将之前保存的区块按优化后路径写入临时文件中
                if grid.countItems() > 0:
                    curGlvIndex = outputBlock(
                        f, grid, curGlvIndex, glvFiles, isTopSide,mirrorX)
                f.write(line+LINE_BREAK)

    # 将原始钻带文件备份为.bak文件, 用生成的临时钻带替换原始钻带
    os.rename(name+ext, name+'.bak')
    os.rename(name+'.tmp', name+ext)
    showinfo(title='优化完成', message='以下程序已经完成路径优化并添加参数:\n'+name+ext)
