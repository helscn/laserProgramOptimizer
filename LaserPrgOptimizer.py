#!/usr/bin/python
# -*-coding:utf-8-*-

# Compatible with python2 and python3
from __future__ import print_function
import sys
import os
import re

if sys.version_info[0] == 2:
    input = raw_input
    import sys
    reload(sys)
    sys.setdefaultencoding('utf8')


# DocString type: epydoc

class GridData(object):
    """根据数据的位置坐标保存二维矩阵数据，所有数据会按照X,Y坐标保存在等间距分隔的Cell单元格中"""

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
        根据数据的坐标位置添加至对应的Grid表子单元中
        @param data: 需要增加的数据，其中数据保存时的x,y坐标由创建对象时的xParser,yParser解析器对data解析
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
        删除Grid边缘区域中不包含任何元素的空行、空列，以减小数据占用空间，注意优化后数据的索引号可能发生变更
        注意Grid中没有任何元素时至少会保留一个空的Cell
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

    def showGrid(self):
        print(self.__gridData)


def parseNumber(s, length=3, lead_zero=False):
    """
    根据Excellon中的文本坐标数据识别转换为实际坐标值
    @param length: 数据位数长度，当lead_zero为True时代表整数位数，当lead_zero为False(即trail_zero)时代表小数位数
    @type length: int
    @param lead_zero: 设置数据是否为前导零格式，值为Fasle时表示数据为后补零格式，此时length指定了数据的小数位数
    @type lead_zero: bool
    @return: 根据文本数据识别出的实际数值
    @rtype: float
    """
    s = re.sub(r'(^\s+|\s+$)', '', s)
    if '.' in s:
        return float(s)
    if s[0] == '-':
        sign = -1
        s = s[1:]
    else:
        sign = 1
    if lead_zero:
        if len(s) < length:
            v = sign * float(s) * pow(10, length-len(s))
        else:
            v = sign * float(s[0:length]+'.'+s[length:])
    else:
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
    regBlock = re.compile(r'N\d+G1X(-?\d+)Y(-?\d+)')
    result = regBlock.match(s)
    if result:
        BlockX, BlockY = result.groups()
        return (-1*parseNumber(BlockY), parseNumber(BlockX))


def optimizeBlockOrder(grid, clockwise=True):
    """
    将GridData中保存的三菱机区块指令按照从外向内加工路径依次输出的生成器
    @param grid: 保存三菱机区块数据的GridData对象
    @param clockwise: 输出区块时是否按顺时针顺序进行输出，值为True时从左下角按顺时针输出，值为False时从右下角按逆时针输出
    @return :该函数为生成器函数，每次yield一个区块，直至取出所有区块后清空grid中的数据
    @rtype: None
    """
    while grid.countItems() > 0:
        minX = 0
        minY = 0
        maxX = grid.width-1
        maxY = grid.height-1
        if maxX == 0 or maxY == 0:
            # 只有单行和单列时的特殊情况
            if clockwise:
                # 正面钻带加工顺序
                for y in range(maxY+1):
                    for x in range(maxX+1):
                        if grid.countItems(x, y) > 0:
                            yield grid.popItem(x, y, 0)
            else:
                # 反面钻带加工顺序
                for y in range(maxY+1):
                    for x in range(maxX, -1, -1):
                        if grid.countItems(x, y) > 0:
                            yield grid.popItem(x, y, 0)
        else:
            if clockwise:
                # 正面钻带加工顺序
                for y in range(minY, maxY):
                    if grid.countItems(minX, y) > 0:
                        yield grid.popItem(minX, y, 0)
                for x in range(minX, maxX):
                    if grid.countItems(x, maxY) > 0:
                        yield grid.popItem(x, maxY, 0)
                for y in range(maxY, minY, -1):
                    if grid.countItems(maxX, y) > 0:
                        yield grid.popItem(maxX, y, 0)
                for x in range(maxX, minX, -1):
                    if grid.countItems(x, minY) > 0:
                        yield grid.popItem(x, minY, 0)
            else:
                # 反面钻带加工顺序
                for y in range(minY, maxY):
                    if grid.countItems(maxX, y) > 0:
                        yield grid.popItem(maxX, y, 0)
                for x in range(maxX, minX, -1):
                    if grid.countItems(x, maxY) > 0:
                        yield grid.popItem(x, maxY, 0)
                for y in range(maxY, minY, -1):
                    if grid.countItems(minX, y) > 0:
                        yield grid.popItem(minX, y, 0)
                for x in range(minX, maxX):
                    if grid.countItems(x, minY) > 0:
                        yield grid.popItem(x, minY, 0)
        # 清除外围取出数据后的空Cell
        if grid.countItems(column=0) == 0 and grid.countItems(column=maxX) == 0 and grid.countItems(row=0) == 0 and grid.countItems(row=maxY) == 0:
            grid.optimizeGrid()
    grid.delAllItems()


def getGlvFileIndex(N, glvFiles):
    """
    根据当前区块编号返回区块所在的.glv文件索引
    @param N: 需要查询所在GLV文件中的区块编号
    @type N: int
    @param glvFiles: 保存每个GLV文件启始区块编号的列表，列表的索引编号代表GLV的文件编号，值代表该GLV文件中第一个区块编号
    @type glvFiles: list
    @return: 返回区块N所在的GLV文件编号，值为0表示在第一个glv文件，值为1表示在第二个glv文件，依此类推
    @rtype: int
    """
    for i in range(len(glvFiles)-1, -1, -1):
        if N >= glvFiles[i]:
            return i


def outputBlock(f, gridData, curGlvIndex, glvFiles, isTopSide):
    """将GridData中的区块按优化后路径保存到文件中"""
    reBlockN = re.compile(r'N(\d+)G1X-?\d+Y-?\d+')
    for block in optimizeBlockOrder(grid, isTopSide):
        N = int(reBlockN.match(block).groups()[0])
        glvFileIndex = getGlvFileIndex(N, glvFiles)
        if curGlvIndex != glvFileIndex:
            curGlvIndex = glvFileIndex
            f.write('M9'+str(curGlvIndex).zfill(2)+'\n')
        f.write(block+'\n')
        f.write('M300\n')
    return curGlvIndex


def encode(s, encoding='gbk'):
    """对于Python2.x需在Windows环境下将文本编码转换为GBK"""
    if sys.version_info[0] == 2:
        return s.encode('gbk')
    else:
        return s


def checkPrg(isTopSide, prg):
    """检查镭射钻带是否满足要求"""
    # 判断是否为三菱机加工钻带
    if prg[0] != '%':
        print(encode('加工钻带无法识别，请确认是否为三菱机加工钻带。'))
        sys.exit(1)

    # 判断是否使用回形加工转换
    if '(BEST DIVISION:SP1_DIV)' not in prg:
        print(encode('请使用SP1_DIV回形加工方法转换钻带!'))
        sys.exit(1)

    # 判断扫描区域大小设置是否为30mm*30mm
    if '(Area:X=30.000,Y=30.000)' not in prg:
        print(encode('请使用30mm*30mm扫描区域大小转换钻带!'))
        sys.exit(1)

    # 判断正面是否关闭X-Mirror进行转换
    if isTopSide and '(X MIRROR:ON)' in prg:
        print(encode('正面钻带需关闭X Mirror设置进行转换!'))
        sys.exit(1)

    # 判断反面是否有使用X-Mirror转换
    if (not isTopSide) and '(X MIRROR:OFF)' in prg:
        print(encode('反面钻带需使用X Mirror设置进行转换!'))
        sys.exit(1)

    # 判断镭射钻带是否已经优化过加工路径
    if '(Drilling Path Optimized)' in prg:
        print(encode('加工钻带已经优化过加工路径!'))
        sys.exit(1)


if __name__ == '__main__':
    # 获取钻带文件名
    if len(sys.argv) > 1:
        filePath = sys.argv[1]
    else:
        filePath = input(encode('请输入钻带程序名：'))
    name, ext = os.path.splitext(filePath)
    ext = '.prg'
    if not os.path.isfile(name+ext):
        print(encode('钻带程序未找到: '+name+ext))
        sys.exit(1)

    # 从钻带文件名中判断钻带的面次
    isTopSide = re.search(r'lsr(\d\d)(\d\d)', name)
    if isTopSide:
        isTopSide = isTopSide.groups()
    else:
        print(encode('无法识别钻带面次!'))
        sys.exit(1)
    if isTopSide[0] < isTopSide[1]:
        isTopSide = True
    else:
        isTopSide = False

    # 读取钻带内容并按行保存至prg的列表中，删除每行头尾的空字符和换行符
    prg = []
    with open(name+ext) as f:
        for line in f:
            prg.append(line.strip())

    # 检查钻带是否符合钻带转换要求
    checkPrg(isTopSide, prg)

    # 在钻带程式头添加参数
    flag = True
    while flag:
        cond = input(encode('请输入生产板板厚(默认为2mil):'))
        cond = cond.lower().replace('mil', '').replace('2.0', '2').strip()
        if not cond:
            cond = '2'
        if cond == '2' or cond == '2.3':
            flag = False
        else:
            print(encode('板厚大小不正确！\n'))
    if isTopSide:
        cond = r"M100(1st-ldd8um-{0}mil-core-2'4mil)".format(cond)
    else:
        cond = r"M100(2nd-ldd8um-{0}mil-core-2'4mil)".format(cond)
    print(encode('程式头添加参数：'+cond))
    prg.insert(1, cond)

    # 钻带末尾添加执行路径优化的备注
    print(encode('正在优化钻带加工路径:'), name+ext)
    prg.insert(-1, '(Drilling Path Optimized)')

    # 重写镭射机加工程序至临时文件中
    curTool = 0                 # 当前区块的刀具编号
    curBlock = 0                # 当前的区块编号
    curGlvIndex = 0             # 当前区块所在的GLV文件编号
    glvFiles = [1]              # 保存每个GLV文件中起始区块编号的列表，处理过程中根据M90x指令自动识别更新
    grid = GridData(30, 30)     # 按照30mm*30mm的间隔划分每个回形加工路径间隔
    grid.posParser = parseBlockXY
    regBlock = re.compile(r'N(\d+)G1X-?\d+Y-?\d+')
    regTool = re.compile(r'M1(0[1-9]|[1-4]\d|50)')
    regGlvIndex = re.compile(r'M9(\d\d)')
    flagIndex = -1

    with open(name+'.tmp', 'w') as f:
        for line in prg:
            if regTool.match(line):
                # 识别刀具切换指令
                toolNum = int(regTool.match(line).groups()[0])
                # 将T03-T23合并为T02
                if 2 < toolNum < 24:
                    toolNum = 2
                if toolNum != curTool:
                    if grid.countItems() > 0:
                        curGlvIndex = outputBlock(
                            f, grid, curGlvIndex, glvFiles, isTopSide)
                    curTool = toolNum
                    f.write('M1'+str(curTool).zfill(2)+'\n')
            elif regBlock.match(line):
                # 识别区块指令
                curBlock = int(regBlock.match(line).groups()[0])
                grid.addItem(line)
                # 如果前一个指令为Glv切换指令，则更新glv区块文件域值列表
                if flagIndex > -1:
                    if curBlock < glvFiles[flagIndex]:
                        glvFiles[flagIndex] = curBlock
                    flagIndex = -1
            elif line == 'M300':
                # 识别区块加工执行指令，输出优化后区块路径时可以自动添加，故删除原M300指令
                pass
            elif regGlvIndex.match(line):
                # 识别Glv数据文件切换指令
                glvIndex = int(regGlvIndex.match(line).groups()[0])
                while glvIndex >= len(glvFiles):
                    glvFiles.append(999999)
                flagIndex = glvIndex
            else:
                if grid.countItems() > 0:
                    curGlvIndex = outputBlock(
                        f, grid, curGlvIndex, glvFiles, isTopSide)
                f.write(line+'\n')

    # 将原始钻带文件备份为.bak文件, 用生成的临时钻带替换原始钻带
    os.rename(name+ext, name+'.bak')
    os.rename(name+'.tmp', name+ext)
    print(encode('\n==== 钻带优化已经完成! ====\n'))
    