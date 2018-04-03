# encoding: utf-8
import os
import time
import traceback

"""
   作者：吴杰春
微信/QQ：179819930
"""
try:
    import xlrd
    import xlwt
except:
    print('提示: 你还没有装xlrd xlwt库文件!正在为你安装...')
    print('\t如若无法自动安装, 请手动在cmd输入命令后回车安装')
    print('命令: ')
    print('\tpip install xlrd')
    print('\tpip install xlwt')
    os.system('pip install xlrd')
    os.system('pip install xlwt')
    import xlrd
    import xlwt


# 分钟转时间
def m2h(minutes):
    """
    :param minutes: int - 分钟 - 61
    :return:        str - 分钟对应的时间 - '01:01'
    """
    return '{0:02d}:{1:02d}'.format(minutes//60, minutes%60)
# 时间转分钟
def h2m(time):
    """
    :param time: str - 时间字符串 - '01:01'
    :return:     int - 分钟 - 61
    """
    try:
        return int(time.split(':')[0]) * 60 + int(time.split(':')[1])
    except:
        return int(time.split('-')[0]) * 60 + int(time.split(':')[1])

# 根据学生名字在数组中找到它的下标
def fid(members, name):
    """
    在members中找名为name的学生, 返回下标

    :param members: list - 学生数组
    :param name:    str  - 要找学生的姓名
    :return:        int  - 下标
                    None - 代表找不到'
    """
    # 遍历学生数组
    for n in range(len(members)):
        # 从首/末两端比较姓名
        if members[n].name == name:
            return n
        elif members[-n-1].name == name:
            return len(members)-n-1
    # 找不到
    return None

# 异常退出
def error(meg):
    print(meg)
    input('回车后, 结束程序!')
    exit(1)

# 读取xls文件
def readXlsFile(fileName):
    try:
        xls = xlrd.open_workbook(fileName)
        sheet = xls.sheet_by_index(0)
        xls_list = []
        for r in range(sheet.nrows):
            xls_list.append(sheet.row_values(r))
        return xls_list
    except:
        error('警告: {} 读取失败!'.format(fileName))


# 根据list, 输出名为 xlsName 的xls文件
def outputXLS(xlsName, List):
    wbook = xlwt.Workbook()
    wsheet = wbook.add_sheet(xlsName.replace('.xls', ''))
    # 自动换行 垂直居中 水平居中
    style = xlwt.easyxf('align: wrap on,vert centre, horiz center')
    try:
        for rIndex, rValue in enumerate(List):
            for cIndex, content in enumerate(rValue):
                wsheet.write(rIndex, cIndex, content, style)
        # 针对无课表做特殊列宽设置
        if '无课表' in List[0][0]:
            # 设置列的宽度
            wsheet.col(0).width = 2800
            for colId in range(1, len(List[1])):
                wsheet.col(colId).width= 6000

        wbook.save(xlsName)
    except PermissionError:
        print('警告: {} 可能正在被使用中!'.format(xlsName))
        input('      请关闭后回车')
        try:
            for rIndex, rValue in enumerate(List):
                for cIndex, content in enumerate(rValue):
                    wsheet.write(rIndex, cIndex, content, style)
            # 针对无课表做特殊列宽设置
            if '无课表' in List[0][0]:
                # 设置列的宽度
                wsheet.col(0).width = 3000
                for colId in range(1, len(List[1])):
                    wsheet.col(colId).width= 5800
        except:
            print('警告: 无法保存 {}!'.format(xlsName))

# 定义社团类
class Club:
    name = None
    members = None

    # 根据学生名字在数组中找到它的下标
    def fid(self, name):
        """
        :param name: str  - 要找学生的姓名
        :return:     int  - 下标
                     None - 代表找不到'
        """
        # 遍历学生数组
        for n in range(len(self.members)):
            # 从首/末两端比较姓名
            if self.members[n].name == name:
                return n
            elif self.members[-(n+1)].name == name:
                return len(self.members) - n - 1
        # 找不到
        return None

    # 添加成员
    def addMember(self):
        print('\n-----------')
        print('提示: 输入 break 时, 退出添加\n')
        while True:
            tmpMem = Member()
            tmpMem.name = input('姓名: ')
            if tmpMem.name == 'break':
                break
            tmpMem.sex = input('性别: ')
            if tmpMem.sex == 'break':
                break
            tmpMem.phoNum = input('手机: ')
            if tmpMem.phoNum == 'break':
                break
            tmpMem.adr = input('籍贯: ')
            if tmpMem.adr == 'break':
                break
            tmpMem.dept = input('部门: ')
            if tmpMem.dept == 'break':
                break
            tmpMem.job = input('职务: ')
            if tmpMem.job == 'break':
                break
            tmpMem.id = input('学号: ')
            if tmpMem.id == 'break':
                break
            tmpMem.squad = input('班级: ')
            if tmpMem.squad == 'break':
                break
            tmpMem.grade = input('年级: ')
            if tmpMem.grade == 'break':
                break
            tmpMem.buld = input('楼栋: ')
            if tmpMem.buld == 'break':
                break
            tmpMem.dorPla = input('门牌: ')
            if tmpMem.dorPla == 'break':
                break
            tmpMem.buld = tmpMem.buld.replace('栋', '')
            # 在刚多出的位置, 添加 tmpMem 到 members
            self.members.append(tmpMem)
            print('提示: 成功添加 ' + tmpMem.name)
            print()
        input('\n提示: 请回车以继续!')
        print('-------------\n')

    # 移除成员
    def delMember(self):
        print('\n-----------')
        print('提示: 输入 break 时, 退出移除\n')
        while True:
            rmStuName = input('姓名: ')
            if rmStuName == 'break':
                break
            index = self.fid(rmStuName)
            if index != None:
                del self.members[index]
                print('提示: 成功移除 ' + rmStuName)
            else:
                print('警告: 查无"' + rmStuName + '"')
            print()
        input('提示: 请回车以继续!')
        print('-------------\n')

    # 浏览成员
    def browMember(self):
        print('\n-----------')
        print('提示: 输入 break 时, 退出浏览; 输入all，浏览全部\n')
        while True:
            stuName = input('姓名: ')
            if stuName == 'break':
                break
            if stuName == 'all':
                for m in self.members:
                    m.showInfo()
                    print()
                continue
            index = self.fid(stuName)
            if index != None:
                self.members[index].showInfo()
            else:
                print('提示: 查无 "' + stuName + '"')
            print()
        input('\n提示: 请回车以继续!')
        print('-------------\n')

    # 输出信息
    def outputMain(self):
        # 输出通讯录
        def outputAdrBok():
            l_xlsAdreBok = []  # 新通讯录列表
            l_xlsAdreBok.append(['部门', '职务', '姓名', '性别', '学号', '长号', '短号', '籍贯', '宿舍'])

            for deptKey in ['理事会', '组织部', '秘书部', '外联部', '宣传部', '培训部']:
                n = 1
                for m in self.members:
                    if m.dept == deptKey:
                        if n == 1:
                            if deptKey != '理事会':
                                l_xlsAdreBok.append([])
                            l_xlsAdreBok.append([m.dept, m.job, m.name, m.sex, m.id, m.longPhoNum, m.shortPhoNum, m.adr, m.buld + m.dorPla])
                            n = n + 1
                        else:
                            l_xlsAdreBok.append(['', m.job, m.name, m.sex, m.id, m.longPhoNum, m.shortPhoNum, m.adr, m.buld + m.dorPla])
                            n = n + 1
            try:
                os.remove(data.xls_AddressBbook.replace('xlsx', '备份').replace('xls', '备份'))
            except:
                pass

            try:
                os.rename(data.xls_AddressBbook, data.xls_AddressBbook.replace('xlsx', '备份').replace('xls', '备份'))
            except PermissionError:
                print('警告: {} 正在被使用中'.format(data.xls_AddressBbook))
                print('      请关闭后重新操作')
                input('      请回车已继续!')
                return ''
            outputXLS(data.xls_AddressBbook.replace('xlsx', 'xls'), l_xlsAdreBok)
            print('提示: {} 已被更新!'.format(data.xls_AddressBbook))
            input('\n提示: 请回车以继续!')
            print('-------------\n')

        # 制作xls课表需要的list
        def makeXlsSch(xlsSchName, nameList):
            # 将二维列表转置
            def reSetXY(List):
                newList = [['', '', '', '', '', '', ''] for _ in range(5)]
                # 7天
                for x in range(7):
                    # 每天5节
                    for y in range(5):
                        newList[y][x] = List[x][y]
                return newList

            # xls课表内容
            l_xlsSch = [
                [xlsSchName.replace('.xls', '') + '_单周'],  # 标题
                [[''], ['星期一'], ['星期二'], ['星期三'], ['星期四'], ['星期五'], ['星期六'], ['星期日']],
                [['第一大节\n8:10-9:30'], [''], [''], [''], [''], [''], [''], [''], ['']],
                [['第二大节\n9:50-12:00'], [''], [''], [''], [''], [''], [''], [''], ['']],
                [['第三大节\n14:00-16:10'], [''], [''], [''], [''], [''], [''], [''], ['']],
                [['第四大节\n16:40-18:00'], [''], [''], [''], [''], [''], [''], [''], ['']],
                [['第五大节\n19:00-20:20'], [''], [''], [''], [''], [''], [''], [''], ['']],
                [''],
                [xlsSchName.replace('.xls', '') + '_双周'],  # 标题
                [[''], ['星期一'], ['星期二'], ['星期三'], ['星期四'], ['星期五'], ['星期六'], ['星期日']],
                [['第一大节\n8:10-9:30'], [''], [''], [''], [''], [''], [''], [''], ['']],
                [['第二大节\n9:50-12:00'], [''], [''], [''], [''], [''], [''], [''], ['']],
                [['第三大节\n14:00-16:10'], [''], [''], [''], [''], [''], [''], [''], ['']],
                [['第四大节\n16:40-18:00'], [''], [''], [''], [''], [''], [''], [''], ['']],
                [['第五大节\n19:00-20:20'], [''], [''], [''], [''], [''], [''], [''], ['']]
            ]
            schList = []
            try:
                for name in nameList:
                    schList.append([self.members[self.fid(name)].name, reSetXY(self.members[self.fid(name)].schedule[0]),
                            reSetXY(self.members[self.fid(name)].schedule[1])])
                #schList = [[self.members[self.fid(name)].name, reSetXY(self.members[self.fid(name)].schedule[0]),
                #            reSetXY(self.members[self.fid(name)].schedule[1])] for name in nameList]
            except:
                error('\n警告: {0} 的课表制作失败!\n可能原因: \n\t1.没有{0}的课表! \n\t2.{0}的课表格式不正确!\n\t3.{0}使用了别人的课表, 缺没改名字!\n解决方法: 让{0}重新发送从教务系统下载的课表!'.format(name))

            for mSchList in schList:
                mName = mSchList[0]
                mSch_s = mSchList[1]
                mSch_d = mSchList[2]
                # 单周部分
                for index_p, part in enumerate(mSch_s):
                    # print('第',index_p,'部分')
                    for index_wd, weekday in enumerate(part):
                        # print('星期',index_wd,weekday)
                        if weekday == '0':  # 无课
                            if l_xlsSch[index_p + 2][index_wd + 1][0] == '':
                                l_xlsSch[index_p + 2][index_wd + 1][0] = mName
                            else:
                                l_xlsSch[index_p + 2][index_wd + 1][0] = l_xlsSch[index_p + 2][index_wd + 1][
                                                                             0] + '、' + mName
                # 双周部分
                for index_p, part in enumerate(mSch_d):
                    # print('第',index_p,'部分')
                    for index_wd, weekday in enumerate(part):
                        # print('星期',index_wd,weekday)
                        if weekday == '0':  # 无课
                            if l_xlsSch[index_p + 10][index_wd + 1][0] == '':
                                l_xlsSch[index_p + 10][index_wd + 1][0] = mName
                            else:
                                l_xlsSch[index_p + 10][index_wd + 1][0] = l_xlsSch[index_p + 10][index_wd + 1][
                                                                                 0] + '、' + mName

            # 输出课表
            outputXLS(xlsSchName.replace('xlsx', 'xls'), l_xlsSch)
            print('提示: 成功制作 {} !'.format(xlsSchName.replace('xlsx', 'xls')))

        # 输出课表
        def outputSch():
            print('\n-----------')
            print('模式: 1-个人  2-部门  3-大X  4-全体')
            choice_1 = input('选择: ')
            if choice_1 == '1':  # 个人模式
                nameList = []
                print('\n-----------')
                print('模式: 1-循环名字  2-全部个人')
                choice_2 = input('选择: ')
                if choice_2 == '1':  # 循环名字
                    print('提示: 先输入多个名字, 后输入 break 进行查询')
                    # 循环获取要查询的人的名字
                    while True:
                        tmpName = input('姓名: ')
                        if tmpName == 'break':
                            break
                        if self.fid(tmpName) != None:
                            nameList.append(tmpName)
                            print('提示: {} 可查询'.format(tmpName))
                        else:
                            print('提示: 查无"{}"'.format(tmpName))
                        print()

                    print()
                    for name in nameList:
                        makeXlsSch(name + '_个人无课表.xls', [name])
                else:  # 全部个人
                    for m in self.members:
                        makeXlsSch(m.name + '_个人无课表.xls', [m.name])

            elif choice_1 == '2':  # 部门模式
                deptList = []
                print('\n-----------')
                print('模式: 1-循环部门  2-全部部门')
                choice_2 = input('选择: ')
                if choice_2 == '1':
                    print('提示: 先输入多个部门, 后输入 break 进行查询')
                    while True:
                        tmpDept = input('部门: ')
                        if tmpDept == 'break':
                            break
                        if tmpDept in ['理事会', '组织部', '秘书部', '外联部', '宣传部', '培训部']:
                            deptList.append(tmpDept)
                            print('提示: {} 可查询'.format(tmpDept))
                        else:
                            print('提示: 查无"{}"'.format(tmpDept))
                        print()
                else:
                    deptList = ['理事会', '组织部', '秘书部', '外联部', '宣传部', '培训部']
                print()
                for tmpDept in deptList:
                    makeXlsSch(tmpDept + '_部门无课表.xls', [m.name for m in self.members if m.dept == tmpDept])

            elif choice_1 == '3':  # 大X模式
                gradetList = []
                print('\n-----------')
                print('模式: 1-循环大X  2-全部大X')
                choice_2 = input('选择: ')
                if choice_2 == '1':
                    print('提示: 先输入多个年级, 后输入 break 行查询')
                    while True:
                        tmpGrade = input('年级: ')
                        if tmpGrade == 'break':
                            break
                        if tmpGrade in ['大一', '大二', '大三']:
                            gradetList.append(tmpGrade)
                            print('提示: {} 可查询'.format(tmpGrade))
                        else:
                            print('提示: 查无"{}"'.format(tmpGrade))
                        print()
                else:
                    gradetList = ['大一', '大二', '大三']
                print()
                for tmpGrade in gradetList:
                    makeXlsSch(tmpGrade + '_大X无课表.xls', [m.name for m in self.members if m.grade == tmpGrade])

            elif choice_1 == '4':  # 红会全体
                print()
                makeXlsSch('红会全体无课表.xls', [m.name for m in self.members])

            input('\n提示: 请回车以继续!')
            print('-------------\n')

        print('\n-----------')
        print('提示: 输入 break 时, 退出输出\n')
        print('输出: 1-通讯录  2-xls课表')
        outputChoice = input('选择: ')
        if outputChoice == 'break':
            input('\n提示: 请回车以继续!')
            print('-------------\n')
            return ''
        if outputChoice == '1':
            # 输出通讯录
            outputAdrBok()
        else:
            # 输出课表
            outputSch()

    # 安排活动
    def arrnageAct(self):
        li = []
        for m in self.members:
            li.append(m.toJson())

        with open('data.js', 'w', encoding='utf-8') as data:
            data.write('var data = '+str(li))
        cm = '''start activity.html'''.format()
        os.system(cm)
        input('回车以继续\n')

# 定义member类
class Member:
    # 实例化类对象时, 初始化对象的基本属性
    def __init__(self):
        self.name = 'null'  # 姓名
        self.grade = 'null'  # 年级
        self.job = 'null'  # 职务
        self.id = 'null'  # 学号
        self.sex = 'null'  # 性别
        self.dept = 'null'  # 部门
        self.longPhoNum = 'null'  # 长号
        self.shortPhoNum = 'null'  # 短号
        self.buld = 'null'  # 楼栋
        self.dorPla = 'null'  # 门牌号
        self.adr = 'null'  # 籍贯
        self.schedule = ''

    # 输出个人的信息
    def showInfo(self):
        print('姓名:', self.name)
        print('性别: {0:8s} \t籍贯: {1:9s}'.format(self.sex, self.adr))
        print('年级: {0:9s} \t学号: {1:9s} '.format(self.grade, self.id))
        print('短号: {0:9s} \t长号: {1:9s}'.format(self.shortPhoNum, self.longPhoNum))
        print('部门: {0:6s} \t职务: {1:9s}'.format(self.dept, self.job))
        print('楼栋: {0:8s} \t门牌: {1:9s}'.format(self.buld, self.dorPla))
    def toJson(self):
        di = {}
        di['name'] = self.name  # 姓名
        di['grade'] = self.grade # 年级
        di['job'] = self.job # 职务
        di['id'] = self.id  # 学号
        di['sex'] = self.sex  # 性别
        di['dept'] = self.dept # 部门
        di['longPhoNum'] = self.longPhoNum # 长号
        di['shortPhoNum'] = self.shortPhoNum  # 短号
        di['buld'] = self.buld  # 楼栋
        di['dorPla'] = self.dorPla  # 门牌号
        di['adr'] = self.adr # 籍贯
        # 默认空闲区间: 早上6点到晚上23点
        classTimeMinute = {0:[7*60+10, 7*60+50], 1:[490, 570], 2:[590, 720], 3:[840, 970], 4:[1000, 1080], 5:[1140, 1220]}
        li = [
            [[360, 1380], [360, 1380], [360, 1380], [360, 1380], [360, 1380], [360, 1380], [360, 1380]] # 单周,
            ,[[360, 1380], [360, 1380], [360, 1380], [360, 1380], [360, 1380], [360, 1380], [360, 1380]] # 双周
        ]
        # 周末没有早读
        for weektype in range(2):
            for workday in li[weektype][0:5:]:
                workday += classTimeMinute[0][::]

        for weekType, timeList in enumerate(self.schedule):
            for dayId, day in enumerate(timeList):
                for partId, partStatue in enumerate(day):
                    classTimeId = partId+1
                    # 无课
                    if partStatue == '1':
                        li[weekType][dayId] += classTimeMinute[classTimeId]
                li[weekType][dayId].sort()
                tmp = li[weekType][dayId][::]
                #print(tmp)
                li[weekType][dayId] = [[tmp[i], tmp[i+1]] for i in range(0, len(tmp), 2)]
                
        di['free'] = li
        return di
# 定义数据类 数据加载等
class loadData:
    # 初始化
    def __init__(self):
        # xls_AddressBbook  # 通讯录.xls
        # txt_jwxt          # 课表数字.txt

        # 数据文件名
        self.xls_AddressBbook = ''
        self.txt_InfoOfAct = ''
        self.txt_jwxt = ''
        # 文件读取状态
        self.status_loadAdrBok = False
        self.status_loadSch = False

        # 获取数据文件名
        self.getDataFileNname()

    # 在获取数据文件名_当前目录下存在的
    def getDataFileNname(self):
        # 获取列表中的第一个数据
        def getFistContent(li, fileName):
            if len(li) == 1:
                return li[0]
            else:
                print(li)
                error('警告: {}数据文件多于一个'.format(fileName))
                return ''

        self.xls_AddressBbook = getFistContent([f for f in os.listdir(os.getcwd()) if '通讯' in f and 'xls' in f], '通讯')
        self.txt_jwxt = getFistContent([f for f in os.listdir(os.getcwd()) if '教务' in f], '教务')

    # 加载通讯录
    def loadAddressBbook(self):
        members = []
        try:
            # 读取通讯录
            l_AdrBbk = readXlsFile(self.xls_AddressBbook)
            # 遍历表格长度,找出标准长度(最长行)
            stdLen = 0  # 标准长度
            for rowData in l_AdrBbk:
                if len(rowData) > stdLen:
                    stdLen = len(rowData)

            # 清除空行
            while ['', '', '', '', '', '', '', '', ''] in l_AdrBbk:
                l_AdrBbk.remove(['', '', '', '', '', '', '', '', ''])
            # 遍历表格, 把数据转为str类型
            for row in range(len(l_AdrBbk)):  # row: 表格中的行下标(第几行)
                # 当长度不够, 以''补够长度
                for _ in range(stdLen - len(l_AdrBbk[row])):
                    l_AdrBbk[row] = l_AdrBbk[row] + ['']

                # 把每个元素都转为str类型
                for ele in range(len(l_AdrBbk[row])):  # ele: 行中的元素下标(第几个元素)
                    l_AdrBbk[row][ele] = str(l_AdrBbk[row][ele]).replace(' ', '')

            # 删除标题行  2017年红会通讯录
            for ele in l_AdrBbk[0]:
                if '通讯录' in ele:
                    del l_AdrBbk[0]
                    break

            # 从数据起始行中, 找出成员属性 在表格 对应的列下标
            c_name = c_job = c_id = c_sex = c_longPhoNum = c_dormi = -1
            d_info = {'职务':-1, '姓名':-1, '性别':-1, '学号':-1, '长号':-1, '宿舍':-1}
            for index, ele in enumerate(l_AdrBbk[0]):
                if '职务' in ele:
                    d_info['职务'] = index
                elif '姓名' in ele:
                    d_info['姓名'] = index
                elif '性别' in ele:
                    d_info['性别'] = index
                elif '学号' in ele:
                    d_info['学号'] = index
                elif '长号' in ele:
                    d_info['长号'] = index
                elif '宿舍' in ele:
                    d_info['宿舍'] = index
            for k, v in d_info.items():
                if v == -1:
                    error('警告: 通讯录缺少{}栏'.format(k))

            # 删除导引行    部门 职务 姓名 性别 学号 电话号码 短号 籍贯 宿舍号 班级
            del l_AdrBbk[0]
            each_dept = ''  # 部门名
            # 获成员的各种属性
            for index, row in enumerate(l_AdrBbk):
                members.append(Member())
                m = members[-1]
                # 第一列不为'',为部门时 ( 第一列不为空就是部门名字 )
                # 组织部 小红   第一列为部门名
                #       小明   第一列为''

                if row[0] != '':
                    each_dept = row[0]
                # 根据学号判断年级
                preNo = row[d_info['学号']][0:2]
                begin = time.mktime(time.strptime("09 10 20{}".format(preNo),"%m %d %Y"))
                now = time.time()
                year = (now-begin)/60/60/24/365
                if year < 1:
                    m.grade = '大一'
                elif year < 2:
                    m.grade = '大二'
                elif year < 3:
                    m.grade = '大三'
                elif year < 4:
                    m.grade = '大四'

                m.name = row[d_info['姓名']]
                m.sex = row[d_info['性别']]
                m.id = row[d_info['学号']].replace('.0', '')
                m.longPhoNum = row[d_info['长号']].replace('.0', '')
                m.dept = each_dept
                m.job = row[d_info['职务']]
                try:
                    m.buld = row[d_info['宿舍']].split('栋')[0] + '栋'
                    m.dorPla = row[d_info['宿舍']].split('栋')[1]
                except:
                    error('警告: {}的宿舍信息不是规范的"xx栋xx"'.format(m.name))

            return True, members

        # 加载失败
        except PermissionError:
            traceback.print_exc()
            print('信息: 某程序正在使用该文件')
            error('警告: 无法读取 通讯录\n      你什么都干不了!')

        except Exception as e:
            traceback.print_exc()
            print('信息:', e.args)
            error('警告: 无法读取 通讯录\n      你什么都干不了!')

    # 从教务系统的课表中提取课表信息
    def loadOfficalTimeSchedules(self, members):
        try:
            # 教务系统下的xls课表
            path = os.getcwd() + '\\' + self.txt_jwxt + '\\'
            allTimeScheduleNames = [f for f in os.listdir(path) if 'xls' in f]

            for TimeScheduleName in allTimeScheduleNames:

                # 循环两次制作单双周课表数字
                dataOfTimeSchedule = readXlsFile(path + TimeScheduleName)
                # 学生名字
                title = dataOfTimeSchedule[0][0]
                while '  ' in title:
                    title = title.replace('  ', ' ')
                try:
                    stuName = dataOfTimeSchedule[0][0].split(' ')[1]
                except:
                    error('警告: 无法从{}中提取姓名!'.format(TimeScheduleName))
                # 课程
                classes = dataOfTimeSchedule[3:8]
                if fid(members, stuName) != None:
                    try:
                        tmpMem = members[fid(members, stuName)]
                        # 设置默认单双周课表数字
                        tmpMem.schedule = [[], []]
                        # 单周
                        tmpMem.schedule[0] = [[], [], [], [], [], [], []]
                        # 双周
                        tmpMem.schedule[1] = [[], [], [], [], [], [], []]

                        for index, type in enumerate(['single', 'double']):
                            # 周一 - 周五
                            for day in range(7):
                                for bigpart in range(5):
                                    if '双周' in classes[bigpart][day + 1]:
                                        # print('双周有课')
                                        if type == 'single':
                                            tmpMem.schedule[index][day].append('0')
                                            # print(members[fid(stuName)].schedule[index][day])
                                            # input('single')
                                        elif type == 'double':
                                            tmpMem.schedule[index][day].append('1')
                                            # print(members[fid(stuName)].schedule[index][day])
                                            # input('double')
                                    elif '单周' in classes[bigpart][day + 1]:
                                        # print('单周有课')
                                        if type == 'single':
                                            tmpMem.schedule[index][day].append('1')
                                            # input('single')
                                        elif type == 'double':
                                            tmpMem.schedule[index][day].append('0')
                                            # input('double')
                                    elif classes[bigpart][day + 1] != ' ':
                                        tmpMem.schedule[index][day].append('1')
                                    else:
                                        tmpMem.schedule[index][day].append('0')
                    except:
                        error('警告: 无法从{}中提取课表数字!'.format(TimeScheduleName))
                else:
                    error('警告: 通讯录中找不到{}!'.format(stuName))

            return True, members
        # 无法记教务系统的通讯录文件
        except:
            traceback.print_exc()
            return False, members

    # 从文件读取数据信息
    def loadDataFromFile(self):
        # members = []   学生列表/数组
        # 有通讯录, 才能加载 课表数字/活动信息
        self.status_loadAdrBok, members = self.loadAddressBbook()  # 加载 通讯录 文件
        if self.status_loadAdrBok:
            # 活动信息 使用了 课表 的内容, 所以存在课表, 才能加载 活动信息
            self.status_loadSch, members = self.loadOfficalTimeSchedules(members)  # 加载 课表 文件
        return members

# 程序头
if __name__ == '__main__':
    # 红会类
    redClub = Club()
    # 数据
    data = loadData()
    redClub.members = data.loadDataFromFile()
    while True:
        print('┏━━━━━┓')
        print('┃  社团管理┃')
        print('┃1.添加成员┃')
        print('┃2.移除成员┃')
        print('┃3.浏览成员┃')
        print('┃4.输出信息┃')
        print('┃5.活动安排┃')
        print('┃6.退出工具┃')
        print('┗━━━━━┛')
        choice = input('选择: ')
        while choice not in ['1', '2', '3', '4', '5', '6']:
            choice = input('选择: ')

        if choice == '1':
            # 添加成员
            redClub.addMember()
        elif choice == '2':
            # 删除成员
            redClub.delMember()
        elif choice == '3':
            # 浏览成员
            redClub.browMember()
        elif choice == '4':
            # 输出信息
            redClub.outputMain()
        elif choice == '5':
            # 安排活动
            redClub.arrnageAct()
        elif choice == '6':
            print('\n提示: 程序结束')
            break
