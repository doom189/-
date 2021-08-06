#!/usr/bin/python3

#region 引入模块
import xlwings as xw  # Excel模块
import os  # 文件目录模块
import re  # 原生正则表达式
import time  # 时间模块
import datetime  # 日期时间模块
import sys  # 系统
#endregion 引入模块 尾巴

#region 周几转换
result = {
    "0": 7,
    "1": 1,
    "2": 2,
    "3": 3,
    "4": 4,
    "5": 5,
    "6": 6
}
#endregion 周几转换 尾巴

#region 验算传入项是否存在
def GetDicVal(Dic, *Keys):
    临时字典 = Dic.copy()
    for Key in Keys:
        if Key in 临时字典:
            临时字典 = 临时字典[Key]
            if type(临时字典) != dict:
                return 临时字典
    return False
#endregion 验算传入项是否存在 尾巴

#region 多层字典填值
def SetDicVal(Dic, Val, *Keys):
    临时字典 = Dic
    键长度 = len(Keys)
    计数项 = 0
    for Index in range(0, 键长度 - 1):
        计数项 += 1
        if Keys[Index] not in 临时字典:
            临时字典[Keys[Index]] = {}
        临时字典 = 临时字典[Keys[Index]]
    临时字典[Keys[键长度 - 1]] = Val
#endregion 多层字典填值 尾巴

#region 文本转时间格式
def Str2Time(StrTime):
    return time.strptime('{}'.format(StrTime), r"%H:%M")
#endregion 文本转时间格式 尾巴

#region 文本转日期格式
def Str2Date(StrDate):
    return time.strptime('{}'.format(StrDate), r"%Y-%m-%d")
#endregion 文本转日期格式 尾巴

#region 日期减一天
def 日期减一天(日期文本):
    日期 = datetime.date.fromisoformat(日期文本)
    日期 += datetime.timedelta(seconds=-1)
    return 日期.isoformat()
#endregion 日期减一天 尾巴

#region 获取日期是周几
def Date2WeekDay(DateObj):
    return result.get(time.strftime('%w', DateObj))
#endregion 获取日期是周几 尾巴

#region 获取字典键不报错
def GetVal(Dic, Key):
    if Key in Dic:
        return Dic[Key]
    else:
        return ''
#endregion 获取字典键不报错 尾巴

#region 获取字典键不报错
def GetFloat(Dic, Key):
    if Key in Dic:
        return Dic[Key]
    else:
        return 0
#endregion 获取字典键不报错 尾巴

#region 计算满半个小时的加班
def JiSuanJiaBan(JieShu, KaiShi):
    return (
        int(
            (
                (JieShu.tm_hour - KaiShi.tm_hour) * 60  # 小时转分钟
                + (JieShu.tm_min - KaiShi.tm_min)  # 加上分钟数
            ) / 30  # 计算有几个半小时
        ) * 30  # 取整后还原半小时个数
    ) / 60  # 转成小时
#endregion 计算满半个小时的加班 尾巴

#region 智能判断当天中午上下班
def ShangXiaBan(员工工号, 日期, 刷卡时间):
    工号日期主键 = f"{员工工号}丨{日期}"  # 工号+日期
    if 工号日期主键 in ShuaKaData:
        BanCi = RenYuanBanCi[员工工号]  # 取得班次名
        if BanCi == '两班倒':
            if GetVal(ShuaKaData[工号日期主键], "1下班") == '':
                ShuaKaData[工号日期主键]["1下班"] = 刷卡时间
                return
            ShuaKaData[工号日期主键]["2上班"] = 刷卡时间
        if BanCi == '责任制3笔':
            ShuaKaData[工号日期主键]["下午上班"] = 刷卡时间
            return  # 结束
        if GetVal(ShuaKaData[工号日期主键], "上午下班") == '':
            ShuaKaData[工号日期主键]["上午下班"] = 刷卡时间
            return
        ShuaKaData[工号日期主键]["下午上班"] = 刷卡时间
#endregion 智能判断当天中午上下班 尾巴

#region 两班倒自动判断
def 两班倒上班判断(工号日期主键, 刷卡时间, 刷卡数据):
    if 刷卡时间 >= Str2Time(GetDicVal(BanCiDic, '生产白班', '1', '最早上班')) \
        and Str2Time(GetDicVal(BanCiDic, '生产白班', '1', '上班时间')):
        SetDicVal(ShuaKaData, 刷卡数据, 工号日期主键, '1上班')
    elif 刷卡时间 >= Str2Time(GetDicVal(BanCiDic, '生产夜班', '1', '最早上班')) \
        and Str2Time(GetDicVal(BanCiDic, '生产夜班', '1', '上班时间')):
        SetDicVal(ShuaKaData, 刷卡数据, 工号日期主键, '1上班')
    else:
        SetDicVal(ShuaKaData, f"{刷卡数据}丨两班倒判断错误", 工号日期主键, '1上班')
#endregion 两班倒自动判断 尾巴

ShuaKaData = {}  # 人员刷卡数据汇总空字典

#region 读取班次信息生成字典
# 打开Excel程序，默认设置：Excel程序不可见，只打开不新建工作薄
app = xw.App(visible=False, add_book=False)
app.display_alerts = False  # 关闭错误警告
app.screen_updating = False  # 屏幕更新关闭
# 文件位置：filepath，打开test文档，然后保存，关闭，结束程序
wb = app.books.open(r'.\班次信息维护.xlsx')  # 只读打开路径
print('正在打开读取班次信息维护表0')
sht = wb.sheets[0]  # 获取表
LastRow = sht.range('A65536').end('up').row  # 获取最后一行

读取数据 = sht.range(f"A2:L{LastRow}").value

BanCiDic = {}  # 班次空字典
# 逐行读取班次信息
for row in 读取数据:
    BanCi = row[0]  # 班次名 职员班
    ShiDuan = row[1]  # 时段 上午/下午

    # 判断班次是否存在
    if (BanCi not in BanCiDic):
        BanCiDic[BanCi] = {}  # 不存在就创建班次的空字典
    # 判断班次是否存在 尾巴

    # 判断时段是否存在
    if (ShiDuan not in BanCiDic[BanCi]):
        BanCiDic[BanCi][ShiDuan] = {}  # 不存在就创建时段的空字典
    # 判断时段是否存在 尾巴

    BanCiDic[BanCi][ShiDuan]['最早上班'] = row[2]
    BanCiDic[BanCi][ShiDuan]['上班时间'] = row[3]
    BanCiDic[BanCi][ShiDuan]['下班时间'] = row[5]
    BanCiDic[BanCi][ShiDuan]['最晚下班'] = row[6]
# 逐行读取班次信息 尾巴

#region 逐行读取人员所在班次
print('正在打开读取班次信息维护表1')
sht = wb.sheets[1]  # 获取表
LastRow = sht.range('A65536').end('up').row  # 获取最后一行

读取数据 = sht.range(f"A2:L{LastRow}").value

RenYuanBanCi = {}  # 人员班次对照
for row in 读取数据:
    YGGH = row[0]  # 员工工号 02284
    BanCi = row[2]  # 班次名 职员班

    RenYuanBanCi[YGGH] = BanCi  # 字典增加班次名
#endregion 逐行读取人员所在班次 尾巴

wb.close()
app.quit()  # 退出
#endregion 读取班次信息生成字典 尾巴

#region 循环读取目录下的文件和目录
# root 所指的是当前正在遍历的这个文件夹的本身的地址
# dirs 是一个 list ，内容是该文件夹中所有的目录的名字(不包括子目录)
# files 同样是 list , 内容是该文件夹中所有的文件(不包括子目录)
for root, dirs, files in os.walk(".", topdown=False):
    # 循环获取所有文件
    for name in files:
        if re.match(r'.*?\d{1,2}月打卡记录\.xlsx?$', name, re.M):
            filepath = os.path.join(root, name)  # 输出路径
            print('当前打开的文件是', filepath)
            # 打开Excel程序，默认设置：Excel程序不可见，只打开不新建工作薄
            app = xw.App(visible=False, add_book=False)
            app.display_alerts = False  # 关闭错误警告
            app.screen_updating = False  # 屏幕更新关闭
            # 文件位置：filepath，打开test文档，然后保存，关闭，结束程序
            wb = app.books.open(filepath)  # 只读打开路径
            sht = wb.sheets[0]  # 获取表
            LastRow = sht.range('A65536').end('up').row  # 获取最后一行

            读取数据 = sht.range(f"A2:L{LastRow}").value

            # 循环读取打卡记录并比较打卡时间数据
            rowcnt = 0
            总数 = len(读取数据)
            for row in 读取数据:
                rowcnt += 1
                print('正在读取记录\r', f"{rowcnt}/{总数}", end='', flush=True)
                YGGH = row[0]  # 员工工号(文本)
                YGXM = row[1]  # 员工姓名
                ShuaKaRiQi = row[2]  # 刷卡日期
                ShuaKa = row[3]  # 刷卡时间
                try:
                    BanCi = RenYuanBanCi[YGGH]  # 取得班次名
                except:
                    # print("Unexpected error:", sys.exc_info()[0])
                    continue

                # 转换刷卡时间用以比较
                ShuaKaShiJian = Str2Time(ShuaKa)
                ShiJian = ShuaKaShiJian.tm_hour  # 获取时间
                WeekNum = Date2WeekDay(Str2Date(ShuaKaRiQi))  # 周几

                ShiDuan = '下午'  # 时段
                工号日期主键 = f"{YGGH}丨{ShuaKaRiQi}"
                # 构建工号+日期数据结构
                if 工号日期主键 not in ShuaKaData:
                    ShuaKaData[工号日期主键] = {
                        "日期": ShuaKaRiQi,
                        "员工工号": YGGH,
                        "员工姓名": YGXM,
                        "今天周几": WeekNum
                    }  # 新建当天刷卡数据
                # 构建工号+日期数据结构 尾巴
                #region 自动两班倒
                if BanCi == '两班倒':
                    if not GetDicVal(ShuaKaData, 工号日期主键, "1上班"):
                        两班倒上班判断(工号日期主键, ShuaKaShiJian, ShuaKa)
                        continue  # 跳过
                    elif ShiJian > 23 or ShiJian > 0 and ShiJian < 1:
                        # 生产夜班中间下班和2上班卡
                        ShuaKaRiQi = 日期减一天(ShuaKaRiQi)
                        工号日期主键 = f"{YGGH}丨{ShuaKaRiQi}"
                        if not GetDicVal(ShuaKaData, 工号日期主键, "1下班"):
                            SetDicVal(ShuaKaData, ShuaKa, 工号日期主键, "1下班")
                        else:
                            SetDicVal(ShuaKaData, ShuaKa, 工号日期主键, "2上班")
                        continue  # 跳过
                    elif ShiJian > 1 and ShiJian < 10 and GetDicVal(ShuaKaData, 工号日期主键, "2上班"):
                        # 时段2下班卡
                        SetDicVal(ShuaKaData, ShuaKa, 工号日期主键, "2下班")
                        continue  # 跳过
                    elif ShiJian <= 11:
                        # 白班上午上班卡
                        SetDicVal(ShuaKaData, ShuaKa, 工号日期主键, "1上班")
                        continue  # 跳过
                    elif ShiJian < 13:
                        # 白班中午自动判断
                        if not GetDicVal(ShuaKaData, 工号日期主键, "1下班"):
                            SetDicVal(ShuaKaData, ShuaKa, 工号日期主键, "1下班")
                        else:
                            SetDicVal(ShuaKaData, ShuaKa, 工号日期主键, "2上班")
                        continue  # 跳过
                    elif ShiJian > 16 or ShiJian > 20:
                        SetDicVal(ShuaKaData, ShuaKa, 工号日期主键, "2下班")
                #endregion 自动两班倒 尾巴
                # 刷卡时间判断上下午
                elif ShiJian < 11:
                    ShiDuan = '上午'  # 上班卡
                elif ShiJian < 13:
                    ShiDuan = '智能中间判断'
                if BanCi == '责任制2笔':
                    ShiDuan = '全天'
                # 刷卡时间判断上下午 尾巴

                if ShiDuan == '智能中间判断':
                    ShangXiaBan(YGGH, ShuaKaRiQi, ShuaKa)
                    continue  # 跳过

                # 判断是否是周六日
                if WeekNum > 5:
                    # print(YGGH, ShuaKaRiQi, '周', WeekNum, '加班', ShuaKa)
                    # 取第一笔打卡记录
                    if ShiDuan + "上班" not in ShuaKaData[工号日期主键]:
                        ShuaKaData[工号日期主键][ShiDuan + "上班"] = ShuaKa
                    else:
                        ShuaKaData[工号日期主键][ShiDuan + "下班"] = ShuaKa

                        # 计算加班时长
                        JiaBan = JiSuanJiaBan(ShuaKaShiJian, Str2Time(ShuaKaData[工号日期主键][ShiDuan + "上班"]))
                        if JiaBan > 0:
                            ShuaKaData[工号日期主键]["加班"] = GetFloat(ShuaKaData[工号日期主键], "加班") + JiaBan
                        # 计算加班时长 尾巴
                    # 取第一笔打卡记录 尾巴
                    continue  # 跳过本行
                # 判断是否是周六日 尾巴

                # 判断班次是否存在
                if BanCi in BanCiDic:
                    # 判断时段是否存在
                    if ShiDuan in BanCiDic[BanCi]:
                        # 平时请假记录表判断(过滤打卡错误情况)
                        # 节假日补班记录表判断(周末补班转正常班)
                        XiaBanShiJian = Str2Time(BanCiDic[BanCi][ShiDuan]['下班时间'])
                        # 判断该时段的上班打卡区间
                        if ShuaKaShiJian >= Str2Time(BanCiDic[BanCi][ShiDuan]['最早上班']) \
                                and ShuaKaShiJian <= Str2Time(BanCiDic[BanCi][ShiDuan]['上班时间']):
                            # 上班打卡了
                            ShuaKaData[工号日期主键][ShiDuan + "上班"] = ShuaKa
                            # print(YGGH, ShuaKaRiQi, '周', WeekNum, ShiDuan, '上班卡', ShuaKa)
                        # 判断该时段的上班打卡区间 尾巴

                        # 判断下班打卡区间
                        elif ShuaKaShiJian >= XiaBanShiJian \
                                and ShuaKaShiJian <= Str2Time(BanCiDic[BanCi][ShiDuan]['最晚下班']):
                            # 下班打卡了
                            ShuaKaData[工号日期主键][ShiDuan + "下班"] = ShuaKa

                            # 计算加班时长
                            JiaBan = JiSuanJiaBan(ShuaKaShiJian, XiaBanShiJian)
                            if JiaBan > 0:
                                ShuaKaData[工号日期主键]["加班"] = JiaBan
                            # 计算加班时长 尾巴
                            # print(YGGH, ShuaKaRiQi, '周', WeekNum, ShiDuan, '下班卡', ShuaKa)
                        # 判断下班打卡区间 尾巴

                        # 例外情况测试 1. 迟到; 2. 请假; 3. 漏打卡;
                        else:
                            # 迟到
                            if ShuaKaShiJian > Str2Time(BanCiDic[BanCi][ShiDuan]['上班时间']) \
                                    and ShuaKaShiJian < XiaBanShiJian \
                                    and ShiDuan + "上班" not in ShuaKaData[工号日期主键]:
                                # print("上班迟到咯! 如果请假请自行处理")
                                ShuaKaData[工号日期主键][ShiDuan + "上班"] = f"{ShuaKa}丨迟到"
                                continue  # 迟到跳出
                            # 迟到 尾巴

                            # 早退
                            if ShuaKaShiJian > Str2Time(BanCiDic[BanCi][ShiDuan]['上班时间']) \
                                    and ShuaKaShiJian < XiaBanShiJian \
                                    and ShiDuan + "下班" not in ShuaKaData[工号日期主键]:
                                # print("早退, 如果请假请自行处理")
                                ShuaKaData[工号日期主键][ShiDuan + "下班"] = f"{ShuaKa}丨早退"
                                continue  # 早退跳出
                            # 早退 尾巴

                            # print(YGGH, YGXM, ShuaKaRiQi, '周', WeekNum, ShiDuan, '例外情况', ShuaKa)
                        # 例外情况测试 尾巴
                    # 判断时段是否存在 尾巴
                # 判断班次是否存在 尾巴
            # 循环读取打卡记录并比较打卡时间数据 尾巴
            wb.close()  # 关闭工作簿
            app.quit()  # 退出

            # print(ShuaKaData)  # 输出刷卡数据

            # 输出刷卡数据到Excel
            print('\n在写结果表test了哦!')
            app = xw.App(visible=False, add_book=False)
            wb = app.books.open(r'.\test.xlsx')
            sht = wb.sheets[0]  # 获取表
            rowcnt = 1  # 初始化行
            总数 = len(ShuaKaData)
            填表数据 = []
            for Key in ShuaKaData:
                rowcnt += 1  # 行数+1
                print('\r', f"{rowcnt - 1}/{总数}", end='', flush=True)
                临时数据 = [
                    GetVal(ShuaKaData[Key], '日期'),
                    f"'{GetVal(ShuaKaData[Key], '员工工号')}",
                    GetVal(ShuaKaData[Key], '员工姓名'),
                    GetVal(ShuaKaData[Key], '上午上班') + GetVal(ShuaKaData[Key], '全天上班') + GetVal(ShuaKaData[Key], '1上班'),
                    GetVal(ShuaKaData[Key], '上午下班') + GetVal(ShuaKaData[Key], '1下班'),
                    GetVal(ShuaKaData[Key], '下午上班') + GetVal(ShuaKaData[Key], '2上班'),
                    GetVal(ShuaKaData[Key], '下午下班') + GetVal(ShuaKaData[Key], '全天下班') + GetVal(ShuaKaData[Key], '2下班'),
                    GetVal(ShuaKaData[Key], '加班'),
                    f"=WEEKDAY($A{rowcnt + 1},2)"
                ]
                填表数据.append(临时数据)
                # print(rowcnt, Key)
            sht.range('A2').value = 填表数据
            sht.range((2, 1), (rowcnt, 1)).number_format = "yyyy-mm-dd;@"
            sht.range((2, 2), (rowcnt, 3)).number_format = "@"
            sht.range((2, 4), (rowcnt, 7)).number_format = "hh:mm;@"
            sht.range((2, 8), (rowcnt, 8)).number_format = "G/通用格式"
            for col in range(1, 7):
                sht.range((1, col), (rowcnt, col)).autofit()  # 自动列宽
            # 输出刷卡数据到Excel 尾巴

            wb.save()  # 保存工作簿
            wb.close()  # 关闭工作簿
            app.quit()  # 退出
            print('输出完成')
            sys.exit()
  # 结束程序
    # 循环获取所有文件 尾巴
#endregion 循环读取目录下的文件和目录 尾巴
