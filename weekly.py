import datetime, calendar
import pandas as pd
import os

"""日期"""
sunday = datetime.datetime.today()
while sunday.weekday() != calendar.SUNDAY:
    sunday -= datetime.timedelta(days=1)
monday = sunday - datetime.timedelta(days=6)
prieod = "{}年{}月{}日-{}月{}日".format(
    monday.year, monday.month, monday.day, sunday.month, sunday.day
)


def huanbi(rate):
    """判断环比上涨或下降"""
    change = "下降" if rate < 0 else "上涨"
    return "{}{}%".format(change, rate)


def detail(row):
    """成交面积、套数及环比情况"""
    return "成交{:.2f}万㎡，环比{}，成交{:.0f}套，环比{}".format(
        row['已售面积']/1e4, huanbi(row['面积环比(%)']),
        row['已售套数(套)'], huanbi(row['套数环比(%)'])
    )


"""商品房"""
shangpinfang = pd.read_excel('templates.xlsx', '商品房', index_col='板块')
shangpinfang_quanshi = "全市商品房" + detail(shangpinfang.ix['合计'])
shangpinfang_lishui = "溧水区商品房" + detail(shangpinfang.ix['溧水'])

"""商品住宅"""
zhuzhai = pd.read_excel('templates.xlsx', '住宅', index_col='板块')
zhuzhai_quanshi = "商品住宅" + detail(zhuzhai.ix['合计'])
zhuzhai_lishui = "商品住宅" + detail(zhuzhai.ix['溧水'])

"""输出结果"""
result = "{}，{}，其中{}；本周，{}，其中{}。".format(
    prieod,
    shangpinfang_quanshi, zhuzhai_quanshi,
    shangpinfang_lishui, zhuzhai_lishui
)
with open("{}.txt".format(prieod), "w+") as f:
    f.write(result)
os.system('notepad {}.txt'.format(prieod))
