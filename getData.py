import pandas as pd
from datetime import datetime

def test1():
    data = pd.read_excel('收款明细表样表.xls')
    print(data)
    # data_users = data['报销人']


def test2():
    data = pd.read_excel('收款明细表样表.xls')
    columns = list(data)
    # print(columns)
    data = pd.DataFrame(data)
    a = data.iloc[:, :columns.index('抄送人员')+1]

    print(a.values[1])
    b = data.loc[:,'通知人员':'抄送人员']
    c = list(a.iloc[0])
    print(c[0].strftime('%Y-%m-%d'))

def test3():
    data = pd.read_excel('付款明细表样表.xls')
    columns = list(data)    # 获取列名
    # print(columns)
    data = pd.DataFrame(data)
    a = data.iloc[:, :columns.index('收款从属')+1]  # 截取有效数据
    user_index = columns.index('报销人')
    data_dict = {}
    print(user_index)
    for var in a.values:
        # print(var[user_index], var)
        if var[user_index] in data_dict.keys():
            data_dict[var[user_index]].append(list(var))
        else:
            data_dict[var[user_index]] = [list(var)]
    print(len(data_dict))
    for key, value in data_dict.items():
        print(key, value)


def test4():
    data = pd.read_excel('付款明细表样表.xls')
    df = data['报销人']
    users = df.drop_duplicates()    # 去重
    print(users)
    print(len(users))
    # df = pd.DataFrame(data)
    # print(df)

if __name__ == '__main__':
    test3()
    test4()
    # datetime.today().strftime()