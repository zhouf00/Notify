import pandas as pd

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

    data = pd.DataFrame(pd.read_excel('付款明细表样表.xls'))
    columns = list(data)  # 获取列名

    a = data.iloc[:, columns.index('日期'):columns.index('收款从属')+1]  # 截取有效数据
    to_index = columns.index('报销人')
    cc = '抄送人员'
    if cc in columns:
        cc_index = columns.index(cc)
    else:
        cc_index = None
    data_dict = {}
    for var in a.values:
        if var[to_index] in data_dict.keys():
            data_dict[var[to_index]]['value'].append(list(var[:to_index]))
        else:
            if cc_index:
                data_dict[var[to_index]] = {
                    'to': var[to_index],
                    'cc': var[cc_index],
                    'value': [list(var[:to_index])]
                }
            else:
                data_dict[var[to_index]] = {
                    'to': var[to_index],
                    'cc': None,
                    'value': [list(var[:to_index])]
                }
    for key, value in data_dict.items():
        print(key, value)
    return data_dict.items()


def test4():
    data = pd.read_excel('付款明细表样表.xls')
    df = data['报销人']
    users = df.drop_duplicates()    # 去重
    print(users)
    print(len(users))
    # df = pd.DataFrame(data)
    # print(df)

if __name__ == '__main__':
    a = test3()
    for key, value in a:
        for var in value['value']:
            print(value['to'], value['cc'], var)
        value.pop('value')
    print(a)
    # datetime.today().strftime()