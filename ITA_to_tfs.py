import pandas as pd
import os
import sys
import csv

et_files=[et_file for et_file in (os.listdir(os.path.split(os.path.realpath(__file__))[0])) if et_file.endswith(".et")]
if len(et_files) > 1:
    print("error！目录下有超过1个.et文件，请只保留一个")
    sys.exit(1)
elif len(et_files) ==0:
    print("error！目录下没有.et文件，请保留一个")
    sys.exit(1)
else:
    et_file=et_files[0]
try:
    df = pd.read_excel(et_file, sheet_name="测试用例")
except ValueError:
    print("sheet名称须命名为测试用例")
    sys.exit(1)


df_columns=df.columns.tolist()
#处理表格
#1.第一列修改内容为测试用例
df[df_columns[0]]="测试用例"
#2.修改关联需求列的位置
related_requirements_ID=df[df_columns[2]]
df.drop(labels=[df_columns[2]], axis=1,inplace = True)
df.insert(12, df_columns[2], related_requirements_ID)
#3.删除测试阶段列
df.drop(labels=[df_columns[6]], axis=1,inplace = True)
#4.删除资产标签列
df.drop(labels=[df_columns[9]], axis=1,inplace = True)
#5.修改前置条件列位置
pre_condition=df[df_columns[10]]
df.drop(labels=[df_columns[10]], axis=1,inplace = True)
df.insert(12, df_columns[10], pre_condition)
#6.为设计者添加"tfsx\"
def add_tfsx(designer):
    if not designer.startswith("tfsx\\"):
        designer="tfsx\\"+designer
    return designer
df["设计者"]=df["设计者"].apply(add_tfsx)

#7修改表头名称
df_columns=df.columns.tolist()
expected_columns=["类型","测试用例编号","用例描述","用例优先级","交易类型","测试类型","用例属性","[步骤]描述","[步骤]预期","关联需求编号","ATP编号","ATP执行方式","前置条件","设计者"]
dict_comlumns=dict(zip(df_columns,expected_columns))
df=df.rename(columns=dict_comlumns,errors="raise")
#8.保存到文件
df.to_csv('tfsx.csv', index=False, header=True)
print("运行成功，转换成功文件名为tfsx.csv")



with open('tfsx.csv', 'r',encoding="utf-8") as f,open('tfsx_new.csv', 'w+',encoding="utf-8",newline='') as new_f:
    reader = csv.reader(f)
    csv_writer = csv.writer(new_f)
    for row in reader:
        new_row=[]
        for content in row:
            content=content.replace("\n","")
            new_row.append(content.strip("\r\n"))
        csv_writer.writerow(new_row)
os.remove('tfsx.csv')
os.rename('tfsx_new.csv','tfsx.csv')