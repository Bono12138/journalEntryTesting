import pandas as pd
import os

# 新建列表，存放文件名
file_list_1 = []
file_list_2 = []
# 新建列表，存放每个文件数据框（每一个excel读取后存放在数据框）
dfs_1 = pd.DataFrame() #存储试算平衡表合并结果
dfs_2 = pd.DataFrame() #存储序时账合并结果
dfs_3 = pd.DataFrame()
i = 0
j = 0
# 遍历
# 设置余额表存放目录
TBfilePath = r'C:\Users\LV333KW\Desktop\桌面20220527\爬虫项目\JEtesting\test\余额表'
# 设置序时账存放目录
JEfilePath = r'C:\Users\LV333KW\Desktop\桌面20220527\爬虫项目\JEtesting\test\序时账'
'''读取全部试算平衡表，合并后筛选出科目号码长度为4的结果'''
i = 0
for root, dirs, files in os.walk(TBfilePath):
    for file in files:
        file_path = os.path.join(root, file)   # 使用os.path.join(dirpath, name)得到全路径
        file_name = file.split('.')[0]
        print(root, file,"进度：",str(i+1)+"/25")
        TB = pd.read_excel(file_path, skiprows=7, header=None, names= ["科目名称","科目编码","C","D","E","科目层级","G","H","I","J","K","L",
                "M","N","O","P","本期借方发生额","R","S","本期贷方发生额","U","V","累计借方发生额","X","Y","累计贷方发生额","AA","AB","AC","科目全称","AE"]) 
                # 将excel转换成DataFrame（在这里根据余额表实际情况更新列数）
        TB_Filter = TB[TB['科目编码'].astype('str').str.len() == 4].copy() #只筛选出科目号码长度为4的结果
        TB_Sum = TB_Filter[["科目编码","科目层级","累计借方发生额","累计贷方发生额","科目全称"]].copy() #只筛选"科目号","累计借方发生额","累计贷方发生额"三列进行保留
        TB_Sum.insert(0,"基金名",file_name.strip("20221231余额报表")) #插入一列基金名
        i += 1
        if i == 1:
            dfs_1 = TB_Sum
        else:
            dfs_1 = pd.concat([dfs_1,TB_Sum],axis=0)
'''读取全部序时账，合并后筛选出科目号码长度为4的结果'''
dfs_2 = pd.DataFrame() #存储序时账合并结果
j = 0
JEfilePath = r'C:\Users\LV333KW\Desktop\桌面20220527\爬虫项目\JEtesting\test\序时账'
for root, dirs, files in os.walk(JEfilePath):
    for file in files:
        file_path = os.path.join(root, file)  # 使用os.path.join(dirpath, name)得到全路径
        file_name = file.split('.')[0]
        print(root, file, "进度：",str(j+1)+"/25")
        JE_dic = pd.read_excel(file_path,sheet_name=None, skiprows=3)  # 将excel转换成DataFrame
        JE = pd.DataFrame()
        for i in JE_dic:
            JE = pd.concat([JE,JE_dic[i]])
        JE_copy = JE.copy()
        JE_copy = JE_copy[JE_copy['凭证类型'] != "科目体系切换凭证"].copy() #筛选,剔除全部科目体系切换凭证的结果
        JE_copy = JE_copy.reset_index(drop=True) #重置序号，防止dataframe序号重复报错      
        JE_copy['JE_借方发生额'] = JE_copy.apply(lambda x: x.本位币金额 if x.借贷 == '借' else 0, axis=1)
        JE_copy['JE_贷方发生额'] = JE_copy.apply(lambda x: x.本位币金额 if x.借贷 == '贷' else 0, axis=1)
        JE_Sum = JE_copy.groupby(['科目代码'], as_index=False)[['JE_借方发生额', 'JE_贷方发生额']].sum()
        JE_Sum['基金名'] = JE_copy['投资组合']
        JE_Sum['科目号'] = JE_Sum['科目代码'].astype('str').apply(lambda x: x[0:4]).tolist()
        j += 1
        if j == 1:
            dfs_2 = JE_Sum
        else:
            dfs_2 = pd.concat([dfs_2,JE_Sum],axis=0)

'''将结果生成到指定的worksheet'''
from pandas import ExcelWriter
resultFilePath = r"C:\Users\LV333KW\Desktop\桌面20220527\爬虫项目\JEtesting\JE与TB合并结果与比对2.xlsx"
with ExcelWriter(resultFilePath) as writer:
	dfs_1.to_excel(writer,sheet_name='TB合并结果')
	dfs_2.to_excel(writer,sheet_name='JE合并结果')
	writer.save()
