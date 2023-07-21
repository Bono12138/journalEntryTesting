# journalEntryTesting
# 余额表和序时账数据处理

## 项目概述

该项目的目标是读取存放在不同目录下的多个Excel文件(余额表和序时账),进行数据提取、合并和处理,最终生成结果汇总Excel文件。

## 代码说明

main.py中主要完成以下工作:

1. 定义存放余额表和序时账的目录路径。

2. 遍历读取每个Excel文件,加载为pandas DataFrame。  

3. 对DataFrame进行列提取和筛选,只保留需要的列。

4. 将多个DataFrame按照基金名进行合并。

5. 将合并后的DataFrame输出为结果汇总Excel文件。

## 环境需求

- Python 3.6+
- pandas
- openpyxl

## 目录结构

    ├── main.py 
    ├── input/
    │   ├── 余额表/
    │   ├── 序时账/
    ├── output/
        ├── 结果.xlsx

## 使用说明

1. 将余额表和序时账Excel文件分别放入input/余额表/和input/序时账/目录下

2. 运行main.py

3. 在output目录下找到结果汇总Excel文件  

## TODO

- [ ] 添加命令行参数,支持自定义输入和输出目录
- [ ] 支持parquet等其他格式的输出  
- [ ] 添加日志功能

欢迎提交issue或PR来完善项目!
