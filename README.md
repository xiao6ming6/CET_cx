# CET批量查询脚本
**所有脚本均只能在官网开放查询时使用，其余时间不可使用**
## 一、报考科目查询
cet2.py 提供报考科目查询服务
需要读取数据：
 data/data.xlsx excel表中第一列填充姓名，第二列填充身份证号
输出数据：
 data/cet_data.txt 文本中第一列是姓名，第二列是报考科目
## 二、成绩查询
cetcx.py 提供成绩查询服务
需要读取数据：
 data/data.xlsx excel表中第一列填充姓名，第二列填充身份证号
 data/cet_data.txt 文本中第一列是姓名，第二列报考科目
输出数据：
 data/cet_data.txt 文本中第一列是姓名，第二列是科目，第三列是总分，后面是各科目分数
 ## 三、注意事项
 需提前安装好chrome和对应的py驱动
