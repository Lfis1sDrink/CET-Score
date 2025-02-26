# CET-Score
批量查询CET4/6的成绩，读取xls文件中的[姓名]列和[身份证号]列，查询对应的成绩写到xls文件对应的成绩列

## How to Use
**students.xlsx**和**query.py**文件放在同级目录下，保证**students.xlsx**中有[姓名]列和[身份证号]列，注意[身份证号]列的单元格格式。然后执行
```
python query.py
```
等待即可，默认每个请求之间延迟3s。
