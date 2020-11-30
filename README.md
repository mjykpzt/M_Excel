# M_Excel
对openpyxl的简单封装


```python
from tem_mould import Make_tem


#生成体温表的例子

if __name__ == '__main__':
    path = r"" #路径
    sheet_name="" #sheet名称
    col=0 # 初始坐标
    row=0# 初始坐标
    num  = 0 #生成体温的次数
    s = Make_tem.TemperatureXlsx(path,sheet_name,row,col)
    for i in range(num):
        b = [i, "name" + str(i), str(i) + "x"] #基础数据
        s.add(b + [2018, "dsadsa"], 21)
        s.exe()

    s.save()
```
