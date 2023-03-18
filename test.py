import xlwings as xw
wb = xw.Book('报价表（顺诚）3月6.xlsm')
#name={'0':'联盛','1':'绿叶','2':'地龙'}
#keys=[0,1,2]
name=["联盛","绿叶","地龙"]
sheet=xw.sheets['Sheet1']
b=1
a=1
for x in name:
    exec("temp%s = wb.sheets['%s']"%(a,x))
    a=a+1
for i in range(1,3):
    for row in range(41, 45):
        row_str = str(row)
        exec("start_value = temp%s.range('A' + row_str).value"%i)
        exec("sheet['A%s'].value = start_value"%b)
        b=b+1
        print(start_value)   