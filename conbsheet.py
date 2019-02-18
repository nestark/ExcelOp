import openpyxl
import Getworkbook
a=Getworkbook.getFileName('.')
print('以上文件已读取，正在处理中')
wb2=openpyxl.Workbook()
ws2=wb2.active
ws2.title='每日数据'
n=1
for worksheet in a:
    wb1=openpyxl.load_workbook(worksheet)
    ws1=wb1.active
    time_range = ws1['B4:B16']
    value_range = ws1['E4:E16']

    if ws1.title[12:14]=='温度':
        for i in range(n, n + 13):
            ws2.cell(row=i,column=1).value = time_range[i-n][0].value
            ws2.cell(row=i, column=2).value = ws1.title[4:7]
            ws2.cell(row=i,column=3).value=value_range[i-n][0].value
    else:
        for i in range(n, n + 13):
            ws2.cell(row=i, column=4).value = value_range[i-n][0].value
        n+=13

ws2.insert_rows(0)
ws2['A1']='日期'
ws2['B1']='机房'
ws2['C1']='温度'
ws2['D1']='湿度'

wb2.save('summary.xlsx')
print('数据汇总已保存在summary.xlsx')
