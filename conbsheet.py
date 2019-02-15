import openpyxl
import Getworkbook
Getworkbook.getFileName()
wb1=openpyxl.load_workbook("35F_34东-温湿度_湿度趋势分析_2019-02-01至2019-02-13.xlsx")
ws1=wb1.active
cell_range = ws1['A4':'E16']