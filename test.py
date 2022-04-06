from openpyxl import Workbook, load_workbook
from openpyxl.styles import *


import warnings
warnings.filterwarnings('ignore')

'''
需要三张表
sections - 款号清单
products - 单品清单
erp - 实际库存
'''

wb1 = load_workbook('products.xlsx')
wb2 = load_workbook('sections.xlsx')
wb3 = load_workbook('erp.xlsx')

ws1 = wb1['Sheet0']
ws2 = wb2['Sheet1']
ws3 = wb3['download']

a = ws2.cell(1034, 4).value
b = ws3.cell(2, 6).value
print(a)
print(b)
if a == b:
    print('check')