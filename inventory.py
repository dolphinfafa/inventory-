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

wb1 = load_workbook('单品.xlsx')
wb2 = load_workbook('款号.xlsx')
wb3 = load_workbook('库存.xlsx')

baishazhou = '武汉总仓'
wansongyuan = '经济万松园店'
ezhou = '仓储鄂州阳光货仓店'
fuxin = '仓储富鑫常青店'
sifang = '仓储肆方光谷店'

ws1 = wb1.active
ws2 = wb2.active
ws3 = wb3.active

def product_search(warehouse):
    i1 = 2
    invent = []
    warehouse = warehouse
    while(True):
        code = ws1.cell(i1,3).value
        if code == None:
            break

        number = section_search(code, warehouse)
        name = ws1.cell(i1, 2).value
        invent.append([name, code, number])
        i1 = i1 +1
        # print('prodect get:')
        # print(number)
    return(invent)


def section_search(code, warehouse):
    i2 = 2
    warehouse = warehouse
    while(True):
        section = ws2.cell(i2,1).value
        if section == None:
            break
        for x in range(3,7):
            section = ws2.cell(i2, x).value
            if section == code:
                number = 0
                for x in range(3,7):
                    check = ws2.cell(i2, x).value
                    if check != None:
                        x = erp_search(check, warehouse)
                        number = number + x
                # print('section return:')
                # print(number)
                return(number)

        i2 = i2 + 1


def erp_search(check, warehouse):
    i3 = 2
    number = 0
    warehouse = warehouse
    while(True):
        erp_section = ws3.cell(i3, 3).value
        if erp_section == None:
            break
        print('check is :'+check)
        print(erp_section)
        if erp_section == check and ws3.cell(i3, 1).value == warehouse:
            x = ws3.cell(i3, 9).value
            number = number + int(x)
            print(x)
        i3 = i3 + 1
        # print('erp return:')
        # print(number)
    return(number)


def new_invent(invent, warehouse):
    wb = Workbook()
    ws = wb.active
    x = 0
    for i in invent:
        ws.cell(x+2, 1).value = invent[x][0]
        ws.cell(x+2, 2).value = invent[x][1]
        ws.cell(x+2, 3).value = invent[x][2]
        x = x + 1
    name = warehouse+'.xlsx'
    wb.save(name)
  


if __name__ == '__main__':
    invent = product_search(baishazhou)
    new_invent(invent, baishazhou)

    invent = product_search(wansongyuan)
    new_invent(invent, wansongyuan)

    invent = product_search(ezhou)
    new_invent(invent, ezhou)

    invent = product_search(fuxin)
    new_invent(invent, fuxin)

    invent = product_search(sifang)
    new_invent(invent, sifang)
