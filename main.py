from openpyxl.reader.excel import load_workbook
import openpyxl
book = load_workbook('spr.xlsx')
# Определяем рабочий лист
ws = book.worksheets[0]
str = []
kol = []
vr = []
Pustie_stroki = []
spisok_index_pust_str = []
# Определяем необходимую ячейку на листе
Cell_A = ws['A']
Cell_B = ws['B']
x = 0
for i in Cell_A:
    vr = [i.value]
    str.append(vr)
for j in Cell_B:
    vr = [j.value]
    kol.append(vr)
    # поиск пустых строк
    if vr[0] is None:
        Pustie_stroki.append(str[x])
    x += 1


Poisk_w = openpyxl.Workbook()
sh = Poisk_w['Sheet']
sh.title = 'Исключения'
# поиск повторов по пустым строкам
def find_N2():
    delete_a = []
    delete_b = []
    q = 0
    for i in str:
        for j in Pustie_stroki:
            if i == j:
                delete_a.append(str[q])
                delete_b.append(kol[q])
               # print(f'{delete_a[w]}    {delete_b[w]} \n')
                spisok_index_pust_str.append(q)
        q += 1
    sh.cell(row=1, column=1).value = 'Код дирекции (код по заявке)'
    sh.cell(row=1, column=2).value = 'МКБ-10'
    r = 2
    for statN in delete_a:
        sh.cell(row=r, column=1).value = statN[0]
        r += 1
    r = 2
    for statN in delete_b:
        sh.cell(row=r, column=2).value = statN[0]
        r += 1
    #sh['A'] = delete_a
    #sh['B'] = delete_b

find_N2()

# удаление повторов
def FindAndDelete2():
    ws = Poisk_w.create_sheet('Итог')

    spisok_index_pust_str.reverse()
    for i in spisok_index_pust_str:
        del str[i]
        del kol[i]
    r = 1
    for statN in str:
        ws.cell(row=r, column=1).value = statN[0]
        r += 1
    r = 1
    for statN in kol:
        ws.cell(row=r, column=2).value = statN[0]
        r += 1
    #for z in range(0,len(str)):
    #    print(f'{str[z]}    {kol[z]} \n')
FindAndDelete2()

wz = Poisk_w.create_sheet('Пустые МКБ')
wz.cell(row=1, column=1).value = 'Код дирекции (код по заявке)'
wz.cell(row=1, column=2).value = 'МКБ-10'
r = 2
for statN in Pustie_stroki:
    wz.cell(row=r, column=1).value = statN[0]
    r += 1


Poisk_w.save(filename='itog.xlsx')