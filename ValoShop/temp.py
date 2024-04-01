import openpyxl

bd = openpyxl.load_workbook('skins.xlsx')
sbd = bd['skins']
row = 1
column = 'B'
while sbd[column+str(row)].value != None:
    aboba = f'{sbd["B" + str(row)].value}  {sbd["A" + str(row)].value}, '
    f = open('skins.txt', 'a')
    f.write(aboba)
    row+=1
    print(f'{row} |{aboba}')
print('Done')