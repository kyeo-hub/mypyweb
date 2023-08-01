import pandas as pd
import openpyxl as op

df = pd.read_excel("wrk.xlsx")
table = df.to_dict(orient='records')
# print(table)


def write(biaoqian):
    bg = op.load_workbook(r"wrk.xlsx")
    sheet = bg["tushi"]
    for i in range(1,len(biaoqian)):
        print(biaoqian[i])
        sheet.cell(int(biaoqian[i]['located_row/所在行']),int(biaoqian[i]['located_column/所在列']),biaoqian[i]['code/捆包号'])
    bg.save(r"wrk.xlsx")
write(table)