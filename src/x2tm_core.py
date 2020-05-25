#!/usr/bin/python3

import xlrd

def phase_excel(exc, tmp):
    excf = xlrd.open_workbook(exc)
    sheet = excf.sheet_by_index(0)

    out = ""

    for i in range(1, sheet.nrows):
        s = tmp
        for j in range(sheet.ncols):
            s = s.replace(
                "[{}]".format(str(sheet.cell_value(0, j))), str(sheet.cell_value(i, j))
            )
        out += s + "\n"

    out = out.replace(".0", "")
    return out



