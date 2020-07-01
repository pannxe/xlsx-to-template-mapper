#!/usr/bin/python3

import xlrd


def parse_excel(exc, tmp):
    excf = xlrd.open_workbook(exc)
    sheet = excf.sheet_by_index(0)

    out = ""

    for i in range(1, sheet.nrows):
        s = tmp
        for j in range(sheet.ncols):
            rp = str(sheet.cell_value(i, j))
            if rp.endswith(".0"):
                rp = rp[:len(rp)-2]
            s = s.replace(
                f"[{str(sheet.cell_value(0, j))}]", rp
            )
        out += s + "\n"
    return out
