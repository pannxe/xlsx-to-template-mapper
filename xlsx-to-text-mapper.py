import argparse
import xlrd


def phase_excel(exc, tmp, opt):
    excf = xlrd.open_workbook(exc)
    sheet = excf.sheet_by_index(0)
    with open(tmp) as f:
        template = f.read()

    out = ""

    for i in range(1, sheet.nrows):
        s = template
        for j in range(sheet.ncols):
            s = s.replace(
                "[{}]".format(str(sheet.cell_value(0, j))), str(sheet.cell_value(i, j))
            )
        out += s + "\n\n"

    out = out.replace(".0", "")
    with open(opt, "w") as f:
        f.write(out)


def main():
    parser = argparse.ArgumentParser()
    parser.add_argument(
        "--excel", "-x", help="path to excel file", default="./data.xlsx"
    )
    parser.add_argument(
        "--template", "-t", help="path to template file", default="./template.txt"
    )
    parser.add_argument(
        "--output", "-o", help="path to output file", default="./output.txt"
    )
    args = parser.parse_args()
    phase_excel(args.excel, args.template, args.output)


if __name__ == "__main__":
    main()
