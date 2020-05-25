#!/usr/bin/python3

import argparse
import x2tm_core 
import os

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

    while True:
        print("\n")
        exists_exc = os.path.exists(args.excel)
        exists_tmp = os.path.exists(args.template)
        if exists_exc and exists_tmp:
            break
        if not exists_exc:
            print(">> Excel file {} not found".format(args.excel))
        if not exists_tmp:
            print(">> Template file {} not found".format(args.template))
        try_again = input("\nDo you want to try again? [y/n] : ").lower()
        if try_again == "n":
            exit(0)
    with open(args.template) as f:
        tmp = f.read()
    out = x2tm_core.phase_excel(args.excel, tmp)
    with open(args.output, "w") as f:
        f.write(out)
    print("Done")


if __name__ == "__main__":
    main()