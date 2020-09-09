#!/usr/bin/python
# coding=UTF-8

__author__ = 'Kingson'

import xlrd
import sys
# reload(sys)
# sys.setdefaultencoding('utf-8')

def exportStrings(sourceFile, sourceSheet, saveFile, fromRow = 0):
    data = xlrd.open_workbook(sourceFile)
    table = data.sheet_by_index(int(sourceSheet))

    if (int(fromRow) == 0):
        for i in range(table.ncols):
            for y in range(table.nrows):
                if y == 0:
                    header = str(table.col(i)[y].value) + "\n"
                    with open(saveFile, "a") as f:
                        f.write(header)
                else:
                    base = str(table.col(0)[y].value)
                    value = str(table.col(i)[y].value)
                    if base != "" and value != "":
                        content = '"' + base + '"' + " = " + '"' + value + '"' + ";" + "\n"
                        with open(saveFile, "a") as f:
                            f.write(content)
    else:
        for i in range(table.ncols):
            header = str(table.col(i)[0].value) + "\n"
            with open(saveFile, "a") as f:
                f.write(header)
            for y in range(int(fromRow) - 1, table.nrows):
                base = str(table.col(0)[y].value)
                value = str(table.col(i)[y].value)
                if base != "" and value != "":
                    content = '"' + base + '"' + " = " + '"' + value + '"' + ";" + "\n"
                    with open(saveFile, "a") as f:
                        f.write(content)

open_file = sys.argv[1]
open_sheet = sys.argv[2]
save_file = sys.argv[3]
from_row = sys.argv[4]

exportStrings(open_file, open_sheet, save_file, from_row)
