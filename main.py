#!/usr/bin/env python

from openpyxl import load_workbook
import sys, datetime

CURR_STORE=""

def process_row(row):
    global CURR_STORE

    a = row[0].value
    if a != CURR_STORE and type(a) is not datetime.datetime:
        CURR_STORE = a
        return

    if type(a) is datetime.datetime:
        print CURR_STORE, ",", ",".join([str(cell.value) for cell in row])


def main(fn="data/sample.xlsx"):
    wb = load_workbook(filename=fn, read_only=True)
    ws = wb['Sheet1'] # ws is now an IterableWorksheet

    for row in ws.rows:
        process_row(row)

if __name__ == '__main__':
    if len(sys.argv) != 2:
        print "usage: {} filename".format(sys.argv[0])
    else:
        main(sys.argv[1])
