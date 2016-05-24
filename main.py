#!/usr/bin/env python

from openpyxl import load_workbook
import sys, datetime

CURR_STORE=""
CURR_ROW=0

def datetype(val):
    return type(val) is datetime.datetime or type(val) is datetime.date

def print_row(row):
    global CURR_ROW

    if CURR_ROW > 14:
        return

    CURR_ROW = CURR_ROW + 1
    for cell in row:
        value = cell.value
#        if not isinstance(value, unicode):
#            value = unicode(value, encoding='utf-8', errors='replace')
        print value, type(value), repr(value)
    print

def process_row(row):
    global CURR_STORE, CURR_ROW

    if CURR_ROW > 14:
        #return
        pass

    a = row[0].value

    if a != CURR_STORE and not datetype(a):
        CURR_STORE = a
        return

    if datetype(a):
        CURR_ROW = CURR_ROW + 1
        print CURR_STORE, ",", ",".join([str(cell.value) for cell in row])


def main(fn="data/sample.xlsx"):
    wb = load_workbook(filename=fn, read_only=True)
    ws = wb['Report'] # ws is now an IterableWorksheet

    for row in ws.rows:
        process_row(row)
        #print_row(row)

if __name__ == '__main__':
    if len(sys.argv) != 2:
        print "usage: {} filename".format(sys.argv[0])
    else:
        main(sys.argv[1])
