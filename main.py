#!/usr/bin/env python

from openpyxl import load_workbook
from openpyxl import styles
import os, re, sys, datetime

CURR_STORE=""
CURR_ROW=0

def datetype(val):
    return type(val) is datetime.datetime or type(val) is datetime.date

def print_row(row, wk, yr):
    global CURR_ROW

    if CURR_ROW > 14:
        return

    CURR_ROW = CURR_ROW + 1
    for cell in row:
        value = cell.value
#        if not isinstance(value, unicode):
#            value = unicode(value, encoding='utf-8', errors='replace')
        print yr, wk, value, type(value), repr(value)
    print

def process_row(row, wk, yr):
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
        print "{},{},{},{}".format(CURR_STORE, yr, wk, ",".join([str(cell.value) for cell in row]))


def process_file(fn, wk, yr):
    global CURR_ROW
    CURR_ROW = 0

    wb = load_workbook(filename=fn)
    ws = wb['Report']

    # put back date formats on Column A
    col_a = ws.column_dimensions['A']
    #col_a = ws.range('A1:A{}'.format(ws.get_highest_row()))
    col_a.number_format = 'm/d/yy'
    #wb.save(fn)

    # process rows
    for row in ws.rows:
        process_row(row, wk, yr)
        #print_row(row, wk, yr)

def parse_filename(fn):
    p = re.compile(r'TSC [Ww]eek (?P<wk>[0-9]+) (?P<yr>[0-9]+).xlsx')
    m = p.search(fn)
    return m.group("wk"), m.group("yr")


def main(d="data/"):
    for subdir, dirs, files in os.walk(d):
        for file in files:
            filepath = subdir + os.sep + file
            if filepath.endswith(".xlsx"):
                wk, yr = parse_filename(file)
                process_file(filepath, wk, yr)

if __name__ == '__main__':
    if len(sys.argv) != 2:
        print "usage: {} directory".format(sys.argv[0])
    else:
        main(sys.argv[1])
