# -*- coding:utf-8 -*-

__author__ = 'LIUHUAN'

import xlrd
import csv
import sys
reload(sys)
sys.setdefaultencoding( "utf-8" )


def csv_from_excel():
    wb = xlrd.open_workbook('test.xls')
    sh = wb.sheet_by_name('Sheet1')
    your_csv_file = open('your_csv_file.csv', 'wb')
    wr = csv.writer(your_csv_file, quoting=csv.QUOTE_ALL)

    for rownum in xrange(sh.nrows):
        wr.writerow(sh.row_values(rownum))

    your_csv_file.close()


def csv_to_dic():
    with open('your_csv_file.csv') as f:
        f_csv = csv.DictReader(f)
        import json
        for row in f_csv:
            print json.dumps(row, ensure_ascii=False)


if __name__ == '__main__':
    csv_from_excel()
    csv_to_dic()
