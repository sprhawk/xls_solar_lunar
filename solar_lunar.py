# -*- coding: utf-8 -*-


from xlrd import open_workbook
from xlwt import Workbook

from lunardate import LunarDate

def process_workbook(workbook):
    s = workbook.sheet_by_index(0)
    print(s.name)

    out_book = Workbook()
    sheet1 = out_book.add_sheet("Sheet 1")

    cells = s.row(0)

    for i in range(3):
        sheet1.write(0, i, cells[i].value)
    sheet1.write(0, 3, u"阴历")
    for i in range(5):
        sheet1.write(0, 4 + i, unicode(2014 + i) + u" 阳历")

    for row in range(1, s.nrows):
        for col in range(3):
            value = s.cell(row, col).value
            sheet1.write(row, col, value)

        value = s.cell(row, 2).value
        valstr = str(int(value))
        y = int(valstr[:4])
        m = int(valstr[4:6])
        d = int(valstr[5:7])
        try:
            lunar_date = LunarDate.fromSolarDate(y, m, d)
            sheet1.write(row, 3, u"{}-{}-{}".format(lunar_date.year, lunar_date.month, lunar_date.day))
            for i in range(5):
                lunar_date.year = 2014 + i
                try:
                    d = lunar_date.toSolarDate()
                    sheet1.write(row, 4 + i, u"{}-{}-{}".format(d.year, d.month, d.day))
                except Exception as e:
                    print(e)
                    print(u"错误的日期:" + valstr)
                    print(lunar_date)
                    sheet1.write(row, 4 + i, u"错误")
        except Exception as e:
            print(e)
            print(u"错误的日期:" + valstr)
            sheet1.write(row, 3, u"错误的日期" + valstr)

    return out_book


if __name__ == "__main__":
    import sys
    if len(sys.argv) < 2:
        print("Usage: {} input.xls [output.xls]".format(sys.argv[0]))
        sys.exit(1)
    workbook =  open_workbook(sys.argv[1], on_demand=True)
    book = process_workbook(workbook)
    if len(sys.argv) < 3:
        output_name = sys.argv[1] + "_output.xls"
    else:
        output_name = sys.argv[2]
    book.save(output_name)
