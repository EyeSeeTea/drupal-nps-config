#!/usr/bin/env python3

"""
List the contents of the given column.
"""

import sys
import openpyxl


def main():
    if not 2 <= len(sys.argv) <= 3:
        sys.exit('usage: %s <column> [<file>]' % sys.argv[0])

    col = sys.argv[1]
    fname = (sys.argv[2] if len(sys.argv) >= 3 else
             'ECDD1950_Substances considered (2022_11_21)TLR.xlsx')

    column = openpyxl.load_workbook(fname)['Full Sheet'][col]

    for cell in column:
        value = cell.value
        hl = cell.hyperlink.target if cell.hyperlink else ''
        print(repr(value) + ((' -> ' + remove_query_string(hl)) if hl else ''))


def remove_query_string(link):
    # 'CriticalReview_5FPB22.pdf?ua=1' -> 'CriticalReview_5FPB22.pdf'
    qs_pos = link.rfind('?')
    return link[:qs_pos] if qs_pos > 0 else link



if __name__ == '__main__':
    main()
