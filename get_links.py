#!/usr/bin/env python3

"""
Print the links to download from the spreadsheet with all the documents
from the ECDD repository.
"""

# In an email, Thomas Le Ruez (letho@who.int) says:
#
#   There are still some aspects to update, such as the drug classes (in
#   columns D, E and F) which might result in some changes as it is not
#   a simple answer.
#
#   The documents to be downloaded are found in column L - Pre-review /
#   critical review report. I have filtered this column to remove the
#   blanks.
#
# But it seems they want *all* the pdfs, from columns L to V.

import openpyxl


def main():
    fname = 'ECDD1950_Substances considered (2022_11_21)TLR.xlsx'

    sheet = openpyxl.load_workbook(fname)['Full Sheet']

    #print_trs_where_review_exists(sheet)  # if we want only those TRSs

    print_all_links(sheet)


def print_all_links(sheet):
    for row in sheet:
        for c in 'LMNOPQRSTUV':
            print_link(row, c)


def print_link(row, c):
    "Print the link at a given column (if any)"
    hl = row[ord(c.upper()) - ord('A')].hyperlink
    if hl:
        link = remove_query_string(hl.target)
        print(link.replace('www.who.int', 'origin.who.int'))


def print_trs_where_review_exists(sheet):
    # L -> Pre-review / critical review report
    # V -> Final meeting report (WHO Technical Report)
    for col_L, col_V in zip(sheet['L'], sheet['V']):
        if col_L.value and col_V.hyperlink:
            link = remove_query_string(col_V.hyperlink.target)

            print(link.replace('www.who.int', 'origin.who.int'))
            # Or to just see the file name:
            #   print('%-30s %s' % (col_V.value, link.split('/')[-1]))


def remove_query_string(link):
    # 'CriticalReview_5FPB22.pdf?ua=1' -> 'CriticalReview_5FPB22.pdf'
    qs_pos = link.rfind('?')
    return link[:qs_pos] if qs_pos > 0 else link



if __name__ == '__main__':
    main()
