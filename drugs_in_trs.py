#!/usr/bin/env python3

"""
Get drug information from a WHO TRS report.
"""

# This program is used with the text file that comes from something like:
#   pdftotext -layout -nopgbrk 9789240001848-eng.pdf
# and then editing that file to leave only the relevant parts, well formatted.

import sys
import re
from collections import OrderedDict


def drugs_in_trs(trs):
    "Return a dict with all the sections (and contents) for every drug"
    # The idea is to use this function from other programs that use
    # this file as a module.
    data = extract_categories(trs)
    drugs = {}
    for d in data.values():
        drugs.update(d)
    return drugs


def extract_categories(fname):
    "Return dict of drug categories and their content (dicts too)"
    categories = OrderedDict()

    txt = open(fname).read()

    parts = re.split(r'\n\d+\.\d+\s+(.*)', txt)

    if len(parts) == 1:  # no categories
        categories[''] = extract_drugs(parts[0])
        return categories

    parts = parts[1:]

    for i in range(len(parts) // 2):
        name = parts[2*i].strip()
        categories[name] = extract_drugs('\n' + parts[2*i+1].strip())
        # we add the '\n' so we can use the same safe pattern to identify
        # drug names later on with r'\n\d+\.\d+\.\d+\s+(.*)'

    return categories


def extract_drugs(txt):
    "Return a dict of drugs and their content (dicts too)"
    parts = re.split(r'\n\d+\.\d+\.\d+\s+(.*)', txt)[1:]

    drugs = OrderedDict()

    for i in range(len(parts) // 2):
        name = parts[2*i].strip()

        section = parts[2*i+1]

        while section.startswith('\n  '):
            pos = section.find('\n', 1)  # next line
            name += ' ' + section[:pos].strip()
            section = section[pos:]

        drugs[name] = extract_sections(section)

    return drugs


def extract_sections(txt):
    "Return dict of sections and their content (strings)"
    sections = OrderedDict()

    for part in txt.split('\n\n'):
        name, content = part.lstrip().split('\n', 1)

        sections[name] = format_paragraphs(content)

    return sections


def format_paragraphs(txt):
    "Format the texts from pdftotext file like normal paragraphs"
    # 'Sentence one. Sentence\n'
    # 'two. All the same paragraph.\n'
    # '    New sentence in a new paragraph.\n'
    # ->
    # 'Sentence one. Sentence two. All the same paragraph\n'
    # 'New sentence in a new paragraph.\n'
    paragraphs = txt.split('\n    ')

    return '\n'.join(p.strip().replace('\n', ' ') for p in paragraphs)


def print_all():
    import sys
    import colors

    if len(sys.argv) != 2:
        sys.exit('usage: %s TEXT_FILE' % sys.argv[0])

    fname = sys.argv[1]

    data = extract_categories(fname)

    # One way of showing what we have:
    for cat, content_cat in data.items():
        print(colors.yellow(cat))

        for drug, content_drug in content_cat.items():
            print(colors.green(drug))

            for section, content_section in content_drug.items():
                print(colors.magenta(section))
                #print(content_section)  # uncomment to print contents too


if __name__ == '__main__':
    print_all()
