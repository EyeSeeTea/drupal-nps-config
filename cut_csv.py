#!/usr/bin/env python3

"""
Cut a csv file into pieces.
"""

import sys
import os


class Dumper:
    "Class to dump text into new files named like an original one"

    def __init__(self, fname):
        self.prefix, self.extension = fname.rsplit('.', 1)  # 'a.o' -> 'a', 'o'
        self.nfiles = 0

    def dump(self, txt):
        fname = f'{self.prefix}_{self.nfiles}.{self.extension}'
        assert not os.path.exists(fname), f'file already exists: {fname}'
        print(f'Writing: {fname}')
        with open(fname, 'wt') as fout:
            fout.write(txt)
        self.nfiles += 1


def main():
    if len(sys.argv) < 2:
        sys.exit('usage: %s <csv file> [<rows per file>]' % sys.argv[0])

    fname = sys.argv[1]
    assert fname.endswith('.csv'), f'input file may not be a csv: {fname}'

    if len(sys.argv) < 3:
        cut_per_size(fname)
    else:
        n = int(sys.argv[2])
        cut_per_number(fname, n)


def cut_per_size(fname, size=20e3):
    "Divide the contents of a csv file into files with at most the given size"
    dumper = Dumper(fname)

    fin = open(fname)
    header = fin.readline()

    buffer = header
    while True:
        line = fin.readline()
        if not line:
            break

        if len(buffer) + len(line) < size:  # normal increment
            buffer += line
        elif buffer != header:  # we passed the max size, and we had data
            dumper.dump(buffer)
            buffer = header + line
        else:  # we passed the max size with just one line!
            print(f'In the {i}-th file, line is bigger than {size} bytes!')
            dumper.dump(header + line)
            buffer = header

    if buffer != header:
        dumper.dump(buffer + line)


def cut_per_number(fname, n):
    "Divide the contents of a csv file into files with n rows each"
    prefix = fname.rsplit('.', 1)[0]  # without the extension

    for i, line in enumerate(open(fname)):
        if i == 0:
            header = line
            fout = open(f'{prefix}_0.csv', 'wt')
            fout.write(header)
        elif (i - 1) % n == 0:
            fout = open(f'{prefix}_{i // n}.csv', 'wt')
            fout.write(header)
            fout.write(line)
        else:
            fout.write(line)



if __name__ == '__main__':
    main()
