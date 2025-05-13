#!/usr/bin/env python3

# File: fn2txt.py
# Version: 1.0 2025-05-13
# Copyright 2024-2025 Jasper Habicht (mail(at)jasperhabicht.de).
#
# This work may be distributed and/or modified under the
# conditions of the LaTeX Project Public License version 1.3c,
# available at http://www.latex-project.org/lppl/.
#
# This file is part of the `zchinr' package (The Work in LPPL)
# and all files in that bundle must be distributed together.
#
# This work has the LPPL maintenance status `maintained'.

'''Extract all footnotes from a .tex file and store in a .txt file.'''
import re
import sys

out_txt = ''

if __name__ == '__main__':
    file_in = 'input.tex'
    if len(sys.argv) > 1:
        file_in = sys.argv[1]

    with open(file_in, 'r', encoding='utf-8') as f:
        file_data = f.read()

    file_data = re.sub(r'\\(zhs|emph|textbf)\{(.*?)\}', r'\2', file_data)
    file_data = re.sub(r'\\url\{(.*?)\}', r'<\1>', file_data)
    file_data = re.sub(r'(\\,|~)', r' ', file_data)

    file_data = re.sub(r'---', '\u2014', file_data)
    file_data = re.sub(r'--', '\u2013', file_data)
    file_data = re.sub(r"''", '\u201D', file_data)
    file_data = re.sub(r"'", '\u2019', file_data)
    file_data = re.sub(r'``', '\u201C', file_data)
    file_data = re.sub(r'`', '\u2018', file_data)
    file_data = re.sub(r',,', '\u201E', file_data)

    list_fns = re.findall(r'\\footnote\{([^\}]+?)\}', file_data)

    for fn in list_fns:
        out_txt += str(fn) + '\n\n'

    # write file
    file_out = 'output.txt'
    if len(sys.argv) > 2:
        file_out = sys.argv[2]
    with open(file_out, 'w', encoding='utf-8') as file:
        file.write(out_txt)
