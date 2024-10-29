#!/usr/bin/env python3

# File: doc2tex.py
# Version: 1.6 2024-10-23
# Copyright 2024 Jasper Habicht (mail(at)jasperhabicht.de).
#
# This work may be distributed and/or modified under the
# conditions of the LaTeX Project Public License version 1.3c,
# available at http://www.latex-project.org/lppl/.
#
# This file is part of the `zchinr' package (The Work in LPPL)
# and all files in that bundle must be distributed together.
#
# This work has the LPPL maintenance status `maintained'.

'''Convert a .docx file with an ZChinR article to its semantical .tex equivalent.'''
import re
import sys
import zipfile

bold_styles = []
italic_styles = []
section_styles = {}
list_styles = {}

def get_xmlpart(parent, part):
    '''Extract XML parts from .docx file.'''
    with zipfile.ZipFile(parent, 'r') as archive:
        if zipfile.Path(archive, f'word/{part}.xml').is_file():
            with archive.open(f'word/{part}.xml', 'r') as part_file:
                part_file_data = part_file.read()
            return part_file_data.decode('utf-8')
    return False

# process structure
def process_structure(file_name):
    '''Extract styles and footnotes and process docment.'''
    # read file parts
    document_data = get_xmlpart(file_name, 'document')
    footnotes_data = get_xmlpart(file_name, 'footnotes')
    styles_data = get_xmlpart(file_name, 'styles')
    count = { 
        'bold': 0,
        'italic': 0,
        'section': 0,
        'list': 0,
        'documentation': 0,
        'documentation row': 0
    }
    # collect styles
    if styles_data:
        style_ids = re.findall(r'<w:style(\s[^>]+)?\sw:styleId="(.*?)"', styles_data)
        # identify styles as bold, italic or level
        for style in style_ids:
            style_properties = re.search(rf'<w:style(\s[^>]+)?\sw:styleId="{re.escape(style[1])}"(\s[^>]+)?>(.*?)<\/w:style>', styles_data)
            if re.search(r'<w:b\/>', style_properties[3]):
                bold_styles.append(style[1])
            if re.search(r'<w:1\/>', style_properties[3]):
                italic_styles.append(style[1])
            if re.search(r'<w:outlineLvl\s.*?\/>', style_properties[3]):
                section_level = re.search(r'<w:outlineLvl(\s[^>]+)?\sw:val="(.*?)"', style_properties[3])
                section_styles[style[1]] = section_level[2]
            if re.search(r'<w:ilvl\s.*?\/>', style_properties[3]):
                list_level = re.search(r'<w:ilvl(\s[^>]+)?\sw:val="(.*?)"', style_properties[3])
                list_styles[style[1]] = list_level[2]
    # collect footnotes
    footnote_nodes = {}
    if footnotes_data:
        footnote_ids = re.findall(r'<w:footnote(\s[^>]+)?\sw:id="(.*?)"', footnotes_data)
        for footnote in footnote_ids:
            footnote_node = re.search(rf'<w:footnote(\s[^>]+)?\sw:id="{re.escape(footnote[1])}"(\s[^>]+)?>.*?<\/w:footnote>', footnotes_data)
            footnote_nodes[footnote[1]] = f'\\footnote\u007b{process_p_nodes(footnote_node[0], count, True)}\u007d'
    print(f'{len(footnote_nodes)} footnotes found.')
    # filter body part
    result = re.search(r'<w:body(\s[^>]+)?>(.*?)<\/w:body>', document_data)[2]
    result = process_nodes(result, count)
    # replace footnotes inline
    if footnote_nodes:
        result = re.sub(r'<w:footnoteReference(\s[^>]+)?\sw:id="(.*?)"(\s[^>]+)?\/>', lambda m: footnote_nodes.get(m.group(2)), result)
    print(f'{count['documentation']} documentations found.')
    print(f'{count['documentation row']} documentation rows found.')
    print(f'{count['bold']} bold found.')
    print(f'{count['italic']} italic found.')
    print(f'{count['section']} section found.')
    print(f'{count['list']} list found.')
    return result

def select_level(level):
    '''Select hierarchy level of p nodes.'''
    if level == '0':
        append_before = '\\section{'
    elif level == '1':
        append_before = '\\subsection{'
    elif level == '2':
        append_before = '\\subsubsection{'
    elif level == '3':
        append_before = '\\paragraph{'
    else:
        append_before = '\\subparagraph{'
    return append_before

# process nodes
def process_p_nodes(data, count, ignore_footnotes = False):
    '''Find all w:p nodes in a string and process them accordingly. Consider headers, bold and italic.'''
    result = ''
    # process w:p nodes
    paragraphs = re.findall(r'(<w:p(\s[^>]+)?>.*?<\/w:p>)', data)
    for p in paragraphs:
        paragraph = ''
        # process node style
        append_before_p = ''
        append_after_p = ''
        p_style = re.search(r'<w:pStyle(\s[^>]+)?\sw:val="(.*?)"', p[0])
        is_section = False
        if p_style:
            if p_style[2] in section_styles:
                is_section = True
                append_before_p = select_level(section_styles[p_style[2]])
                append_after_p = '}'
                count['section'] += 1
            if p_style[2] in bold_styles and is_section is False:
                append_before_p += '\\textbf{'
                append_after_p += '}'
                count['bold'] += 1
            if p_style[2] in italic_styles:
                append_before_p += '\\emph{'
                append_after_p += '}'
                count['italic'] += 1
            if p_style[2] in list_styles and is_section is False:
                append_before_p = '<zchinr:item>' + append_before_p
                append_after_p += '</zchinr:item>'
                count['list'] += 1
        # process node properties
        p_properties = re.search(r'<w:pPr(\s[^>]+)?>(.*?)<\/w:pPr>', p[0])
        if p_properties:
            if re.search(r'<w:outlineLvl\s.*?\/>', p_properties[2]):
                is_section = True
                level = re.search(r'<w:outlineLvl(\s[^>]+)?\sw:val="(.*?)"', p_properties[2])
                append_before_p = select_level(level)
                append_after_p = '}'
                count['section'] += 1
            if re.search(r'<w:b\/>', p_properties[2]) and is_section is False:
                append_before_p += '\\textbf{'
                append_after_p += '}'
                count['bold'] += 1
            if re.search(r'<w:i\/>', p_properties[2]):
                append_before_p += '\\emph{'
                append_after_p += '}'
                count['italic'] += 1
            if re.search(r'<w:ilvl\s.*?\/>', p_properties[2]) and is_section is False:
                append_before_p = '<zchinr:item>' + append_before_p
                append_after_p += '</zchinr:item>'
                count['list'] += 1
        # process w:r nodes
        runs = re.findall(r'(<w:r(\s[^>]+)?>.*?<\/w:r>)', p[0])
        for r in runs:
            run = ''
            # process node style
            append_before_r = ''
            append_after_r = ''
            r_style = re.search(r'<w:r_style(\s[^>]+)?\sw:val="(.*?)"', r[0])
            if r_style:
                if r_style[2] in bold_styles and is_section is False:
                    append_before_r += '\\textbf{'
                    append_after_r += '}'
                    count['bold'] += 1
                if r_style[2] in italic_styles:
                    append_before_r += '\\emph{'
                    append_after_r += '}'
                    count['italic'] += 1
            # process node properties
            r_properties = re.search(r'<w:rPr(\s[^>]+)?>(.*?)<\/w:rPr>', r[0])
            if r_properties:
                if re.search(r'<w:b\/>', r_properties[2]) and is_section is False:
                    append_before_r += '\\textbf{'
                    append_after_r += '}'
                    count['bold'] += 1
                if re.search(r'<w:i\/>', r_properties[2]):
                    append_before_r += '\\emph{'
                    append_after_r += '}'
                    count['italic'] += 1
        # process w:t nodes
            if ignore_footnotes:
                texts = re.findall(r'(<w:t(\s[^>]+)?>(.*?)<\/w:t>)', r[0])
            else:
                texts = re.findall(r'(<w:t(\s[^>]+)?>(.*?)<\/w:t>|<w:footnoteReference(\s[^>]+)?\/>)', r[0])
            for t in texts:
                if re.search(r'<w:t(\s[^>]+)?>.*?<\/w:t>', t[0]):
                    run += t[2]
                else:
                    run += t[0]
            paragraph += append_before_r + run + append_after_r
        result += append_before_p + paragraph + append_after_p + '\n\n'
    return result

def process_tbl_nodes(data, count):
    '''Find all w:tbl nodes in a string and process them accordingly. Output as documentation environment.'''
    result = ''
    tables = re.findall(r'(<w:tbl(\s[^>]+)?>.*?<\/w:tbl>)', data)
    for tbl in tables:
        result += '\n\n\\begin{documentation}\n'
        rows = re.findall(r'(<w:tr(\s[^>]+)?>.*?<\/w:tr>)', tbl[0])
        for tr in rows:
            cells = re.findall(r'(<w:tc(\s[^>]+)?>.*?<\/w:tc>)', tr[0])
            result += process_p_nodes(cells[0][0], count)
            for tc in cells[1:]:
                result += '<zchinr:cellsep/>' + process_p_nodes(tc[0], count)
            result += '<zchinr:rowsep/>'
            count['documentation row'] += 1
        result += '\\end{documentation}\n\n'
        count['documentation'] += 1
    return result

def process_nodes(data, count):
    '''Find all w:p and w:tbl nodes in a string and process them accordingly. Remove empty w:p nodes first.'''
    result = ''
    data = re.sub(r'<w:p(\s[^>]+)?\/>', r'', data)
    nodes = re.findall(r'(<w:tbl(\s[^>]+)?>.*?<\/w:tbl>|<w:p(\s[^>]+)?>.*?<\/w:p>)', data)
    for node in nodes:
        if re.search(r'<w:tbl(\s[^>]+)?>.*?<\/w:tbl>', node[0]):
            result += process_tbl_nodes(node[0], count)
        else:
            result += process_p_nodes(node[0], count)
    return result

def replace_endash(string):
    '''Replace dashes between numbers, but only if there is only one.'''
    if len(re.findall(r'\-', string)) > 1:
        return string
    return re.sub(r'(\d)\-(\d)', r'\1--\2', string)

def reduce_emph(string):
    '''Join subsequent emph commands.'''
    if re.search(r'\\emph\{.*?\}[ \t\f]*\\emph\{.*?\}', string):
        return reduce_emph(re.sub(r'\\emph\{(.*?)\}([ \t\f]*)\\emph\{(.*?)\}', r'\\emph{\1\2\3}', string))
    return string

def reduce_textbf(string):
    '''Join subsequent textbf commands.'''
    if re.search(r'\\textbf\{.*?\}[ \t\f]*\\textbf\{.*?\}', string):
        return reduce_textbf(re.sub(r'\\textbf\{(.*?)\}([ \t\f]*)\\textbf\{(.*?)\}', r'\\textbf{\1\2\3}', string))
    return string

if __name__ == '__main__':
    file_in = 'input.docx'
    if len(sys.argv) > 1:
        file_in = sys.argv[1]

    # process file data
    file_data = process_structure(file_in)

    # escape ampersand, less than, greater than, number sigh, dollar and percent
    file_data = re.sub(r'&amp;', '&', file_data)
    file_data = re.sub(r'&lt;', '<', file_data)
    file_data = re.sub(r'&gt;', '>', file_data)
    file_data = re.sub(r'\$', r'\\$', file_data)
    file_data = re.sub(r'#', r'\\#', file_data)
    file_data = re.sub(r'&', r'\\&', file_data)
    file_data = re.sub(r'%', r'\\%', file_data)

    # format urls
    file_data = re.sub(r'<?((www\.[-a-zA-Z\d]+\.[^\s]+\/|http:\/\/|https:\/\/)[^\s>]+)>?', r'\\url{\1}', file_data)

    # tidy up empty, nested and subsequent macros
    file_data = re.sub(r'\\emph\{\\emph\{(.*?)\}(\s*)\}', r'\\emph{\1\2}', file_data)
    file_data = re.sub(r'\\textbf\{\\textbf\{(.*?)\}(\s*)\}', r'\\textbf{\1\2}', file_data)
    file_data = re.sub(r'\\(emph|textbf)\{\\footnote\{(.*?)\}(\s*)\}', r'\\footnote{\2}\3', file_data)
    file_data = reduce_emph(file_data)
    file_data = reduce_textbf(file_data)
    file_data = re.sub(r'\\(emph|textbf)\{\s*\}', '', file_data)
    file_data = re.sub(r'\\(emph|textbf)\{(.*?)\s+\}', r'\\\1{\2} ', file_data)
    file_data = re.sub(r'\\(emph|textbf)\{\s+(.*?)\}', r' \\\1{\2}', file_data)
    file_data = re.sub(r'\\footnote\{(.*?)\s+\}', r'\\footnote{\1}', file_data)
    file_data = re.sub(r'\\footnote\{\s+(.*?)\}', r'\\footnote{\1}', file_data)

    # replace endash
    file_data = re.sub(r'[\d\-]+', lambda m: replace_endash(m.group()), file_data)
    file_data = re.sub(r'\s\-\s', ' -- ', file_data)

    # add thin space to abbreviations
    file_data = re.sub(r'\b([a-zA-Z])\.([a-zA-Z])\.([a-zA-Z])\.', r'\1.\\,\2.\\,\3.', file_data)
    file_data = re.sub(r'\b([a-zA-Z])\.([a-zA-Z]{1,2})\.', r'\1.\\,\2.', file_data)
    file_data = re.sub(r'(\d)%', r'\1\\,%', file_data)

    # add non-breakable space
    file_data = re.sub(r'(§§?|Artt?\.|Abs\.|Bd\.|Vol\.|S\.|pp?\.|Nr\.|No\.|Fn\.|Rn\.|Sec\.|sec\.|lit\.)\s(\d+)\s(ff?\.)', r'\1~\2~\3', file_data)
    file_data = re.sub(r'(§§?|Artt?\.|Abs\.|Bd\.|Vol\.|S\.|pp?\.|Nr\.|No\.|Fn\.|Rn\.|Sec\.|sec\.|lit\.)\s(\d+)', r'\1~\2', file_data)

    # replace row and cell separators
    file_data = re.sub(r'\s+<zchinr:cellsep\/>', r' & \n', file_data)
    file_data = re.sub(r'\s+<zchinr:rowsep\/>', r' \\\\ \n\n', file_data)

    # replace lists with items
    file_data = re.sub(r'((<zchinr:item>.*?<\/zchinr:item>\s*)+)', r'\\begin{itemize}\n\1\\end{itemize}\n\n', file_data)
    file_data = re.sub(r'<zchinr:item>(.*?)<\/zchinr:item>\n*', r'\\item \1\n', file_data)

    # process typography
    file_data = re.sub(r'\u00a0', '~', file_data)
    file_data = re.sub(r'\u201c\u2018', '``{}`', file_data)
    file_data = re.sub(r'\u2018\u201c', '`{}``', file_data)
    file_data = re.sub(r'\u201d\u2019', "''{}'", file_data)
    file_data = re.sub(r'\u2019\u201d', "'{}''", file_data)
    file_data = re.sub(r'\u201e\u201a', ',,{},', file_data)
    file_data = re.sub(r'\u201a\u201e', ',{},,', file_data)
    file_data = re.sub(r'\u201c', '``', file_data)
    file_data = re.sub(r'\u201d', "''", file_data)
    file_data = re.sub(r'\u201e', ',,', file_data)
    file_data = re.sub(r'\u2018', '`', file_data)
    file_data = re.sub(r'\u2019', "'", file_data)
    file_data = re.sub(r'\u201a', ',', file_data)
    file_data = re.sub(r'\u2026', r'\\ldots{}', file_data)
    file_data = re.sub(r'\.\.\.', r'\\ldots{}', file_data)
    file_data = re.sub(r'\u2013', '--', file_data)
    file_data = re.sub(r'\u2014', '---', file_data)
    file_data = re.sub(r'!`', '!{}`', file_data)
    file_data = re.sub(r'\?`', '?{}`', file_data)

    # process spaces
    file_data = re.sub(r'~\n', '\n', file_data)
    file_data = re.sub(r'\n{3,}', '\n\n', file_data)
    file_data = re.sub(r'~[ \t\f]', '~', file_data)
    file_data = re.sub(r'[ \t\f]~', '~', file_data)

    # process chinese characters
    file_data = re.sub(r'([\u3000-\u303F\u4e00-\u9fff\uFF00-\uFFEF]+)', r'\\zhs{\1}', file_data)

    # footnote references
    # currently not supported

    # write file
    file_out = 'output.tex'
    if len(sys.argv) > 2:
        file_out = sys.argv[2]
    with open(file_out, 'w', encoding='utf-8') as file:
        file.write(file_data)
