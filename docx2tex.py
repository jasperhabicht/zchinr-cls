#!/usr/bin/env python3
'''Convert a .docx file with an ZChinR article to its semantical .tex equivalent. Version 1.0: 2024-10.15.'''
import zipfile
import re
import sys

def get_xmlpart(parent, part):
    '''Extract XML parts from .docx file'''
    with zipfile.ZipFile(parent, 'r') as archive:
        with archive.open(f'word/{part}.xml', 'r') as part_file:
            part_filedata = part_file.read()
    return part_filedata.decode('utf-8')

def process_nodes(data):
    '''Find all w:p and w:tbl nodes in a string and process them accordingly. Remove empty w:p nodes first.'''
    data = re.sub(r'<w:p(\s[^>]+)?\/>', r'', data)
    nodes = re.findall(r'(<w:tbl(\s[^>]+)?>.*?<\/w:tbl>|<w:p(\s[^>]+)?>.*?<\/w:p>)', data)
    result = ''
    for node in nodes:
        if re.search(r'<w:tbl(\s[^>]+)?>.*?<\/w:tbl>', node[0]):
            result += process_tbl_nodes(node[0])
        else:
            result += process_p_nodes(node[0])
    return result

def process_tbl_nodes(data):
    '''Find all w:tbl nodes in a string and process them accordingly. Output as documentation environment.'''
    result = ''
    tables = re.findall(r'(<w:tbl(\s[^>]+)?>.*?<\/w:tbl>)', data)
    for tbl in tables:
        result += '\n\n\\begin{documentation}\n'
        rows = re.findall(r'(<w:tr(\s[^>]+)?>.*?<\/w:tr>)', tbl[0])
        for tr in rows:
            cells = re.findall(r'(<w:tc(\s[^>]+)?>.*?<\/w:tc>)', tr[0])
            result += process_p_nodes(cells[0][0])
            for tc in cells[1:]:
                result += ' <x:cellsep/> ' + process_p_nodes(tc[0])
            result += ' // \n'
        result += '\\end{documentation}\n\n'
    return result

# process nodes
def process_p_nodes(data, ignore_footnotes = False):
    '''Find all w:p nodes in a string and process them accordingly. Consider headers, bold and italic.'''
    result = ''
    paragraphs = re.findall(r'(<w:p(\s[^>]+)?>.*?<\/w:p>)', data)
    for p in paragraphs:
        paragraph = ''
        append_before_p = ''
        append_after_p = ''
        p_style = re.search(r'<w:pStyle(\s[^>]+)?\sw:val="(.*?)"', p[0])
        is_section = False
        if p_style:
            if p_style[2] in level_styles:
                is_section = True
                if level_styles[p_style[2]] == '0':
                    append_before_p += '\\section\u007b'
                    append_after_p += '\u007d'
                elif level_styles[p_style[2]] == '1':
                    append_before_p += '\\subsection\u007b'
                    append_after_p += '\u007d'
                elif level_styles[p_style[2]] == '2':
                    append_before_p += '\\subsubsection\u007b'
                    append_after_p += '\u007d'
                elif level_styles[p_style[2]] == '3':
                    append_before_p += '\\paragraph\u007b'
                    append_after_p += '\u007d'
                else:
                    append_before_p += '\\subparagraph\u007b'
                    append_after_p += '\u007d'
            if p_style[2] in bold_styles and is_section is False:
                append_before_p += '\\textbf\u007b'
                append_after_p += '\u007d'
            if p_style[2] in italic_styles:
                append_before_p += '\\emph\u007b'
                append_after_p += '\u007d'
        p_properties = re.search(r'<w:pPr(\s[^>]+)?>(.*?)<\/w:pPr>', p[0])
        if p_properties:
            if re.search(r'<w:outlineLvl(>|\s)', p_properties[2]):
                is_section = True
                level = re.search(r'<w:outlineLvl(\s[^>]+)?\sw:val="(.*?)"', p_properties[2])
                if level == '0':
                    append_before_p += '\\section\u007b'
                    append_after_p += '\u007d'
                elif level == '1':
                    append_before_p += '\\subsection\u007b'
                    append_after_p += '\u007d'
                elif level == '2':
                    append_before_p += '\\subsubsection\u007b'
                    append_after_p += '\u007d'
                elif level == '3':
                    append_before_p += '\\paragraph\u007b'
                    append_after_p += '\u007d'
                else:
                    append_before_p += '\\subparagraph\u007b'
                    append_after_p += '\u007d'
            if re.search(r'<w:b\/>', p_properties[2]) and is_section is False:
                append_before_p += '\\textbf\u007b'
                append_after_p += '\u007d'
            if re.search(r'<w:i\/>', p_properties[2]):
                append_before_p += '\\emph\u007b'
                append_after_p += '\u007d'
        # process runs
        runs = re.findall(r'(<w:r(\s[^>]+)?>.*?<\/w:r>)', p[0])
        for r in runs:
            run = ''
            append_before_r = ''
            append_after_r = ''
            r_style = re.search(r'<w:r_style(\s[^>]+)?\sw:val="(.*?)"', r[0])
            if r_style:
                if r_style[2] in bold_styles and is_section is False:
                    append_before_r += '\\textbf\u007b'
                    append_after_r += '\u007d'
                if r_style[2] in italic_styles:
                    append_before_r += '\\emph\u007b'
                    append_after_r += '\u007d'
            r_properties = re.search(r'<w:rPr(\s[^>]+)?>(.*?)<\/w:rPr>', r[0])
            if r_properties:
                if re.search(r'<w:b\/>', r_properties[2]) and is_section is False:
                    append_before_r += '\\textbf\u007b'
                    append_after_r += '\u007d'
                if re.search(r'<w:i\/>', r_properties[2]):
                    append_before_r += '\\emph\u007b'
                    append_after_r += '\u007d'
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

def replace_endash(string):
    '''Replace dashes between numbers, but only if there is only one.'''
    if len(re.findall(r'\-', string)) > 1:
        return string
    else:
        return re.sub(r'(\d)\-(\d)', r'\1--\2', string)

if __name__ == '__main__':
    if len(sys.argv) > 1:
        filename = sys.argv[1]
    else:
        filename = 'input.docx'

    # read file parts
    document_data = get_xmlpart(filename, 'document')
    footnotes_data = get_xmlpart(filename, 'footnotes')
    styles_data = get_xmlpart(filename, 'styles')

    # collect styles
    style_ids = re.findall(r'<w:style(\s[^>]+)?\sw:styleId="(.*?)"', styles_data)
    bold_styles = []
    italic_styles = []
    level_styles = {}

    # identify styles as bold, italic or level
    for i in style_ids:
        style = re.search(rf'<w:style(\s[^>]+)?\sw:styleId="{re.escape(i[1])}"(\s[^>]+)?>(.*?)<\/w:style>', styles_data)
        if re.search(r'<w:b\/>', style[3]):
            bold_styles.append(i[1])
        if re.search(r'<w:1\/>', style[3]):
            italic_styles.append(i[1])
        if re.search(r'<w:outlineLvl(>|\s)', style[3]):
            section_level = re.search(r'<w:outlineLvl(\s[^>]+)?\sw:val="(.*?)"', style[3])
            level_styles[i[1]] = section_level[2]

    # collect footnotes
    footnote_ids = re.findall(r'<w:footnote(\s[^>]+)?\sw:id="(.*?)"', footnotes_data)
    footnote_nodes = {}
    for i in footnote_ids:
        footnote_node = re.search(rf'<w:footnote(\s[^>]+)?\sw:id="{re.escape(i[1])}"(\s[^>]+)?>.*?<\/w:footnote>', footnotes_data)
        footnote_nodes[i[1]] = f'\\footnote\u007b{process_p_nodes(footnote_node[0], True)}\u007d'

    # filter body part
    filedata = re.search(r'<w:body(\s[^>]+)?>(.*?)<\/w:body>', document_data)[2]
    filedata = process_nodes(filedata)

    # replace footnotes inline
    if footnote_nodes:
        filedata = re.sub(r'<w:footnoteReference(\s[^>]+)?\sw:id="(.*?)"(\s[^>]+)?\/>', lambda m: footnote_nodes.get(m.group(2)), filedata)

    # add thin space to abbreviations
    filedata = re.sub(r'\b([a-zA-Z])\.([a-zA-Z])\.([a-zA-Z])\.', r'\1.\\,\2.\\,\3.', filedata)
    filedata = re.sub(r'\b([a-zA-Z])\.([a-zA-Z]{1,2})\.', r'\1.\\,\2.', filedata)
    filedata = re.sub(r'(\d)%', r'\1\\,%', filedata)

    # replace endash
    filedata = re.sub(r'[\d\-]+', lambda m: replace_endash(m.group()), filedata)
    filedata = re.sub(r'\s-\s', ' -- ', filedata)

    # add non-breakable space
    filedata = re.sub(r'(§§?|Artt?\.|Abs\.|Bd\.|Vol\.|S\.|pp?\.|Nr\.|No\.|Fn\.|Rn\.|Sec\.|sec\.|lit\.)\s(\d+)\s(ff?.)', r'\1~\2~\3', filedata)
    filedata = re.sub(r'(§§?|Artt?\.|Abs\.|Bd\.|Vol\.|S\.|pp?\.|Nr\.|No\.|Fn\.|Rn\.|Sec\.|sec\.|lit\.)\s(\d+)', r'\1~\2', filedata)

    # escape ampersand, less than, greater than, number sigh, dollar and percent
    filedata = re.sub(r'&amp;', '&', filedata)
    filedata = re.sub(r'&lt;', '<', filedata)
    filedata = re.sub(r'&gt;', '>', filedata)
    filedata = re.sub(r'#', r'\\#', filedata)
    filedata = re.sub(r'&', r'\\&', filedata)
    filedata = re.sub(r'\$', r'\\$', filedata)
    filedata = re.sub(r'%', r'\\%', filedata)

    # urls
    filedata = re.sub(r'((www\.[-a-zA-Z\d]+\.[^\s]+\/|http:\/\/|https:\/\/)[^\s]+)', r'\\url{\1}', filedata)

    # tidy up
    filedata = re.sub(r'\\(emph|textbf)\{\s*\}', '', filedata)
    filedata = re.sub(r'\\(emph|textbf)\{\s+(.*?)\}', r' \\\1{\2}', filedata)
    filedata = re.sub(r'\\(emph|textbf)\{(.*?)\s+\}', r'\\\1{\2} ', filedata)
    filedata = re.sub(r'\\(emph|textbf)\{\\footnote\{(.*?)\}\}', r'\\footnote{\2}', filedata)
    filedata = re.sub(r'\\footnote\{\s+(.*?)\}', r'\\footnote{\1}', filedata)
    filedata = re.sub(r'\\footnote\{(.*?)\s+\}', r'\\footnote{\1}', filedata)

    # replace cell separators
    filedata = re.sub(r'<x:cellsep\/>', r'&', filedata)

    # process typography
    filedata = re.sub(r'\u00a0', '~', filedata)
    filedata = re.sub(r'\u201c\u2018', '``{}`', filedata)
    filedata = re.sub(r'\u2018\u201c', '`{}``', filedata)
    filedata = re.sub(r'\u201d\u2019', "''{}'", filedata)
    filedata = re.sub(r'\u2019\u201d', "'{}''", filedata)
    filedata = re.sub(r'\u201e\u201a', ',,{},', filedata)
    filedata = re.sub(r'\u201a\u201e', ',{},,', filedata)
    filedata = re.sub(r'\u201c', '``', filedata)
    filedata = re.sub(r'\u201d', "''", filedata)
    filedata = re.sub(r'\u201e', ',,', filedata)
    filedata = re.sub(r'\u2018', '`', filedata)
    filedata = re.sub(r'\u2019', "'", filedata)
    filedata = re.sub(r'\u201a', ',', filedata)
    filedata = re.sub(r'\u2026', r'\\ldots{}', filedata)
    filedata = re.sub(r'\.\.\.', r'\\ldots{}', filedata)
    filedata = re.sub(r'\u2013', '--', filedata)
    filedata = re.sub(r'\u2014', '---', filedata)

    filedata = re.sub('!`', '!{}`', filedata)
    filedata = re.sub('?`', '?{}`', filedata)

    # process spaces
    filedata = re.sub(r'~\n', '\n', filedata)
    filedata = re.sub(r'\n{3,}', '\n\n', filedata)
    filedata = re.sub(r'~[ \t\f]', '~', filedata)
    filedata = re.sub(r'[ \t\f]~', '~', filedata)

    # process chinese characters
    # filedata = re.sub(r'([\p{Han}\x{3000}-\x{303F}\x{FF00}-\x{FFEF}]+)', r'\\zhs{\1}', filedata)

    # itemize / enumerate : w:ilvl

    # write file
    with open('output.tex', 'w', encoding='utf-8') as file:
        file.write(filedata)
