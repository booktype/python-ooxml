# -*- coding: utf-8 -*-

import six
import collections

from .doc import Paragraph

def is_header(doc, name):

    if name == '':
        return True

    default_font_size = 0

    if doc.default_style:
        if 'sz' in doc.default_style.rpr:
            default_font_size = int(doc.default_style.rpr['sz']) / 2

    for style_id in doc.used_styles:
        style = doc.styles.get_by_id(style_id)

        if hasattr(style, 'rpr') and 'sz' in style.rpr:
            font_size = int(style.rpr['sz']) / 2

            if font_size <= default_font_size:
                continue

            if style.style_id == name:
                return True

    return False


def find_important(doc, headers):

    default_font_size = 0

    if doc.default_style:
        if 'sz' in doc.default_style.rpr:
            default_font_size = int(doc.default_style.rpr['sz']) / 2

    def _big_enough(block):
        # This is something which should be increased
        return block['weight'] > 50

    def _filter_font_sizes(sizes):
        for sz, value in sizes:
            if sz > default_font_size:
                yield (sz, value)

        return 

    most_common_fonts = doc.used_font_size.most_common()

    if len(most_common_fonts) > 0:
        most_common = most_common_fonts[0][0]
    else:
        most_common = -1

    # remove from the list font size which is mostly used
    # this could result in a lot of errors
    _list_of_fonts = [fnt for fnt in doc.used_font_size.items() if fnt[0] != most_common] 
    HEADERS_IMPORTANCE = [[el] for el in reversed(collections.OrderedDict(sorted(_filter_font_sizes(_list_of_fonts), key=lambda t: t[0])))]

    for style_id in doc.used_styles:
        style = doc.styles.get_by_id(style_id)

        if hasattr(style, 'rpr') and 'sz' in style.rpr:
            font_size = int(style.rpr['sz']) / 2

            if font_size <= default_font_size:
                continue

            for i in range(len(HEADERS_IMPORTANCE)):
                if HEADERS_IMPORTANCE[i][0] < font_size:
                    if i == 0:
                        HEADERS_IMPORTANCE = [[font_size, style_id]]+HEADERS_IMPORTANCE
                    else:
                        HEADERS_IMPORTANCE[i-1].append(style_id)
                        break
                elif HEADERS_IMPORTANCE[i][0] == font_size:
                    HEADERS_IMPORTANCE[i].append(style_id)

    for hdrs in HEADERS_IMPORTANCE:
        lst = []
        for header in headers:
            if header['name'] == '' and ('font_size' in header and header['font_size'] == hdrs[0]):
                lst.append({'name': '', 'weight': header['weight'], 'index': header['index'], 'font_size': header['font_size']})
            elif header['name'] != '' and header['name'] in hdrs[1:]:
                lst.append({'name': header['name'], 'weight': header['weight'], 'index': header['index']})                
            else:
                if len(lst) > 0:
                    lst[-1]['weight'] += header['weight']

        if len(lst) > 1:
            big_blocks = [el for el in lst if _big_enough(el)]
            if len(big_blocks) > 0:
                return lst

    return None 


def mark_headers(doc, markers):
    selected = []

    for style in markers:
        if is_header(doc, style['name']):
            if style['name'] == '':
                selected.append({'name': '', 'index': style['index'], 'weight': style['weight'], 'font_size': style['font_size']})
            else:
                selected.append({'name': style['name'], 'index': style['index'], 'weight': style['weight']})
        else:
            if len(selected) > 0:
                selected[-1]['weight'] += style['weight']

    return selected


def mark_styles(doc, elements):
    markers = [{'name': '', 'weight': 0, 'index': 0, 'font_size': 0}]

    for pos, elem in enumerate(elements):
        try:
            style = doc.styles.get_by_id(elem.style_id)
        except AttributeError:
            style = None

        if isinstance(elem, Paragraph):
            if elem.is_dropcap():
                continue

        weight = 0

        if hasattr(elem, 'elements'):
            try:
                content = ''.join([e.value() for e in elem.elements])

                if len(content.strip()) == 0:
                    continue

                weight = len(content)
            except:
                pass

        font_size = -1
        if style:
            markers.append({'name': elem.style_id, 'weight': weight, 'index': pos})
        else:
            if hasattr(elem, 'rpr') and 'sz' in elem.rpr:
                markers.append({'name': '', 'weight': weight, 'index': pos, 'font_size': int(elem.rpr['sz'])/2})
                font_size = int(elem.rpr['sz'])/2

        if hasattr(elem, 'elements'):
            for e in elem.elements:
                # TODO
                # check if this is empty element
                if hasattr(e, 'rpr') and 'sz' in e.rpr:
                    fnt_size = int(e.rpr['sz'])/2                        
                    if fnt_size != font_size:
                        markers.append({'name': '', 'weight': weight, 'index': pos, 'font_size': fnt_size})

    return markers


def split_document(doc):
    markers = mark_styles(doc, doc.elements)
    headers = mark_headers(doc, markers)
    important = find_important(doc, headers)

    return important


def parse_html_string(s):
    from lxml import html

    utf8_parser = html.HTMLParser(encoding='utf-8')
    html_tree = html.document_fromstring(s , parser=utf8_parser)

    return html_tree


def get_chapters(doc):
    from lxml import etree
    from . import serialize

    def _serialize_chapter(els):
        s =  serialize.serialize_elements(doc, els)

        if s.startswith(six.b('<div/>')):
            return ('', six.b('<body></body>'))

        root = parse_html_string(s[5:-6])
        body = root.find('.//body')

        if body[0].tag in ['h1', 'h2', 'h3', 'h4', 'h5', 'h6']:            
            body[0].tag = 'h1'
            # get text content of first header
            chapter_title = body[0].text_content().strip()
            # clears it up and set new content
            # this is when we have different html tags in the header
            # this also clears attributes
            body[0].clear()
            body[0].text = chapter_title
        else:
            title =  etree.Element('h1')
            title.text = 'Unknown'
            chapter_title = ''

            body.insert(0, title)

        return (chapter_title, etree.tostring(body, pretty_print=True, encoding="utf-8", xml_declaration=False))

    chapters = split_document(doc)

    export_chapters = []

    if chapters:
    # first, everything before the first chapter
        if len(chapters) > 0:
            if chapters[0]['index'] != 0:
                export_chapters.append(_serialize_chapter(doc.elements[:chapters[0]['index']-1]))

            if len(chapters) > 1 and chapters[0]['index'] == chapters[1]['index']:
                chapters = chapters[1:]

        for n in range(len(chapters)-1):
            if chapters[n]['index'] == chapters[n+1]['index']-1:
                _html = _serialize_chapter([doc.elements[chapters[n]['index']]])
            else:
                _html = _serialize_chapter(doc.elements[chapters[n]['index']:chapters[n+1]['index']-1])

            export_chapters.append(_html)

        export_chapters.append(_serialize_chapter(doc.elements[chapters[-1]['index']:]))
    else:
        export_chapters.append(_serialize_chapter(doc.elements))

    return export_chapters