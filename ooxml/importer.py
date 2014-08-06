# -*- coding: utf-8 -*-

import six

HEADERS_IMPORTANCE = [
    ('heading 1', 'Title'),
    ('heading 2', 'Subtitle'),
    ('heading 3', ),
    ('heading 4', ),
    ('heading 5', ),
    ('heading 6', ),
    ('heading 7', )
]

def is_header(doc, name):
    if name in ['heading 1', 'heading 2', 'heading 3', 'heading 4', 'heading 5', 'Title', 'Subtitle']:
        return True

    return False


def find_important(headers):

    def _big_enough(block):
        return block['weight'] > 100

    for hdrs in HEADERS_IMPORTANCE:
        lst = []       

        for header in headers:
            if header['name'] in hdrs:
                lst.append({'name': header['name'], 'weight': header['weight'], 'index': header['index']})
            else:
                if len(lst) > 0:
                    lst[-1]['weight'] += header['weight']

        if len(lst) > 1:
            big_blocks = [el for el in lst if _big_enough(el)]

            if len(big_blocks) > 1:
                return lst

    return None 


def mark_headers(doc, markers):
    selected = []

    for style in markers:
        if is_header(doc, style['name']):
            selected.append({'name': style['name'], 'index': style['index'], 'weight': style['weight']})
        else:
            if len(selected) > 0:
                selected[-1]['weight'] += style['weight']

    return selected


def mark_styles(doc, elements):
    markers = [{'name': '', 'weight': 0, 'index': 0}]

    for pos, elem in enumerate(elements):
        try:
            style = doc.styles[elem.style_id]
        except AttributeError:
            style = None

        if style:
            markers.append({'name': style.name, 'weight': 0, 'index': pos})

        try:
            content = ''.join([e.value() for e in elem.elements])
            markers[-1]['weight'] += len(content)
        except:
            # breaks for tables, ...
            pass

    return markers


def split_document(doc):
    markers = mark_styles(doc, doc.elements)

    headers = mark_headers(doc, markers)
    
    important = find_important(headers)

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

        for n in range(len(chapters)-1):
            export_chapters.append(_serialize_chapter(doc.elements[chapters[n]['index']:chapters[n+1]['index']-1]))

        export_chapters.append(_serialize_chapter(doc.elements[chapters[-1]['index']:]))
    else:
        export_chapters.append(_serialize_chapter(doc.elements))

    return export_chapters