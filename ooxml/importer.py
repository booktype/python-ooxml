# -*- coding: utf-8 -*-

import six
import collections
import logging

from .doc import (Paragraph, Table, TableCell, Link, TextBox, TOC, Break)

logger = logging.getLogger('ooxml')

POSSIBLE_HEADER_SIZE = 24
BIG_ENOUGH = 100
MAXIMUM_FRONTMATTER_HEADER = 300
MAXIMUM_FRONTMATTER_TIMES = 5
MAXIMUM_EAT_MARKER = 100
MINIMUM_USED_STYLES = 5


def _calculate(doc, elem, style_id):
    weight = 0
    value = elem.value()

    if hasattr(elem, 'style_id'):
        style_id = elem.style_id

    if value:
        if type(value) in [type(u' '), type(' ')]:
            weight += len(value.strip())

            if hasattr(elem, 'rpr') and 'sz' in elem.rpr:
                font_size = int(elem.rpr['sz'])/2
                doc.usage_font_size[font_size] += weight
            elif style_id is not None:
                font_size = -1
                # should check all styles
                for style in doc.get_styles(style_id):
                    if 'sz' in style.rpr:
                        font_size = int(style.rpr['sz'])/2
                        break

                if font_size != -1:
                    doc.usage_font_size[font_size] += weight
            else:
                st = doc.styles.get_by_id(style_id, 'paragraph')
                font_size = -1

                for style in doc.get_styles(st.style_id):
                    if 'sz' in style.rpr:
                        font_size = int(style.rpr['sz'])/2
                        break

                if font_size != -1:
                    doc.usage_font_size[font_size] += weight
                else:
                    if doc.default_style:
                        if 'sz' in doc.default_style.rpr:
                            font_size = int(doc.default_style.rpr['sz'])/2
                            doc.usage_font_size[font_size] += weight

        if isinstance(elem, Table):
            for column in value:
                for cell in column:
                    weight += _calculate(doc, cell, style_id)

        if isinstance(elem, TableCell) or isinstance(elem, Link) or isinstance(elem, TextBox):
            for el in value:
                weight += _calculate(doc, el, style_id)

    if hasattr(elem, 'elements'):
        for e in elem.elements:
            weight += _calculate(doc, e, style_id)

    return weight


def calculate_weight(doc, elem):
    weight = _calculate(doc, elem, None)
    # TODO
    # - update global list of font sizes
    return weight


def is_header(doc, name):
    if name == '':
        return True

    return name in doc.possible_headers


def find_important(doc, headers):
    default_font_size = 0

    def _big_enough(block):
        # This is something which should be increased
        return block['weight'] > BIG_ENOUGH

    HEADERS_IMPORTANCE = [[el] for el in doc.possible_headers]

    # this might not be needed anymore 
    for style_id in doc.used_styles:
        style = doc.styles.get_by_id(style_id)
        font_size = style.get_font_size()

        if font_size != -1:
            for i in range(len(HEADERS_IMPORTANCE)):
                if HEADERS_IMPORTANCE[i][0] < font_size:
                    if i == 0:
                        HEADERS_IMPORTANCE = [[font_size, style_id]]+HEADERS_IMPORTANCE
                    else:
                        HEADERS_IMPORTANCE[i-1].append(style_id)
                        break
                elif HEADERS_IMPORTANCE[i][0] == font_size:
                    HEADERS_IMPORTANCE[i].append(style_id)    

    # find TOC
    def _find_toc():
        for idx, header in enumerate(headers):
            if 'is_toc' in header:
                return (idx, header)

        return None

    found_toc = _find_toc()
    frontmatter = []

    if found_toc is not None:        
        idx = found_toc[0]
        frontmatter = headers[:idx]
        headers = headers[idx+1:]
    else:
        maximum = 0
        for idx, header in enumerate(headers):
            if maximum > MAXIMUM_FRONTMATTER_HEADER:
                break
            if maximum * MAXIMUM_FRONTMATTER_TIMES > header['weight']:
                frontmatter = headers[:idx-1]
                headers = headers[idx:]
                break
            else:
                if maximum < header['weight']:
                    maximum = header['weight']

    n = 0

    while n < len(headers)-1:
        if headers[n]['weight'] < MAXIMUM_EAT_MARKER and headers[0]['font_size'] == headers[n+1]['font_size']:
            headers[n+1]['weight'] += headers[n]['weight']
            headers[n+1]['index'] = headers[n]['index']
            del headers[n]
        else:
            n += 1

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

            if len(big_blocks) > 1:
                return frontmatter+lst

    return None 


def mark_headers(doc, markers):
    selected = []

    for style in markers:
        font_size = style.get('font_size', 0)

        if style['name'] != '':
            st = doc.styles.get_by_id(style['name'])
            font_size = st.get_font_size()

        if is_header(doc, font_size) or style['name'] == 'ContentsHeading':
            if style['name'] == '':
                selected.append({'name': '', 'index': style['index'], 'weight': style['weight'], 'font_size': style['font_size']})
            else:
                selected.append({'name': style['name'], 'index': style['index'], 'weight': style['weight'], 'font_size': font_size})

            if style['name'] == 'ContentsHeading':
                selected[-1]['is_toc'] = True
        else:
            if len(selected) > 0:
                selected[-1]['weight'] += style['weight']

    return selected


def mark_styles(doc, elements):
    """
    Checks all elements and creates a list of diferent markers for styles or different elements.
    """

    not_using_styles = False

    # there is a better way also
    logger.info('[MARK_STYLES]')
    logger.info('  used types = %s', str(doc.used_styles))
    logger.info('  headers = %s', str(doc.possible_headers))

    # check here how much the styles are being used in possible headers
    # maybe they are not used that much

    if len(doc.used_styles) < 5 and len(doc.possible_headers) < 5:
        not_using_styles = True
        doc.possible_headers = [POSSIBLE_HEADER_SIZE] + doc.possible_headers
        logger.info('   => not using styles')

    markers = [{'name': '', 'weight': 0, 'index': 0, 'font_size': 0}]

    for pos, elem in enumerate(elements):
        try:
            style = doc.styles.get_by_id(elem.style_id)
        except AttributeError:
            style = None

        if isinstance(elem, Paragraph):
            # Ignore if it is dropcap. 
            # What we should do here is also calculate weight, which we do not do at the moment.
            if elem.is_dropcap():
                continue

            for el in elem.elements:
                # insert it inside
                if isinstance(el, Break):
                    pass
                    markers.append({'name': '', 'weight': 0, 'index': pos, 'font_size': 0, 'page_break': True})

        weight = calculate_weight(doc, elem)
        if weight == 0:
            continue

        font_size = -1
        has_found = False

        if isinstance(elem, TOC):
            markers.append({'name': '', 'weight': weight, 'index': pos, 'font_size': fnt_size, 'is_toc': True})
            continue                

        if style:
            markers.append({'name': elem.style_id, 'weight': weight, 'index': pos})

            if elem.style_id == 'ContentsHeading':
                markers[-1]['is_toc'] = True

            has_found = True
        else:
            if hasattr(elem, 'rpr') and 'sz' in elem.rpr:
                markers.append({'name': '', 'weight': weight, 'index': pos, 'font_size': int(elem.rpr['sz'])/2})
                font_size = int(elem.rpr['sz'])/2
                has_found = True

            # Prolog, Teil 1

            if not_using_styles:
                if hasattr(elem, 'rpr') and ('jc' in elem.ppr or 'b' in elem.rpr):
                    elements[pos].possible_header = True

                    markers.append({'name': '', 'weight': weight+100, 'index': pos, 'font_size': POSSIBLE_HEADER_SIZE})
                    font_size = POSSIBLE_HEADER_SIZE
                    has_found = True

        if hasattr(elem, 'elements'):
            for e in elem.elements:
                # TODO
                # check if this is empty element
                if hasattr(e, 'rpr') and 'sz' in e.rpr:
                    fnt_size = int(e.rpr['sz'])/2                        
                    if fnt_size != font_size:
                        markers.append({'name': '', 'weight': weight, 'index': pos, 'font_size': fnt_size})
                        has_found = True

        if not has_found and len(markers) > 0:
            markers[-1]['weight'] += weight

    return markers


def split_document(doc): 
    markers = mark_styles(doc, doc.elements)

    doc._calculate_possible_headers()

    if len(doc.used_styles) < MINIMUM_USED_STYLES and len(doc.possible_headers) < 5: # not sure if possible_headers is available at all
        logger.info('[SPLIT_DOCUMENT]')
        logger.info('  used styles = %s', doc.used_styles)
        logger.info('  possible headers = %s', doc.possible_headers)

        doc.possible_headers = [POSSIBLE_HEADER_SIZE] + doc.possible_headers

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

    def _serialize_chapter(idx, els):        
        options = {'empty_paragraph_as_nbsp': True}

        if len(doc.possible_text) > 0:
            options['scale_to_size'] = doc.possible_text[0]

        s =  serialize.serialize_elements(doc, els, options=options)

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
            
            # This is temp
            #body[0].clear()
            body[0].text = chapter_title
        else:
            title =  etree.Element('h1')
            if idx == 0:
                title.text = _'Frontmatter'
                chapter_title = 'Frontmatter'
            else:
                title.text = 'Unknown'
                chapter_title = ''

            body.insert(0, title)

        return (chapter_title, etree.tostring(body, pretty_print=True, encoding="utf-8", xml_declaration=False))

    chapters = split_document(doc)

    export_chapters = []
    idx = 0

    if chapters:
        # first, everything before the first chapter
        if len(chapters) > 0:
            if chapters[0]['index'] != 0:
                export_chapters.append(_serialize_chapter(idx, doc.elements[:chapters[0]['index']-1]))
                idx += 1

            if len(chapters) > 1 and chapters[0]['index'] == chapters[1]['index']:
                chapters = chapters[1:]

        for n in range(len(chapters)-1):
            if chapters[n]['index'] == chapters[n+1]['index']-1:
                _html = _serialize_chapter(idx, [doc.elements[chapters[n]['index']]])
            else:
                _html = _serialize_chapter(idx, doc.elements[chapters[n]['index']:chapters[n+1]['index']-1])
            idx += 1

            export_chapters.append(_html)

        export_chapters.append(_serialize_chapter(idx, doc.elements[chapters[-1]['index']:]))
    else:
        export_chapters.append(_serialize_chapter(idx, doc.elements))

    return export_chapters