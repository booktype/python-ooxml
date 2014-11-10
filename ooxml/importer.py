# -*- coding: utf-8 -*-

"""Try to import and recognised different chapters in the document.

.. moduleauthor:: Aleksandar Erkalovic <aerkalov@gmail.com>

"""

import six
import collections
import logging

from .doc import (Paragraph, Table, TableCell, Link, TextBox, TOC, Break)
from .serialize import _get_font_size

logger = logging.getLogger('ooxml')

POSSIBLE_HEADER_SIZE = 24


def text_length(elem):
    """Returns length of the content in this element.

    Return value is not correct but it is **good enough***.
    """

    if not elem:
        return 0

    value = elem.value()

    try:
        value = len(value)
    except:
        value = 0

    try:
        for a in elem.elements:
            value += len(a.value())
    except:
        pass

    return value


def parse_html_string(s):
    from lxml import html

    utf8_parser = html.HTMLParser(encoding='utf-8')
    html_tree = html.document_fromstring(s , parser=utf8_parser)

    return html_tree


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


DEFAULT_OPTIONS = {
    'minimum_used_styles': 5,
    'minimum_possible_headers': 5,
    'not_using_styles': True,
    'header_as_text': True, # font size used to mark header
    'header_as_text_in_elements': True, # dont size used to mark header in elements
    'header_as_text_length': 30,
    'header_as_text_length_minimum': 5,
    'header_as_bold_centered': False,
    'big_enough_for_block': 100,
    'find_toc': True,
    'separate_frontmatter_h1': False,
    'maximum_frontmatter_header': 300,
    'maximum_frontmatter_times': 5,
    'squash_frontmatter': True,
    'maximum_eat_marker': 100,
    'squash_small_blocks': True

}

class ImporterContext:
    def __init__(self, options=None):
        self.options = dict(DEFAULT_OPTIONS)

        if options:
            self.options.update(options)


def find_important(ctx, doc, headers):
    default_font_size = 0

    def _big_enough(block):
        # This is something which should be increased
        return block['weight'] > ctx.options['big_enough_for_block']

    HEADERS_IMPORTANCE = [[el] for el in doc.possible_headers]

    # this might not be needed anymore 
    for style_id in doc.used_styles:
        style = doc.styles.get_by_id(style_id)
        font_size = _get_font_size(doc, style)

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

    frontmatter = []

    if ctx.options['find_toc']:
        found_toc = _find_toc()
    else:
        found_toc = None

    if found_toc is not None:        
        idx = found_toc[0]
        frontmatter = headers[:idx]
        headers = headers[idx+1:]
    else:
        if ctx.options['separate_frontmatter_h1']:
            for idx, header in enumerate(headers):
                if header.get('name', '') == 'Heading1':
                    if idx > 0:
                        frontmatter = headers[:idx-1]
                        headers = headers[idx:]
                    break

        if ctx.options['squash_frontmatter']:
            maximum = 0
            for idx, header in enumerate(headers):
                if maximum > ctx.options['maximum_frontmatter_header']:
                    break
                if maximum * ctx.options['maximum_frontmatter_times'] > header['weight']:
                    frontmatter = headers[:idx-1]
                    headers = headers[idx:]
                    break
                else:
                    if maximum < header['weight']:
                        maximum = header['weight']

    n = 0

    # COMMENT 1
    if ctx.options['squash_small_blocks']:
        while n < len(headers)-1:
            if headers[n]['weight'] < ctx.options['maximum_eat_marker'] and headers[0]['font_size'] == headers[n+1]['font_size']:
                headers[n+1]['weight'] += headers[n]['weight']
                headers[n+1]['index'] = headers[n]['index']
                del headers[n]
            else:
                n += 1

    for hdrs in HEADERS_IMPORTANCE:        
        lst = []
        for header in headers:
            # if header.get('page_break', False) == True:
            #     lst.append({'name': '', 'weight': 100000, 'index': header['index']})
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


def mark_headers(ctx, doc, markers):
    selected = []

    for style in markers:
        font_size = style.get('font_size', 0)

        if style['name'] != '':
            st = doc.styles.get_by_id(style['name'])
            font_size = _get_font_size(doc, st)

        if is_header(doc, font_size) or style['name'] == 'ContentsHeading':
            if style['name'] == '':
                selected.append({'name': '', 'index': style['index'], 'weight': style['weight'], 'font_size': style['font_size']})
            else:
                selected.append({'name': style['name'], 'index': style['index'], 'weight': style['weight'], 'font_size': font_size})

            if style['name'] == 'ContentsHeading':
                selected[-1]['is_toc'] = True
        else:
            # if style.get('page_break', False) == True:
            #     selected.append({'name': '', 'page_break': True, 'index': style['index'], 'weight': 0, 'font_size': 0})
            # else:
            if len(selected) > 0:
                selected[-1]['weight'] += style['weight']

    return selected


def mark_styles(ctx, doc, elements):
    """
    Checks all elements and creates a list of diferent markers for styles or different elements.
    """

    not_using_styles = False

    if ctx.options['not_using_styles']:
        if len(doc.used_styles) < ctx.options['minimum_used_styles'] and len(doc.possible_headers) < ctx.options['minimum_possible_headers']:
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
            if ctx.options['header_as_text']:
                if hasattr(elem, 'rpr') and 'sz' in elem.rpr:

                    t_length = text_length(elem)

                    if  t_length < ctx.options['header_as_text_length'] and t_length >= ctx.options['header_as_text_length_minimum']:
                        markers.append({'name': '', 'weight': weight, 'index': pos, 'font_size': int(elem.rpr['sz']) / 2})
                        font_size = int(elem.rpr['sz'])/2
                        has_found = True

            if ctx.options['header_as_bold_centered']:
                if not_using_styles:
                    if hasattr(elem, 'rpr') and ('jc' in elem.ppr or 'b' in elem.rpr or 'i' in elem.rpr):
                        if text_length(elem) < 30:

                            elements[pos].possible_header = True

                            markers.append({'name': '', 'weight': weight+100, 'index': pos, 'font_size': POSSIBLE_HEADER_SIZE})
                            font_size = POSSIBLE_HEADER_SIZE
                            has_found = True

        if ctx.options['header_as_text_in_elements']:
            if hasattr(elem, 'elements'):
                for e in elem.elements:
                    # TODO
                    # check if this is empty element
                    if hasattr(e, 'rpr') and 'sz' in e.rpr:
                        t_length = text_length(elem)
                        if  t_length < ctx.options['header_as_text_length'] and t_length >= ctx.options['header_as_text_length_minimum']: 
                            fnt_size = int(e.rpr['sz'])/2

                            if fnt_size != font_size:
                                markers.append({'name': '', 'weight': weight, 'index': pos, 'font_size': fnt_size})
                                has_found = True

        if not has_found and len(markers) > 0:
            markers[-1]['weight'] += weight

    return markers


def split_document(ctx, doc): 
    markers = mark_styles(ctx, doc, doc.elements)
    doc._calculate_possible_headers()

    if ctx.options['separate_frontmatter_h1']:
        style = doc.styles.get_by_id('berschrift1')

        if style:
            font_size = style.get_font_size()

            for value in doc.possible_headers[:]:
                if value > font_size:
                    doc.possible_headers.remove(value)

            for value in doc.possible_headers_style[:]:
                if value > font_size:
                    doc.possible_headers_style.remove(value)

    if ctx.options['not_using_styles']:
        if len(doc.used_styles) < ctx.options['minimum_used_styles'] and len(doc.possible_headers) < 5:
            doc.possible_headers = [POSSIBLE_HEADER_SIZE] + doc.possible_headers

    headers = mark_headers(ctx, doc, markers)

    important = find_important(ctx, doc, headers)

    return important


def get_chapters(doc, options=None, serialize_options=None):
    from lxml import etree
    from . import serialize

    context = ImporterContext(options)

    def _serialize_chapter(idx, els, is_frontmatter):        
        options = {'empty_paragraph_as_nbsp': True}

        if len(doc.possible_text) > 0:
            options['scale_to_size'] = doc.possible_text[0]

        s =  serialize.serialize_elements(doc, els, options=serialize_options)

        if s.startswith(six.b('<div/>')):
            return ('', six.b('<body></body>'))

        root = parse_html_string(s[5:-6])
        body = root.find('.//body')
        chapter_title = ''

        if not is_frontmatter:
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
                if idx > 0:  
                    title =  etree.Element('h1')

                    title.text = 'Unknown'
                    chapter_title = ''
                    body.insert(0, title)
        else:
            h1_headers = body.find('.//h1')
            if h1_headers is not None:
                for h1 in h1_headers:
                    h1.tag = 'h2'

        return (chapter_title, etree.tostring(body, pretty_print=True, encoding="utf-8", xml_declaration=False))

    chapters = split_document(context, doc)

    export_chapters = []
    idx = 0

    if chapters:
        # first, everything before the first chapter
        if len(chapters) > 0:
            if chapters[0]['index'] != 0:
                # The idea is that front matter should not have a chapter title
                chap = _serialize_chapter(idx, doc.elements[:chapters[0]['index']-1], True)
                export_chapters.append((u'', chap[1]))
                idx += 1

            if len(chapters) > 1 and chapters[0]['index'] == chapters[1]['index']:
                chapters = chapters[1:]

        for n in range(len(chapters)-1):
            if chapters[n]['index'] == chapters[n+1]['index']-1:
                _html = _serialize_chapter(idx, [doc.elements[chapters[n]['index']]], False)
            else:
                _html = _serialize_chapter(idx, doc.elements[chapters[n]['index']:chapters[n+1]['index']-1], False)
            idx += 1

            export_chapters.append(_html)

        export_chapters.append(_serialize_chapter(idx, doc.elements[chapters[-1]['index']:], False))
    else:
        export_chapters.append(_serialize_chapter(idx, doc.elements, False))

    return export_chapters