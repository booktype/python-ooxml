# -*- coding: utf-8 -*-

"""Serialize OOXML Document structure to HTML.

.. moduleauthor:: Aleksandar Erkalovic <aerkalov@gmail.com>

Hooks:

* page_break
* math
* a
* img
* footnote
* ul
* ol
* li
* p
* td
* tr
* table
* textbox 
* h

.. code-block:: python

    def hook_paragraph(ctx, document, elem):
        pass

"""

import os.path
import six
import collections

from lxml import etree
from . import doc


def _get_based_on(styles, name):
    for _, values in styles.items():
        if values.based_on == name:
            return values
    return None


def _get_numbering(document, numid, ilvl):
    try:
        return document.numbering[numid][ilvl]['numFmt']
    except:
        return 'bullet'


def _get_numbering_tag(fmt):
    if fmt == 'decimal':
        return 'ol'

    return 'ul'


def _get_parent(root):
    elem = root

    while True:
        elem = elem.getparent()

        if elem.tag in ['ul', 'ol']:
            return elem


# Serialize elements

def serialize_break(ctx, document, elem, root):
    _div = etree.SubElement(root, 'span')
    _div.set('style', 'page-break-after: always;')

    fire_hooks(ctx, document, elem, _div, ctx.get_hook('page_break'))

    return root


def serialize_math(ctx, document, elem, root):
    _div = etree.SubElement(root, 'span')
    _div.set('style', 'border: 1px solid red')
    _div.text = 'We do not support Math blocks at the moment.'

    fire_hooks(ctx, document, elem, _div, ctx.get_hook('math'))

    return root

def serialize_link(ctx, document, elem, root):
    _a = etree.SubElement(root, 'a')

    # this should not take just first element
    if len(elem.elements) > 0:
        _a.text = elem.elements[0].value()
    _a.set('href', document.relationships[elem.rid]['target'])

    fire_hooks(ctx, document, elem, _a, ctx.get_hook('a'))

    return root


def serialize_image(ctx, document, elem, root):            
    img_src = document.relationships[elem.rid]['target']

    img_name, img_extension = os.path.splitext(img_src)

    _img = etree.SubElement(root, 'img')
    # make path configurable
    _img.set('src', 'static/{}{}'.format(elem.rid, img_extension))

    fire_hooks(ctx, document, elem, _img, ctx.get_hook('img'))

    return root


def close_list(ctx, root):

    try:
        n = len(ctx.in_list)
        if n <= 0:
            return root

        elem = root

        while n > 0:
            while True:
                if elem.tag in ['ul', 'ol', 'td']:                    
                    elem = elem.getparent()
                    break

                elem = elem.getparent()

            n -= 1

        ctx.in_list = []

        return elem
    except:
        return None


def open_list(ctx, document, par, root, elem):
    _ls = None

    if par.ilvl != ctx.ilvl or par.numid != ctx.numid:
        # start

        if ctx.ilvl is not None and (par.ilvl > ctx.ilvl):
            fmt = _get_numbering(document, par.numid, par.ilvl)

            if par.ilvl > 0:
                # get last <li> in <ul>
                # could be nicer
                _b = list(root)[-1]
                _ls = etree.SubElement(_b, _get_numbering_tag(fmt))
                root = _ls
            else:
                _ls = etree.SubElement(root, _get_numbering_tag(fmt))
                root = _ls

            fire_hooks(ctx, document, par, _ls, ctx.get_hook(_get_numbering_tag(fmt)))

            ctx.in_list.append((par.numid, par.ilvl))
        elif ctx.ilvl is not None and par.ilvl < ctx.ilvl:
            fmt = _get_numbering(document, ctx.numid, ctx.ilvl)

            try:
                while True:
                    numid, ilvl = ctx.in_list[-1]

                    if numid == par.numid and ilvl == par.ilvl:
                        break

                    root = _get_parent(root)
                    ctx.in_list.pop()
            except:
                pass

#        if ctx.numid is not None and par.numid > ctx.numid:
#            if ctx.numid != None:   
        if par.numid > ctx.numid:
            fmt = _get_numbering(document, par.numid, par.ilvl)
            _ls = etree.SubElement(root, _get_numbering_tag(fmt))
            fire_hooks(ctx, document, par, _ls, ctx.get_hook(_get_numbering_tag(fmt)))

            ctx.in_list.append((par.numid, par.ilvl))
            root = _ls
                
    ctx.ilvl = par.ilvl
    ctx.numid = par.numid

    _a = etree.SubElement(root, 'li')
    _a.text = elem.text

    for a in list(elem):
        _a.append(a)

    fire_hooks(ctx, document, par, _a, ctx.get_hook('li'))

    return root
    

def fire_hooks(ctx, document, elem, element, hooks):
    if not hooks:
        return

    for hook in hooks:
        hook(ctx, document, elem, element)


def has_style(node):
    elements = ['b', 'i', 'u', 'strike', 'color', 'jc', 'sz', 'ind', 'superscript', 'subscript']

    return any([True for elem in elements if elem in node.rpr])


def get_style_fontsize(node):
    if 'sz' in node.rpr:
        return int(node.rpr['sz']) / 2

    return 0

def get_style_css(ctx, node, embed=True):
    style = []

    if not node:
        return

    # temporarily
    if not embed:
        if 'b' in node.rpr:
            style.append('font-weight: bold')

        if 'i' in node.rpr:
            style.append('font-style: italic')

        if 'u' in node.rpr:
            style.append('text-decoration: underline')

    if 'strike' in node.rpr:
        style.append('text-decoration: line-through')

    if 'color' in node.rpr:
        if node.rpr['color'] != '000000':
            style.append('color: #{}'.format(node.rpr['color']))

    if 'jc' in node.ppr:
        # left right both
        align = node.ppr['jc']
        if align.lower() == 'both':
            align = 'justify'

        style.append('text-align: {}'.format(align))

    if 'sz' in node.rpr:
        size = int(node.rpr['sz']) / 2

        if ctx.options['scale_to_size']:
            scale = int(round(size) / ctx.options['scale_to_size'] * 100.0)
            style.append('font-size: {}%'.format(scale))
        else:
            style.append('font-size: {}pt'.format(size))

    if 'ind' in node.ppr:
        if 'left' in node.ppr['ind']:
            size = int(node.ppr['ind']['left']) / 10
            style.append('margin-left: {}px'.format(size))

    if len(style) == 0:
        return ''

    return '; '.join(style) + ';'


def get_style(document, elem):
    try:
        return document.styles.get_by_id(elem.style_id)
    except AttributeError:
        return None


def get_style_name(style):
    return style.style_id


def get_all_styles(document, style):
    classes = []

    while True:
        classes.insert(0, get_style_name(style))

        if style.based_on:
            style = document.styles.get_by_id(style.based_on)
        else:
            break

    return classes


def get_css_classes(document, style):
    return ' '.join([st.lower() for st in get_all_styles(document, style)])


def serialize_paragraph(ctx, document, par, root, embed=True):
    style = get_style(document, par)

    elem = etree.Element('p')

    _style = get_style_css(ctx, par)

    if _style != '':
        elem.set('style', _style)

    if style:
        elem.set('class', get_css_classes(document, style))

    max_font_size = get_style_fontsize(par)

    if style:
        max_font_size = style.get_font_size()

        
    for el in par.elements:
        _serializer =  ctx.get_serializer(el)

        if _serializer:
            _serializer(ctx, document, el, elem)

        if isinstance(el, doc.Text):
            children = list(elem)
            _text_style = get_style_css(ctx, el)

            if get_style_fontsize(el) > max_font_size:
                max_font_size = get_style_fontsize(el)

            if 'superscript' in el.rpr:
                new_element = etree.Element('sup')
                new_element.text = el.value()
            elif 'subscript' in el.rpr:
                new_element = etree.Element('sub')
                new_element.text = el.value()               
            elif 'b' in el.rpr or 'i' in el.rpr or 'u' in el.rpr:                
                new_element = None
                _element = None

                def _add_formatting(f, new_element, _element):
                    if f in el.rpr:
                        _t = etree.Element(f)

                        if new_element is not None:
                            _element.append(_t)
                            _element = _t
                        else:
                            new_element = _t
                            _element = new_element

                    return new_element, _element

                new_element, _element = _add_formatting('b', new_element, _element)
                new_element, _element = _add_formatting('i', new_element, _element)
                new_element, _element = _add_formatting('u', new_element, _element)
                
                _element.text = el.value()
            else:
                new_element = etree.Element('span')
                new_element.text = el.value()

            if _text_style != '':
                new_element.set('style', _text_style)
            # else:
            #     # maybe we don't need this
            #     new_element.set('class', 'noformat')

            was_inserted = False


            if len(children) > 0:
                _child_style = children[-1].get('style') or ''
                
                if new_element.tag == children[-1].tag and _text_style == _child_style and children[-1].tail is None:
                    txt = children[-1].text or ''
                    txt2 = new_element.text or ''
                    children[-1].text = u'{}{}'.format(txt, txt2)  
                    was_inserted = True

                if _style == '' and _text_style == '' and new_element.tag == 'span':
                    _e = children[-1]

                    txt = _e.tail or ''
                    _e.tail = u'{}{}'.format(txt, new_element.text)
                    was_inserted = True

            if not was_inserted:
                if _style == '' and _text_style == '' and new_element.tag == 'span':
                    txt = elem.text or ''
                    elem.text = u'{}{}'.format(txt, new_element.text)
                else:
                    elem.append(new_element)
    
    if not par.is_dropcap():
        if style:
            # todo:
            # - missing list of heading styles somehwhere
            if ctx.header.is_header(par, max_font_size, elem):
                elem.tag = ctx.header.get_header(par, style, elem)
                if par.ilvl == None:        
                    root = close_list(ctx, root)
                    ctx.ilvl, ctx.numid = None, None

                if root is not None:
                    root.append(elem)

                fire_hooks(ctx, document, par, elem, ctx.get_hook('h'))
                return root
        else:
#            if max_font_size > ctx.header.default_font_size:
            if True:
                if ctx.header.is_header(par, max_font_size, elem):
                    if elem.text != '' and len(list(elem)) != 0:
                        elem.tag = ctx.header.get_header(par, max_font_size, elem)

                        if par.ilvl == None:        
                            root = close_list(ctx, root)
                            ctx.ilvl, ctx.numid = None, None


                        if root is not None:
                            root.append(elem)

                        fire_hooks(ctx, document, par, elem, ctx.get_hook('h'))
                        return root

    if len(list(elem)) == 0 and elem.text is None:
        if ctx.options['empty_paragraph_as_nbsp']:
            elem.append(etree.Entity('nbsp'))

    # Indentation is different. We are starting or closing list.
    if par.ilvl != None:        
        root = open_list(ctx, document, par, root, elem)
        return root
    else:
        root = close_list(ctx, root)
        ctx.ilvl, ctx.numid = None, None

    # Add new elements to our root element.
    if root is not None:
        root.append(elem)

    fire_hooks(ctx, document, par, elem, ctx.get_hook('p'))

    return root


def serialize_symbol(ctx, document, el, root):
    span = etree.SubElement(root, 'span')
    span.text = el.value()

    fire_hooks(ctx, document, el, span, ctx.get_hook('symbol'))

    return root


def serialize_footnote(ctx, document, el, root):
    p_foot = document.footnotes[el.rid]

    p = etree.Element('p')
    foot_doc = serialize_paragraph(ctx, document, p_foot, p)
    # must put content of the footter somewhere
    footnote_num = el.rid

    if el.rid not in ctx.footnote_list:
        ctx.footnote_id += 1
        ctx.footnote_list[el.rid] = ctx.footnote_id

    footnote_num = ctx.footnote_list[el.rid]

    note = etree.SubElement(root, 'sup')
    link = etree.SubElement(note, 'a')
    link.set('href', '#')
    link.text = u'{}'.format(footnote_num)

    fire_hooks(ctx, document, el, note, ctx.get_hook('footnote'))

    return root


def serialize_table(ctx, document, table, root):
    if ctx.ilvl != None:
        root = close_list(ctx, root)
        ctx.ilvl, ctx.numid = None, None

    _table = etree.SubElement(root, 'table')
    _table.set('border', '1')
    _table.set('width', '100%')

    style = get_style(document, table)

    if style:
        _table.set('class', get_css_classes(document, style))

    for rows in table.rows:
        _tr = etree.SubElement(_table, 'tr')

        for cell in rows:
            _td = etree.SubElement(_tr, 'td')

            if cell.grid_span != 1:
                _td.set('colspan', str(cell.grid_span))

            if cell.row_span != 1:
                _td.set('rowspan', str(cell.row_span))

            for elem in cell.elements:
                if isinstance(elem, doc.Paragraph):
                    _ser = ctx.get_serializer(elem)
                    _td = _ser(ctx, document, elem, _td, embed=False)

            if ctx.ilvl != None:
#                root = close_list(ctx, root)
                _td = close_list(ctx, _td)

                ctx.ilvl, ctx.numid = None, None

            fire_hooks(ctx, document, table, _td, ctx.get_hook('td'))
        fire_hooks(ctx, document, table, _td, ctx.get_hook('tr'))

    fire_hooks(ctx, document, table, _table, ctx.get_hook('table'))

    return root    


def serialize_textbox(ctx, document, txtbox, root):
    _div = etree.SubElement(root, 'div')
    _div.set('class', 'textbox')

    for elem in txtbox.elements:
        _ser = ctx.get_serializer(elem)

        if _ser:
            _ser(ctx, document, elem, _div)

    fire_hooks(ctx, document, txtbox, _div, ctx.get_hook('textbox'))

    return root

# Header Context

class HeaderContext:
    def __init__(self):
        self.doc = None

        self.default_font_size = 0
        self.headers_sizes = []

    def init(self, doc):
        self.doc = doc

        if doc.default_style:
            self.default_font_size = get_style_fontsize(doc.default_style)

        def _filter_font_sizes(sizes):
            for sz, value in sizes:
                if sz > self.default_font_size:
                    yield (sz, value)

            return 

        most_common_fonts = doc.used_font_size.most_common()
        most_common = -1
        if len(most_common_fonts) > 0:
            if most_common_fonts[0][1] > 4: # this is huge gable, we should calculate it better
                most_common = most_common_fonts[0][0]

        _list_of_fonts = [fnt for fnt in doc.used_font_size.items() if fnt[0] != most_common] 
        self.header_sizes = [[el] for el in reversed(collections.OrderedDict(sorted(_filter_font_sizes(_list_of_fonts), key=lambda t: t[0])))]

        for style_id in doc.used_styles:
            style = doc.styles.get_by_id(style_id)

            if hasattr(style, 'rpr') and 'sz' in style.rpr:
                font_size = int(style.rpr['sz']) / 2

                if font_size <= self.default_font_size:
                    continue

                for i in range(len(self.header_sizes)):
                    if self.header_sizes[i][0] < font_size:
                        if i == 0:
                            self.header_sizes = [[font_size]]+self.header_sizes
                        else:
                            self.header_sizes[i-1].append(style_id)
                            break
                    elif self.header_sizes[i][0] == font_size:
                        self.header_sizes[i].append(style_id)


    def is_header(self, elem, style, node):
        if hasattr(elem, 'possible_header'):
            if elem.possible_header:
                return True

        if not style:
            return False


        if hasattr(style, 'style_id'):
            return style.get_font_size() in self.doc.possible_headers
        else:
            return style in self.doc.possible_headers
        # for st in self.header_sizes:
        #     if style:
        #         if hasattr(style, 'style_id'):
        #             if style.style_id in st:
        #                 return True
        #         else:
        #             if style in st:
        #                 return True

        # return False        


    def get_header(self, elem, style, node):

        font_size = style

        if hasattr(elem, 'possible_header'):
            if elem.possible_header:
                return 'h1'

        if not style:
            return 'h6'

        if hasattr(style, 'style_id'):
            font_size = style.get_font_size() 

        try:
            return 'h{}'.format(self.doc.possible_headers.index(font_size)+1)
        except ValueError:
            return 'h6'

        # for idx, st in enumerate(self.header_sizes):
        #     if style:
        #         if hasattr(style, 'style_id'):
        #             if style.style_id in st:
        #                 n = idx + 1
        #                 if n > 6:
        #                     n = 6
        #                 return 'h{}'.format(n)
        #         else:
        #             if style in st:
        #                 n = idx + 1
        #                 if n > 6:
        #                     n = 6
        #                 return 'h{}'.format(n)

        # return 'h2'



# Default options

DEFAULT_OPTIONS = {
    'serializers': {
        doc.Paragraph: serialize_paragraph,
        doc.Table: serialize_table,
        doc.Link: serialize_link,
        doc.Image: serialize_image,
        doc.Math: serialize_math,
        doc.Break: serialize_break,
        doc.TextBox: serialize_textbox,
        doc.Footnote: serialize_footnote,
        doc.Symbol: serialize_symbol
    },

    'hooks': {},
    'header': HeaderContext,
    'scale_to_size': None,
    'empty_paragraph_as_nbsp': False
}


class Context:
    def __init__(self, document,options=None):
        self.options = dict(DEFAULT_OPTIONS)

        if options:
            if 'serializers' in options:
                self.options['serializers'].update(options['serializers'])

            if 'hooks' in options:
                self.options['hooks'].update(options['hooks'])

            # this is not that good way of updating options

            if 'header' in options:
                self.options['header'] = options['header']

            if 'scale_to_size' in options:
                self.options['scale_to_size'] = options['scale_to_size']

            if 'empty_paragraph_as_nbsp' in options:
                self.options['empty_paragraph_as_nbsp'] = options['empty_paragraph_as_nbsp']

        self.reset()
        self.header.init(document)

    def get_hook(self, name):
        return self.options['hooks'].get(name, None)

    def get_serializer(self, node):
        return self.options['serializers'].get(type(node), None)

        if type(node) in self.options['serializers']:
            return self.options['serializers'][type(node)]

        return None

    def reset(self):
        self.ilvl = None
        self.numid = None

        self.footnote_id = 0
        self.footnote_list = {}

        self.in_list = []
        self.header = self.options['header']()

# Serialize style into CSS

def serialize_styles(document, prefix='', options=None):
    all_styles = []
    css_content = ''

    # get all used styles
    for style_id in document.used_styles:
        all_styles += get_all_styles(document, document.styles.get_by_id(style_id))

    # get all default styles

    def _generate(ctx, style_id, n):
        style = document.styles.get_by_id(style_id)
        style_css = get_style_css(ctx, style, embed=False)

        return "{} {{ {} }}\n".format(",".join(['{} {}'.format(prefix, x) for x in n]), style_css)

    ctx = Context(document, options=options)

    for style_type, style_id in six.iteritems(document.styles.default_styles):
        if style_type == 'table':
            n = ["table"]
            css_content += _generate(ctx, style_id, n)
        elif style_type == 'paragraph':
            n = ["p", "div", "span"]
            css_content += _generate(ctx, style_id, n)            
        elif style_type == 'character':
            n = ["span"]
            css_content += _generate(ctx, style_id, n)            
        elif style_type == 'numbering':
            n = ["ul", "li"]
            css_content += _generate(ctx, style_id, n)

    # DEFAULT STYLES SEPARATELY!

    # get style content for all styles
    for style_id in set(all_styles):
        style = document.styles.get_by_id(style_id)
        css_content += "{0} .{1} {{ {2} }}\n\n".format(prefix, style_id.lower(), get_style_css(ctx, style, embed=False))

    return css_content

# Serialize list of elements into HTML

def serialize_elements(document, elements, options=None):
    ctx = Context(document, options)

    tree_root = root = etree.Element('div')

    for elem in elements:
        _ser = ctx.get_serializer(elem)

        if _ser:
            root = _ser(ctx, document, elem, root)

    # TODO:
    # - create footnotes now

    return etree.tostring(tree_root, pretty_print=True, encoding="utf-8", xml_declaration=False)


def serialize(document, options=None):    
    return serialize_elements(document, document.elements, options)
