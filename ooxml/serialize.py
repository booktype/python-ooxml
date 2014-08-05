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

.. code-block:: python

    def hook_paragraph(ctx, document, elem):
        pass

"""

import os.path

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

    fire_hooks(ctx, document, _div, ctx.get_hook('page_break'))

    return root


def serialize_math(ctx, document, elem, root):
    _div = etree.SubElement(root, 'span')
    _div.set('style', 'border: 1px solid red')
    _div.text = 'We do not support Math blocks at the moment.'

    fire_hooks(ctx, document, _div, ctx.get_hook('math'))

    return root

def serialize_link(ctx, document, elem, root):
    _a = etree.SubElement(root, 'a')

    # this should not take just first element
    if len(elem.elements) > 0:
        _a.text = elem.elements[0].value()
    _a.set('href', document.relationships[elem.rid]['target'])

    fire_hooks(ctx, document, _a, ctx.get_hook('a'))

    return root


def serialize_image(ctx, document, elem, root):            
    img_src = document.relationships[elem.rid]['target']

    img_name, img_extension = os.path.splitext(img_src)

    _img = etree.SubElement(root, 'img')
    # make path configurable
    _img.set('src', 'static/{}{}'.format(elem.rid, img_extension))

    fire_hooks(ctx, document, _img, ctx.get_hook('img'))

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


def fire_hooks(ctx, document, element, hooks):
    if not hooks:
        return

    for hook in hooks:
        hook(ctx, document, element)


def get_style(node):
    style = []

    if not node:
        return

    if 'b' in node.rpr:
        style.append('font-weight: bold')

    if 'i' in node.rpr:
        style.append('font-style: italic')

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
        style.append('font-size: {}pt'.format(size))

    if 'ind' in node.ppr:
        if 'left' in node.ppr['ind']:
            size = int(node.ppr['ind']['left']) / 10
            style.append('margin-left: {}px'.format(size))

    if len(style) == 0:
        return ''

    return '; '.join(style) + ';'


def serialize_paragraph(ctx, document, par, root, embed=True):
    try:
        # style_id can be none
        style = document.styles[par.style_id]
    except AttributeError:
        style = None

    elem = etree.Element('p')

    _style = get_style(par)

    if _style != '':
        elem.set('style', _style)

    # This is just for debugging purposes at the moment
    if style:
        elem.set('data-style', par.style_id)
        
    for el in par.elements:   
        if isinstance(el, doc.TextBox):
            ctx.get_serializer(el)(ctx, document, el, elem)

        if isinstance(el, doc.Break):
            ctx.get_serializer(el)(ctx, document, el, elem)

        if isinstance(el, doc.Math):
            # this is just wrong
            # must go trough each element and serialize it separately            
            ctx.get_serializer(el)(ctx, document, el, elem)

        if isinstance(el, doc.Link):
            # this is just wrong
            # must go trough each element and serialize it separately            
            ctx.get_serializer(el)(ctx, document, el, elem)

        if isinstance(el, doc.Text):
            _style = get_style(el)

            if _style != '':
                # check if it is the same style
                # and then just add at the end
                _s = None
                children = list(elem)

                if len(children) > 0:
                    _s = children[-1].get('style')

                if _s == _style:
                    txt = children[-1].text or ''
                    # this extra space is creating problems
                    children[-1].text = u'{} {}'.format(txt, el.value())
                else:
                    span = etree.SubElement(elem, 'span')
                    span.text = el.value()                    
                    span.set('style', _style)                
            else:
                children = list(elem)

                if len(children) == 0:   
                    txt = elem.text or ''
                    # TODO
                    # extra space here
                    elem.text = u'{} {}'.format(txt, el.value())
                else:
                    txt = children[-1].tail or ''
                    children[-1].tail = u'{} {}'.format(txt, el.value())
    
        if isinstance(el, doc.Image):
            ctx.get_serializer(el)(ctx, document, el, elem)

        if isinstance(el, doc.Footnote):
            p_foot = document.footnotes[el.rid]

            p = etree.Element('p')
            foot_doc = serialize_paragraph(ctx, document, p_foot, p)
            # must put content of the footter somewhere
            footnote_num = el.rid

            if el.rid not in ctx.footnote_list:
                ctx.footnote_id += 1
                ctx.footnote_list[el.rid] = ctx.footnote_id

            footnote_num = ctx.footnote_list[el.rid]

            note = etree.SubElement(elem, 'sup')
            link = etree.SubElement(note, 'a')
            link.set('href', '#')
            link.text = u'{}'.format(footnote_num)

            fire_hooks(ctx, document, note, ctx.get_hook('footnote'))

        if isinstance(el, doc.Symbol):
            span = etree.SubElement(elem, 'span')
            span.text = el.value()

    if style:
        try:
            # style_id can be none
            style2 = document.styles.get(style.based_on, None)
        except AttributeError:
            style2 = None        

        # todo:
        # - missing fire hook for heading
        # - missing list of heading styles somehwhere
        if style.name in ['heading 1', 'Title']:
            elem.tag = 'h1'
            root.append(elem)

            return root

        if style.name in ['heading 2', 'Subtitle']:
            elem.tag = 'h2'
            root.append(elem)

            return root

    if par.ilvl != None:
        _ls = None

        if par.ilvl != ctx.ilvl or par.numid != ctx.numid:
            # start
            if par.ilvl > ctx.ilvl:
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

                fire_hooks(ctx, document, _ls, ctx.get_hook(_get_numbering_tag(fmt)))

                ctx.in_list.append((par.numid, par.ilvl))
            elif par.ilvl < ctx.ilvl:
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

            if par.numid > ctx.numid:
                if ctx.numid != None:   
                    # not sure about this
                    fmt = _get_numbering(document, par.numid, par.ilvl)
                    _ls = etree.SubElement(root, _get_numbering_tag(fmt))
                    fire_hooks(ctx, document, _ls, ctx.get_hook(_get_numbering_tag(fmt)))

                    ctx.in_list.append((par.numid, par.ilvl))
                    root = _ls
                    
        ctx.ilvl = par.ilvl
        ctx.numid = par.numid

        _a = etree.SubElement(root, 'li')
        _a.text = elem.text

        for a in list(elem):
            _a.append(a)

        fire_hooks(ctx, document, _a, ctx.get_hook('li'))

        return root
    else:
        root = close_list(ctx, root)
        ctx.ilvl, ctx.numid = None, None

    if root is not None:
        root.append(elem)

    fire_hooks(ctx, document, elem, ctx.get_hook('p'))

    return root


def serialize_table(ctx, document, table, root):
    _table = etree.SubElement(root, 'table')
    _table.set('border', '1')
    _table.set('width', '100%')

    for rows in table.rows:
        _tr = etree.SubElement(_table, 'tr')

        for columns in rows:
            _td = etree.SubElement(_tr, 'td')

            for elem in columns:
                if isinstance(elem, doc.Paragraph):
                    _ser = ctx.get_serializer(elem)

                    _td = _ser(ctx, document, elem, _td, embed=False)

            root = close_list(ctx, root)
            ctx.ilvl, ctx.numid = None, None

            fire_hooks(ctx, document, _td, ctx.get_hook('td'))

        fire_hooks(ctx, document, _td, ctx.get_hook('tr'))

    fire_hooks(ctx, document, _table, ctx.get_hook('table'))

    return root    


def serialize_textbox(ctx, document, txtbox, root):
    _div = etree.SubElement(root, 'div')
    _div.set('class', 'textbox')

    for elem in txtbox.elements:
        _ser = ctx.get_serializer(elem)

        if _ser:
            _ser(ctx, document, elem, _div)

    fire_hooks(ctx, document, _div, ctx.get_hook('textbox'))

    return root

# Default options

DEFAULT_OPTIONS = {
    'serializers': {
        doc.Paragraph: serialize_paragraph,
        doc.Table: serialize_table,
        doc.Link: serialize_link,
        doc.Image: serialize_image,
        doc.Math: serialize_math,
        doc.Break: serialize_break,
        doc.TextBox: serialize_textbox
    },

    'hooks': {}
}


class Context:
    def __init__(self, options=None):
        self.options = dict(DEFAULT_OPTIONS)

        if options:
            if 'serializers' in options:
                self.options['serializers'].update(options['serializers'])

            if 'hooks' in options:
                self.options['hooks'].update(options['hooks'])

        self.reset()

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

# Serialize style into CSS

def serialize_styles(document):
    for name, style in document.styles.iteritems():
        _css = ''

        based_style = style

        while True:
            try:
                based_style = document.styles[based_style.based_on]
                _css = get_style(based_style) + _css
            except:
                break

        _css += get_style(style)


# Serialize list of elements into HTML

def serialize_elements(document, elements, options=None):
    ctx = Context(options)

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
