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
import six

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

            fire_hooks(ctx, document, _ls, ctx.get_hook(_get_numbering_tag(fmt)))

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
    

def fire_hooks(ctx, document, element, hooks):
    if not hooks:
        return

    for hook in hooks:
        hook(ctx, document, element)


def has_style(node):
    elements = ['b', 'i', 'u', 'strike', 'color', 'jc', 'sz', 'ind', 'superscript', 'subscript']

    return any([True for elem in elements if elem in node.rpr])


def get_style(node):
    style = []

    if not node:
        return

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
        _serializer =  ctx.get_serializer(el)

        if _serializer:
            _serializer(ctx, document, el, elem)

        if isinstance(el, doc.Text):
            children = list(elem)
            _text_style = get_style(el)

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
            else:
                new_element.set('class', 'noformat')

            was_inserted = False


            if len(children) > 0:
                _child_style = children[-1].get('style') or ''

                if new_element.tag == children[-1].tag and _text_style == _child_style and children[-1].tail == '':
                    txt = children[-1].text or ''
                    children[-1].text = u'{}{}'.format(txt, new_element.text)                    
                    was_inserted = True

                if _style == '' and _text_style == '' and new_element.tag == 'span':
                    _e = children[-1]

                    txt = _e.tail or ''
                    _e.tail = u'{}{}'.format(txt, new_element.text)
                    was_inserted = True

            if not was_inserted:
                if _style == '' and _text_style == '' and new_element.tag == 'span' and elem.tail == '':
                    txt = elem.text or ''
                    elem.text = u'{}{}'.format(txt, new_element.text)
                else:
                    elem.append(new_element)
    
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
            if root is not None:
                root.append(elem)

            return root

        if style.name in ['heading 2', 'Subtitle']:
            elem.tag = 'h2'
            if root is not None:
                root.append(elem)

            return root

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

    fire_hooks(ctx, document, elem, ctx.get_hook('p'))

    return root


def serialize_symbol(ctx, document, el, root):
    span = etree.SubElement(root, 'span')
    span.text = el.value()

    fire_hooks(ctx, document, span, ctx.get_hook('symbol'))

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

    fire_hooks(ctx, document, note, ctx.get_hook('footnote'))

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
        doc.TextBox: serialize_textbox,
        doc.Footnote: serialize_footnote,
        doc.Symbol: serialize_symbol
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
    for name, style in six.iteritems(document.styles):
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
