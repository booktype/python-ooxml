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
import math

from lxml import etree
from . import doc


###############################################################################
## TEMP FUNCTIONS
###############################################################################

def _get_font_size(document, style):
    """Get font size defined for this style.

    It will try to get font size from it's parent style if it is not defined by original style. 

    :Args:
      - document (:class:`ooxml.doc.Document`): Document object
      - style (:class:`ooxml.doc.Style`): Style object

    :Returns:
      Returns font size as a number. -1 if it can not get font size.
    """

    font_size = style.get_font_size()

    if  font_size == -1:
        if style.based_on:
            based_on = document.styles.get_by_id(style.based_on)
            if based_on:
                return _get_font_size(document, based_on)

    return font_size

def _get_based_on(styles, name):
    for _, values in styles.items():
        if values.based_on == name:
            return values
    return None


def _get_numbering(document, numid, ilvl):
    """Returns type for the list.

    :Returns:
      Returns type for the list. Returns "bullet" by default or in case of an error.
    """

    try:
        abs_num = document.numbering[numid]
        return document.abstruct_numbering[abs_num][ilvl]['numFmt']
    except:
        return 'bullet'


def _get_numbering_tag(fmt):
    """Returns HTML tag defined for this kind of numbering.

    :Args:
      - fmt (str): Type of numbering

    :Returns:
      Returns "ol" for numbered lists and "ul" for everything else.      
    """

    if fmt == 'decimal':
        return 'ol'

    return 'ul'


def _get_parent(root):
    """Returns root element for a list.

    :Args:
      root (Element): lxml element of current location

    :Returns:
      lxml element representing list      
    """

    elem = root

    while True:
        elem = elem.getparent()

        if elem.tag in ['ul', 'ol']:
            return elem


def close_list(ctx, root):
    """Close already opened list if needed.

    This will try to see if it is needed to close already opened list.

    :Args:
      - ctx (:class:`Context`): Context object 
      - root (Element): lxml element representing current position.

    :Returns:
      lxml element where future content should be placed.
    """

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
    """Open list if it is needed and place current element as first member of a list.

    :Args:
      - ctx (:class:`Context`): Context object
      - document (:class:`ooxml.doc.Document`): Document object
      - par (:class:`ooxml.doc.Paragraph`): Paragraph element
      - root (Element): lxml element of current location
      - elem (Element): lxml element representing current element we are trying to insert

    :Returns:
        lxml element where future content should be placed.
    """

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


###############################################################################
## SERIALIZER HOOKS
###############################################################################

def serialize_break(ctx, document, elem, root):
    "Serialize break element."

    if elem.break_type == u'textWrapping':
        _div = etree.SubElement(root, 'br')
    else:
        _div = etree.SubElement(root, 'span')
        if ctx.options['embed_styles']:
            _div.set('style', 'page-break-after: always;')

    fire_hooks(ctx, document, elem, _div, ctx.get_hook('page_break'))

    return root


def serialize_math(ctx, document, elem, root):
    """Serialize math element.

    Math objects are not supported at the moment. This is wht we only show error message.
    """

    _div = etree.SubElement(root, 'span')
    if ctx.options['embed_styles']:
        _div.set('style', 'border: 1px solid red')
    _div.text = 'We do not support Math blocks at the moment.'

    fire_hooks(ctx, document, elem, _div, ctx.get_hook('math'))

    return root

def serialize_link(ctx, document, elem, root):
    """Serilaze link element.

    This works only for external links at the moment.
    """

    _a = etree.SubElement(root, 'a')

    for el in elem.elements:
        _ser = ctx.get_serializer(el)

        if _ser:
            _td = _ser(ctx, document, el, _a)
        else:
            if isinstance(el, doc.Text):
                children = list(_a)

                if len(children) == 0:
                    _text = _a.text or u''

                    _a.text = u'{}{}'.format(_text, el.value())
                else:
                    _text = children[-1].tail or u''

                    children[-1].tail = u'{}{}'.format(_text, el.value())

    if elem.rid in document.relationships:
        _a.set('href', document.relationships[elem.rid].get('target', ''))

    fire_hooks(ctx, document, elem, _a, ctx.get_hook('a'))

    return root


def serialize_image(ctx, document, elem, root):
    """Serialize image element.

    This is not abstract enough.
    """

    _img = etree.SubElement(root, 'img')
    # make path configurable

    if elem.rid in document.relationships:
        img_src = document.relationships[elem.rid].get('target', '')
        img_name, img_extension = os.path.splitext(img_src)

        _img.set('src', 'static/{}{}'.format(elem.rid, img_extension))

    fire_hooks(ctx, document, elem, _img, ctx.get_hook('img'))

    return root


def fire_hooks(ctx, document, elem, element, hooks):
    """Fire hooks on newly created element.

    For each newly created element we will try to find defined hooks and execute them.

    :Args:
      - ctx (:class:`Context`): Context object
      - document (:class:`ooxml.doc.Document`): Document object
      - elem (:class:`ooxml.doc.Element`): Element which we serialized
      - element (Element): lxml element which we created
      - hooks (list): List of hooks
    """

    if not hooks:
        return

    for hook in hooks:
        hook(ctx, document, elem, element)


def has_style(node):    
    """Tells us if node element has defined styling.

    :Args:
      - node (:class:`ooxml.doc.Element`): Element

    :Returns:
      True or False

    """

    elements = ['b', 'i', 'u', 'strike', 'color', 'jc', 'sz', 'ind', 'superscript', 'subscript', 'small_caps']

    return any([True for elem in elements if elem in node.rpr])


def get_style_fontsize(node):
    """Returns font size defined by this element.

    :Args:
      - node (:class:`ooxml.doc.Element`): Node element

    :Returns:
      Font size as int number or 0 if it is not defined

    """
    if hasattr(node, 'rpr'):
        if 'sz' in node.rpr:
            return int(node.rpr['sz']) / 2

    return 0


def get_style_css(ctx, node, embed=True, fontsize=-1):
    """Returns as string defined CSS for this node.

    Defined CSS can be different if it is embeded or no. When it is embeded styling 
    for bold,italic and underline will not be defined with CSS. In that case we
    use defined tags <b>,<i>,<u> from the content.

    :Args:
      - ctx (:class:`Context`): Context object
      - node (:class:`ooxml.doc.Element`): Node element
      - embed (book): True by default.

    :Returns:
      Returns as string defined CSS for this node
    """

    style = []

    if not node:
        return

    if fontsize in [-1, 2]:
        if 'sz' in node.rpr:
            size = int(node.rpr['sz']) / 2

            if ctx.options['embed_fontsize']:
                if ctx.options['scale_to_size']:                
                    multiplier = size-ctx.options['scale_to_size']
                    scale = 100 + int(math.trunc(8.3*multiplier))
                    style.append('font-size: {}%'.format(scale))
                else:
                    style.append('font-size: {}pt'.format(size))

    if fontsize in [-1, 1]:
        # temporarily
        if not embed:
            if 'b' in node.rpr:
                style.append('font-weight: bold')

            if 'i' in node.rpr:
                style.append('font-style: italic')

            if 'u' in node.rpr:
                style.append('text-decoration: underline')

        if 'small_caps' in node.rpr:
            style.append('font-variant: small-caps')

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

        if 'ind' in node.ppr:
            if 'left' in node.ppr['ind']:
                size = int(node.ppr['ind']['left']) / 10
                style.append('margin-left: {}px'.format(size))

            if 'right' in node.ppr['ind']:
                size = int(node.ppr['ind']['right']) / 10
                style.append('margin-right: {}px'.format(size))

            if 'first_line' in node.ppr['ind']:
                size = int(node.ppr['ind']['first_line']) / 10
                style.append('text-indent: {}px'.format(size))

    if len(style) == 0:
        return ''

    return '; '.join(style) + ';'


def get_style(document, elem):
    """Get the style for this node element.

    :Args:
      - document (:class:`ooxml.doc.Document`): Document object
      - elem (:class:`ooxml.doc.Element`): Node element

    :Returns:
      Returns :class:`ooxml.doc.Style` object or None if it is not found.
    """

    try:
        return document.styles.get_by_id(elem.style_id)
    except AttributeError:
        return None


def get_style_name(style):
    """Returns style if for specific style.
    """

    return style.style_id


def get_all_styles(document, style):
    """Returns list of styles on which specified style is based on.

    :Args:
      - document (:class:`ooxml.doc.Document`): Document object
      - style (:class:`ooxml.doc.Style`): Style object

    :Returns:
      List of style objects.
    """

    classes = []

    while True:
        classes.insert(0, get_style_name(style))

        if style.based_on:
            style = document.styles.get_by_id(style.based_on)
        else:
            break

    return classes


def get_css_classes(document, style):
    """Returns CSS classes for this style.

    This function will check all the styles specified style is based on and return their CSS classes.

    :Args:
      - document (:class:`ooxml.doc.Document`): Document object
      - style (:class:`ooxml.doc.Style`): Style object

    :Returns:
      String representing all the CSS classes for this element.

    >>> get_css_classes(doc, st)
    'header1 normal'
    """
    lst = [st.lower() for st in get_all_styles(document, style)[-1:]] + \
        ['{}-fontsize'.format(st.lower()) for st in get_all_styles(document, style)[-1:]]

    return ' '.join(lst)


def serialize_paragraph(ctx, document, par, root, embed=True):
    """Serializes paragraph element.

    This is the most important serializer of them all.    
    """

    style = get_style(document, par)

    elem = etree.Element('p')

    if ctx.options['embed_styles']:
        _style = get_style_css(ctx, par)

        if _style != '':
            elem.set('style', _style)

    else:
        _style = ''

    if style:
        elem.set('class', get_css_classes(document, style))

    max_font_size = get_style_fontsize(par)

    if style:
        max_font_size = _get_font_size(document, style)

        
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

                for comment_id in ctx.opened_comments:
                    document.comments[comment_id].text += ' ' + el.value()
            else:
                new_element = etree.Element('span')
                new_element.text = el.value()

                for comment_id in ctx.opened_comments:
                    document.comments[comment_id].text += ' ' + el.value()


            if ctx.options['embed_styles']:
                if _text_style != '' and _style != _text_style:
                    new_element.set('style', _text_style)

            # This is for situations when style has options and
            # text is trying to unset them
            # else:            
            #     new_element.set('class', 'noformat')

            was_inserted = False


            if len(children) > 0:
                _child_style = children[-1].get('style') or ''

                if new_element.tag == children[-1].tag and (_text_style == _child_style or _child_style == '') and children[-1].tail is None:
                    txt = children[-1].text or ''
                    txt2 = new_element.text or ''
                    children[-1].text = u'{}{}'.format(txt, txt2)  
                    was_inserted = True

                if not was_inserted:
#                    if _style == '' and _text_style == '' and new_element.tag == 'span':
                    if _style == _text_style  and new_element.tag == 'span':

                        _e = children[-1]

                        txt = _e.tail or ''
                        _e.tail = u'{}{}'.format(txt, new_element.text)
                        was_inserted = True

            if not was_inserted:
#                if _style == '' and _text_style == '' and new_element.tag == 'span':
                if _style ==  _text_style  and new_element.tag == 'span':

                    txt = elem.text or ''
                    elem.text = u'{}{}'.format(txt, new_element.text)
                else:
                    elem.append(new_element)
    
    if not par.is_dropcap() and par.ilvl == None:
        if style:
            if ctx.header.is_header(par, max_font_size, elem, style=style):
                elem.tag = ctx.header.get_header(par, style, elem)
                if par.ilvl == None:        
                    root = close_list(ctx, root)
                    ctx.ilvl, ctx.numid = None, None

                if root is not None:
                    root.append(elem)

                fire_hooks(ctx, document, par, elem, ctx.get_hook('h'))
                return root
        else:
#            Commented part where we only checked for heading if font size
#            was bigger than default font size. In many cases this did not
#            work out well.
#            if max_font_size > ctx.header.default_font_size:
            if True:
                if ctx.header.is_header(par, max_font_size, elem, style=style):
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
    "Serialize special symbols."

    span = etree.SubElement(root, 'span')
    span.text = el.value()

    fire_hooks(ctx, document, el, span, ctx.get_hook('symbol'))

    return root


def serialize_footnote(ctx, document, el, root):
    "Serializes footnotes."

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


def serialize_comment(ctx, document, el, root):
    "Serializes comment."

    # Check if option is turned on

    if el.comment_type == 'end':
        ctx.opened_comments.remove(el.cid)
    else:
        if el.comment_type != 'reference':
            ctx.opened_comments.append(el.cid)

        if ctx.options['comment_span']:
            link = etree.SubElement(root, 'a')
            link.set('href', '#')
            link.set('class', 'comment-link')    
            link.set('id', 'comment-id-' + el.cid)    

            link.text = ''

            fire_hooks(ctx, document, el, link, ctx.get_hook('comment'))

    return root


def serialize_endnote(ctx, document, el, root):
    "Serializes endnotes."

    footnote_num = el.rid

    if el.rid not in ctx.endnote_list:
        ctx.endnote_id += 1
        ctx.endnote_list[el.rid] = ctx.endnote_id

    footnote_num = ctx.endnote_list[el.rid]

    note = etree.SubElement(root, 'sup')
    link = etree.SubElement(note, 'a')
    link.set('href', '#')
    link.text = u'{}'.format(footnote_num)

    fire_hooks(ctx, document, el, note, ctx.get_hook('endnote'))

    return root


def serialize_smarttag(ctx, document, el, root):
    "Serializes smarttag."

    if ctx.options['smarttag_span']:
        _span = etree.SubElement(root, 'span', {'class': 'smarttag', 'data-smarttag-element': el.element})
    else:
        _span = root

    for elem in el.elements:
        _ser = ctx.get_serializer(elem)

        if _ser:
            _td = _ser(ctx, document, elem, _span)
        else:
            if isinstance(elem, doc.Text):
                children = list(_span)

                if len(children) == 0:
                    _text = _span.text or u''

                    _span.text = u'{}{}'.format(_text, elem.text)
                else:
                    _text = children[-1].tail or u''

                    children[-1].tail = u'{}{}'.format(_text, elem.text)

    fire_hooks(ctx, document, el, _span, ctx.get_hook('smarttag'))

    return root


def serialize_table(ctx, document, table, root):
    """Serializes table element.
    """

    # What we should check really is why do we pass None as root element
    # There is a good chance some content is missing after the import

    if root is None:
        return root

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
    """Serialize textbox element."""

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
    """Header context used for header recognition.

    This is used only for easier recognition of headers used during the import process.
    """

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


    def is_header(self, elem, font_size, node, style=None):
        """Used for checking if specific element is a header or not.

        :Returns:
          True or False
        """
        
        # This logic has been disabled for now. Mark this as header if it has
        # been marked during the parsing or mark.
        # if hasattr(elem, 'possible_header'):
        #     if elem.possible_header:                
        #         return True

        # if not style:
        #     return False

        if hasattr(style, 'style_id'):
            fnt_size = _get_font_size(self.doc, style)

            from .importer import calculate_weight
            weight = calculate_weight(self.doc, elem)
            
            if weight > 50:
                return False

            if fnt_size in self.doc.possible_headers_style:                
                return True

            return font_size in self.doc.possible_headers
        else:
            list_of_sizes = {}
            for el in elem.elements:
                try:
                    fs = get_style_fontsize(el)
                    weight = len(el.value()) if el.value() else 0

                    list_of_sizes[fs] = list_of_sizes.setdefault(fs, 0) + weight
                except:
                    pass

            sorted_list_of_sizes = list(collections.OrderedDict(sorted(list_of_sizes.iteritems(), key=lambda t: t[0])))
            font_size_to_check = font_size

            if len(sorted_list_of_sizes) > 0:
                if sorted_list_of_sizes[0] != font_size:
                    return sorted_list_of_sizes[0] in self.doc.possible_headers

            return font_size in self.doc.possible_headers


    def get_header(self, elem, style, node):
        """Returns HTML tag representing specific header for this element.

        :Returns:
          String representation of HTML tag.

        """
        font_size = style
        if hasattr(elem, 'possible_header'):
            if elem.possible_header:
                return 'h1'

        if not style:
            return 'h6'

        if hasattr(style, 'style_id'):
            font_size = _get_font_size(self.doc, style)

        try:
            if font_size in self.doc.possible_headers_style:
                return 'h{}'.format(self.doc.possible_headers_style.index(font_size)+1)

            return 'h{}'.format(self.doc.possible_headers.index(font_size)+1)
        except ValueError:
            return 'h6'


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
        doc.Comment: serialize_comment,
        doc.Footnote: serialize_footnote,
        doc.Endnote: serialize_endnote,
        doc.Symbol: serialize_symbol,
        doc.SmartTag: serialize_smarttag
    },

    'hooks': {},
    'header': HeaderContext,
    'scale_to_size': None,
    'empty_paragraph_as_nbsp': False,
    'embed_styles': True,
    'embed_fontsize': True,
    'smarttag_span': False,
    'comment_span': False,
    'pretty_print': True
}


class Context:
    """Context object used during the serialization.

    It is used to hold needed information like hooks, serializers, extra options and temporary information.

    :Args:
      - document (:class:`ooxml.doc.Document`): Document object
      - options (dict): Optional dictionary with options

    Options:
      - serializers (dict):!
      - hooks (dict):
      - header (:class:`HeaderContext`): Reference to a class
      - scale_to_size: None is a default option. If defined as int will be used as base font size for the text
      - empty_paragraph_as_nbsp: False is a default option. If True it will insert &nbsp; inside of empty paragraphs
    """

    def __init__(self, document, options=None):
        self.options = dict(DEFAULT_OPTIONS)

        if options:
            for opt_key, opt_value in six.iteritems(options):
                if type(opt_value) == type({}):
                    self.options[opt_key].update(opt_value)
                else:
                    self.options[opt_key] = opt_value

        self.reset()
        self.header.init(document)

    def get_hook(self, name):
        """Get reference to a specific hook.

        :Args:
          - name (str): Hook name

        :Returns:
          List with defined hooks. None if it is not found.
        """

        return self.options['hooks'].get(name, None)

    def get_serializer(self, node):
        """Returns serializer for specific element.

        :Args:
          - node (:class:`ooxml.doc.Element`): Element object 

        :Returns:
          Returns reference to a function which will be used for serialization.
        """

        return self.options['serializers'].get(type(node), None)

        if type(node) in self.options['serializers']:
            return self.options['serializers'][type(node)]

        return None

    def reset(self):
        self.ilvl = None
        self.numid = None

        self.opened_comments = []
        self.footnote_id = 0
        self.footnote_list = {}
        self.endnote_id = 0
        self.endnote_list = {}

        self.in_list = []
        self.header = self.options['header']()

# Serialize style into CSS

###############################################################################
## SERIALIZERS
###############################################################################

def serialize_styles(document, prefix='', options=None):
    """

    :Args:
      - document (:class:`ooxml.doc.Document`): Document object
      - prefix (str): Optional prefix used for 
      - options (dict): Optional dictionary with :class:`Context` options

    :Returns:
        CSS styles as string.

    >>> serialize_styles(doc)
    p { color: red; }

    >>> serialize_styles(doc, '#editor')
    #editor p { color: red; }

    """
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
            n = ["p", "div"] #  n = ["p", "div", "span"]
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
        styles = []

        while True:
            styles.insert(0, style)

            if style.based_on:
                style = document.styles.get_by_id(style.based_on)
            else:
                break

        content = "\n".join([get_style_css(ctx, st, embed=False, fontsize=1) for st in styles])
        css_content += "{0} .{1} {{ {2} }}\n\n".format(prefix, style_id.lower(), content)

        content = "\n".join([get_style_css(ctx, st, embed=False, fontsize=2) for st in styles])
        css_content += "{0} .{1}-fontsize {{ {2} }}\n\n".format(prefix, style_id.lower(), content)

    return css_content

# Serialize list of elements into HTML

def serialize_elements(document, elements, options=None):
    """Serialize list of elements into HTML string.

    :Args:
      - document (:class:`ooxml.doc.Document`): Document object
      - elements (list): List of elements
      - options (dict): Optional dictionary with :class:`Context` options

    :Returns:
      Returns HTML representation of the document.
    """    
    ctx = Context(document, options)

    tree_root = root = etree.Element('div')

    for elem in elements:
        _ser = ctx.get_serializer(elem)

        if _ser:
            root = _ser(ctx, document, elem, root)

    # TODO:
    # - create footnotes now

    return etree.tostring(tree_root, pretty_print=ctx.options.get('pretty_print', True), encoding="utf-8", xml_declaration=False)


def serialize(document, options=None):
    """Serialize entire document into HTML string.

    :Args:
      - document (:class:`ooxml.doc.Document`): Document object
      - options (dict): Optional dictionary with :class:`Context` options

    :Returns:
      Returns HTML representation of the document.
    """
    return serialize_elements(document, document.elements, options)
