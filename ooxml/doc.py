# -*- coding: utf-8 -*-

"""Different OOXML Document elements.

.. moduleauthor:: Aleksandar Erkalovic <aerkalov@gmail.com>

"""

import six
import collections


class Style(object):
    """Style object represent OOXML defined style.

    Style is defined by unique identifier. Besides unique identifier style also has defined
    descriptive name. Styles can also be based on another style.
    """

    def __init__(self):
        self.reset()

    def reset(self):
        self.style_id = ''
        self.style_type = ''
        self.is_default = False
        self.name = ''
        self.based_on = ''

        self.rpr = {}
        self.ppr = {}

    def get_font_size(self):
        """Returns font size for this style. 

        Does not check definition in the parent styles.

        :Returns:
          Returns font size as integer. Returns -1 if font size is not defined for this style.
        """
        if 'sz' in self.rpr:
            return int(self.rpr['sz'])/2

        return -1


class StylesCollection:
    """Collection of defined styles.

    """

    def __init__(self):
        self.reset()

    def get_by_name(self, name, style_type = None):
        """Find style by it's descriptive name.

        :Returns:
          Returns found style of type :class:`ooxml.doc.Style`.
        """
        for st in self.styles.values():
            if st:
                if st.name == name:
                    return st

        if style_type and not st:
            st = self.styles.get(self.default_styles[style_type], None)            
        return st

    def get_by_id(self, style_id, style_type = None):
        """Find style by it's unique identifier

        :Returns:
          Returns found style of type :class:`ooxml.doc.Style`.
        """

        for st in self.styles.values():
            if st:
                if st.style_id == style_id:
                    return st

        if style_type:
            return self.styles.get(self.default_styles[style_type], None)
        return None
    
    def reset(self):
        self.styles = {}
        self.default_styles = {}


class Document(object):
    "Represents OOXML document."

    def __init__(self):
        super(Document, self).__init__()

        self.reset()

    def add_style_as_used(self, name):
        if name not in self.used_styles:
            self.used_styles.append(name)

    def add_font_as_used(self, sz):
        fsz = int(sz) / 2
        self.used_font_size[fsz] += 1

    def get_styles(self, name):
        styles = []
        while True:                     
            style = self.styles.get_by_id(name)

            styles.append(style)

            if style.based_on == '':
                return styles

            name = style.based_on

    def _calculate_possible_headers(self):
        _headers = []
        _text = []
        max_count = sum(six.itervalues(self.usage_font_size))

        from .serialize import _get_font_size

        for name in self.used_styles:
            _style = self.styles.get_by_id(name)
            font_size = _get_font_size(self, _style)

            if font_size != -1 and font_size not in _headers:
                _headers.append(font_size)

        self.possible_headers_style = [x for x in reversed(sorted(_headers))]

        _text_list = collections.Counter()

        for font_size, amount in six.iteritems(self.usage_font_size):
            if float(amount) / max_count <= 0.1:                
                if font_size not in _headers:
                    _headers.append(font_size)
            else:
                # This will require some cleanup
                _text.append(font_size)
                _text_list[font_size] = amount

        self.possible_headers = [x for x in reversed(sorted(_headers))]
        self.possible_text = [x for x in reversed(sorted(_text))]

        # remove all possible headers which are bigger than biggest normal font size
        if len(self.possible_text) > 0:
            for value in self.possible_headers[:]:
                if self.possible_text[0] >= value:
                    self.possible_headers.remove(value)
#                    self.possible_headers_style.remove(value)

        _mc = _text_list.most_common(1)

        if len(_mc) > 0:
            self.base_font_size = _mc[0][0]

    def reset(self):
        self.elements = []
        self.relationships = {}
        self.footnotes = {}
        self.endnotes = {}
        self.comments = {}
        self.numbering = {}
        self.abstruct_numbering = {}
        self.styles = StylesCollection()
        self.default_style = None
        self.used_styles = []
        self.used_font_size = collections.Counter()

        self.usage_font_size = collections.Counter()
        self.possible_headers_style = []
        self.possible_headers = []
        self.possible_text = []
        self.base_font_size = -1


class CommentContent:
    def __init__(self, cid):
        self.cid = cid
        self.text = ''
        self.elements = []
        self.date = None
        self.author = None


class Element(object):
    "Basic element paresed in the OOXML document."

    def reset(self):
        pass

    def value(self):
        return None

 
class Paragraph(Element):
    """Represents basic paragraph element in OOXML document.

    Paragraph can also hold other elements. Besides that, list items and dropcaps are also defined by
    this element.
    """
    def __init__(self):
        super(Paragraph, self).__init__()

        self.reset()


    def reset(self):
        self.elements = []

        # Should not be here
        self.numid = None
        self.ilvl = None

        self.rpr = {}
        self.ppr = {}

        self.possible_header = False

    def is_dropcap(self):
        return 'dropcap' in self.ppr and self.ppr['dropcap']


class Text(Element):
    "Represents Text element which can be found inside of other Paragraph elements."

    def __init__(self, text='', parent=None):
        super(Text, self).__init__()

        self.text = text
        self.rpr = {}
        self.ppr = {}
        self.parent = None


    def value(self):
        return self.text


class Link(Element):
    "Represents link element holding reference to internal or external link."

    def __init__(self, rid):
        super(Link, self).__init__()

        self.elements = []
        self.rid = rid
        self.rpr = {}
        self.ppr = {}


    def value(self):
        return self.elements
#        return ''.join(elem.value() for elem in self.elements)


class Image(Element):
    "Represent image element."

    def __init__(self, rid):
        super(Image, self).__init__()

        self.rid = rid

    def value(self):
        return self.rid


class TableCell(Element):
    "Represent one cell in a table."

    def __init__(self):
        super(TableCell, self).__init__()

        self.grid_span = 1
        self.row_span = 1
        self.vmerge = None        
        self.elements = []


    def value(self):
        return self.elements


class Table(Element):
    "Represents table element."

    def __init__(self):
        super(Table, self).__init__()

        self.rows = []

    def value(self):
        return self.rows


class Comment(Element):
    "Represents comment element."

    def __init__(self, cid, comment_type):
        super(Comment, self).__init__()

        self.cid = cid
        self.comment_type = comment_type

    def value(self):
        return self.cid


class Footnote(Element):
    "Represents footnote element."

    def __init__(self, rid):
        super(Footnote, self).__init__()

        self.rid = rid

    def value(self):
        return self.rid


class Endnote(Element):
    "Represents endnote element."

    def __init__(self, rid):
        super(Endnote, self).__init__()

        self.rid = rid

    def value(self):
        return self.rid


class TextBox(Element):
    "Represents TextBox element."

    def __init__(self, elements):
        super(TextBox, self).__init__()

        self.elements = elements

    def value(self):
        return self.elements


class Symbol(Element):
    """Represents special symbol element.

    For some symbols it can do transformation into unicode element.
    """

    SYMBOLS = {
        'F020': u'\u0020',
        'F021': u'\u270F',        
        'F022': u'\u2702',
        'F023': u'\u2701',
        'F024': u'\u1F453',
        'F025': u'\u1F514',
        'F026': u'\u1F4D6',
        'F028': u'\u260E',
        'F029': u'\u2706',
        'F02A': u'\u2709',
        'F037': u'\u2328',
        'F03E': u'\u2707',
        'F03F': u'\u270D',
        'F041': u'\u270C',
        'F045': u'\u261C',
        'F046': u'\u261E',
        'F047': u'\u261D',
        'F048': u'\u261F',
        'F04A': u'\u263A',
        'F04C': u'\u2639',
        'F04E': u'\u2620',
        'F04F': u'\u2690',
        'F051': u'\u2708',
        'F052': u'\u263C',
        'F054': u'\u2744',
        'F056': u'\u271E',
        'F058': u'\u2720',
        'F059': u'\u2721',
        'F05A': u'\u262A',
        'F05B': u'\u262F',
        'F05C': u'\u0950',
        'F05D': u'\u2638',
        'F06A': u'\u0026',
        'F06B': u'\u0026',
        'F06C': u'\u25CF',
        'F06D': u'\u274D',
        'F06E': u'\u25A0',
        'F06F': u'\u25A1',
        'F071': u'\u2751',
        'F072': u'\u2752',
        'F073': u'\u2B27',
        'F074': u'\u29EB',
        'F075': u'\u25C6',
        'F076': u'\u2756',
        'F077': u'\u2B25',
        'F078': u'\u2327',
        'F079': u'\u2353',
        'F07A': u'\u2318',
        'F07B': u'\u2740',
        'F07C': u'\u273F',
        'F07D': u'\u275D',
        'F07E': u'\u275E',
        'F0A4': u'\u25C9',
        'F0A5': u'\u25CE',
        'F0A7': u'\u25AA',
        'F0A8': u'\u25FB',
        'F0AA': u'\u2726',
        'F0AB': u'\u2605',
        'F0AC': u'\u2736',
        'F0AE': u'\u2739',
        'F0AF': u'\u2735',
        'F0B1': u'\u2316',
        'F0B2': u'\u27E1',
        'F0B3': u'\u2311',
        'F0B5': u'\u272A',
        'F0B6': u'\u2730',
        'F0D5': u'\u232B',
        'F0D6': u'\u2326',
        'F0D8': u'\u27A2',
        'F0DC': u'\u27B2',
        'F0E8': u'\u2794',
        'F0EF': u'\u21E6',
        'F0F0': u'\u21E8',
        'F0F1': u'\u21E7',
        'F0F2': u'\u21E9',
        'F0F3': u'\u2B04',
        'F0F4': u'\u21F3',
        'F0F5': u'\u2B00',
        'F0F6': u'\u2B01',
        'F0F7': u'\u2B03',
        'F0F8': u'\u2B02',
        'F0F9': u'\u25AD',
        'F0FA': u'\u25AB',
        'F0FB': u'\u2717',
        'F0FC': u'\u2713',
        'F0FD': u'\u2612',
        'F0FE': u'\u2611'
    }

    def __init__(self, font='Wingdings', character=''):
        self.font = font
        self.character = character

    def value(self):
        return self.SYMBOLS.get(self.character, self.character)


class TOC(Element):
    """Represents Table Of Contents element.

    We don't do much with this element at the moment."""    
    pass


class Break(Element):
    """Represents break element.

    At the moment we support Line Break (textWrapping) and Page Break (page).
    """

    def __init__(self, break_type='textWrapping'):
        self.break_type = break_type

    def value(self):
        return self.break_type


class Math(Element):
    """Represents Math element.

    Math elements are not supported at the moment. We just parse them and create empty element."""

    def __init__(self):
        pass

    def value(self):
        return ''


class SmartTag(Element):
    "Represents SmartTag element."

    def __init__(self):
        self.elements = []
        self.element = ''

    def value(self):
        return ''
