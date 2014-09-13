# -*- coding: utf-8 -*-

"""Different OOXML Document elements.

.. moduleauthor:: Aleksandar Erkalovic <aerkalov@gmail.com>

"""

class Style(object):
    def __init__(self):
        self.reset()

    def reset(self):
        self.style_id = ''
        self.style_type = ''
        self.is_default = False
        self.named = ''
        self.based_on = ''

        self.rpr = {}
        self.ppr = {}


class StylesCollection:
    def __init__(self):
        self.reset()

    def get_by_name(self, name, style_type = None):
        st = self.styles.get(style_id, None)
        if style_type and not st:
            st = self.styles.get(self.default_style[style_type], None)
        return st

    def get_by_id(self, style_id, style_type = None):
        for st in self.styles.values():
            if st:
                if st.style_id == style_id:
                    return st

        if style_type:
            return self.styles.get(self.default_style[style_type], None)
        return None
    
    def reset(self):
        self.styles = {}
        self.default_styles = {}


class Document(object):
    def __init__(self):
        super(Document, self).__init__()

        self.reset()

    def add_style_as_used(self, name):
        if name not in self.used_styles:
            self.used_styles.append(name)

    def add_font_as_used(self, sz):
        fsz = int(sz) / 2
        self.used_font_size[fsz] = self.used_font_size.setdefault(fsz, 0) + 1

    def get_styles(self, name):
        styles = []

        while True:                    
            style = self.get_style_by_name(name)

            styles.append(style)

            if style.based_on == '':
                return styles

            name = style.based_on

    def reset(self):
        self.elements = []
        self.relationships = {}
        self.footnotes = {}
        self.styles = StylesCollection()
        self.default_style = None
        self.used_styles = []
        self.used_font_size = {}


class Element(object):
    def reset(self):
        pass

    def value(self):
        pass

 
class Paragraph(Element):
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


class Text(Element):
    def __init__(self, text = ''):
        super(Text, self).__init__()

        self.text = text
        self.rpr = {}
        self.ppr = {}


    def value(self):
        return self.text

class Link(Element):
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
    def __init__(self, rid):
        super(Image, self).__init__()

        self.rid = rid

    def value(self):
        return self.rid

class Table(Element):
    def __init__(self):
        super(Table, self).__init__()

        self.rows = []

    def value(self):
        return self.rows


class Footnote(Element):
    def __init__(self, rid):
        super(Footnote, self).__init__()

        self.rid = rid

    def value(self):
        return self.rid

class TextBox(Element):
    def __init__(self, elements):
        super(TextBox, self).__init__()

        self.elements = elements

    def value(self):
        return self.elements


class Symbol(Element):
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


class Break(Element):
    def __init__(self, break_type):
        self.break_type = break_type

    def value(self):
        return self.break_type


class Math(Element):
    def __init__(self):
        pass

    def value(self):
        return ''

