# -*- coding: utf-8 -*-

"""Parse OOXML structure.

.. moduleauthor:: Aleksandar Erkalovic <aerkalov@gmail.com>

"""

import zipfile
import logging

from lxml import etree

from . import doc, NAMESPACES


logger = logging.getLogger('ooxml')


def _name(name):
    return name.format(**NAMESPACES)


def parse_previous_properties(doc, paragraph, prop):
    if not paragraph:
        return

    style = prop.find(_name('{{{w}}}rStyle'))

    if style is not None:
        paragraph.rpr['style'] = style.attrib[_name('{{{w}}}val')]
        doc.add_style_as_used(paragraph.rpr['style'])

    color = prop.find(_name('{{{w}}}color'))

    if color is not None:
        paragraph.rpr['color'] = color.attrib[_name('{{{w}}}val')]

    rtl = prop.find(_name('{{{w}}}rtl'))

    if rtl is not None:
        paragraph.rpr['rtl'] = rtl.attrib[_name('{{{w}}}val')]

    sz = prop.find(_name('{{{w}}}sz'))

    if sz is not None:
        paragraph.rpr['sz'] = sz.attrib[_name('{{{w}}}val')]
        doc.add_font_as_used(paragraph.rpr['sz'])

    # parse bold
    b = prop.find(_name('{{{w}}}b'))

    if b is not None:        
        # todo
        # check b = on and not off
        paragraph.rpr['b'] = True

    # parse italic
    i = prop.find(_name('{{{w}}}i'))

    if i is not None:        
        # todo
        # check b = on and not off
        paragraph.rpr['i'] = True

    # parse underline
    u = prop.find(_name('{{{w}}}u'))

    if u is not None:        
        # todo
        # check b = on and not off
        paragraph.rpr['u'] = True

    # parse underline
    strike = prop.find(_name('{{{w}}}strike'))

    if strike is not None:        
        # todo
        # check b = on and not off
        paragraph.rpr['strike'] = True

    vert_align = prop.find(_name('{{{w}}}vertAlign'))

    if vert_align is not None:
        value = vert_align.attrib[_name('{{{w}}}val')]

        if value == 'superscript':
            paragraph.rpr['superscript'] = True

        if value == 'subscript':
            paragraph.rpr['subscript'] = True


def parse_paragraph_properties(doc, paragraph, prop):
    if not paragraph:
        return

    style = prop.find(_name('{{{w}}}pStyle'))

    if style is not None:
        paragraph.style_id = style.attrib[_name('{{{w}}}val')]
        doc.add_style_as_used(paragraph.style_id)

    numpr = prop.find(_name('{{{w}}}numPr'))

    if numpr is not None:
        ilvl = numpr.find(_name('{{{w}}}ilvl'))

        if ilvl is not None:
            paragraph.ilvl = int(ilvl.attrib[_name('{{{w}}}val')])

        numid = numpr.find(_name('{{{w}}}numId'))

        if numid is not None:
            paragraph.numid = int(numid.attrib[_name('{{{w}}}val')])

    rpr = prop.find(_name('{{{w}}}rPr'))

    if rpr is not None:
        parse_previous_properties(doc, paragraph, rpr)

    jc = prop.find(_name('{{{w}}}jc'))

    if jc is not None:
        paragraph.ppr['jc'] = jc.attrib[_name('{{{w}}}val')]

    ind = prop.find(_name('{{{w}}}ind'))

    if ind is not None:
        paragraph.ppr['ind'] = {}
        
        if _name('{{{w}}}left') in ind.attrib:
            paragraph.ppr['ind']['left'] = ind.attrib[_name('{{{w}}}left')]


    # w:ind - left leftChars right hanging firstLine

def parse_drawing(document, container, elem):
    blip = elem.xpath('.//a:blip', namespaces=NAMESPACES)[0]        
    _rid =  blip.attrib[_name('{{{r}}}embed')]

    img = doc.Image(_rid)
    container.elements.append(img)


def parse_footnote(document, container, elem):
    _rid =  elem.attrib[_name('{{{w}}}id')]

    foot = doc.Footnote(_rid)
    container.elements.append(foot)


def parse_alternate(document, container, elem):
    txtbx = elem.find('.//'+_name('{{{w}}}txbxContent'))
    paragraphs = []

    if txtbx is None:
        return

    for el in txtbx:
        if el.tag == _name('{{{w}}}p'):
            paragraphs.append(parse_paragraph(document, el))

    textbox = doc.TextBox(paragraphs)
    container.elements.append(textbox)
    

def parse_text(document, container, element):
    txt = None

    alternate = element.find(_name('{{{mc}}}AlternateContent'))

    if alternate is not None:
        parse_alternate(document, container, alternate)

    br = element.find(_name('{{{w}}}br'))

    if br is not None:
        if _name('{{{w}}}type') in br.attrib:
            _type = br.attrib[_name('{{{w}}}type')]        

            brk = doc.Break(_type)
            container.elements.append(brk)

    t = element.find(_name('{{{w}}}t'))

    if t is not None:
        txt = doc.Text(t.text)

        container.elements.append(txt)

    rpr = element.find(_name('{{{w}}}rPr'))

    if rpr is not None:
        parse_previous_properties(document, txt, rpr)

    r = element.find(_name('{{{w}}}r'))

    if r is not None:
        parse_text(document, container, r)

    foot = element.find(_name('{{{w}}}footnoteReference'))

    if foot is not None:
        parse_footnote(document, container, foot)

    sym = element.find(_name('{{{w}}}sym'))

    if sym is not None:
        _font = sym.attrib[_name('{{{w}}}font')]
        _char = sym.attrib[_name('{{{w}}}char')]

        container.elements.append(doc.Symbol(font=_font, character=_char))

    image = element.find(_name('{{{w}}}drawing'))

    if image is not None:
        parse_drawing(document, container, image)

    return


def parse_paragraph(document, par):
    paragraph = doc.Paragraph()    
    paragraph.document = document

    for elem in par:
        if elem.tag == _name('{{{w}}}pPr'):
            parse_paragraph_properties(document, paragraph, elem)

        if elem.tag == _name('{{{w}}}r'):
            parse_text(document, paragraph, elem)      

        if elem.tag == _name('{{{m}}}oMath'):
            _m = doc.Math()
            paragraph.elements.append(_m)

        if elem.tag == _name('{{{m}}}oMathPara'):
            _m = doc.Math()
            paragraph.elements.append(_m)

        if elem.tag == _name('{{{w}}}hyperlink'):
            try:
                t = doc.Link(elem.attrib[_name('{{{r}}}id')])

                parse_text(document, t, elem)            

                paragraph.elements.append(t)            
            except:
                logger.error('Error with with hyperlink [%s].', str(elem.attrib.items()))

    return paragraph


def parse_table_properties(doc, table, prop):
    if not table:
        return

    style = prop.find(_name('{{{w}}}tblStyle'))

    if style is not None:
        table.style_id = style.attrib[_name('{{{w}}}val')]
        doc.add_style_as_used(table.style_id)


def parse_table_column_properties(doc, cell, prop):
    if not cell:
        return

    grid = prop.find(_name('{{{w}}}gridSpan'))

    if grid is not None:
        cell.grid_span = int(grid.attrib[_name('{{{w}}}val')])


    vmerge = prop.find(_name('{{{w}}}vMerge'))

    if vmerge is not None:
        if _name('{{{w}}}val') in vmerge.attrib:
            cell.vmerge = vmerge.attrib[_name('{{{w}}}val')]
        else:
            cell.vmerge = ""


def parse_table(document, tbl):

    def _change(rows, pos_x):
        if len(rows) == 1:
            return rows

        count_x = 1

        for x in rows[-1]:
            if count_x == pos_x:
                x.row_span += 1

            count_x += x.grid_span

        return rows

    table = doc.Table()

    tbl_pr = tbl.find(_name('{{{w}}}tblPr'))

    if tbl_pr is not None:
        parse_table_properties(document, table, tbl_pr)

    for tr in tbl.xpath('./w:tr', namespaces=NAMESPACES):
        columns = []
        pos_x = 0

        for tc in tr.xpath('./w:tc', namespaces=NAMESPACES):            
            cell = doc.TableCell()

            tc_pr = tc.find(_name('{{{w}}}tcPr'))

            if tc_pr is not None:
                parse_table_column_properties(doc, cell, tc_pr)

            # maybe after
            pos_x += cell.grid_span

            if cell.vmerge is not None and cell.vmerge == "":
                table.rows = _change(table.rows, pos_x)
            else:
                for p in tc.xpath('./w:p', namespaces=NAMESPACES):
                    cell.elements.append(parse_paragraph(document, p))

                columns.append(cell)

        table.rows.append(columns)

    return table


def parse_document(xmlcontent):
    document = etree.fromstring(xmlcontent)

    body = document.xpath('.//w:body', namespaces=NAMESPACES)[0]

    document = doc.Document()

    for elem in body:
        if elem.tag == _name('{{{w}}}p'):
            document.elements.append(parse_paragraph(document, elem))

        if elem.tag == _name('{{{w}}}tbl'):
            document.elements.append(parse_table(document, elem))

    return document


def parse_relationship(document, xmlcontent):
    doc = etree.fromstring(xmlcontent)

    for elem in doc:
        if elem.tag == _name('{{{pr}}}Relationship'):
            rel = {'target': elem.attrib['Target'],
                   'type': elem.attrib['Type'],
                   'target_mode': elem.attrib.get('TargetMode', 'Internal')}

            document.relationships[elem.attrib['Id']] = rel


def parse_style(document, xmlcontent):
    styles = etree.fromstring(xmlcontent)

    # parse default styles
    default_rpr = styles.find(_name('{{{w}}}rPrDefault'))

    _r = styles.xpath('.//w:rPrDefault', namespaces=NAMESPACES)

    if len(_r) > 0:
        rpr = _r[0].find(_name('{{{w}}}rPr'))

        if rpr is not None:
            st = doc.Style()
            parse_previous_properties(document, st, rpr)
            document.default_style = st

    # rest of the styles
    for style in styles.xpath('.//w:style', namespaces=NAMESPACES):
        st = doc.Style()

        st.style_id = style.attrib[_name('{{{w}}}styleId')]

        style_type = style.attrib[_name('{{{w}}}type')]        
        if style_type is not None:
            st.style_type = style_type

        if _name('{{{w}}}default') in style.attrib:
            is_default = style.attrib[_name('{{{w}}}default')]
            if is_default is not None:
                st.is_default = is_default == '1'

        name = style.find(_name('{{{w}}}name'))
        if name is not None:
            st.name = name.attrib[_name('{{{w}}}val')]

        based_on = style.find(_name('{{{w}}}basedOn'))

        if based_on is not None:
            st.based_on = based_on.attrib[_name('{{{w}}}val')]

        document.styles.styles[st.style_id] = st

        if st.is_default:
            document.styles.default_styles[st.style_type] = st.style_id

        rpr = style.find(_name('{{{w}}}rPr'))

        if rpr is not None:
            parse_previous_properties(document, st, rpr)


        ppr = style.find(_name('{{{w}}}pPr'))

        if ppr is not None:
           parse_paragraph_properties(document, st, ppr)


def parse_footnotes(document, xmlcontent):
    styles = etree.fromstring(xmlcontent)
    document.footnotes = {}

    for style in styles.xpath('.//w:footnote', namespaces=NAMESPACES):
        _type = style.attrib.get(_name('{{{w}}}type'), None)

        # don't know what to do with these now
        if _type in ['separator', 'continuationSeparator', 'continuationNotice']:
            continue

        p = parse_paragraph(document, style.find(_name('{{{w}}}p')))

        document.footnotes[style.attrib[_name('{{{w}}}id')]] = p


def parse_numbering(document, xmlcontent):
    numbering = etree.fromstring(xmlcontent)
    document.numbering = {}

    for abstruct_num in numbering.xpath('.//w:abstractNum', namespaces=NAMESPACES):
        numb = {}

        for lvl in abstruct_num.xpath('./w:lvl', namespaces=NAMESPACES):
            ilvl = int(lvl.attrib[_name('{{{w}}}ilvl')])

            fmt = lvl.find(_name('{{{w}}}numFmt'))
            numb[ilvl] = {'numFmt': fmt.attrib[_name('{{{w}}}val')]}

        document.numbering[abstruct_num.attrib[_name('{{{w}}}abstractNumId')]] = numb


def parse_from_file(file_object):
    logger.info('Parsing %s file.', file_object.file_name)

    # Read the files
    doc_content = file_object.read_file('document.xml')
    
    # Parse the document
    document = parse_document(doc_content)

    try:    
        style_content = file_object.read_file('styles.xml')
        parse_style(document, style_content)        
    except KeyError:
        logger.warning('Could not read styles.')

    try:        
        doc_rel_content = file_object.read_file('_rels/document.xml.rels')
        parse_relationship(document, doc_rel_content)
    except KeyError:
        logger.warning('Could not read relationships.')

    try:    
        footnotes_content = file_object.read_file('footnotes.xml')
        parse_footnotes(document, footnotes_content)    
    except KeyError:
        logger.warning('Could not read footnotes.')

    try:
        numbering_content = file_object.read_file('numbering.xml')
        parse_numbering(document, numbering_content)    
    except KeyError:
        logger.warning('Could not read numbering.')

    return document

