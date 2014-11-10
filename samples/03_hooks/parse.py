import six
import logging

import ooxml
from ooxml import parse, serialize, importer

logging.basicConfig(filename='ooxml.log', level=logging.INFO)

def check_for_header(ctx, document, el, elem):
    if hasattr(el, 'style_id'):
        if el.style_id == 'Title':
            elem.tag = 'h1'

def check_for_quote(ctx, document, el, elem):
    if hasattr(el, 'style_id'):
        if el.style_id == 'Quote':
            elem.set('class', elem.get('class', '') + ' our_quote')

file_name = '../files/03_hooks.docx'
dfile = ooxml.read_from_file(file_name)

opts = {
    'hooks': {
       'p': [check_for_quote, check_for_header]
    }
}

six.print_(serialize.serialize(dfile.document, opts))

