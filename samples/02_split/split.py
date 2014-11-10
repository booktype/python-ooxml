import six
import logging

import ooxml
from ooxml import parse, serialize, importer

logging.basicConfig(filename='ooxml.log', level=logging.INFO)

file_name = '../files/02_split.docx'

dfile = ooxml.read_from_file(file_name)

chapters = importer.get_chapters(dfile.document)

for title, content in chapters:
    six.print_('====================================================================')
    six.print_(title)
    six.print_('====================================================================')
    six.print_(content)

