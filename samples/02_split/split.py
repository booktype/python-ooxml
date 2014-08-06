import sys
import six
import logging

from lxml import etree

import ooxml
from ooxml import parse, serialize, importer

logging.basicConfig(filename='ooxml.log', level=logging.INFO)

if len(sys.argv) > 1:
    file_name = sys.argv[1]

    dfile = ooxml.read_from_file(file_name)

    chapters = importer.get_chapters(dfile.document)

    for title, content in chapters:
        six.print_('====================================================================')
        six.print_(title)
        six.print_('====================================================================')
        six.print_(content)

