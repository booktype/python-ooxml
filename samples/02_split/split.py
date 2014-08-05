import sys
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
        print '===================================================================='
        print title
        print '===================================================================='
        print content

