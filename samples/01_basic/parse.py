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

    six.print_(serialize.serialize(dfile.document))
    six.print_(serialize.serialize_styles(dfile.document))
