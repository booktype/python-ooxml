import sys
import six
import logging

import ooxml
from ooxml import parse, serialize, importer

logging.basicConfig(filename='ooxml.log', level=logging.INFO)


if len(sys.argv) > 1:
    file_name = sys.argv[1]

    dfile = ooxml.read_from_file(file_name)

    six.print_("\n-[HTML]-----------------------------\n")
    six.print_(serialize.serialize(dfile.document))

    six.print_("\n-[CSS STYLE]------------------------\n")
    six.print_(serialize.serialize_styles(dfile.document))

    six.print_("\n-[USED STYLES]----------------------\n")
    six.print_(dfile.document.used_styles)

    six.print_("\n-[USED FONT SIZES]------------------\n")
    six.print_(dfile.document.used_font_size)

