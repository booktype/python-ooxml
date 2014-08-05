# -*- coding: utf-8 -*-

"""

.. moduleauthor:: Aleksandar Erkalovic <aerkalov@gmail.com>

"""

import zipfile

from .parse import parse_from_file


class DOCXFile(object):
    """DOCXFile represents the .docx File.

    This class is the main interface for communication with .docx file.

    .. code-block:: python

        dfile = DOCXFile('document.docx')
        dfile.parse()
        dfile.close()

    .. note::
        API interface is still work in progress.

    """

    def __init__(self, file_name):
        self.file_name = file_name

        self.reset()

    @property
    def document(self):
        return self._doc

    def reset(self):
        "Resets the values."

        self.zf = zipfile.ZipFile(self.file_name, 'r')
        self._doc = None

    def parse(self):
        self._doc = parse_from_file(self)

    def read_file(self, file_name):
        return self.zf.open('word/{}'.format(file_name)).read()

    def close(self):
        self.zf.close()
