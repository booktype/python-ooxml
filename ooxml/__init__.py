# -*- coding: utf-8 -*-

VERSION = '0.12'

NAMESPACES = {
    'mo': 'http://schemas.microsoft.com/office/mac/office/2008/main',
    'o':  'urn:schemas-microsoft-com:office:office',
    've': 'http://schemas.openxmlformats.org/markup-compatibility/2006',
    # Text Content
    'w':   'http://schemas.openxmlformats.org/wordprocessingml/2006/main',
    'w10': 'urn:schemas-microsoft-com:office:word',
    'wne': 'http://schemas.microsoft.com/office/word/2006/wordml',
    # Drawing
    'a':   'http://schemas.openxmlformats.org/drawingml/2006/main',
    'm':   'http://schemas.openxmlformats.org/officeDocument/2006/math',
    'mv':  'urn:schemas-microsoft-com:mac:vml',
    'mc':  'http://schemas.openxmlformats.org/markup-compatibility/2006',
    'mo':  'http://schemas.microsoft.com/office/mac/office/2008/main',
    'pic': 'http://schemas.openxmlformats.org/drawingml/2006/picture',
    'v':   'urn:schemas-microsoft-com:vml',
    'wp':  ('http://schemas.openxmlformats.org/drawingml/2006/wordprocessing'
            'Drawing'),
    # Properties (core and extended)
    'cp':  ('http://schemas.openxmlformats.org/package/2006/metadata/core-pr'
            'operties'),
    'dc':  'http://purl.org/dc/elements/1.1/',
    'ep':  ('http://schemas.openxmlformats.org/officeDocument/2006/extended-'
            'properties'),
    'xsi': 'http://www.w3.org/2001/XMLSchema-instance',
    # Content Types
    'ct':  'http://schemas.openxmlformats.org/package/2006/content-types',
    # Package Relationships
    'r':  ('http://schemas.openxmlformats.org/officeDocument/2006/relationships'),
    'pr':  'http://schemas.openxmlformats.org/package/2006/relationships',
    # Dublin Core document properties
    'dcmitype': 'http://purl.org/dc/dcmitype/',
    'dcterms':  'http://purl.org/dc/terms/'}


def read_from_file(file_name):
    """Parser OOXML file and returns parsed document.
    
    :Args:
      - file_name (str): Path to OOXML file

    :Returns:
      Returns object of type :class:`ooxml.docx.DOCXFile`.
    """
    from .docxfile import DOCXFile

    dfile = DOCXFile(file_name)
    dfile.parse()

    return dfile
