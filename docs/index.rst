.. Python-OOXML documentation master file, created by
   sphinx-quickstart on Wed Aug  6 13:14:38 2014.
   You can adapt this file completely to your liking, but it should at least
   contain the root `toctree` directive.

Welcome to Python-OOXML's documentation!
========================================

Python-OOXML is a Python library for parsing Office Open XML files.

Homepage: https://github.com/booktype/python-ooxml

Contents:

.. toctree::
   :maxdepth: 2

   api

Usage
=====

Parsing
-------

.. code-block:: python

    import sys
    import logging

    from lxml import etree

    import ooxml
    from ooxml import parse, serialize, importer

    logging.basicConfig(filename='ooxml.log', level=logging.INFO)

    if len(sys.argv) > 1:
        file_name = sys.argv[1]

        dfile = ooxml.read_from_file(file_name)

        print serialize.serialize(dfile.document)
        print serialize.serialize_styles(dfile.document)    

Indices and tables
==================

* :ref:`genindex`
* :ref:`modindex`
* :ref:`search`

