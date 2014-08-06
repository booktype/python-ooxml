import unittest
import six

from ooxml.parse import parse_relationship

content_valid = six.b('''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId3" Type="http://schemas.microsoft.com/office/2007/relationships/stylesWithEffects" Target="stylesWithEffects.xml"/><Relationship Id="rId4" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/settings" Target="settings.xml"/><Relationship Id="rId5" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/webSettings" Target="webSettings.xml"/><Relationship Id="rId6" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="media/image1.jpeg"/><Relationship Id="rId7" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/fontTable" Target="fontTable.xml"/><Relationship Id="rId8" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme" Target="theme/theme1.xml"/><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/numbering" Target="numbering.xml"/><Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/></Relationships>''')
content_external = six.b('''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/settings" Target="settings.xml"/><Relationship Id="rId4" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/webSettings" Target="webSettings.xml"/><Relationship Id="rId5" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink" Target="http://www.google.com/" TargetMode="External"/><Relationship Id="rId6" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/fontTable" Target="fontTable.xml"/><Relationship Id="rId7" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme" Target="theme/theme1.xml"/><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/><Relationship Id="rId2" Type="http://schemas.microsoft.com/office/2007/relationships/stylesWithEffects" Target="stylesWithEffects.xml"/></Relationships>''')
content_invalid = six.b('<>')

# document mockup
class Document:
    def __init__(self):
        self.relationships = {}


class TestParseRelationship(unittest.TestCase):
    def setUp(self):
        self.document = Document()

    def tearDown(self):
        pass

    def _assertRelationship(self, r_id, r_type, r_target, r_mode):
        self.assertEqual(self.document.relationships[r_id]['type'], r_type)
        self.assertEqual(self.document.relationships[r_id]['target'], r_target)
        self.assertEqual(self.document.relationships[r_id]['target_mode'], r_mode)

    def test_parse(self):
        "Make sure we parse entire relationship file."

        parse_relationship(self.document, content_valid)

        self.assertEqual(len(self.document.relationships), 8)

    def test_parse_invalid(self):
        "Parsing invalid content."

        from lxml import etree 

        with self.assertRaises(etree.XMLSyntaxError):
            parse_relationship(self.document, content_invalid)

    def test_parse_external(self):
        "Did we parse External reference"

        parse_relationship(self.document, content_external)

        self._assertRelationship('rId5', 
            'http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink',
            'http://www.google.com/', 
            'External')


    def test_read_values(self):
        "Make sure we read all the values for rId6."

        parse_relationship(self.document, content_valid)

        self.assertEqual(len(self.document.relationships['rId6'].keys()), 3)

    def test_read_content(self):
        "Make sure content for rId6 is correct"

        parse_relationship(self.document, content_valid)

        self._assertRelationship('rId6', 
            'http://schemas.openxmlformats.org/officeDocument/2006/relationships/image',
            'media/image1.jpeg',
            'Internal')



if __name__ == '__main__':
    unittest.main()
