import unittest

from mock import patch, call, Mock, MagicMock, ANY

from ooxml.serialize import serialize_elements, serialize_break, serialize_link

from lxml import etree

def _render(elem):
    return etree.tostring(elem, pretty_print=False, encoding="utf-8", xml_declaration=False)

# serialize_link


def _create_relationship(relationship):
    def _getitem(m):
        return relationship

    relationships = MagicMock(spec_set=dict)
    relationships.__getitem__.side_effect = _getitem

    return relationships


def _create_context(MockContext):
    instance = MockContext.return_value
    instance.get_hook.return_value = None

    return instance


def _create_elem(elements):
    def _getitem(n):
        return elements[n]

    elem = MagicMock()
    elem.elements = MagicMock(spec_set=dict)
    elem.elements.__getitem__.side_effect = _getitem
    elem.elements.__len__.return_value = len(elements)

    return elem


class TestSerializeLinkElements(unittest.TestCase):
    def setUp(self):
        self.doc = Mock()
        self.doc.relationships = _create_relationship({'target': 'http://www.google.com/'})

        self.root = etree.Element('div')

    @patch('ooxml.serialize.fire_hooks')
    @patch('ooxml.serialize.Context')    
    def test_creation_empty(self, MockContext, fire_mock):
        context = _create_context(MockContext)

        elem = _create_elem([])

        ret = serialize_link(context, self.doc, elem, self.root)
        self.assertEqual(ret, self.root)        

    @patch('ooxml.serialize.fire_hooks')
    @patch('ooxml.serialize.Context')    
    def test_creation_one(self, MockContext, fire_mock):
        context = _create_context(MockContext)

        link = _create_elem([])
        link.value.return_value = 'link'

        elem = _create_elem([link])

        ret = serialize_link(context, self.doc, elem, self.root)
        self.assertEqual(ret, self.root)        

    @patch('ooxml.serialize.fire_hooks')
    @patch('ooxml.serialize.Context')    
    def test_content(self, MockContext, fire_mock):
        context = _create_context(MockContext)

        link = _create_elem([])
        link.value.return_value = 'link'

        elem = _create_elem([link])

        ret = serialize_link(context, self.doc, elem, self.root)
        self.assertEqual(_render(ret), '<div><a href="http://www.google.com/">link</a></div>')

    @patch('ooxml.serialize.fire_hooks')
    @patch('ooxml.serialize.Context')    
    def test_hooks(self, MockContext, fire_mock):
        context = _create_context(MockContext)

        link = _create_elem([])
        link.value.return_value = 'link'

        elem = _create_elem([link])

        serialize_link(context, self.doc, elem, self.root)
        fire_mock.assert_called_once_with(context, self.doc, ANY, None)


# serialize_break

class TestSerializeBreakElements(unittest.TestCase):
    def setUp(self):
        self.doc = Mock()
        self.root = etree.Element('div')

    @patch('ooxml.serialize.fire_hooks')
    @patch('ooxml.serialize.Context')    
    def test_creation(self, MockContext, fire_mock):
        instance = MockContext.return_value
        instance.get_hook.return_value = None

        ret = serialize_break(instance, self.doc, None, self.root)
        self.assertEqual(ret, self.root)        

    @patch('ooxml.serialize.fire_hooks')
    @patch('ooxml.serialize.Context')    
    def test_content(self, MockContext, fire_mock):
        instance = MockContext.return_value
        instance.get_hook.return_value = None

        ret = serialize_break(instance, self.doc, None, self.root)
        self.assertEqual(_render(ret), '<div><span style="page-break-after: always;"/></div>')

    @patch('ooxml.serialize.fire_hooks')
    @patch('ooxml.serialize.Context')    
    def test_fire_hooks(self, MockContext, fire_mock):
        instance = MockContext.return_value
        instance.get_hook.return_value = None

        serialize_break(instance, self.doc, None, self.root)
        fire_mock.assert_called_once_with(instance, self.doc, ANY, None)


# serialize_elements

class TestSerializeElements(unittest.TestCase):
    def setUp(self):
        self.doc = Mock()

    def tearDown(self):
        pass

    @patch('ooxml.serialize.Context')
    def test_serialize(self, MockContext):
        instance = MockContext.return_value
        instance.get_serializer.return_value = None

        self.assertEqual(serialize_elements(self.doc, [1, 2, 3]), "<div/>\n")
        self.assertEqual(serialize_elements(self.doc, []), "<div/>\n")

    @patch('ooxml.serialize.Context')
    def test_serialize_something(self, MockContext):
        def _func(ctx, document, elem, root):
            return etree.SubElement(root, 'p')

        instance = MockContext.return_value
        instance.get_serializer.return_value = _func

        self.assertEqual(serialize_elements(self.doc, [1]), "<div>\n  <p/>\n</div>\n")
        instance.get_serializer.assert_called_with(1)

    @patch('ooxml.serialize.Context')
    def test_serialize_calls(self, MockContext):
        instance = MockContext.return_value
        instance.get_serializer.return_value = None

        serialize_elements(self.doc, [1,2,3])
        self.assertEqual(instance.get_serializer.call_args_list, [call(1), call(2), call(3)])


if __name__ == '__main__':
    unittest.main()
