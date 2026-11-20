"""Word metadata may be renumbered on save, but its contents must survive."""
import tempfile
import unittest
from pathlib import Path

import fixtures
import office_batch
import word_custom_xml as custom


def store(number, text):
    identifier = '{11111111-0000-0000-0000-%012d}' % number
    properties = f'<ds:datastoreItem xmlns:ds="{custom.DATASTORE[1:-1]}" ds:itemID="{identifier}"/>'
    return custom.Store(f'<example>{text}</example>'.encode(),properties.encode(),'application/xml',
                        'application/vnd.openxmlformats-officedocument.customXmlProperties+xml')


class WordCustomXmlTests(unittest.TestCase):
    def setUp(self):
        directory = tempfile.TemporaryDirectory()
        self.addCleanup(directory.cleanup)
        folder = Path(directory.name)
        fixtures.make(folder)
        self.base = office_batch.read_office((folder/'sample-template.docx').read_bytes(), '.docx', editable=True)

    def test_change_of_numbering_does_not_change_store_identity(self):
        before = self.base.copy()
        after = self.base.copy()
        custom.attach_stores(before, [store(1,'First'),store(2,'Second')])
        custom.attach_stores(after, [store(2,'Second'),store(1,'First')])
        self.assertNotEqual(before['customXml/item1.xml'],after['customXml/item1.xml'])
        office_batch.verify_protected(before, after)
        after['customXml/item1.xml'] = b'<example>Changed</example>'
        with self.assertRaisesRegex(ValueError, '自訂 XML'):
            office_batch.verify_protected(before, after)

    def test_merge_restores_second_source_metadata_with_relationships(self):
        first = self.base.copy()
        second = self.base.copy()
        custom.attach_stores(first, [store(1,'First')])
        custom.attach_stores(second, [store(2,'Second')])
        merge = first.copy()
        expected = custom.merge_plan([first,second])
        custom.preserve_after_merge(merge,expected,[first,second])
        custom.verify_same(expected,custom.stores(merge))
        office_batch.core.validate_package(merge, 'docx')
        office_batch.verify_protected(first,merge,word_stores=expected)
        # The semantic custom XML comparison never weakens the label check.
        merge['docMetadata/LabelInfo.xml'] = b'<example/>'
        with self.assertRaisesRegex(ValueError, '敏感度標籤資料'):
            office_batch.verify_protected(first,merge,word_stores=expected)

    def test_conflicting_store_ids_are_refused(self):
        first = self.base.copy()
        second = self.base.copy()
        custom.attach_stores(first, [store(1,'First')])
        custom.attach_stores(second, [store(1,'Different')])
        with self.assertRaisesRegex(ValueError, '識別碼'):
            custom.merge_plan([first,second])
