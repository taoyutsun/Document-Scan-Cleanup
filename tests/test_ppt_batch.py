"""Verify selected cleanup across classic and modern PowerPoint parts."""
import tempfile
import unittest
from pathlib import Path
from unittest.mock import patch

from lxml import etree

import document_cleanup_core as core
import fixtures
import ppt_batch as batch
import ppt_structure as opc


class PptBatchTests(unittest.TestCase):
    def setUp(self):
        directory = tempfile.TemporaryDirectory()
        self.addCleanup(directory.cleanup)
        self.folder = Path(directory.name)
        fixtures.make(self.folder)
        self.source = self.folder / 'sample-template.pptx'
        self.data = batch.read_presentation(self.source.read_bytes())

    def add_part(self, owner, name, kind, relationship, content, explicit=False):
        self.data[name] = content.encode()
        types = opc.parse(self.data['[Content_Types].xml'])
        etree.SubElement(types, '{' + fixtures.CT + '}Override', PartName='/' + name, ContentType=kind)
        self.data['[Content_Types].xml'] = opc.xml(types)
        rel_path = opc.relpath(owner)
        rels = opc.parse(self.data[rel_path])
        rid = f'rAnnotation{len(rels)}'
        etree.SubElement(rels, '{' + opc.REL + '}Relationship', Id=rid, Type=relationship, Target='/' + name)
        self.data[rel_path] = opc.xml(rels)
        if explicit:
            slide = opc.parse(self.data[owner])
            extensions = etree.SubElement(slide, '{' + opc.P + '}extLst')
            extension = etree.SubElement(extensions, '{' + opc.P + '}ext', uri='{G}')
            etree.SubElement(extension, batch.MODERN_COMMENT, {batch.RID: rid})
            self.data[owner] = opc.xml(slide)

    def test_classic_and_modern_comment_removal_preserves_notes(self):
        notes = 'ppt/notesSlides/notesSlide1.xml'
        classic = 'ppt/comments/comment1.xml'
        modern = 'ppt/comments/modern.xml'
        authors = 'ppt/authors.xml'
        slide = 'ppt/slides/slide1.xml'
        self.add_part(slide, notes, next(iter(batch.NOTES)), opc.R + '/notesSlide', f'<p:notes xmlns:p="{opc.P}"/>')
        namespace = 'http://schemas.microsoft.com/office/powerpoint/2018/8/main'
        relationship = 'http://schemas.microsoft.com/office/2018/10/relationships'
        self.add_part(slide, classic, 'application/vnd.openxmlformats-officedocument.presentationml.comments+xml',
                      opc.R + '/comments', f'<p:cmLst xmlns:p="{opc.P}"/>')
        self.add_part(slide, modern, 'application/vnd.ms-powerpoint.comments+xml', relationship + '/comments',
                      f'<p:cmLst xmlns:p="{namespace}"/>', explicit=True)
        self.add_part(batch.PRESENTATION, authors, 'application/vnd.ms-powerpoint.authors+xml', relationship + '/authors',
                      f'<p:authorLst xmlns:p="{namespace}"/>')
        self.source.write_bytes(batch.write_package(self.data))
        original = self.source.read_bytes()
        result = batch.process([self.source], self.folder / 'comments', batch.Options(comments=True))
        output = batch.read_presentation(Path(result['output']).read_bytes())
        self.assertEqual(output[notes], self.data[notes])
        self.assertEqual(output['ppt/media/image1.png'], self.data['ppt/media/image1.png'])
        self.assertFalse({classic, modern, authors} & set(output))
        self.assertIsNone(opc.parse(output[slide]).find('.//' + batch.MODERN_COMMENT))
        self.assertEqual(self.source.read_bytes(), original)

    def test_checked_keywords_and_stale_source(self):
        report = core.scan_file(self.source, ['Public slide'])
        selected = next(f for f in report['findings'] if f['cleanable'])
        job = batch.KeywordSelection((selected['id'],), ('Public slide',), report['source_sha256'], 'keyword', False)
        result = batch.process([self.source], self.folder / 'combined', batch.Options(notes=True), keywords={self.source: job})
        self.assertEqual(result['keyword_items'], 1)
        after = core.scan_file(result['output'], ['Public slide'])
        self.assertEqual(after['match_count'], report['match_count'] - 1)
        self.source.write_bytes(self.source.read_bytes() + b'changed after scan')
        with self.assertRaisesRegex(core.DocumentError, '掃描後已變更'):
            batch.process([self.source], self.folder / 'stale', batch.Options(notes=True), keywords={self.source: job})
        self.assertFalse((self.folder / 'stale').exists())

    def label_fixture(self, enabled, removed, label_id='synthetic-label'):
        part = 'docMetadata/LabelInfo.xml'
        self.add_part('', part, 'application/vnd.ms-office.classificationlabels+xml', batch.LABEL_REL,
                      f'<labelList xmlns="{batch.LABEL_NS[1:-1]}"><label id="{label_id}" siteId="synthetic-site" '
                      f'enabled="{enabled}" removed="{removed}" method="Standard"/></labelList>')
        return self.data.copy()

    def test_removed_label_is_accepted_and_active_label_selects_base(self):
        clean = self.data.copy()
        removed = self.label_fixture('0', '1')
        self.assertFalse(batch.check_native_input(removed))
        self.data = clean.copy()
        active = self.label_fixture('true', 'false')
        paths = [Path('first.pptx'), Path('second.pptx')]
        base, expected = batch.native_plan([removed, active], paths)
        self.assertEqual(base, 1)
        self.assertEqual(expected, batch.active_labels(active))
        self.assertEqual(len(expected), 1)
        self.data = clean.copy()
        different = self.label_fixture('1', '0', 'different-label')
        with self.assertRaisesRegex(ValueError, '標籤不一致'):
            batch.native_plan([active, different], paths)

    def test_native_output_that_loses_a_label_is_not_published(self):
        unlabelled = self.data.copy()
        self.label_fixture('1', '0')
        slide = 'ppt/slides/slide1.xml'
        root = opc.parse(self.data[slide])
        root.set('show', '0')
        self.data[slide] = opc.xml(root)
        self.source.write_bytes(batch.write_package(self.data))
        with patch.object(batch, 'native_transform', return_value=unlabelled), patch.object(
                batch, 'inspect', side_effect=[{'slides': 6, 'hidden': 1}, {'slides': 5, 'hidden': 0}]):
            with self.assertRaisesRegex(ValueError, '另存時變更了有效敏感度標籤'):
                batch.process([self.source], self.folder / 'withheld', batch.Options(hidden=True))
        self.assertFalse((self.folder / 'withheld').exists())
