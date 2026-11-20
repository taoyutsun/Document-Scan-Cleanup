"""Check per-format routing and cleanup of mixed Office attachments."""
import tempfile
import unittest
from pathlib import Path

import fixtures
import office_annotations as annotations
import office_batch as batch
import office_merge
import ppt_batch
import ppt_structure as opc


class OfficeBatchTests(unittest.TestCase):
    def test_merge_keeps_word_and_excel_as_separate_jobs(self):
        files = ['1.pptx', 'word.docx', '2.pptx', 'book.xlsx', 'extra.pdf']
        tasks, skipped = batch.plan(files, batch.Actions(merge=True, word_comments=True, excel_notes=True), {})
        self.assertEqual([(job.kind, job.paths, job.merge) for job in tasks],
                         [('PowerPoint', ('1.pptx', '2.pptx'), True), ('Word', ('word.docx',), False),
                          ('Excel', ('book.xlsx',), False)])
        self.assertEqual(skipped[0]['source'], ['extra.pdf'])
        tasks, _ = batch.plan(['1.pptx', 'word.docx'], batch.Actions(merge=True, word_comments=True), {})
        self.assertEqual([job.kind for job in tasks], ['Word'])

    def test_word_and_excel_merge_options_are_independent(self):
        paths = ['b.docx', 'b.xlsx', 'a.docx', 'a.xlsx', 'p.pptx']
        tasks, skipped = batch.plan(paths, batch.Actions(word_merge=True, excel_merge=True), {})
        self.assertEqual(tasks, [batch.Task(('b.docx','a.docx'), 'Word', True),
                                 batch.Task(('b.xlsx','a.xlsx'), 'Excel', True)])
        self.assertEqual(skipped[0]['source'], ['p.pptx'])
        tasks, skipped = batch.plan(paths + ['q.pptx'], batch.Actions(merge=True, word_merge=True, excel_merge=True), {})
        self.assertEqual([task.kind for task in tasks], ['Word','Excel','PowerPoint'])
        self.assertTrue(all(task.merge for task in tasks))
        self.assertFalse(skipped)

    def test_colliding_excel_names_do_not_take_a_later_sheet_name(self):
        def workbook(*names):
            sheets = ''.join(f'<sheet name="{name}"/>' for name in names)
            return {'xl/workbook.xml': f'<workbook xmlns="{fixtures.S}"><sheets>{sheets}</sheets></workbook>'.encode()}
        first = workbook('Summary')
        next_book = workbook('Summary', 'Summary (2)')
        self.assertEqual(office_merge.excel_name_plan([first, next_book]),
                         [['Summary'], ['Summary (3)','Summary (2)']])

    def test_word_removes_comments_from_body_and_header_keeps_revisions(self):
        with tempfile.TemporaryDirectory() as temp:
            folder = Path(temp)
            fixtures.make(folder)
            source = folder / 'sample-template.docx'
            data = batch.read_office(source.read_bytes(), '.docx', editable=True)
            W = fixtures.W
            data['word/document.xml'] = f'''<w:document xmlns:w="{W}"><w:body><w:p>
                <w:commentRangeStart w:id="0"/><w:ins w:id="1"><w:r><w:t>Current text</w:t></w:r></w:ins>
                <w:commentRangeEnd w:id="0"/><w:r><w:commentReference w:id="0"/></w:r>
                </w:p></w:body></w:document>'''.encode()
            data['word/footer1.xml'] = f'''<w:ftr xmlns:w="{W}"><w:p>
                <w:r><w:commentReference w:id="0"/></w:r></w:p></w:ftr>'''.encode()
            data['word/comments.xml'] = f'<w:comments xmlns:w="{W}"><w:comment w:id="0"/></w:comments>'.encode()
            types = opc.parse(data['[Content_Types].xml'])
            types.append(opc.parse(f'<Override xmlns="{fixtures.CT}" PartName="/word/comments.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.comments+xml"/>'.encode()))
            data['[Content_Types].xml'] = opc.xml(types)
            rels = opc.parse(data['word/_rels/document.xml.rels'])
            rels.append(opc.parse(f'<Relationship xmlns="{opc.REL}" Id="rComment" Type="{opc.R}/comments" Target="comments.xml"/>'.encode()))
            data['word/_rels/document.xml.rels'] = opc.xml(rels)
            source.write_bytes(ppt_batch.write_package(data))
            before = source.read_bytes()
            task = batch.Task((source,), 'Word')
            report = batch.process(task, folder/'out', batch.Actions(word_comments=True), {})
            result = batch.read_office(Path(report['output']).read_bytes(), '.docx', editable=True)
            self.assertEqual(report['after']['revision_marks'], 1)
            self.assertEqual(report['after']['comments'], 0)
            self.assertNotIn('word/comments.xml', result)
            self.assertNotIn(b'commentReference', result['word/footer1.xml'])
            self.assertIn(b'Current text', result['word/document.xml'])
            self.assertEqual(source.read_bytes(), before)

    def test_excel_removes_note_shape_without_changing_other_vml_shapes(self):
        with tempfile.TemporaryDirectory() as temp:
            folder = Path(temp)
            fixtures.make(folder)
            source = next(folder.rglob('hidden.xlsx'))
            data = batch.read_office(source.read_bytes(), '.xlsx', editable=True)
            sheet = 'xl/worksheets/sheet1.xml'
            notes = 'xl/comments1.xml'
            drawing = 'xl/drawings/vmlDrawing1.vml'
            data[notes] = f'''<comments xmlns="{fixtures.S}"><authors><author>Example</author></authors>
                <commentList><comment ref="A1" authorId="0"><text><t>Note</t></text></comment></commentList></comments>'''.encode()
            data[drawing] = b'''<xml xmlns:v="urn:schemas-microsoft-com:vml" xmlns:x="urn:schemas-microsoft-com:office:excel">
                <v:shape id="note"><x:ClientData ObjectType="Note"><x:Row>0</x:Row><x:Column>0</x:Column></x:ClientData></v:shape>
                <v:shape id="control"><x:ClientData ObjectType="Button"/></v:shape></xml>'''
            data[opc.relpath(sheet)] = fixtures.relationships([
                ('rNote','comments','../comments1.xml'), ('rVml','vmlDrawing','../drawings/vmlDrawing1.vml'),
            ]).encode()
            types = opc.parse(data['[Content_Types].xml'])
            types.append(opc.parse(f'<Override xmlns="{fixtures.CT}" PartName="/{sheet}" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>'.encode()))
            types.append(opc.parse(f'<Override xmlns="{fixtures.CT}" PartName="/{notes}" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.comments+xml"/>'.encode()))
            data['[Content_Types].xml'] = opc.xml(types)
            original_sheet = data[sheet]
            annotations.remove_excel_annotations(data, threaded=False, notes=True)
            self.assertEqual(data[sheet], original_sheet)
            self.assertNotIn(notes, data)
            self.assertNotIn(b'id="note"', data[drawing])
            self.assertIn(b'id="control"', data[drawing])
            self.assertEqual(annotations.hidden_sheets(data), [{'name':'Hidden','state':'hidden'}])
