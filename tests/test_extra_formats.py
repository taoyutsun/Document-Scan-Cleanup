import sys,unittest,tempfile,io,copy,json
from pathlib import Path
from zipfile import ZipFile
from email.message import EmailMessage
from email.parser import BytesParser
from email import policy
import pikepdf
import pypdfium2 as pdfium
from PIL import ImageChops,ImageDraw
sys.path.insert(0,str(Path(__file__).resolve().parents[1]))
import document_cleanup_core as c
import fixtures

def pdf_bytes(content):
    doc=pikepdf.Pdf.new();page=doc.add_blank_page(page_size=(400,300))
    page.Resources=pikepdf.Dictionary(Font=pikepdf.Dictionary(F1=doc.make_indirect(pikepdf.Dictionary(Type=pikepdf.Name('/Font'),Subtype=pikepdf.Name('/Type1'),BaseFont=pikepdf.Name('/Helvetica')))))
    page.Contents=pikepdf.Stream(doc,content)
    out=io.BytesIO();doc.save(out);doc.close();return out.getvalue()

class ExtraTests(unittest.TestCase):
    def setUp(self):self.temp=tempfile.TemporaryDirectory();self.p=Path(self.temp.name);fixtures.make(self.p)
    def tearDown(self):self.temp.cleanup()
    def xlsx(self,change):
        data,_=c.package((self.p/'fixtures/hidden.xlsx').read_bytes());change(data);p=self.p/'book.xlsx'
        with ZipFile(p,'w') as z:
            for n,b in data.items():z.writestr(n,b)
        return p
    def test_xlsx_shared_text_selected_cell_only(self):
        def modify(d):d['xl/worksheets/sheet1.xml']=d['xl/worksheets/sheet1.xml'].replace(b'</row>',b'<c r="C1" t="s"><v>0</v></c><c r="D1"><f>1+2</f><v>3</v></c></row>')
        p=self.xlsx(modify);r=c.scan_file(p);fid=next(f['id'] for f in r['findings'] if f['path']=='cell:B1')
        result=c.clean_file(p,self.p/'out',[fid]);d,_=c.package(Path(result['output']).read_bytes());root=c.old.parse(d['xl/worksheets/sheet1.xml'])
        cells={e.get('r'):e for e in root.iter('{'+c.S+'}c')}
        self.assertEqual(cells['B1'].get('t'),'inlineStr');self.assertEqual(cells['C1'].get('t'),'s')
        self.assertEqual(cells['D1'].find('{'+c.S+'}f').text,'1+2')
        self.assertTrue(any(f['location'].endswith('C1') for f in result['after']['findings']))
    def test_xlsx_protected_sheet_and_formula_readonly(self):
        def modify(d):d['xl/worksheets/sheet1.xml']=d['xl/worksheets/sheet1.xml'].replace(b'</worksheet>',b'<sheetProtection sheet="1"/></worksheet>')
        p=self.xlsx(modify);r=c.scan_file(p)
        self.assertFalse(any(f['cleanable'] for f in r['findings']))
        with self.assertRaises(c.DocumentError):c.clean_file(p,self.p/'out')
    def test_eml_cleanup_keeps_attachment_and_addresses(self):
        m=EmailMessage();m['From']='author@example.invalid';m['To']='external@example.invalid';m['Subject']='Confidential business'
        m.set_content('Public line\nConfidential disclaimer\nKeep line')
        m.add_attachment(b'Confidential original',maintype='application',subtype='octet-stream',filename='original.txt')
        p=self.p/'mail.eml';p.write_bytes(m.as_bytes(policy=policy.SMTP));result=c.clean_file(p,self.p/'out')
        after=BytesParser(policy=policy.default).parsebytes(Path(result['output']).read_bytes())
        self.assertEqual(str(after['From']),str(m['From']));self.assertEqual(next(after.iter_attachments()).get_payload(decode=True),b'Confidential original')
        self.assertIn('Keep line',after.get_body(preferencelist=('plain',)).get_content())
        self.assertTrue(all(not f['cleanable'] for f in result['after']['findings']))
    def test_eml_html_split_word_and_styles(self):
        m=EmailMessage();m['Subject']='Business';m.set_content('<html><head><style>p{color:red}</style></head><body><p style="font-size:12px">Confi<b><i>dential</i></b> note</p><p>Keep</p></body></html>',subtype='html')
        p=self.p/'html.eml';p.write_bytes(m.as_bytes());result=c.clean_file(p,self.p/'out');after=BytesParser(policy=policy.default).parsebytes(Path(result['output']).read_bytes())
        text=after.get_content();self.assertEqual(result['after']['match_count'],0);self.assertIn('font-size:12px',text);self.assertIn('p{color:red}',text);self.assertIn('Keep',text)
    def test_eml_dkim_refused(self):
        m=EmailMessage();m['Subject']='Confidential';m['DKIM-Signature']='v=1; fake=test';m.set_content('Confidential')
        p=self.p/'signed.eml';p.write_bytes(m.as_bytes());self.assertFalse(c.scan_file(p)['can_clean'])
        with self.assertRaises(c.DocumentError):c.clean_file(p,self.p/'out')
    def test_eml_block_removes_line_only(self):
        m=EmailMessage();m['Subject']='Business';m.set_content('Keep before\nConfidential long disclaimer\nKeep after')
        p=self.p/'block.eml';p.write_bytes(m.as_bytes());result=c.clean_file(p,self.p/'out',mode='block')
        text=BytesParser(policy=policy.default).parsebytes(Path(result['output']).read_bytes()).get_content()
        self.assertIn('Keep before',text);self.assertIn('Keep after',text);self.assertNotIn('long disclaimer',text)
    def test_pdf_removes_text_without_moving_following_text(self):
        raw=pdf_bytes(b'BT /F1 16 Tf 1 0 0 1 40 200 Tm 1 Tc 3 Tw (Confidential ) Tj (Keep after) Tj ET BT /F1 14 Tf 1 0 0 1 40 100 Tm (Keep lower) Tj ET')
        p=self.p/'sample.pdf';p.write_bytes(raw);r=c.clean_file(p,self.p/'out');self.assertEqual(r['after']['match_count'],0)
        with pdfium.PdfDocument(raw) as a,pdfium.PdfDocument(r['output']) as b:
            pa=a[0];pb=b[0];ba=pa.render(scale=2);bb=pb.render(scale=2);ia=ba.to_pil().convert('RGB');ib=bb.to_pil().convert('RGB')
            diff=ImageChops.difference(ia,ib);draw=ImageDraw.Draw(diff)
            for x0,y0,x1,y1 in r['before']['findings'][0]['bounds']:draw.rectangle((x0*2-3,(300-y1)*2-3,x1*2+3,(300-y0)*2+3),fill=0)
            self.assertIsNone(diff.getbbox());ba.close();bb.close();pa.close();pb.close()
        with pikepdf.open(r['output']) as doc:self.assertNotIn(b'Confidential',doc.pages[0].Contents.read_bytes())
    def test_pdf_split_tj_and_metadata(self):
        raw=pdf_bytes(b'BT /F1 14 Tf 40 100 Td [(Confi) -10 (dential)] TJ ET')
        out=io.BytesIO()
        with pikepdf.open(io.BytesIO(raw)) as doc:doc.docinfo['/Title']='Confidential title';doc.save(out)
        p=self.p/'split.pdf';p.write_bytes(out.getvalue());r=c.clean_file(p,self.p/'out');self.assertEqual(r['before']['match_count'],2);self.assertEqual(r['after']['match_count'],0)
    def test_pdf_signed_and_encrypted_refused(self):
        raw=pdf_bytes(b'BT /F1 14 Tf 40 100 Td (Confidential) Tj ET');out=io.BytesIO()
        with pikepdf.open(io.BytesIO(raw)) as doc:doc.Root.AcroForm=pikepdf.Dictionary(Fields=pikepdf.Array([doc.make_indirect(pikepdf.Dictionary(FT=pikepdf.Name('/Sig')))]));doc.save(out)
        p=self.p/'signed.pdf';p.write_bytes(out.getvalue());self.assertFalse(c.scan_file(p)['can_clean'])
        with self.assertRaises(c.DocumentError):c.clean_file(p,self.p/'out')
        out=io.BytesIO()
        with pikepdf.open(io.BytesIO(raw)) as doc:doc.save(out,encryption=pikepdf.Encryption(owner='owner',user=''))
        p=self.p/'encrypted.pdf';p.write_bytes(out.getvalue());self.assertFalse(c.scan_file(p)['can_clean'])
    def test_text_encodings_and_structured_formats(self):
        for i,encoding in enumerate(['utf-8','utf-8-sig','utf-16-le','utf-16-be','cp950']):
            text=('\ufeff' if encoding.startswith('utf-16') else '')+'Confidential 測試\r\nKeep\r\n';p=self.p/f'e{i}.md';p.write_bytes(text.encode(encoding))
            r=c.clean_file(p,self.p/f'out{i}');self.assertEqual(Path(r['output']).read_bytes(),text.replace('Confidential','').encode(encoding))
        p=self.p/'data.csv';p.write_text('Title,Value\n"Confidential note",12\n',encoding='utf-8');r=c.clean_file(p,self.p/'csv')
        self.assertEqual(Path(r['output']).read_text(encoding='utf-8'),'Title,Value\n" note",12\n')
        with self.assertRaises(c.DocumentError):c.clean_file(p,self.p/'badcsv',mode='block')
        p=self.p/'data.json';p.write_text('{"note":"Confidential"}',encoding='utf-8');self.assertFalse(c.scan_file(p)['can_clean'])
    def test_legacy_binary_refusal_before_office(self):
        p=self.p/'bad.doc';p.write_bytes(b'not a binary Office file')
        with self.assertRaises(c.DocumentError):c.scan_file(p)
    def test_author_metadata(self):
        import app_metadata as m
        self.assertEqual(m.AUTHOR,'Arthur Tao');self.assertEqual(m.LICENSE_NAME,'MIT');self.assertTrue((m.SOURCE_DIRECTORY/'LICENSE').exists())

if __name__=='__main__':unittest.main(verbosity=2)
