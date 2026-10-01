import sys,json,tempfile,unittest,copy
from pathlib import Path
from zipfile import ZipFile,ZIP_DEFLATED
from email.message import EmailMessage
from lxml import etree as E

PROJECT=Path(__file__).resolve().parents[1]
sys.path.insert(0,str(PROJECT))
import document_cleanup_core as c
import document_cleanup_app as app
import fixtures
_fixture_temp=tempfile.TemporaryDirectory()
BASE=Path(_fixture_temp.name)
fixtures.make(BASE)

def save_package(path,data):
    with ZipFile(path,'w',ZIP_DEFLATED) as z:
        for n,b in data.items():z.writestr(n,b)

class Tests(unittest.TestCase):
    def setUp(self):self.temp=tempfile.TemporaryDirectory();self.p=Path(self.temp.name)
    def tearDown(self):self.temp.cleanup()
    def make_variant(self,source,name,mutate):
        data,_=c.package(source.read_bytes());mutate(data);out=self.p/name;save_package(out,data);return out

    def test_docx_all_words_and_styles(self):
        p=BASE/'sample-template.docx';before=c.digest(p.read_bytes());r=c.clean_file(p,self.p)
        self.assertGreater(r['before']['match_count'],0);self.assertEqual(r['after']['match_count'],0)
        self.assertEqual(before,c.digest(p.read_bytes()))
        old,_=c.package(p.read_bytes());new,_=c.package(Path(r['output']).read_bytes())
        self.assertEqual(old['word/document.xml'],new['word/document.xml'])
        self.assertEqual(old['word/settings.xml'],new['word/settings.xml'])
        for n in old:
            if n.startswith('word/media/'):self.assertEqual(old[n],new[n])
        styles=c.old.parse(new['word/styles.xml']);ids={e.get('{'+c.W+'}styleId') for e in styles if c.local(e)=='style'}
        for n,b in new.items():
            if n.startswith('word/') and n.endswith('.xml'):
                for e in c.old.parse(b).iter():
                    if c.local(e) in {'pStyle','rStyle','link'}:
                        v=e.get('{'+c.W+'}val')
                        if v and v.startswith('CleanStyle_'):self.assertIn(v,ids)

    def test_pptx_all_and_pruning(self):
        p=BASE/'sample-template.pptx';before=c.digest(p.read_bytes());r=c.clean_file(p,self.p,prune=True)
        self.assertEqual(r['after']['match_count'],0);self.assertEqual(r['pruning']['after']['slides'],6)
        self.assertEqual(r['pruning']['after']['masters'],1)
        self.assertEqual(before,c.digest(p.read_bytes()))
        d,_=c.package(Path(r['output']).read_bytes());c.old.validate(d)

    def test_selected_only_preserves_unselected(self):
        p=BASE/'sample-template.pptx';r=c.scan_file(p)
        f=next(f for f in r['findings'] if f['category']=='文字段落' and f['cleanable'])
        out=c.clean_file(p,self.p,[f['id']])
        self.assertGreater(out['after']['match_count'],0)
        self.assertEqual(len(out['changes']),1)

    def test_split_run_and_block(self):
        def mutate(data):
            root=c.old.parse(data['word/document.xml']);body=root.find('{'+c.W+'}body')
            p=E.Element('{'+c.W+'}p')
            for text in ['Keep ','Confi','dential',' tail']:
                r=E.SubElement(p,'{'+c.W+'}r');E.SubElement(r,'{'+c.W+'}t').text=text
            body.insert(0,p);data['word/document.xml']=c.old.xml(root)
        p=self.make_variant(BASE/'sample-template.docx','split.docx',mutate)
        r=c.scan_file(p);f=next(f for f in r['findings'] if f['part']=='word/document.xml')
        first=c.clean_file(p,self.p/'one',[f['id']]);data,_=c.package(Path(first['output']).read_bytes())
        text=''.join(c.old.parse(data['word/document.xml']).itertext());self.assertIn('Keep  tail',text)
        second=c.clean_file(p,self.p/'two',[f['id']],mode='block');data,_=c.package(Path(second['output']).read_bytes())
        text=''.join(c.old.parse(data['word/document.xml']).itertext());self.assertNotIn('Keep',text)

    def test_enabled_word_protection_stops(self):
        def mutate(data):
            root=c.old.parse(data['word/settings.xml']);root.find('{'+c.W+'}documentProtection').set('{'+c.W+'}enforcement','1')
            data['word/settings.xml']=c.old.xml(root)
        p=self.make_variant(BASE/'sample-template.docx','locked.docx',mutate)
        r=c.scan_file(p);self.assertFalse(r['can_clean']);self.assertFalse(any(f['cleanable'] for f in r['findings']))
        with self.assertRaises(c.DocumentError):c.clean_file(p,self.p/'out')
        self.assertFalse((self.p/'out').exists())

    def test_formal_label_preserved(self):
        def mutate(data):
            ns='http://schemas.openxmlformats.org/officeDocument/2006/custom-properties'
            r=E.Element('{'+ns+'}Properties',nsmap={None:ns,'vt':c.old.VT})
            e=E.SubElement(r,'{'+ns+'}property',name='MSIP_Label_test_Name',pid='2',fmtid='{D5CDD505-2E9C-101B-9397-08002B2CF9AE}')
            E.SubElement(e,'{'+c.old.VT+'}lpwstr').text='Highly Confidential'
            data['docProps/custom.xml']=c.old.xml(r)
            r=c.old.parse(data['_rels/.rels']);E.SubElement(r,'{'+c.old.REL+'}Relationship',Id='rCustom',Type=c.old.R+'/custom-properties',Target='docProps/custom.xml');data['_rels/.rels']=c.old.xml(r)
            r=c.old.parse(data['[Content_Types].xml']);E.SubElement(r,'{http://schemas.openxmlformats.org/package/2006/content-types}Override',PartName='/docProps/custom.xml',ContentType='application/vnd.openxmlformats-officedocument.custom-properties+xml');data['[Content_Types].xml']=c.old.xml(r)
        p=self.make_variant(BASE/'sample-template.docx','labeled.docx',mutate)
        r=c.clean_file(p,self.p/'out');a,_=c.package(p.read_bytes());b,_=c.package(Path(r['output']).read_bytes())
        self.assertEqual(a['docProps/custom.xml'],b['docProps/custom.xml'])
        self.assertGreater(r['after']['match_count'],0)
        self.assertFalse(any(f['cleanable'] for f in r['after']['findings']))

    def test_fields_preserved(self):
        def mutate(data):
            r=c.old.parse(data['word/document.xml']);p=E.Element('{'+c.W+'}p');run=E.SubElement(p,'{'+c.W+'}r')
            E.SubElement(run,'{'+c.W+'}instrText').text='STYLEREF ConfidentialStyle';r.find('{'+c.W+'}body').insert(0,p);data['word/document.xml']=c.old.xml(r)
        p=self.make_variant(BASE/'sample-template.docx','field.docx',mutate);r=c.clean_file(p,self.p/'out')
        self.assertTrue(any(f['category']=='欄位指令（只掃描）' for f in r['after']['findings']))

    def test_bad_selection_no_output(self):
        with self.assertRaises(c.DocumentError):c.clean_file(BASE/'sample-template.docx',self.p/'out',['invalid'])
        self.assertFalse((self.p/'out').exists())

    def test_source_change_no_output(self):
        with self.assertRaises(c.DocumentError):c.clean_file(BASE/'sample-template.docx',self.p/'out',expected_hash='changed')
        self.assertFalse((self.p/'out').exists())

    def test_no_overwrite(self):
        r=c.clean_file(BASE/'sample-template.docx',self.p)
        h=c.digest(Path(r['output']).read_bytes())
        with self.assertRaises(c.DocumentError):c.clean_file(BASE/'sample-template.docx',self.p)
        self.assertEqual(c.digest(Path(r['output']).read_bytes()),h)

    def test_txt_bom_crlf_and_encoding(self):
        p=BASE/'fixtures/text.txt';r=c.clean_file(p,self.p)
        raw=Path(r['output']).read_bytes();self.assertTrue(raw.startswith(b'\xef\xbb\xbf'))
        self.assertIn(b'\r\n',raw);self.assertEqual(r['after']['match_count'],0)

    def test_xlsx_hidden_shared_rich_cell(self):
        r=c.scan_file(BASE/'fixtures/hidden.xlsx')
        self.assertTrue(any('B1' in f['location'] for f in r['findings']));self.assertTrue(r['can_clean'])
        clean=c.clean_file(BASE/'fixtures/hidden.xlsx',self.p)
        self.assertEqual(clean['after']['match_count'],0)

    def test_eml_and_attachment(self):
        r=c.scan_file(BASE/'fixtures/message.eml');self.assertEqual(r['match_count'],3)
        self.assertTrue(any('attachment.txt' in f['location'] for f in r['findings']))
        self.assertTrue(any(f['cleanable'] for f in r['findings']))
        self.assertFalse(any(f['cleanable'] for f in r['findings'] if '附件' in f['location']))

    def test_signed_package(self):
        p=self.make_variant(BASE/'sample-template.docx','signed.docx',lambda data:data.update({'_xmlsignatures/sig1.xml':b'<Signature/>'}))
        with self.assertRaises(c.DocumentError):c.clean_file(p,self.p/'out')
        self.assertFalse((self.p/'out').exists())

    def test_broken_package(self):
        p=self.make_variant(BASE/'sample-template.pptx','broken.pptx',lambda data:data.pop('ppt/slides/slide1.xml'))
        with self.assertRaises(c.DocumentError):c.scan_file(p)

    def test_literal_keywords(self):
        p=self.p/'literal.txt';p.write_text('C++ public',encoding='utf-8');r=c.clean_file(p,self.p/'out',words=['C++'])
        self.assertEqual(r['before']['match_count'],1);self.assertEqual(r['after']['match_count'],0)

    def test_gui_hidden(self):
        r=app.gui([str(BASE/'sample-template.docx')],True)
        self.assertEqual(r['version'],c.VERSION);self.assertEqual(r['files'],1)
        self.assertTrue(r['clean_disabled_initially']);self.assertEqual(len(r['buttons']),7)

if __name__=='__main__':unittest.main(verbosity=2)
