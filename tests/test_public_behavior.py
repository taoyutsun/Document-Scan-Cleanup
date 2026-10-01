"""Regression coverage for custom cleanup rules and unrelated active templates."""
import sys,tempfile,unittest
from pathlib import Path
from zipfile import ZipFile,ZIP_DEFLATED

sys.path.insert(0,str(Path(__file__).resolve().parents[1]))
import document_cleanup_core as c
import fixtures

class PublicBehaviorTests(unittest.TestCase):
    def test_custom_literal_keyword_preserves_other_content(self):
        with tempfile.TemporaryDirectory() as directory:
            base=Path(directory);source=base/'handoff.txt'
            source.write_text('Project Meridian draft\nConfidential unrelated note\nProject Mars final\n',encoding='utf-8')
            before=source.read_bytes()
            result=c.clean_file(source,base/'out',words=['project meridian'])
            self.assertEqual(result['before']['match_count'],1)
            self.assertEqual(result['after']['match_count'],0)
            self.assertEqual(source.read_bytes(),before)
            text=Path(result['output']).read_text(encoding='utf-8')
            self.assertIn('Confidential unrelated note',text)
            self.assertIn('Project Mars final',text)
            self.assertNotIn('Project Meridian',text)

    def test_pruning_preserves_unselected_active_theme_names(self):
        with tempfile.TemporaryDirectory() as directory:
            base=Path(directory);fixtures.make(base)
            source=base/'sample-template.pptx';parts,_=c.package(source.read_bytes())
            parts['ppt/slides/slide1.xml']=parts['ppt/slides/slide1.xml'].replace(b'Public slide',b'REMOVE-ME')
            with ZipFile(source,'w',ZIP_DEFLATED) as package:
                for name,data in parts.items():package.writestr(name,data)
            scan=c.scan_file(source,words=['REMOVE-ME'])
            self.assertEqual(scan['match_count'],1)
            result=c.clean_file(source,base/'out',[scan['findings'][0]['id']],words=['REMOVE-ME'],prune=True)
            cleaned,_=c.package(Path(result['output']).read_bytes())
            self.assertEqual(cleaned['ppt/theme/theme3.xml'],parts['ppt/theme/theme3.xml'])
            self.assertEqual(cleaned['ppt/slideMasters/slideMaster3.xml'],parts['ppt/slideMasters/slideMaster3.xml'])
            self.assertEqual(result['pruning']['after']['masters'],1)
            self.assertEqual(result['after']['match_count'],0)

if __name__=='__main__':unittest.main()
