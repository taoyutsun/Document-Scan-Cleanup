import json
import os
import subprocess
import sys
import tempfile
import unittest
from pathlib import Path
from unittest.mock import patch

ROOT=Path(__file__).resolve().parents[1]
sys.path.insert(0,str(ROOT))
import keyword_profiles as p
import document_cleanup_core as core
import document_cleanup_app as app
sys.path.insert(0,str(ROOT/'scripts'))
import audit_public
from zipfile import ZipFile


class ProfileTests(unittest.TestCase):
    def setUp(self):
        self.temp=tempfile.TemporaryDirectory();self.base=Path(self.temp.name)
        self.store=p.ProfileStore(self.base/'user-settings')

    def tearDown(self):self.temp.cleanup()

    def test_first_start_does_not_write_settings(self):
        identifier,profile,warning=self.store.startup()
        self.assertIsNone(identifier);self.assertEqual(profile['keywords'],core.DEFAULT_WORDS)
        self.assertEqual(warning,'');self.assertEqual(self.store.list_profiles(),([],0))
        self.assertFalse(self.store.directory.exists())

    def test_saved_startup_survives_reload_and_profile_update(self):
        identifier,profile=self.store.save('Personal example',['Project Cedar','範例詞','Project Cedar'])
        self.assertEqual(profile['keywords'],['Project Cedar','範例詞'])
        self.store.set_startup(identifier)
        self.store.save('Renamed',['Project Elm'],identifier)
        fresh=p.ProfileStore(self.store.directory)
        loaded_id,loaded,warning=fresh.startup()
        self.assertEqual(loaded_id,identifier);self.assertEqual(loaded['keywords'],['Project Elm'])
        self.assertEqual(loaded['name'],'Renamed');self.assertEqual(warning,'')

    def test_import_copies_and_preserves_original(self):
        source=self.base/'external.json'
        source.write_text('\ufeff'+json.dumps({'schema_version':1,'name':'External','keywords':['Sample term']}),encoding='utf-8')
        before=source.read_bytes();identifier,profile=self.store.import_profile(source)
        source.unlink()
        self.assertEqual(self.store.load(identifier),profile)
        self.assertTrue(before.startswith(b'\xef\xbb\xbf'))
        self.assertFalse(self.store.settings.exists())

    def test_validation_rejects_bad_types_versions_and_ui_separators(self):
        cases=[{'schema_version':True,'name':'X','keywords':['X']},
               {'schema_version':2,'name':'X','keywords':['X']},
               {'schema_version':1,'name':'','keywords':['X']},
               {'schema_version':1,'name':'X','keywords':'X'},
               {'schema_version':1,'name':'X','keywords':[1]},
               {'schema_version':1,'name':'X','keywords':[]},
               {'schema_version':1,'name':'X','keywords':['X;Y']},
               {'schema_version':1,'name':'X','keywords':['X\nY']},
               {'schema_version':1,'name':'X','keywords':['X'*151]},
               {'schema_version':1,'name':'X','keywords':[str(i) for i in range(101)]}]
        for data in cases:
            with self.subTest(data=data),self.assertRaises(p.ProfileError):p.validate_profile(data)

    def test_bad_import_does_not_create_settings(self):
        source=self.base/'bad.json';source.write_bytes(b'{broken')
        with self.assertRaises(p.ProfileError):self.store.import_profile(source)
        self.assertFalse(self.store.directory.exists())
        source.write_bytes(b'x'*(p.MAX_PROFILE_BYTES+1))
        with self.assertRaises(p.ProfileError):p.load_profile(source)

    def test_corrupt_startup_falls_back_without_overwriting(self):
        self.store.directory.mkdir();self.store.settings.write_bytes(b'{broken')
        _,profile,warning=self.store.startup()
        self.assertEqual(profile['keywords'],core.DEFAULT_WORDS);self.assertTrue(warning)
        self.assertEqual(self.store.settings.read_bytes(),b'{broken')

    def test_missing_profile_and_path_traversal_are_rejected(self):
        for identifier in ['../outside','..\\outside','a'*31,42]:
            with self.subTest(identifier=identifier),self.assertRaises(p.ProfileError):self.store.profile_path(identifier)
        p.atomic_json(self.store.settings,{'schema_version':1,'startup_profile':'a'*32})
        self.assertTrue(self.store.startup()[2])
        p.atomic_json(self.store.settings,{'schema_version':1,'startup_profile':'../outside'})
        self.assertTrue(self.store.startup()[2])

    def test_atomic_failure_retains_previous_profile(self):
        identifier,_=self.store.save('Original',['Original term'])
        before=self.store.profile_path(identifier).read_bytes()
        with patch.object(p.os,'replace',side_effect=OSError('simulated failure')):
            with self.assertRaises(OSError):self.store.save('Edited',['Edited term'],identifier)
        self.assertEqual(self.store.profile_path(identifier).read_bytes(),before)
        self.assertEqual(list(self.store.profiles.glob('.tmp-*')),[])

    def test_restore_generic_keeps_saved_profile(self):
        identifier,_=self.store.save('Keep',['Sample term']);self.store.set_startup(identifier)
        self.store.set_startup(None)
        self.assertEqual(self.store.startup()[1]['keywords'],core.DEFAULT_WORDS)
        self.assertEqual(self.store.load(identifier)['keywords'],['Sample term'])

    def test_gui_save_and_startup_then_restore(self):
        def actions(ui):
            ui['keywords'].set('Project Cedar;範例詞')
            self.assertTrue(ui['save']());ui['startup']()
        with patch('tkinter.simpledialog.askstring',return_value='Personal example'),patch('tkinter.messagebox.showerror') as errors:
            state=app.gui(smoke_test=True,profile_store=self.store,smoke_callback=actions)
            errors.assert_not_called()
        self.assertEqual(state['keyword_source'],'local_profile')
        self.assertEqual(state['keyword_count'],2);self.assertEqual(len(state['profile_buttons']),6)
        fresh=app.gui(smoke_test=True,profile_store=self.store)
        self.assertEqual(fresh['keyword_sha256'],state['keyword_sha256'])
        restored=app.gui(smoke_test=True,profile_store=self.store,smoke_callback=lambda ui:ui['restore']())
        self.assertEqual(restored['keyword_source'],'generic');self.assertEqual(len(self.store.list_profiles()[0]),1)

    def test_gui_import_and_unsaved_changes_do_not_write_default(self):
        source=self.base/'example.json';p.atomic_json(source,{'schema_version':1,'name':'Imported','keywords':['Example term']})
        def actions(ui):ui['import']();ui['keywords'].set('Unsaved term')
        with patch('tkinter.filedialog.askopenfilename',return_value=str(source)),patch('tkinter.messagebox.showerror') as errors:
            app.gui(smoke_test=True,profile_store=self.store,smoke_callback=actions)
            errors.assert_not_called()
        self.assertFalse(self.store.settings.exists())
        self.assertEqual(self.store.list_profiles()[0][0][1]['keywords'],['Example term'])

    def test_cli_profile_is_explicit_and_ignores_gui_default(self):
        identifier,_=self.store.save('CLI example',['Project Cedar']);self.store.set_startup(identifier)
        source=self.base/'input.txt';source.write_text('Project Cedar\nDraft\n',encoding='utf-8')
        env={**os.environ,'APPDATA':str(self.base/'appdata')}
        gui_store=p.ProfileStore(Path(env['APPDATA'])/'DocumentScanCleanup')
        gui_id,_=gui_store.save('GUI example',['Unrelated term']);gui_store.set_startup(gui_id)
        args=[sys.executable,'-X','utf8',str(ROOT/'document_scan_cleanup.py'),'--clean',str(source),'--output-dir']
        subprocess.run([*args,str(self.base/'explicit'),'--keywords-profile',str(self.store.profile_path(identifier))],check=True,capture_output=True,env=env)
        subprocess.run([*args,str(self.base/'generic')],check=True,capture_output=True,env=env)
        self.assertEqual((self.base/'explicit/001/input.cleaned.txt').read_text(encoding='utf-8'),'\nDraft\n')
        self.assertEqual((self.base/'generic/001/input.cleaned.txt').read_text(encoding='utf-8'),'Project Cedar\n\n')

    def test_public_allowlist_excludes_local_profiles_even_inside_docs(self):
        private=self.base/'docs/profiles/example.keywords.json';private.parent.mkdir(parents=True)
        private.write_text('{}',encoding='utf-8')
        settings=self.base/'docs/keyword-settings.json';settings.write_text('{}',encoding='utf-8')
        regular=self.base/'docs/example.md';regular.write_text('Public example',encoding='utf-8')
        with patch.object(audit_public,'ROOT',self.base):
            self.assertEqual(audit_public.public_files(),[regular])

    def test_zip_audit_rejects_local_settings(self):
        target=self.base/'example.zip'
        with ZipFile(target,'w') as archive:archive.writestr('Document-Scan-Cleanup/profiles/example.keywords.json','{}')
        with self.assertRaises(ValueError):audit_public.audit_zip(target)

    def test_reload_discovers_external_profile_and_startup_selection(self):
        def actions(ui):
            identifier,_=self.store.save('New example',['Project Elm'])
            self.store.set_startup(identifier);ui['reload']()
        result=app.gui(smoke_test=True,profile_store=self.store,smoke_callback=actions)
        self.assertEqual(result['profile_name'],'New example')
        self.assertEqual(result['profiles_available'],1)
        self.assertEqual(result['keyword_count'],1)
        self.assertTrue(Path(result['active_profile_file']).is_file())

    def test_reload_cancel_keeps_unsaved_keywords(self):
        def actions(ui):ui['keywords'].set('Unsaved example');ui['reload']()
        with patch('tkinter.messagebox.askyesno',return_value=False):
            result=app.gui(smoke_test=True,profile_store=self.store,smoke_callback=actions)
        self.assertEqual(result['keyword_source'],'custom')
        self.assertEqual(result['keyword_count'],1)
        self.assertFalse(self.store.settings.exists())

    def test_open_profile_folder_uses_actual_user_directory(self):
        with patch('os.startfile') as opened:
            app.gui(smoke_test=True,profile_store=self.store,smoke_callback=lambda ui:ui['open_folder']())
        opened.assert_called_once_with(str(self.store.profiles))
        self.assertTrue(self.store.profiles.is_dir())

    def test_author_links_fit_normal_and_smaller_windows(self):
        for size,scale in [('1140x880',1.333),('900x760',1.333),('900x640',1.333),('1140x880',2.0),('1040x740',1.75)]:
            with self.subTest(size=size,scale=scale):
                result=app.gui(smoke_test=True,profile_store=self.store,smoke_size=size,smoke_scale=scale)
                self.assertTrue(result['layout']['author_links_visible'],result['layout'])
                self.assertGreater(result['layout']['author_links_height'],0)


if __name__=='__main__':unittest.main()
