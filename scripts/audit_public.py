"""Audit allowlisted public files, portable archives and frozen Python code."""
import argparse,marshal,re,types
from pathlib import Path
from zipfile import ZipFile

ROOT=Path(__file__).resolve().parents[1]
TOP={'app_metadata.py','document_scan_cleanup.py','document_cleanup_app.py','document_cleanup_core.py',
     'document_extra_formats.py','document_legacy.py','pdf_text_cleanup.py','ppt_structure.py',
     'README.md','README.en.md','LICENSE','CHANGELOG.md','SECURITY.md','THIRD_PARTY_NOTICES.md',
     'requirements.txt','requirements-dev.txt','.gitignore','.gitattributes'}
FOLDERS={'tests','scripts','docs','third-party-licenses','.github'}
# Personal Windows home paths must not leak into source or frozen code.
# Organization-specific markers belong in a maintainer's private audit, not this repository.
FORBIDDEN=re.compile(r'[A-Za-z]:[\\/]Users[\\/][^\\/\s<>]+',re.I)

def public_files():
    paths=[ROOT/n for n in sorted(TOP) if (ROOT/n).is_file()]
    for folder in sorted(FOLDERS):
        paths += [p for p in (ROOT/folder).rglob('*') if p.is_file() and '__pycache__' not in p.parts and p.suffix!='.pyc']
    return sorted(paths,key=lambda p:p.relative_to(ROOT).as_posix())

def scan_text(data,location):
    text=data.decode('utf-8-sig',errors='replace')
    if FORBIDDEN.search(text):raise ValueError('Private marker found in '+location)

def audit_source():
    for p in public_files():scan_text(p.read_bytes(),p.relative_to(ROOT).as_posix())
    # Additional untracked top-level files are not part of the allowlist and must not be staged.
    import subprocess
    if (ROOT/'.git').is_dir():
        result=subprocess.run(['git','ls-files','-z'],cwd=ROOT,capture_output=True,check=True)
        allowed={p.relative_to(ROOT).as_posix() for p in public_files()}
        unexpected={p for p in result.stdout.decode('utf-8').split('\0') if p}-allowed
        if unexpected:raise ValueError('Unexpected tracked files: '+', '.join(sorted(unexpected)))
    return len(public_files())

def audit_exe(path):
    from PyInstaller.archive.readers import CArchiveReader
    archive=CArchiveReader(str(path));count=0
    def inspect_code(code,where):
        nonlocal count
        count+=1
        scan_text(code.co_filename.encode('utf-8'),where+' filename')
        for value in code.co_consts:
            if isinstance(value,types.CodeType):inspect_code(value,where)
            elif isinstance(value,str):scan_text(value.encode('utf-8'),where+' constant')
    for name,entry in archive.toc.items():
        if entry[-1]=='s':inspect_code(marshal.loads(archive.extract(name)),name)
    for entry_name,entry in archive.toc.items():
        scan_text(entry_name.encode('utf-8'),'archive entry')
        if entry[-1]=='z':
            embedded=archive.open_embedded_archive(entry_name)
            for name in embedded.toc:
                value=embedded.extract(name)
                if isinstance(value,types.CodeType):inspect_code(value,name)
        elif entry_name=='base_library.zip':
            import io
            with ZipFile(io.BytesIO(archive.extract(entry_name))) as library:
                for name in library.namelist():
                    if name.endswith('.pyc'):inspect_code(marshal.loads(library.read(name)[16:]),name)
    return count

def audit_zip(path):
    with ZipFile(path) as z:
        if z.testzip() is not None:raise ValueError('Invalid ZIP integrity.')
        for name in z.namelist():
            if any(part in {'backups','work','outputs','__pycache__','.git'} for part in Path(name).parts):raise ValueError('Private/generated folder in archive: '+name)
            if name.endswith('.exe'):continue
            scan_text(z.read(name),name)
        return len(z.namelist())

def main():
    ap=argparse.ArgumentParser();ap.add_argument('--exe',type=Path);ap.add_argument('--zip',type=Path);args=ap.parse_args()
    result={'public_files_checked':audit_source()}
    if args.exe:result['frozen_code_objects_checked']=audit_exe(args.exe)
    if args.zip:result['zip_entries_checked']=audit_zip(args.zip)
    import json
    print(json.dumps(result,indent=2))

if __name__=='__main__':main()
