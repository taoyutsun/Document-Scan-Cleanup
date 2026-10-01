"""Create a portable ZIP from the public allowlist, never from a recursive workspace copy."""
import hashlib,json,sys
from pathlib import Path
from zipfile import ZipFile,ZIP_DEFLATED
import audit_public
ROOT=Path(__file__).resolve().parents[1];sys.path.insert(0,str(ROOT))
from app_metadata import VERSION

def main():
    exe=ROOT/'dist/DocumentScanCleanup.exe'
    if not exe.exists():raise ValueError('Build the EXE first.')
    audit_public.audit_source();objects=audit_public.audit_exe(exe)
    target=ROOT/'dist'/f'Document-Scan-Cleanup-v{VERSION}-windows-x64.zip'
    files=audit_public.public_files()
    with ZipFile(target,'w',ZIP_DEFLATED) as z:
        z.write(exe,'Document-Scan-Cleanup/DocumentScanCleanup.exe')
        for p in files:z.write(p,'Document-Scan-Cleanup/'+p.relative_to(ROOT).as_posix())
    entries=audit_public.audit_zip(target)
    checks=ROOT/'dist'/'SHA256SUMS.txt'
    checks.write_text('\n'.join(hashlib.sha256(p.read_bytes()).hexdigest()+'  '+p.name for p in [exe,target])+'\n',encoding='utf-8')
    print(json.dumps({'zip':target.name,'files':entries,'frozen_code_objects_checked':objects,'sha256':hashlib.sha256(target.read_bytes()).hexdigest()},indent=2))

if __name__=='__main__':main()
