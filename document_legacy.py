"""Isolated, bounded conversion of legacy Office/RTF files to modern copies."""
import io, json, subprocess, sys, tempfile, re
from functools import lru_cache
from pathlib import Path

CONVERT = {'.doc': '.docx', '.rtf': '.docx', '.ppt': '.pptx', '.pps': '.pptx', '.xls': '.xlsx'}
_cache = {}  # A single conversion is reused so findings match the exact converted copy.

def preflight(raw, extension):
    if extension == '.rtf':
        if not raw.lstrip().startswith(b'{\\rtf'):raise ValueError('不是有效的 RTF。')
        if b'\\object' in raw.lower():raise ValueError('RTF 含內嵌物件，本版不轉檔。')
        return
    import olefile
    if not olefile.isOleFile(io.BytesIO(raw)):raise ValueError('不是舊版 Office 二進位格式。')
    with olefile.OleFileIO(io.BytesIO(raw)) as ole:
        names=['/'.join(n).lower() for n in ole.listdir()]
        if any(any(s in n for s in ['vba','encryptedpackage','encryptioninfo','digitalsignature']) for n in names):
            raise ValueError('舊檔含巨集、加密或簽章，未轉檔。')
        if any('msip_label_' in str(v).lower() for n in ole.listdir() if n[-1].startswith('\x05')
               for v in ole.getproperties(n).values()):
            raise ValueError('舊檔含正式分類標籤，本版不轉檔。')
        if extension=='.xls':
            import struct
            name='Workbook' if ole.exists('Workbook') else 'Book'
            book=ole.openstream(name).read();offset=0
            while offset+4<=len(book):
                record,size=struct.unpack_from('<HH',book,offset);body=book[offset+4:offset+4+size];offset+=4+size
                if record==0x85 and len(body)>5 and body[5] in {1,6}:raise ValueError('舊 Excel 含 XLM 巨集工作表，不轉檔。')
                if record==0x1AE and len(body)>3 and body[2:4]!=b'\x01\x04':raise ValueError('舊 Excel 含外部活頁簿／增益集引用，不自動轉檔。')

def convert(raw, extension):
    import hashlib
    key=(hashlib.sha256(raw).hexdigest(),extension)
    if key in _cache:return _cache[key]
    if len(raw)>100*1024*1024:raise ValueError('舊格式自動轉檔上限 100 MB；較大檔案請先手動另存新版格式。')
    preflight(raw,extension)
    with tempfile.TemporaryDirectory(prefix='document-convert-') as td:
        folder=Path(td);source=folder/('source'+extension);target=folder/('converted'+CONVERT[extension])
        source.write_bytes(raw)
        request=folder/'request.json';response=folder/'response.json'
        request.write_text(json.dumps({'source':str(source),'target':str(target),'response':str(response)}),encoding='utf-8')
        command=[sys.executable,'--convert-worker',str(request)] if getattr(sys,'frozen',False) else [sys.executable,str(Path(__file__).resolve()),'--convert-worker',str(request)]
        try:
            proc=subprocess.run(command,creationflags=subprocess.CREATE_NO_WINDOW,timeout=90,
                                stdout=subprocess.DEVNULL,stderr=subprocess.DEVNULL)
        except subprocess.TimeoutExpired as exc:
            # Only terminate the dedicated COM app PID registered by this conversion worker.
            if response.exists():
                state=json.loads(response.read_text(encoding='utf-8'))
                if state.get('owned_pid'):
                    import win32api
                    try:
                        handle=win32api.OpenProcess(1,False,state['owned_pid']);win32api.TerminateProcess(handle,1);handle.Close()
                    except Exception:pass
            raise ValueError('Office 轉檔超過 90 秒，已停止。請先在 Office 手動另存新版格式。') from exc
        state=json.loads(response.read_text(encoding='utf-8')) if response.exists() else {}
        if proc.returncode or not target.exists():raise ValueError(state.get('error','需要安裝對應的桌面版 Office；此檔無法自動轉換。'))
        output=target.read_bytes()
    if len(output)>600*1024*1024:raise ValueError('轉檔後超過大小限制。')
    if len(_cache)>=4 or sum(map(len,_cache.values()))+len(output)>200*1024*1024:_cache.clear()
    _cache[key]=output
    return output

def worker(request):
    import pythoncom, win32com.client, win32process
    q=json.loads(Path(request).read_text(encoding='utf-8'));source=Path(q['source']);target=q['target'];response=Path(q['response'])
    app=None;doc=None;owned=False;pid=None;pythoncom.CoInitialize()
    try:
        ext=source.suffix.lower()
        prog='Word.Application' if ext in {'.doc','.rtf'} else 'PowerPoint.Application' if ext in {'.ppt','.pps'} else 'Excel.Application'
        image='WINWORD.EXE' if prog=='Word.Application' else 'POWERPNT.EXE' if prog=='PowerPoint.Application' else 'EXCEL.EXE'
        def process_ids():
            inventory=subprocess.run(['tasklist.exe','/FI','IMAGENAME eq '+image,'/FO','CSV','/NH'],capture_output=True,creationflags=subprocess.CREATE_NO_WINDOW,timeout=10)
            if inventory.returncode:raise ValueError('無法確認 Office 程序所有權，未轉檔。')
            return {int(v) for v in re.findall(rb'"'+image.encode()+rb'","(\d+)"',inventory.stdout,re.I)}
        existing=process_ids()
        if prog=='PowerPoint.Application' and existing:raise ValueError('PowerPoint 正在執行，請先手動另存 PPTX，或關閉 PowerPoint 後再轉檔。')
        app=win32com.client.DispatchEx(prog)
        created=process_ids()-existing
        if len(created)!=1:raise ValueError('Office 未建立可確認的獨立轉檔程序，請手動另存新版格式。')
        pid=created.pop()
        owned=True;response.write_text(json.dumps({'owned_pid':pid}),encoding='utf-8');app.AutomationSecurity=3
        if prog=='Word.Application':
            app.Visible=False;app.DisplayAlerts=0;app.Options.UpdateLinksAtOpen=False
            doc=app.Documents.OpenNoRepairDialog(str(source),ConfirmConversions=False,ReadOnly=True,AddToRecentFiles=False,
                PasswordDocument='DocumentCleanup-No-Password',PasswordTemplate='DocumentCleanup-No-Password',Visible=False,OpenAndRepair=False)
            if doc.HasVBProject or doc.ProtectionType!=-1:raise ValueError('文件含巨集或編輯保護，不轉檔。')
            doc.SaveAs2(target,FileFormat=12,AddToRecentFiles=False)
        elif prog=='Excel.Application':
            app.Visible=False;app.DisplayAlerts=False;app.EnableEvents=False;app.AskToUpdateLinks=False
            try:app.Calculation=-4135
            except Exception:pass
            doc=app.Workbooks.Open(str(source),UpdateLinks=0,ReadOnly=True,Password='DocumentCleanup-No-Password',
                WriteResPassword='DocumentCleanup-No-Password',IgnoreReadOnlyRecommended=True,AddToMru=False)
            if doc.HasVBProject or doc.Excel4MacroSheets.Count or doc.ProtectStructure or any(s.ProtectContents for s in doc.Worksheets):
                raise ValueError('活頁簿含巨集或保護，不轉檔。')
            if doc.Connections.Count:raise ValueError('活頁簿含外部資料連線，本版不轉檔。')
            doc.SaveAs(target,FileFormat=51)
        else:
            app.DisplayAlerts=1
            doc=app.Presentations.Open(str(source),ReadOnly=True,Untitled=False,WithWindow=False)
            if doc.HasVBProject:raise ValueError('簡報含巨集，不轉檔。')
            doc.SaveAs(target,24)
        response.write_text(json.dumps({'success':True,'owned_pid':pid}),encoding='utf-8');return 0
    except Exception as exc:
        response.write_text(json.dumps({'error':'Office 轉檔失敗：'+str(exc),'owned_pid':pid if owned else None},ensure_ascii=False),encoding='utf-8');return 1
    finally:
        if doc is not None:
            try:doc.Close() if source.suffix.lower() in {'.ppt','.pps'} else doc.Close(False)
            except Exception:pass
        if app is not None and owned:
            try:app.Quit()
            except Exception:pass
        pythoncom.CoUninitialize()

if __name__=='__main__':raise SystemExit(worker(sys.argv[2]))
