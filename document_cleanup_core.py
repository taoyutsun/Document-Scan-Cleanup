"""Local document scanning and selected OOXML text cleanup. No network APIs."""
from __future__ import annotations
import copy
import hashlib
import io
import json
import os
import re
import tempfile
from datetime import datetime
from email import policy
from email.parser import BytesParser
from html.parser import HTMLParser
from pathlib import Path
from zipfile import ZipFile, BadZipFile
from lxml import etree as E
import ppt_structure as old

from app_metadata import VERSION
DEFAULT_WORDS = ['Highly Confidential', 'Confidential (Restricted)', 'Confidentiality',
                 'Confidential', 'Internal Use Only', 'Draft', '機密', '僅限內部使用', '草稿']
W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'
A = 'http://schemas.openxmlformats.org/drawingml/2006/main'
S = 'http://schemas.openxmlformats.org/spreadsheetml/2006/main'
LEGACY = {'.ppt', '.pps', '.doc', '.xls', '.rtf'}
TEXT_FORMATS = {'.txt', '.md', '.log', '.csv', '.tsv', '.json', '.xml'}
SUPPORTED = {'.pptx', '.docx', '.xlsx', '.pdf', '.eml'} | TEXT_FORMATS | LEGACY
CLEANABLE = SUPPORTED - {'.json', '.xml'}
MAX_SIZE = 600 * 1024 * 1024
class DocumentError(Exception): pass

def pattern(words=None):
    words = DEFAULT_WORDS if words is None else words
    words = list(dict.fromkeys(w.strip() for w in words if w.strip()))
    if not words or len(words) > 100 or any(len(w) > 150 for w in words):
        raise DocumentError('請設定 1～100 個非空白關鍵字，每個最多 150 字。')
    return re.compile('|'.join(re.escape(w) for w in sorted(words, key=len, reverse=True)), re.I), words

def digest(b): return hashlib.sha256(b).hexdigest()
def now(): return datetime.now().isoformat(timespec='seconds')
def local(e): return E.QName(e).localname if isinstance(e.tag, str) else ''
def xpath(tree,path):
    namespaces={}
    for e in tree.getroot().iter():
        if isinstance(e.tag,str):namespaces.update({k:v for k,v in e.nsmap.items() if k})
    return tree.xpath(path,namespaces=namespaces)

def docx_protected(data):
    if 'word/settings.xml' not in data:return False
    p=old.parse(data['word/settings.xml']).find('{'+W+'}documentProtection')
    return p is not None and p.get('{'+W+'}enforcement','1').lower() not in {'0','false','off'}

def package(raw):
    try:
        with ZipFile(io.BytesIO(raw)) as z:
            infos = z.infolist()
            if len(infos) > 25000 or sum(i.file_size for i in infos) > MAX_SIZE:
                raise DocumentError('解壓縮後超過 600 MB 或零件數超過 25,000。')
            if len({i.filename for i in infos}) != len(infos):
                raise DocumentError('ZIP 含重複零件名稱。')
            for i in infos:
                if i.flag_bits & 1 or '\\' in i.filename or i.filename.startswith('/') or '..' in i.filename.split('/'):
                    raise DocumentError('不支援加密或含不正常路徑的檔案。')
            return {i.filename: z.read(i) for i in infos if not i.is_dir()}, infos
    except BadZipFile as exc:
        raise DocumentError('不是可讀取的 Office 檔案，可能已加密或損毀。') from exc

def protected_part(n):
    return (n.startswith('_xmlsignatures/') or 'labelinfo' in n.lower()
            or n.startswith('customXml/') or n.endswith('.rels') or n == '[Content_Types].xml')

def protected_node(n, e):
    if protected_part(n): return True
    chain = [e] + list(e.iterancestors())
    if any(any('msip_label_' in str(v).lower() or 'mip_label_' in str(v).lower()
               for v in a.attrib.values()) for a in chain): return True
    if n == 'docProps/custom.xml' and any(local(a) == 'property' and re.search(r'msip|sensitivity|label', a.get('name',''), re.I) for a in chain): return True
    return False

def editable_attribute(e,k):
    """Only recognized naming surfaces; never blindly rename arbitrary XML identifiers."""
    q=E.QName(k);ns=E.QName(e).namespace;ln=local(e);attr=q.localname
    if q.namespace in {old.R,old.REL}:return False
    if ln in {'cNvPr','docPr'} and attr in {'name','descr','title'}:return True
    if ns==A and ln in {'theme','clrScheme','fontScheme','fmtScheme'} and attr=='name':return True
    if ns==W and ln=='style' and attr=='styleId':return True
    if ns==W and ln in {'name','link','basedOn','next','pStyle','rStyle','tblStyle'} and attr=='val':return True
    if ns=='urn:schemas-microsoft-com:vml' and ln in {'shape','rect','oval','roundrect'} and attr=='id':return True
    return False

def validate_package(data, kind):
    required = {'[Content_Types].xml', '_rels/.rels', {'pptx':'ppt/presentation.xml','docx':'word/document.xml','xlsx':'xl/workbook.xml'}[kind]}
    if not required.issubset(data): raise DocumentError('缺少必要的 Office 零件。')
    for n,b in data.items():
        if not n.endswith(('.xml','.rels')): continue
        root = old.parse(b)
        if n.endswith('.rels'):
            owner = old.relowner(n)
            if owner and owner not in data: raise DocumentError('孤立的關聯檔：'+n)
            ids=[]
            for e in root:
                if not isinstance(e.tag,str): continue
                ids.append(e.get('Id'))
                t=old.internal_target(owner,e)
                if t and t not in data: raise DocumentError('遺失關聯零件：'+t)
            if len(ids)!=len(set(ids)): raise DocumentError('重複的關聯 ID。')
    for e in old.parse(data['[Content_Types].xml']):
        if local(e)=='Override' and e.get('PartName','').lstrip('/') not in data:
            raise DocumentError('內容類型指向不存在的零件。')
    if kind=='pptx': old.validate(data)

def groups(root):
    """Run text grouped by nearest paragraph, including split formatting and VML fallback."""
    mapping={}
    for e in root.iter():
        if local(e) not in {'t','delText','instrText'} or not e.text: continue
        ns=E.QName(e).namespace
        if ns not in {W,A}: continue
        para=next((p for p in e.iterancestors() if local(p)=='p' and E.QName(p).namespace==ns),e)
        key=(para,local(e)=='instrText',local(e)=='delText')
        mapping.setdefault(key,[]).append(e)
    return [(key[0],nodes,key[1],key[2]) for key,nodes in mapping.items()]

def context(text): return ' '.join(text.split())[:500]
def location(n):
    if re.fullmatch(r'ppt/slides/slide\d+\.xml',n): return '投影片 '+re.search(r'(\d+)\.xml$',n).group(1)
    if 'slideMasters/' in n: return '母片 · '+n.rsplit('/',1)[-1]
    if 'slideLayouts/' in n: return '版面配置 · '+n.rsplit('/',1)[-1]
    if n.startswith('word/footer'): return '頁尾 · '+n.rsplit('/',1)[-1]
    if n.startswith('word/header'): return '頁首 · '+n.rsplit('/',1)[-1]
    if n=='word/document.xml': return 'Word 正文／文字方塊'
    if 'notes' in n.lower(): return '備忘稿／註腳 · '+n.rsplit('/',1)[-1]
    if 'comment' in n.lower(): return '註解 · '+n.rsplit('/',1)[-1]
    if n.startswith('docProps/'): return '文件屬性 · '+n.rsplit('/',1)[-1]
    if 'styles' in n or 'numbering' in n: return '樣式／編號設定'
    return n

def add_finding(report, part, path, category, text, pat, cleanable, extra=None):
    matches=list(pat.finditer(text))
    if not matches:return
    identity=json.dumps([part,path,category,text],ensure_ascii=False)
    first=matches[0].start(); excerpt=text[max(0,first-110):first+350] if len(text)>500 else text
    item={'id':digest(identity.encode('utf-8'))[:24], 'part':part, 'path':path,
          'location':location(part), 'category':category, 'context':context(excerpt),
          'matches':[m.group() for m in matches], 'cleanable':bool(cleanable)}
    if extra:item.update(extra)
    report['findings'].append(item)

def scan_office(data, report, pat):
    kind=report['format']; editable=kind in {'pptx','docx'}
    if kind=='xlsx':
        from document_extra_formats import scan_xlsx
        scan_xlsx(data,report,pat)
    for n,b in (data.items() if kind!='xlsx' else []):
        if not n.endswith('.xml'):continue
        root=old.parse(b); tree=root.getroottree(); consumed=set()
        for para,nodes,field,deleted in groups(root):
            text=''.join(e.text or '' for e in nodes); consumed.update(nodes)
            category='欄位指令（只掃描）' if field else ('修訂刪除文字' if deleted else '文字段落')
            allowed=editable and not field and not any(protected_node(n,e) for e in nodes)
            add_finding(report,n,tree.getpath(para)+('/field' if field else '/deleted' if deleted else ''),category,text,pat,allowed)
        for e in root.iter():
            if not isinstance(e.tag,str):continue
            if e.text and e not in consumed:
                add_finding(report,n,tree.getpath(e)+'/text','屬性／其他文字',e.text,pat,
                            editable and not protected_node(n,e) and (n.startswith('docProps/') or n.startswith(('ppt/','word/'))))
            for k,v in e.attrib.items():
                # Namespace URIs and relationship IDs are structural, not text to strip.
                allow=editable and not protected_node(n,e) and editable_attribute(e,k)
                add_finding(report,n,tree.getpath(e)+'/@'+k,'內部名稱／屬性',v,pat,allow)
    media=[n for n in data if n.startswith(('ppt/media/','word/media/','xl/media/'))]
    embedded=[n for n in data if n.startswith(('ppt/embeddings/','word/embeddings/','xl/embeddings/','ppt/activeX/','word/activeX/'))]
    if media:report['uninspected'].append(f'{len(media)} 個圖片零件未做 OCR。')
    if embedded:report['uninspected'].append(f'{len(embedded)} 個內嵌／ActiveX 物件未解析。')
    report['embedded_parts']=embedded
    if any(n.startswith('_xmlsignatures/') for n in data):
        report['can_clean']=False
        report['warnings'].append('檔案有數位簽章，清理功能會拒絕修改。')
        for f in report['findings']:f['cleanable']=False
    if kind=='docx':
        if docx_protected(data):
            report['can_clean']=False
            report['warnings'].append('文件有編輯保護，清理功能會拒絕修改。')
            for f in report['findings']:f['cleanable']=False

class HTMLText(HTMLParser):
    def __init__(self):super().__init__();self.parts=[];self.skip=0
    def handle_starttag(self,tag,attrs):
        if tag in {'style','script'}:self.skip+=1
        if tag in {'br','p','div','tr'}:self.parts.append('\n')
    def handle_endtag(self,tag):
        if tag in {'style','script'}:self.skip=max(0,self.skip-1)
        if tag in {'p','div','tr'}:self.parts.append('\n')
    def handle_data(self,data):
        if not self.skip:self.parts.append(data)

def decode_text(raw):
    encodings=['utf-8-sig'] if raw.startswith(b'\xef\xbb\xbf') else ['utf-16-le'] if raw.startswith(b'\xff\xfe') else ['utf-16-be'] if raw.startswith(b'\xfe\xff') else ['utf-8','cp950']
    for encoding in encodings:
        try:return raw.decode(encoding),encoding
        except UnicodeError:continue
    raise DocumentError('無法可靠辨識文字編碼；請先另存 UTF-8。')

def scan_file(path, words=None, depth=0, budget=None):
    path=Path(path).resolve()
    if path.suffix.lower() not in SUPPORTED:raise DocumentError('不支援此格式：'+path.suffix)
    if path.stat().st_size>MAX_SIZE:raise DocumentError('單檔超過 600 MB。')
    raw=path.read_bytes()
    return scan_bytes(raw,path.suffix.lower(),str(path),words,depth,budget)

def scan_bytes(raw,extension,source,words=None,depth=0,budget=None):
    pat,words=pattern(words)
    report={'version':VERSION,'source':source,'source_sha256':digest(raw),'format':extension.lstrip('.'),
            'time':now(),'keywords':words,'status':'completed','findings':[],'warnings':[],
            'uninspected':[],'can_clean':extension in CLEANABLE}
    if extension in LEGACY:
        from document_legacy import convert, CONVERT
        try:converted=convert(raw,extension)
        except Exception as exc:raise DocumentError(str(exc)) from exc
        child=scan_bytes(converted,CONVERT[extension],source,words,depth,budget)
        child['source_sha256']=digest(raw);child['format']=extension.lstrip('.')
        child['converted_format']=CONVERT[extension].lstrip('.');child['converted_sha256']=digest(converted)
        child['warnings'].append('舊檔已用桌面版 Office 轉成暫存新版副本；清理會另存新版格式，原檔保留。轉檔可能影響版面。')
        return child
    if extension in {'.pptx','.docx','.xlsx'}:
        data,_=package(raw);validate_package(data,report['format']);scan_office(data,report,pat)
    elif extension=='.pdf':
        from pdf_text_cleanup import scan as scan_pdf
        try:scan_pdf(raw,report,pat)
        except Exception as exc:raise DocumentError('PDF 無法掃描：'+str(exc)) from exc
    elif extension in TEXT_FORMATS:
        text,encoding=decode_text(raw);report['encoding']=encoding
        for i,line in enumerate(text.splitlines(keepends=True),1):
            add_finding(report,'text',str(i),'文字行',line,pat,extension in CLEANABLE,{'location':f'第 {i} 行'})
    elif extension=='.eml':
        from document_extra_formats import scan_eml
        scan_eml(raw,report,pat,depth,budget)
    return finish_scan(report)

def finish_scan(report):
    report['match_count']=sum(len(f['matches']) for f in report['findings'])
    report['status']='partial' if report['uninspected'] else 'completed'
    return report

def replace_nodes(nodes,pat,block):
    text=''.join(n.text or '' for n in nodes)
    matches=list(pat.finditer(text))
    if not matches:return
    intervals=[(0,len(text))] if block else [(m.start(),m.end()) for m in matches]
    offset=0
    for n in nodes:
        original=n.text or '';start=offset;offset+=len(original)
        n.text=''.join(ch for i,ch in enumerate(original,start) if not any(a<=i<b for a,b in intervals))

def scrub_office(data,report,selected,mode,pat):
    actions=[]; selected=set(selected); lookup={f['id']:f for f in report['findings']}
    invalid=[fid for fid in selected if fid not in lookup or not lookup[fid]['cleanable']]
    if invalid:raise DocumentError('選取項目含有只可掃描或已失效的結果，請重新掃描。')
    # Rename style IDs and VML IDs with a complete reference map across the package.
    remap={}; theme_remap={}; styles=data.get('word/styles.xml')
    for fid in selected:
        f=lookup[fid]
        if re.fullmatch(r'ppt/theme/[^/]+\.xml',f['part']) and f['category']=='內部名稱／屬性' and f['path'].endswith('/@name'):
            root=old.parse(data[f['part']]);es=xpath(root.getroottree(),f['path'].split('/@')[0])
            if es and local(es[0])=='theme':
                value=es[0].get('name','');replacement=pat.sub('',value).strip()
                theme_remap[value]=replacement or 'CleanName_'+digest((f['part']+f['path']).encode())[:10]
    if styles:
        root=old.parse(styles); known={e.get('{'+W+'}styleId') for e in root if local(e)=='style'}
        for fid in selected:
            f=lookup[fid]
            if f['part']=='word/styles.xml' and f['category']=='內部名稱／屬性' and f['path'].endswith('/@{'+W+'}styleId'):
                es=xpath(root.getroottree(),f['path'].split('/@')[0])
                if es:
                    value=es[0].get('{'+W+'}styleId');new='CleanStyle_'+digest(value.encode())[:12]
                    if new in known and new!=value:raise DocumentError('中性樣式名稱衝突。')
                    remap[value]=new
    # A selected style reference implies renaming its actual definition as well.
    if styles:
        root=old.parse(styles); definitions={e.get('{'+W+'}styleId') for e in root if local(e)=='style'}
        for fid in selected:
            f=lookup[fid]
            if f['category']=='內部名稱／屬性' and f['part'].startswith('word/') and f['path'].endswith('/@{'+W+'}val'):
                e=xpath(old.parse(data[f['part']]).getroottree(),f['path'].split('/@')[0])
                if e:
                    value=e[0].get('{'+W+'}val')
                    if value in definitions:remap[value]='CleanStyle_'+digest(value.encode())[:12]
    # VML shape IDs may be referenced by OLE/drawing properties: synchronize exact refs.
    for fid in selected:
        f=lookup[fid]
        if f['category']=='內部名稱／屬性' and f['path'].endswith('/@id'):
            e=xpath(old.parse(data[f['part']]).getroottree(),f['path'].split('/@')[0])
            if e and E.QName(e[0]).namespace=='urn:schemas-microsoft-com:vml':
                value=e[0].get('id');remap[value]='CleanShape_'+digest(value.encode())[:12]
    bypart={}
    for fid in selected:bypart.setdefault(lookup[fid]['part'],[]).append(lookup[fid])
    modified=set()
    for n,b in list(data.items()):
        if not n.endswith('.xml') or protected_part(n):continue
        root=old.parse(b); tree=root.getroottree();changed=False
        for f in bypart.get(n,[]):
            path=f['path'];category=f['category']
            if category in {'文字段落','修訂刪除文字'}:
                for para,nodes,field,deleted in groups(root):
                    candidate=tree.getpath(para)+('/field' if field else '/deleted' if deleted else '')
                    if candidate==path:
                        replace_nodes(nodes,pat,mode=='block');changed=True;break
            elif category=='內部名稱／屬性':
                ep,attr=path.rsplit('/@',1);elements=xpath(tree,ep)
                if elements:
                    e=elements[0];v=e.get(attr,'')
                    if v in remap:new=remap[v]
                    elif v.startswith('#') and v[1:] in remap:new='#'+remap[v[1:]]
                    else:
                        new=pat.sub('',v).strip()
                        if not new:new='CleanName_'+digest((n+path).encode())[:10]
                    if new!=v:e.set(attr,new);changed=True
            elif category=='屬性／其他文字':
                elements=xpath(tree,path[:-5])
                if elements:
                    e=elements[0];value=e.text or '';new=theme_remap.get(value,pat.sub('',value))
                    if new!=e.text:e.text=new;changed=True
            if category in {'文字段落','修訂刪除文字'}:
                after_text=''.join(e.text or '' for e in nodes)
            elif category=='內部名稱／屬性':after_text=elements[0].get(attr,'') if elements else ''
            else:after_text=elements[0].text or '' if elements else ''
            actions.append({'id':f['id'],'location':f['location'],'category':category,'before':f['context'],
                            'after':context(after_text),'mode':mode})
        if remap:
            for e in root.iter():
                if not isinstance(e.tag,str) or protected_node(n,e):continue
                for k,v in list(e.attrib.items()):
                    if v in remap:e.set(k,remap[v]);changed=True
                    elif v.startswith('#') and v[1:] in remap:e.set(k,'#'+remap[v[1:]]);changed=True
        if n=='docProps/app.xml' and theme_remap:
            for e in root.iter():
                if isinstance(e.tag,str) and e.text in theme_remap:
                    e.text=theme_remap[e.text];changed=True
        if changed:data[n]=old.xml(root);modified.add(n)
    return actions,modified,remap

def strip_allowed_for_verification(root,kind):
    # Compare geometry/style with only the explicitly supported cleanup surfaces masked.
    for e in root.iter():
        if not isinstance(e.tag,str):continue
        ns=E.QName(e).namespace;ln=local(e)
        if ln in {'t','delText'} and ns in {W,A}:e.text=''
        for k in list(e.attrib):
            q=E.QName(k); attr=q.localname
            if attr in {'name','descr','title'} and ln in {'cNvPr','docPr','theme','clrScheme','fontScheme','fmtScheme'}:e.set(k,'')
            if kind=='docx' and ((ns==W and ln in {'style','name','link','basedOn','next','pStyle','rStyle','tblStyle'}) or (ns=='urn:schemas-microsoft-com:vml' and attr=='id')):
                if attr in {'val','styleId','id'}:e.set(k,'')
    return E.tostring(root,method='c14n')

def clean_file(path,directory,selected=None,words=None,mode='keyword',prune=False,expected_hash=None):
    path=Path(path).resolve();extension=path.suffix.lower()
    if extension not in CLEANABLE:raise DocumentError('此格式目前只掃描或尚未支援清理。')
    if extension in {'.csv','.tsv'} and mode=='block':raise DocumentError('CSV／TSV 僅提供關鍵字清理，保留資料列與分隔符號。')
    if mode not in {'keyword','block'}:raise DocumentError('無效的清理模式。')
    report=scan_file(path,words);pat,_=pattern(report['keywords'])
    if extension in {'.csv','.tsv'}:
        reserved={'\r','\n','"',',' if extension=='.csv' else '\t'}
        if any(any(ch in word for ch in reserved) for word in report['keywords']):raise DocumentError('CSV／TSV 關鍵字不可包含引號、分隔符號或換行。')
    raw=path.read_bytes()
    if digest(raw)!=report['source_sha256'] or (expected_hash and digest(raw)!=expected_hash):
        raise DocumentError('檔案在掃描後已變更，請重新掃描。')
    selected=[f['id'] for f in report['findings'] if f['cleanable']] if selected is None else list(selected)
    lookup={f['id']:f for f in report['findings']}
    if not report['can_clean'] or any(fid not in lookup or not lookup[fid]['cleanable'] for fid in selected):
        raise DocumentError('含只掃描／受保護／已失效的項目，未修改。')
    source_raw=raw
    if extension in LEGACY:
        from document_legacy import convert, CONVERT
        raw=convert(raw,extension);extension=CONVERT[extension]
        if digest(raw)!=report['converted_sha256']:raise DocumentError('轉檔副本與掃描時不同，請重新掃描。')
    kind=extension.lstrip('.')
    if not selected and not (prune and extension=='.pptx'):raise DocumentError('沒有選定可清理的項目。')
    changes=[];prune_report=None;verification={};data=None;infos=None
    if extension in {'.pptx','.docx','.xlsx'}:
        data,infos=package(raw);original=data.copy()
        if any(n.startswith('_xmlsignatures/') for n in data):raise DocumentError('數位簽章檔案不修改。')
        if extension=='.docx':
            if E.QName(old.parse(data['word/document.xml'])).namespace!=W:raise DocumentError('Strict Open XML 目前只掃描。')
            if docx_protected(data):
                raise DocumentError('文件有編輯保護，未修改。')
        elif extension=='.pptx':
            if E.QName(old.parse(data['ppt/presentation.xml'])).namespace!=old.P:raise DocumentError('Strict Open XML 目前只掃描。')
        else:
            if E.QName(old.parse(data['xl/workbook.xml'])).namespace!=S:raise DocumentError('Strict Open XML 目前只掃描。')
        if extension=='.xlsx':
            from document_extra_formats import clean_xlsx
            changes,verification=clean_xlsx(data,report,selected,mode,pat)
            modified={n for n in data if data[n]!=original[n]};remap={}
        else:changes,modified,remap=scrub_office(data,report,selected,mode,pat)
        # Check selected removals cannot change shape geometry, paragraph/run styling or page settings.
        for n in modified:
            if n.startswith(('word/','ppt/')):
                a,b=old.parse(original[n]),old.parse(data[n])
                if strip_allowed_for_verification(a,kind)!=strip_allowed_for_verification(b,kind):
                    raise DocumentError('有超出文字／名稱的變更，已停止：'+n)
        verification.update({'layout_geometry_and_styles_preserved':extension!='.xlsx','source_unchanged':True,
                      'media_and_embedded_parts_identical':True,'reference_validation':'passed'}
                      )
        if prune and extension=='.pptx':
            # Prune unused structural dependencies from an intermediate copy, never the original.
            with tempfile.TemporaryDirectory(prefix='document-cleanup-') as td:
                intermediate=Path(td)/'intermediate.pptx'
                with ZipFile(intermediate,'w') as z:
                    for info in infos:
                        if info.filename in data:z.writestr(copy.copy(info),data[info.filename])
                data,_,prune_report=old.analyze(intermediate)
            verification['pruning']=prune_report['verification']
        validate_package(data,kind)
        for n,b in original.items():
            if n.startswith(('word/media/','ppt/media/','xl/media/','word/embeddings/','ppt/embeddings/','xl/embeddings/')) and n in data and data[n]!=b:
                raise DocumentError('圖片或內嵌物件發生變更，已停止。')
        # Every style reference must retain its original target or point to its neutral replacement.
        if extension=='.docx' and remap:
            styles=old.parse(data['word/styles.xml']);ids={e.get('{'+W+'}styleId') for e in styles if local(e)=='style'}
            for value in remap.values():
                if value.startswith('CleanStyle_') and value not in ids:raise DocumentError('樣式引用驗證失敗。')
        outbuf=io.BytesIO()
        with ZipFile(outbuf,'w') as z:
            for info in infos:
                if info.filename in data:z.writestr(copy.copy(info),data[info.filename])
        output=outbuf.getvalue()
        actual,_=package(output)
        if actual!=data:raise DocumentError('輸出內容不一致。')
        validate_package(actual,kind)
    elif extension=='.pdf':
        from pdf_text_cleanup import clean as clean_pdf
        output,changes,verification=clean_pdf(raw,report,selected,pat)
    elif extension=='.eml':
        from document_extra_formats import clean_eml
        output,changes,verification=clean_eml(raw,report,selected,mode,pat)
    else:
        text,encoding=decode_text(raw);lines=text.splitlines(keepends=True)
        lookup={f['id']:f for f in report['findings']}
        for fid in selected:
            if fid not in lookup:raise DocumentError('選取項目失效。')
            f=lookup[fid];i=int(f['path'])-1;line=lines[i]
            if mode=='block':lines[i]=('\ufeff' if i==0 and line.startswith('\ufeff') else '')+('\r\n' if line.endswith('\r\n') else '\n' if line.endswith('\n') else '')
            else:lines[i]=pat.sub('',line)
            changes.append({'id':fid,'location':f['location'],'before':f['context'],'mode':mode})
        output=''.join(lines).encode(encoding)
        verification={'source_unchanged':True,'encoding_preserved':True}
    if digest(path.read_bytes())!=report['source_sha256']:raise DocumentError('來源在處理期間已變更，未另存。')
    after=scan_bytes(output,extension,'清理後副本',report['keywords'])
    result={'version':VERSION,'source':str(path),'source_sha256':digest(raw),'time':now(),
            'format':report['format'],'output_format':kind,'keywords':report['keywords'],'mode':mode,
            'before':report,'after':after,'changes':changes,'pruning':prune_report,
            'verification':verification,'output_sha256':digest(output)}
    result['source_sha256']=digest(source_raw)
    if path.suffix.lower() in LEGACY:
        result['verification']['legacy_conversion']='Desktop Office -> '+extension+'; review layout before use.'
    if data is not None:
        result['modified_parts']=sorted(n for n in data if n in original and data[n]!=original[n])
        result['deleted_parts']=sorted(set(original)-set(data))
    directory=Path(directory).resolve();directory.mkdir(parents=True,exist_ok=True)
    out=directory/(path.stem+'.cleaned'+extension);rp=out.with_suffix('.report.json');tp=out.with_suffix('.report.txt')
    if any(p.exists() for p in [out,rp,tp]):raise DocumentError('輸出已存在，請改用另一個資料夾。')
    result['output']=str(out);after['source']=str(out)
    created=[]
    try:
        for p,b in [(out,output),(rp,json.dumps(result,ensure_ascii=False,indent=2).encode('utf-8')),
                    (tp,clean_summary(result).encode('utf-8-sig'))]:
            with p.open('xb') as stream:created.append(p);stream.write(b)
    except Exception:
        for p in created:p.unlink(missing_ok=True)
        raise
    return result

def scan_summary(report):
    lines=[f'文件掃描工具 {VERSION}', '來源：'+report['source'],
           f"文字命中：{report['match_count']} 處；結果項目：{len(report['findings'])}。",
           '檢查狀態：'+('已完成文字範圍檢查，部分內容未檢查。' if report['status']=='partial' else '已完成支援範圍檢查。')]
    lines += [f"[{f['id']}] {f['location']} · {f['category']} · {'可清理' if f['cleanable'] else '只掃描'}\n  {f['context']}" for f in report['findings']]
    lines+=['', '未檢查／限制：']+report['uninspected']+report['warnings']
    lines+=['此結果僅涵蓋已檢查範圍；分享副本前仍需確認內容與用途。']
    return '\n'.join(lines)

def clean_summary(result):
    return '\n'.join([f'文件清理工具 {VERSION}', '來源：'+result['source'],'另存：'+result['output'],
                      f"清理前文字命中：{result['before']['match_count']}；清理後：{result['after']['match_count']}。",
                      '原檔保留；已檢查本格式支援的結構與保留內容。',
                      '正式敏感度標籤、欄位指令、未知內部資料與權限保留。',
                      '', '清理後掃描：',scan_summary(result['after'])])

def save_scan(report,directory):
    directory=Path(directory);directory.mkdir(parents=True,exist_ok=True)
    for suffix,text in [('json',json.dumps(report,ensure_ascii=False,indent=2)),('txt',scan_summary(report))]:
        with (directory/('scan.report.'+suffix)).open('x',encoding='utf-8-sig' if suffix=='txt' else 'utf-8') as stream:stream.write(text)
