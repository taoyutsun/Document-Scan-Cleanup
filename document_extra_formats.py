"""Conservative format-specific operations; formulas, attachments and labels stay intact."""
import copy, ctypes, io, re
from email import policy
from email.parser import BytesParser
from lxml import etree as E, html
import document_cleanup_core as c

TEXT = {'.txt','.md','.log','.csv','.tsv','.json','.xml'}

def pdf_objects(page,textpage=None):
    for obj in page.get_objects(max_depth=1):
        if textpage is not None:obj.textpage=textpage
        yield obj

def pdf_state(raw):
    import pikepdf
    with pikepdf.open(io.BytesIO(raw)) as doc:
        signed=any(isinstance(o,pikepdf.Dictionary) and (o.get('/Type')==pikepdf.Name('/Sig') or o.get('/FT')==pikepdf.Name('/Sig')) for o in doc.objects)
        labels=bool(doc.Root.get('/Perms')) or any('msip_label_' in str(k).lower() for k in doc.docinfo)
        if doc.Root.get('/Metadata'):
            labels=labels or bool(re.search(rb'msip_label_|mip_label_|sensitivity',doc.Root.Metadata.read_bytes(),re.I))
        return not(doc.is_encrypted or signed or labels),len(doc.pages)

def scan_xlsx(data,report,pat):
    shared=[];protected=set();report['can_clean']=True
    if 'xl/sharedStrings.xml' in data:
        shared=list(c.old.parse(data['xl/sharedStrings.xml']))
        for i,si in enumerate(shared):
            c.add_finding(report,'xl/sharedStrings.xml',str(i),'共用文字（所有引用，只掃描）',''.join(si.itertext()),pat,False)
    for n,b in data.items():
        if not n.endswith('.xml'):continue
        root=c.old.parse(b);tree=root.getroottree()
        if root.find('{'+c.S+'}sheetProtection') is not None:protected.add(n)
        if root.find('{'+c.S+'}workbookProtection') is not None:report['can_clean']=False
        if re.fullmatch(r'xl/worksheets/[^/]+\.xml',n):
            for cell in root.iter('{'+c.S+'}c'):
                coord=cell.get('r','');v=cell.find('{'+c.S+'}v');formula=cell.find('{'+c.S+'}f');value=''
                if cell.get('t')=='s' and v is not None:
                    try:value=''.join(shared[int(v.text)].itertext())
                    except (ValueError,IndexError):raise c.DocumentError('共用儲存格文字索引無效。')
                elif cell.get('t')=='inlineStr':value=''.join(cell.find('{'+c.S+'}is').itertext())
                elif cell.get('t')=='str' and v is not None:value=v.text or ''
                c.add_finding(report,n,'cell:'+coord,'儲存格文字' if formula is None else '公式結果（只掃描）',value,pat,
                              formula is None and n not in protected,{'location':n+' · '+coord,'default_selected':False})
                if formula is not None:c.add_finding(report,n,'formula:'+coord,'公式（只掃描）',formula.text or '',pat,False,{'location':n+' · '+coord})
            for e in root.iter():
                if c.local(e) in {'oddHeader','evenHeader','firstHeader','oddFooter','evenFooter','firstFooter'}:
                    controls=re.findall(r'&"[^"]*"',e.text or '')
                    edit=n not in protected and not any(pat.search(control) for control in controls)
                    c.add_finding(report,n,tree.getpath(e)+'/text','Excel 頁首頁尾',e.text or '',pat,edit,{'default_selected':False})
        elif n.startswith('xl/comments'):
            for comment in root.iter('{'+c.S+'}comment'):
                text=comment.find('{'+c.S+'}text');value=''.join(text.itertext()) if text is not None else ''
                c.add_finding(report,n,tree.getpath(comment),'Excel 註解',value,pat,not protected,{'location':n+' · '+comment.get('ref',''),'default_selected':False})
        elif n!='xl/sharedStrings.xml':
            for e in root.iter():
                if not isinstance(e.tag,str):continue
                edit=not c.protected_node(n,e) and (n.startswith('docProps/') or (E.QName(e).namespace==c.A and c.local(e)=='t'))
                if e.text:c.add_finding(report,n,tree.getpath(e)+'/text','Excel 屬性／圖形文字',e.text,pat,edit,{'default_selected':False})
                for k,v in e.attrib.items():c.add_finding(report,n,tree.getpath(e)+'/@'+k,'Excel 名稱／屬性（只掃描）',v,pat,False)
    if protected:report['warnings'].append('有保護的工作表只掃描，不清理其儲存格及頁首頁尾。')
    if not report['can_clean']:
        report['warnings'].append('活頁簿含結構保護，未提供清理。')
        for f in report['findings']:f['cleanable']=False

def clean_xlsx(data,report,selected,mode,pat):
    original=copy.deepcopy(data);findings={f['id']:f for f in report['findings']};roots={};changes=[];touched_shared=set()
    shared=list(c.old.parse(data['xl/sharedStrings.xml'])) if 'xl/sharedStrings.xml' in data else []
    for fid in selected:
        f=findings[fid];n=f['part'];root=roots.setdefault(n,c.old.parse(data[n]));tree=root.getroottree()
        if f['path'].startswith('cell:'):
            coord=f['path'][5:];cell=next(e for e in root.iter('{'+c.S+'}c') if e.get('r')==coord)
            if cell.get('t')=='s':
                v=cell.find('{'+c.S+'}v');index=int(v.text);touched_shared.add(index)
                inline=copy.deepcopy(shared[index]);inline.tag='{'+c.S+'}is';cell.remove(v);cell.append(inline);cell.set('t','inlineStr')
                nodes=list(inline.iter('{'+c.S+'}t'))
            elif cell.get('t')=='inlineStr':nodes=list(cell.find('{'+c.S+'}is').iter('{'+c.S+'}t'))
            else:nodes=[cell.find('{'+c.S+'}v')]
            before=''.join(e.text or '' for e in nodes);c.replace_nodes(nodes,pat,mode=='block');after=''.join(e.text or '' for e in nodes)
        elif f['category']=='Excel 註解':
            comment=c.xpath(tree,f['path'])[0];nodes=list(comment.iter('{'+c.S+'}t'))
            before=''.join(e.text or '' for e in nodes);c.replace_nodes(nodes,pat,mode=='block');after=''.join(e.text or '' for e in nodes)
        else:
            e=c.xpath(tree,f['path'][:-5])[0];before=e.text or '';after=pat.sub('',before)
            if mode=='block' and f['category']=='Excel 屬性／圖形文字' and E.QName(e).namespace==c.A:after=''
            e.text=after
        changes.append({'id':fid,'location':f['location'],'before':before,'after':after,'mode':mode})
    for n,root in roots.items():data[n]=c.old.xml(root)
    used=set()
    for n,b in data.items():
        if n.startswith('xl/worksheets/') and n.endswith('.xml'):
            for cell in c.old.parse(b).iter('{'+c.S+'}c'):
                if cell.get('t')=='s':used.add(int(cell.find('{'+c.S+'}v').text))
    if touched_shared-used:
        root=c.old.parse(data['xl/sharedStrings.xml'])
        for index in touched_shared-used:
            for e in root[index].iter('{'+c.S+'}t'):e.text=''
        data['xl/sharedStrings.xml']=c.old.xml(root)
    # Formula nodes, references and numeric/date values must remain identical.
    for n,b in original.items():
        if not n.startswith('xl/worksheets/') or not n.endswith('.xml'):continue
        a=c.old.parse(b);d=c.old.parse(data[n])
        for tag in ['f']:
            if [E.tostring(e) for e in a.iter('{'+c.S+'}'+tag)]!=[E.tostring(e) for e in d.iter('{'+c.S+'}'+tag)]:raise c.DocumentError('Excel 公式發生變更。')
        selected_cells={f['path'][5:] for f in (findings[i] for i in selected) if f['part']==n and f['path'].startswith('cell:')}
        amap={e.get('r'):E.tostring(e) for e in a.iter('{'+c.S+'}c') if e.get('r') not in selected_cells}
        dmap={e.get('r'):E.tostring(e) for e in d.iter('{'+c.S+'}c') if e.get('r') not in selected_cells}
        if amap!=dmap:raise c.DocumentError('未選取的 Excel 儲存格發生變更。')
    return changes,{'source_unchanged':True,'formulas_and_unselected_cells_verified':True,'shared_strings_copied_per_selected_cell':True}

def html_tree(text):return html.document_fromstring(text,parser=html.HTMLParser(no_network=True))
def html_groups(root):
    mapping={}
    def slots(e):
        if not isinstance(e.tag,str) or e.tag.lower() in {'script','style'}:return
        if e.text:yield e,'text'
        for child in e:
            yield from slots(child)
            if child.tail:yield child,'tail'
    for e,attr in slots(root):
        owner=e.getparent() if attr=='tail' else e
        if owner is None:continue
        block=next((p for p in [owner,*owner.iterancestors()] if p.tag in {'p','div','li','td','th','body','h1','h2','h3','pre'}),owner)
        mapping.setdefault(block,[]).append((e,attr))
    return mapping

def clean_html_group(pairs,pat,block):
    text=''.join(getattr(e,attr) or '' for e,attr in pairs);intervals=[(0,len(text))] if block else [m.span() for m in pat.finditer(text)]
    offset=0
    for e,attr in pairs:
        value=getattr(e,attr) or '';start=offset;offset+=len(value)
        setattr(e,attr,''.join(ch for i,ch in enumerate(value,start) if not any(a<=i<b for a,b in intervals)))

def email_parts(msg):
    def walk(part,path='0',attached=False):
        attached=attached or bool(part.get_filename()) or part.get_content_disposition()=='attachment' or part.get_content_type()=='message/rfc822'
        if part.is_multipart() and part.get_content_type()!='message/rfc822':
            for i,child in enumerate(part.iter_parts()):yield from walk(child,path+'.'+str(i),attached)
        else:yield path,part,attached
    return list(walk(msg))

def signed_email(msg):
    return bool(msg.get('DKIM-Signature') or msg.get('ARC-Seal')) or any(p.get_content_type() in {'multipart/signed','multipart/encrypted','application/pgp-encrypted','application/pgp-signature','application/pkcs7-mime','application/x-pkcs7-mime','application/pkcs7-signature','application/x-pkcs7-signature'} for p in msg.walk())

def scan_eml(raw,report,pat,depth,budget):
    msg=BytesParser(policy=policy.default).parsebytes(raw);encapsulated=any(p.get_content_type()=='message/rfc822' for p in msg.walk())
    allowed=not signed_email(msg) and not encapsulated;report['can_clean']=allowed
    c.add_finding(report,'headers','Subject','郵件標題',str(msg.get('Subject','')),pat,allowed,{'default_selected':False})
    if not allowed:report['warnings'].append('郵件含 DKIM／ARC／S/MIME 簽章或加密，只掃描。')
    if encapsulated:report['warnings'].append('郵件含封裝郵件附件，無法保證原始附件序列化一致，本版只掃描。')
    total=[0] if budget is None else budget
    for path,part,attached in email_parts(msg):
        filename=part.get_filename()
        if attached:
            payload=part.get_payload(decode=True);ext=c.Path(filename or '').suffix.lower()
            c.add_finding(report,'attachment',path,'附件名稱（只掃描）',filename or '',pat,False)
            if payload and ext in c.SUPPORTED and ext not in c.LEGACY and depth<2 and len(payload)<=100*1024*1024 and total[0]+len(payload)<=200*1024*1024:
                total[0]+=len(payload)
                try:
                    child=c.scan_bytes(payload,ext,report['source']+' :: '+str(filename),report['keywords'],depth+1,total)
                    for f in child['findings']:
                        f=copy.deepcopy(f);f['id']=c.digest((path+f['id']).encode())[:24];f['cleanable']=False;f['location']='附件 '+str(filename)+' · '+f['location'];report['findings'].append(f)
                    report['uninspected']+=['附件 '+str(filename)+'：'+s for s in child['uninspected']]
                except Exception as exc:report['uninspected'].append('附件無法掃描：'+str(exc))
            else:report['uninspected'].append('附件 '+str(filename)+' 未解析（格式／大小／深度限制）；不自動啟動 Office 解析附件。')
        elif part.get_content_type()=='text/plain':
            for i,line in enumerate(part.get_content().splitlines(keepends=True)):
                c.add_finding(report,'email:'+path,str(i),'郵件純文字行',line,pat,allowed,{'location':'郵件正文 · 第 '+str(i+1)+' 行','default_selected':False})
        elif part.get_content_type()=='text/html':
            root=html_tree(part.get_content());tree=root.getroottree()
            for block,pairs in html_groups(root).items():
                c.add_finding(report,'html:'+path,tree.getpath(block),'郵件 HTML 段落',''.join(getattr(e,a) or '' for e,a in pairs),pat,allowed,{'location':'郵件 HTML 正文','default_selected':False})
            for e in root.iter():
                if not isinstance(e.tag,str):continue
                for attr,value in e.attrib.items():c.add_finding(report,'html:'+path,tree.getpath(e)+'/@'+attr,'郵件 HTML 屬性（只掃描）',value,pat,False)
                if e.tag in {'style','script'}:c.add_finding(report,'html:'+path,tree.getpath(e)+'/text','郵件樣式／指令（只掃描）',e.text or '',pat,False)
    report['uninspected'].append('郵件附件、內嵌圖片保留；附件需另行清理後重新附加。圖片未做 OCR。')

def clean_eml(raw,report,selected,mode,pat):
    msg=BytesParser(policy=policy.default).parsebytes(raw)
    if signed_email(msg) or any(p.get_content_type()=='message/rfc822' for p in msg.walk()):raise c.DocumentError('含簽章、加密或封裝郵件附件的郵件未修改。')
    parts={path:part for path,part,attached in email_parts(msg) if not attached};findings={f['id']:f for f in report['findings']};changes=[]
    grouped={}
    for fid in selected:grouped.setdefault(findings[fid]['part'],[]).append(findings[fid])
    for part_id,items in grouped.items():
        if part_id=='headers':
            old=str(msg.get('Subject',''));new=pat.sub('',old);msg.replace_header('Subject',new)
            changes.append({'id':items[0]['id'],'location':'郵件標題','before':old,'after':new,'mode':'keyword'});continue
        kind,path=part_id.split(':',1);part=parts[path];text=part.get_content()
        if kind=='email':
            lines=text.splitlines(keepends=True)
            for f in items:
                i=int(f['path']);old=lines[i];lines[i]=('' if mode=='block' else pat.sub('',old)).rstrip('\r\n')+('\r\n' if old.endswith('\r\n') else '\n' if old.endswith('\n') else '')
                changes.append({'id':f['id'],'location':f['location'],'before':old,'after':lines[i],'mode':mode})
            text=''.join(lines)
        else:
            root=html_tree(text);tree=root.getroottree();groups=html_groups(root)
            for f in items:
                block=tree.xpath(f['path'])[0];pairs=groups[block];old=''.join(getattr(e,a) or '' for e,a in pairs)
                clean_html_group(pairs,pat,mode=='block');new=''.join(getattr(e,a) or '' for e,a in pairs)
                changes.append({'id':f['id'],'location':f['location'],'before':old,'after':new,'mode':mode})
            text=html.tostring(root,encoding='unicode',method='html')
        # Replace payload only; retain content headers, CID, disposition and all MIME parameters.
        charset=part.get_content_charset() or 'utf-8'
        try:encoded=text.encode(charset)
        except UnicodeEncodeError:raise c.DocumentError('郵件字元無法用原編碼儲存。')
        import base64
        part.set_payload(base64.encodebytes(encoded).decode('ascii'))
        if part.get('Content-Transfer-Encoding'):part.replace_header('Content-Transfer-Encoding','base64')
        else:part['Content-Transfer-Encoding']='base64'
    output=msg.as_bytes(policy=policy.SMTP)
    check=BytesParser(policy=policy.default).parsebytes(output);before=BytesParser(policy=policy.default).parsebytes(raw)
    for key in ['From','To','Cc','Bcc','Date','Message-ID','Reply-To']:
        if [str(v) for v in before.get_all(key,[])]!=[str(v) for v in check.get_all(key,[])]:raise c.DocumentError('郵件寄件／收件資訊發生變更。')
    def attachments(m):return [(p,part.get_filename(),part.get_payload(decode=True),part.get('Content-ID')) for p,part,attached in email_parts(m) if attached or part.get_content_type() not in {'text/plain','text/html'}]
    if attachments(before)!=attachments(check):raise c.DocumentError('郵件附件發生變更。')
    return output,changes,{'source_unchanged':True,'attachment_payloads_and_address_headers_identical':True}
