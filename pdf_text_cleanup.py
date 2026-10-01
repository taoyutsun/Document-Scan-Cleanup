"""Remove selected PDF text-show operations, preserving advance, fonts and other objects."""
import io, re
import pikepdf
from pdfminer.pdfparser import PDFParser
from pdfminer.pdfdocument import PDFDocument
from pdfminer.pdfpage import PDFPage
from pdfminer.pdfinterp import PDFResourceManager
from pdfminer.pdftypes import resolve1
import document_cleanup_core as c

def inspect(raw):
    import logging
    logging.getLogger('pdfminer').setLevel(logging.ERROR)
    miner=PDFDocument(PDFParser(io.BytesIO(raw)));manager=PDFResourceManager();allfonts=[]
    for page in PDFPage.create_pages(miner):
        fonts={}
        for name,ref in resolve1(page.resources.get('Font',{})).items():
            fonts[name]=manager.get_font(getattr(ref,'objid',None),resolve1(ref))
        allfonts.append(fonts)
    results=[]
    with pikepdf.open(io.BytesIO(raw)) as doc:
        for page_index,page in enumerate(doc.pages):
            ops=list(pikepdf.parse_content_stream(page));state={'font':None,'size':0,'tc':0,'tw':0};stack=[];items=[]
            for i,op in enumerate(ops):
                args=op.operands;code=str(op.operator)
                if code=='q':stack.append(state.copy())
                elif code=='Q':state=stack.pop() if stack else state
                elif code=='Tf':state['font']=allfonts[page_index].get(str(args[0]).lstrip('/'));state['size']=float(args[1])
                elif code=='Tc':state['tc']=float(args[0])
                elif code=='Tw':state['tw']=float(args[0])
                if code not in {'Tj','TJ',"'",'"'}:continue
                if code=='"':state['tw']=float(args[0]);state['tc']=float(args[1])
                parts=args[0] if code=='TJ' else [args[-1]];decoded=[];advance=[];font=state['font'];safe=font is not None and state['size']!=0 and not font.is_vertical()
                for part in parts:
                    if not isinstance(part,pikepdf.String):advance.append(part);continue
                    payload=bytes(part)
                    if font is None:decoded.append(payload.decode('latin1'));safe=False;continue
                    codes=list(font.decode(payload));chars=[]
                    for cid in codes:
                        try:chars.append(font.to_unichr(cid))
                        except Exception:chars.append('\ufffd');safe=False
                    decoded.append(''.join(chars))
                    # TJ offsets are in thousandths of a text-space unit. Replace glyphs by
                    # their exact original advance, including character and word spacing.
                    if safe:
                        width=sum(font.char_width(cid) for cid in codes)*1000
                        spacing=len(codes)*state['tc']+(payload.count(b' ') * state['tw'] if not font.is_multibyte() else 0)
                        advance.append(-width-spacing*1000/state['size'])
                items.append({'index':i,'text':''.join(decoded),'safe':safe,'advance':advance,'operator':code,'args':list(args)})
            results.append({'ops':ops,'items':items})
    return results

def scan(raw,report,pat):
    import pypdfium2 as pdfium
    from document_extra_formats import pdf_state,pdf_objects
    allowed,count=pdf_state(raw);report['can_clean']=allowed;report['pages']=count
    if not allowed:report['warnings'].append('PDF 含加密、簽章或受保護分類資訊，只掃描。')
    details=inspect(raw)
    with pdfium.PdfDocument(raw) as viewer:
        for page_index,detail in enumerate(details):
            page=viewer[page_index];tp=page.get_textpage();objects=[o for o in pdf_objects(page,tp) if o.type==pdfium.raw.FPDF_PAGEOBJ_TEXT]
            matched=[]
            for item in detail['items']:
                if not pat.search(item['text']):continue
                matched.extend(m.group().lower() for m in pat.finditer(item['text']))
                boxes=[list(o.get_bounds()) for o in objects if o.extract()==item['text']]
                c.add_finding(report,f'page:{page_index+1}',str(item['index']),'PDF 文字區塊（整個刪除）',item['text'],pat,allowed and item['safe'],
                              {'location':f'PDF 第 {page_index+1} 頁 · 文字區塊 {item["index"]+1}',
                               'full_text':item['text'],'bounds':boxes,'default_selected':False})
            full=tp.get_text_range();counts={w:matched.count(w) for w in matched};unresolved=False
            for m in pat.finditer(full):
                word=m.group().lower()
                if counts.get(word):counts[word]-=1
                else:unresolved=True
            if unresolved:c.add_finding(report,f'page:{page_index+1}','unresolved','PDF 分段／巢狀文字（只掃描）',full,pat,False,{'location':f'PDF 第 {page_index+1} 頁 · 需人工檢查'})
            if not full.strip():report['uninspected'].append(f'PDF 第 {page_index+1} 頁沒有可擷取文字，未做 OCR。')
            tp.close();page.close()
    with pikepdf.open(io.BytesIO(raw)) as doc:
        for key,value in doc.docinfo.items():
            if isinstance(value,pikepdf.String):
                edit=allowed and str(key) in {'/Title','/Author','/Subject','/Keywords','/Creator','/Producer'}
                c.add_finding(report,'metadata',str(key),'PDF 文件屬性',str(value),pat,edit,{'location':'PDF 屬性 · '+str(key),'default_selected':False})
        if doc.Root.get('/Metadata'):c.add_finding(report,'metadata','XMP','PDF XMP（只掃描）',doc.Root.Metadata.read_bytes().decode('utf-8',errors='replace'),pat,False)
        for i,page in enumerate(doc.pages,1):
            for j,annot in enumerate(page.obj.get('/Annots',[])):
                for key in ['/Contents','/T','/Subj']:
                    if annot.get(key):c.add_finding(report,f'page:{i}',f'annot:{j}:{key}','PDF 註解（只掃描）',str(annot[key]),pat,False)
    report['uninspected'].append('PDF 圖片未做 OCR；表單、附件、分段／巢狀內容未完整解析。清理會刪除勾選的整個文字區塊，請雙擊檢視完整範圍。')

def clean(raw,report,selected,pat):
    from document_extra_formats import pdf_state
    if not pdf_state(raw)[0]:raise c.DocumentError('受保護、加密或簽章 PDF 未修改。')
    details=inspect(raw);lookup={f['id']:f for f in report['findings']};changes=[];modified=set();out=io.BytesIO()
    with pikepdf.open(io.BytesIO(raw)) as doc:
        for fid in selected:
            f=lookup[fid]
            if f['part']=='metadata':
                before=str(doc.docinfo[f['path']]);after=pat.sub('',before);doc.docinfo[f['path']]=pikepdf.String(after)
            else:
                page_index=int(f['part'].split(':')[1])-1;index=int(f['path']);detail=details[page_index]
                item=next(o for o in detail['items'] if o['index']==index)
                if not item['safe'] or item['text']!=f['full_text']:raise c.DocumentError('PDF 文字區塊驗證失敗。')
                args=[]
                if item['operator']=='"':
                    args += [pikepdf.ContentStreamInstruction([item['args'][0]],pikepdf.Operator('Tw')),
                             pikepdf.ContentStreamInstruction([item['args'][1]],pikepdf.Operator('Tc'))]
                if item['operator'] in {"'",'"'}:args.append(pikepdf.ContentStreamInstruction([],pikepdf.Operator('T*')))
                args.append(pikepdf.ContentStreamInstruction([pikepdf.Array(item['advance'])],pikepdf.Operator('TJ')))
                detail['ops'][index]=args;modified.add(page_index);before=item['text'];after=''
            changes.append({'id':fid,'location':f['location'],'before':before,'after':after,'mode':'whole_pdf_text_block' if f['part']!='metadata' else 'keyword'})
        for page_index in modified:
            ops=[]
            for op in details[page_index]['ops']:ops.extend(op if isinstance(op,list) else [op])
            doc.pages[page_index].Contents=pikepdf.Stream(doc,pikepdf.unparse_content_stream(ops))
        doc.remove_unreferenced_resources();doc.save(out)
    output=out.getvalue();verified=inspect(output)
    for page_index,(a,b) in enumerate(zip(details,verified)):
        removed={int(lookup[fid]['path']) for fid in selected if lookup[fid]['part']==f'page:{page_index+1}'}
        expected=[item['text'] for item in a['items'] if item['index'] not in removed]
        actual=[item['text'] for item in b['items'] if item['text']]
        if [s for s in expected if s]!=actual:raise c.DocumentError('PDF 保留文字驗證失敗。')
    with pikepdf.open(io.BytesIO(raw)) as before,pikepdf.open(io.BytesIO(output)) as after:
        if len(before.pages)!=len(after.pages):raise c.DocumentError('PDF 頁數發生變更。')
        for a,b in zip(before.pages,after.pages):
            for key in ['/MediaBox','/CropBox','/Rotate']:
                if str(a.obj.get(key))!=str(b.obj.get(key)):raise c.DocumentError('PDF 頁面尺寸或方向變更。')
    return output,changes,{'source_unchanged':True,'pdf_geometry_and_retained_text_verified':True,
                           'fonts_and_non_text_operators_preserved':True,'full_rewrite_no_incremental_save':True}
