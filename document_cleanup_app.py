"""Desktop interface / CLI for the local document cleanup engine."""
import argparse
import json
import queue
import sys
import threading
from datetime import datetime
from pathlib import Path
import document_cleanup_core as core
import app_metadata as meta

def gui(initial=(),smoke_test=False):
    import tkinter as tk
    from tkinter import ttk,filedialog,messagebox
    root=tk.Tk()
    if smoke_test:root.withdraw()
    root.title(meta.APP_NAME+' '+core.VERSION)
    root.geometry('1140x880');root.minsize(900,760)
    paths=list(initial);reports={};errors={};selection=set();rows={};events=queue.Queue();busy=False
    dest_last=None
    title=ttk.Label(root,text='交付文件前，檢查並整理指定內容與範本殘留',font=('Microsoft JhengHei',16))
    title.pack(anchor='w',padx=16,pady=(14,5))
    ttk.Label(root,text='原檔保留。舊 Office／RTF 需桌面版 Office 轉檔；PDF 刪除整個文字物件；郵件附件保留。').pack(anchor='w',padx=16)
    keywords=tk.StringVar(value='; '.join(core.DEFAULT_WORDS))
    entryrow=ttk.Frame(root);entryrow.pack(fill='x',padx=16,pady=10)
    ttk.Label(entryrow,text='關鍵字（分號分隔）：').pack(side='left')
    keyentry=ttk.Entry(entryrow,textvariable=keywords);keyentry.pack(side='left',fill='x',expand=True)
    toolbar=ttk.Frame(root);toolbar.pack(fill='x',padx=16,pady=(0,8))
    all_buttons=[]
    files=ttk.Treeview(root,columns=('format','status'),show='tree headings',height=4)
    files.heading('#0',text='檔案');files.heading('format',text='格式');files.heading('status',text='掃描狀態')
    files.column('#0',width=680);files.column('format',width=70);files.column('status',width=270)
    files.pack(fill='x',padx=16)
    pane=ttk.Frame(root);pane.pack(fill='both',expand=True,padx=16,pady=8)
    hits=ttk.Treeview(pane,columns=('check','file','where','kind','context'),show='headings',height=12)
    for col,text,width in [('check','清理',55),('file','檔案',200),('where','位置',190),('kind','類型',130),('context','命中內容',460)]:
        hits.heading(col,text=text);hits.column(col,width=width,stretch=(col=='context'))
    scroll=ttk.Scrollbar(pane,orient='vertical',command=hits.yview);hits.configure(yscrollcommand=scroll.set)
    hits.pack(side='left',fill='both',expand=True);scroll.pack(side='right',fill='y')
    options=ttk.Frame(root);options.pack(fill='x',padx=16)
    mode=tk.StringVar(value='keyword');prune=tk.BooleanVar(value=True)
    ttk.Label(options,text='文字處理：').pack(side='left')
    ttk.Radiobutton(options,text='只移除關鍵字',variable=mode,value='keyword').pack(side='left',padx=6)
    ttk.Radiobutton(options,text='清空選定段落／文字行',variable=mode,value='block').pack(side='left',padx=6)
    ttk.Checkbutton(options,text='同時清理 PPT 未使用範本',variable=prune).pack(side='left',padx=16)
    ttk.Label(root,text='段落模式保留框架與格式，可能改變換行。名稱／屬性只清理字樣。雙擊結果可查看內容；點第一欄或按空白鍵切換勾選。').pack(anchor='w',padx=16,pady=6)
    log=tk.Text(root,height=8,wrap='word',font=('Microsoft JhengHei',10));log.pack(fill='x',padx=16)
    status=tk.StringVar(value='選擇檔案後按「掃描」。掃描不修改來源文件。')
    ttk.Label(root,textvariable=status).pack(anchor='w',padx=16,pady=10)
    about=ttk.LabelFrame(root,text='關於作者');about.pack(fill='x',padx=16,pady=(0,12))
    ttk.Label(about,text=meta.AUTHOR_DESCRIPTION,wraplength=1080).pack(anchor='w',padx=8,pady=(5,2))
    links=ttk.Frame(about);links.pack(anchor='w',padx=8,pady=(0,5))
    import webbrowser,os
    for label,url in [('亞瑟 ASK 部落格',meta.BLOG_URL),('Facebook',meta.FACEBOOK_URL)]:
        link=ttk.Label(links,text=label,foreground='#175a9a',cursor='hand2');link.pack(side='left',padx=(0,18))
        link.bind('<Button-1>',lambda event,u=url:webbrowser.open(u))
    source=ttk.Label(links,text='檢視原始碼'+('' if meta.SOURCE_REPO_URL else '（本機）'),foreground='#175a9a',cursor='hand2');source.pack(side='left')
    source.bind('<Button-1>',lambda event:webbrowser.open(meta.SOURCE_REPO_URL) if meta.SOURCE_REPO_URL else os.startfile(str(meta.SOURCE_DIRECTORY)))

    def keylist():return [w.strip() for w in keywords.get().split(';') if w.strip()]
    def update_controls():
        for b in all_buttons:b.configure(state='disabled' if busy else 'normal')
        if not busy:
            save_btn.configure(state='normal' if reports or errors else 'disabled')
            can_clean=bool(selection) or (prune.get() and any(r['format']=='pptx' and r['can_clean'] for r in reports.values()))
            clean_btn.configure(state='normal' if can_clean else 'disabled')
            select_btn.configure(state='normal' if rows else 'disabled')
            deselect_btn.configure(state='normal' if rows else 'disabled')
        keyentry.configure(state='disabled' if busy else 'normal')

    def clear_results():
        reports.clear();errors.clear();selection.clear();rows.clear()
        hits.delete(*hits.get_children())
        for i,p in enumerate(paths):
            iid=str(i)
            if files.exists(iid):files.item(iid,values=(Path(p).suffix,'待掃描'))
        update_controls()

    def add():
        chosen=filedialog.askopenfilenames(title='選擇要掃描的文件',filetypes=[('支援文件',' '.join('*'+ext for ext in sorted(core.SUPPORTED))),('所有檔案','*.*')])
        for p in chosen:
            if p not in paths:
                paths.append(p);files.insert('','end',iid=str(len(paths)-1),text=p,values=(Path(p).suffix,'待掃描'))
        clear_results()

    def remove_all():
        paths.clear();files.delete(*files.get_children());clear_results();log.delete('1.0','end')

    def invalidate(*unused):
        if not busy and reports:
            clear_results();status.set('關鍵字已變更，請重新掃描。')

    def refresh_checks():
        for iid,(path,f) in rows.items():
            values=list(hits.item(iid,'values'));values[0]=('[x]' if (path,f['id']) in selection else '[ ]') if f['cleanable'] else '只掃描'
            hits.item(iid,values=values)
        update_controls()

    def toggle(event=None):
        if busy:return
        iid=hits.identify_row(event.y) if event is not None and getattr(event,'num',None)==1 else (hits.selection()[0] if hits.selection() else '')
        if not iid or iid not in rows:return
        path,f=rows[iid]
        if not f['cleanable']:return
        key=(path,f['id'])
        if key in selection:selection.remove(key)
        else:selection.add(key)
        refresh_checks()

    def click(event):
        if hits.identify_column(event.x)=='#1':toggle(event)

    def set_all(value):
        selection.clear()
        if value:selection.update((p,f['id']) for p,f in rows.values() if f['cleanable'])
        refresh_checks()

    def detail(event=None):
        iid=hits.identify_row(event.y) if event else (hits.selection()[0] if hits.selection() else '')
        if iid not in rows:return
        p,f=rows[iid];win=tk.Toplevel(root);win.title('掃描結果內容');win.geometry('760x360')
        text=tk.Text(win,wrap='word',font=('Microsoft JhengHei',11));text.pack(fill='both',expand=True,padx=12,pady=12)
        text.insert('end',p+'\n'+f['location']+'\n類型：'+f['category']+'\n命中字詞：'+', '.join(f['matches'])+'\n\n'+f.get('full_text',f['context'])+'\n\n詳細位置：'+f['part']+'\n'+f['path'])
        text.configure(state='disabled')

    def scan():
        nonlocal busy
        if not paths:return
        try:_,words=core.pattern(keylist())
        except Exception as exc:messagebox.showerror('關鍵字設定',str(exc));return
        clear_results();busy=True;update_controls();log.delete('1.0','end');status.set('正在本機掃描……')
        selected=paths.copy()
        def worker():
            for i,p in enumerate(selected):
                try:events.put(('scan',(i,p,core.scan_file(p,words))))
                except Exception as exc:events.put(('error',(i,p,str(exc))))
            events.put(('done',('scan','')))
        threading.Thread(target=worker,daemon=True).start()

    def destination(prefix):
        chosen=filedialog.askdirectory(title='選擇存放結果的位置（會建立新的資料夾）')
        if not chosen:return None
        return Path(chosen)/(prefix+'_'+datetime.now().strftime('%Y%m%d_%H%M%S_%f'))

    def save():
        nonlocal dest_last
        target=destination('Document_Scan')
        if not target:return
        try:
            for i,(p,r) in enumerate(reports.items(),1):core.save_scan(r,target/f'{i:03d}')
            if errors:
                target.mkdir(parents=True,exist_ok=True)
                (target/'未能掃描的檔案.json').write_text(json.dumps(errors,ensure_ascii=False,indent=2),encoding='utf-8')
            dest_last=target;status.set('掃描報告已存至：'+str(target))
        except Exception as exc:messagebox.showerror('儲存報告',str(exc))

    def clean():
        nonlocal busy
        tasks=[]
        for p,r in reports.items():
            ids=[fid for path,fid in selection if path==p]
            if r['can_clean'] and Path(p).suffix.lower() in core.CLEANABLE and (ids or (prune.get() and Path(p).suffix.lower()=='.pptx')):
                tasks.append((p,r,ids))
        if not tasks:messagebox.showinfo('沒有清理項目','請勾選可清理的結果。PDF／Excel／郵件預設不勾選。');return
        target=destination('Document_Cleanup')
        if not target:return
        busy=True;update_controls();status.set('正在清理選取內容並另存……')
        selected_mode=mode.get();selected_prune=prune.get()
        def worker():
            batch=[]
            for i,(p,r,ids) in enumerate(tasks,1):
                try:
                    result=core.clean_file(p,target/f'{i:03d}',ids,r['keywords'],selected_mode,selected_prune,r['source_sha256'])
                    events.put(('clean',result))
                    batch.append({'source':p,'status':'success','output':result['output']})
                except Exception as exc:
                    batch.append({'source':p,'status':'not_modified','error':str(exc)})
                    events.put(('clean_error',(p,str(exc))))
            try:
                target.mkdir(parents=True,exist_ok=True)
                (target/'batch.report.json').write_text(json.dumps(batch,ensure_ascii=False,indent=2),encoding='utf-8')
            except Exception as exc:events.put(('clean_error',('批次報告',str(exc))))
            successes=sum(item['status']=='success' for item in batch)
            events.put(('done',(f'clean:{successes}:{len(batch)-successes}',str(target))))
        threading.Thread(target=worker,daemon=True).start()

    def poll():
        nonlocal busy,dest_last
        try:
            while True:
                kind,value=events.get_nowait()
                if kind=='scan':
                    i,p,r=value;reports[p]=r
                    label=f"命中 {r['match_count']}；"+('部分未檢查' if r['status']=='partial' else '文字檢查完成')
                    files.item(str(i),values=(Path(p).suffix,label))
                    for f in r['findings']:
                        iid=str(len(rows));rows[iid]=(p,f)
                        chosen=f['cleanable'] and f.get('default_selected',True)
                        if chosen:selection.add((p,f['id']))
                        hits.insert('','end',iid=iid,values=('[x]' if chosen else '[ ]' if f['cleanable'] else '只掃描',Path(p).name,f['location'],f['category'],f['context']))
                    log.insert('end',Path(p).name+'：'+label+'\n')
                    for s in r['uninspected']+r['warnings']:log.insert('end','  '+s+'\n')
                elif kind=='error':
                    i,p,e=value;errors[p]=e;files.item(str(i),values=(Path(p).suffix,'無法掃描'))
                    log.insert('end',p+'\n  無法掃描：'+e+'\n')
                elif kind=='clean':
                    log.insert('end',core.clean_summary(value)+'\n\n')
                elif kind=='clean_error':
                    p,e=value;log.insert('end',p+'\n  未修改：'+e+'\n')
                elif kind=='done':
                    operation,target=value;busy=False;update_controls()
                    if target:
                        dest_last=Path(target)
                        counts=operation.split(':');status.set(f'清理完成：成功 {counts[1]}，未修改 {counts[2]}。結果：'+target)
                    else:status.set(f'掃描完成：{len(reports)} 份成功，{len(errors)} 份無法掃描。可調整勾選後清理。')
                log.see('end')
        except queue.Empty:pass
        root.after(120,poll)

    def close():
        if busy:messagebox.showinfo('正在處理','請等處理完成後再關閉，避免中斷另存。');return
        root.destroy()

    def button(label,fn):
        b=ttk.Button(toolbar,text=label,command=fn);b.pack(side='left',padx=(0,7));all_buttons.append(b);return b
    button('選擇檔案',add);button('清除清單',remove_all);button('掃描',scan)
    select_btn=button('全選可清理項目',lambda:set_all(True));deselect_btn=button('取消勾選',lambda:set_all(False))
    save_btn=button('儲存掃描報告',save);clean_btn=button('清理選取項目並另存',clean)
    hits.bind('<Button-1>',click);hits.bind('<space>',toggle);hits.bind('<Double-1>',detail)
    for i,p in enumerate(paths):files.insert('','end',iid=str(i),text=p,values=(Path(p).suffix,'待掃描'))
    keywords.trace_add('write',invalidate);prune.trace_add('write',lambda *args:update_controls())
    root.protocol('WM_DELETE_WINDOW',close);update_controls()
    if smoke_test:
        root.update_idletasks()
        result={'version':core.VERSION,'title':root.title(),'buttons':[b.cget('text') for b in all_buttons],
                'files':len(files.get_children()),'clean_disabled_initially':str(clean_btn.cget('state'))=='disabled',
                'author':meta.AUTHOR,'license':meta.LICENSE_NAME,'source_link':'public' if meta.SOURCE_REPO_URL else 'local'}
        root.destroy();return result
    poll();root.mainloop()

def main():
    if len(sys.argv)==3 and sys.argv[1]=='--convert-worker':
        from document_legacy import worker
        return worker(sys.argv[2])
    ap=argparse.ArgumentParser(description='指定內容去敏感化與多格式文件範本清理工具')
    ap.add_argument('files',nargs='*');m=ap.add_mutually_exclusive_group()
    m.add_argument('--scan',action='store_true');m.add_argument('--clean',action='store_true')
    ap.add_argument('--output-dir',type=Path);ap.add_argument('--keywords',help='分號分隔的關鍵字')
    ap.add_argument('--mode',choices=['keyword','block'],default='keyword')
    ap.add_argument('--prune',action='store_true',help='同時清理 PPT 未使用範本')
    ap.add_argument('--selection-json',type=Path,help='JSON：來源完整路徑對應要清理的結果 ID 陣列')
    ap.add_argument('--gui',action='store_true');ap.add_argument('--check-gui',type=Path,help=argparse.SUPPRESS)
    args=ap.parse_args()
    if args.check_gui:
        args.check_gui.write_text(json.dumps(gui(args.files,True),ensure_ascii=False,indent=2),encoding='utf-8');return 0
    if args.gui or not(args.scan or args.clean):gui(args.files);return 0
    if not args.files or(args.clean and not args.output_dir):ap.error('請指定檔案；--clean 必須指定 --output-dir。')
    words=[w.strip() for w in args.keywords.split(';') if w.strip()] if args.keywords else None
    selections=json.loads(args.selection_json.read_text(encoding='utf-8-sig')) if args.selection_json else None
    failed=0;failures=[]
    for i,p in enumerate(args.files,1):
        try:
            if args.clean:
                ids=None if selections is None else selections.get(str(Path(p).resolve()),[])
                r=core.clean_file(p,args.output_dir/f'{i:03d}',ids,words,args.mode,args.prune)
                text=core.clean_summary(r)
            else:
                r=core.scan_file(p,words);text=core.scan_summary(r)
                if args.output_dir:core.save_scan(r,args.output_dir/f'{i:03d}')
            if sys.stdout:print(text)
        except Exception as exc:
            failed+=1
            failures.append({'source':str(p),'status':'not_modified','error':str(exc)})
            if sys.stderr:print(str(p)+'：'+str(exc),file=sys.stderr)
    if failures and args.output_dir:
        args.output_dir.mkdir(parents=True,exist_ok=True)
        with (args.output_dir/'errors.report.json').open('x',encoding='utf-8') as stream:json.dump(failures,stream,ensure_ascii=False,indent=2)
    return 1 if failed else 0
