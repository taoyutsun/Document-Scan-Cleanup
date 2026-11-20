"""Tk interface for the PowerPoint tool tab."""
import json
import os
import queue
import re
import threading
import tkinter as tk
from pathlib import Path
from tkinter import messagebox, ttk

import ppt_batch


class BatchView(ttk.Frame):
    def __init__(self, parent, *, add, remove, clear, destination, keyword_jobs, other_busy, update):
        super().__init__(parent, padding=(14, 12))
        self.destination = destination
        self.keyword_jobs = keyword_jobs
        self.other_busy = other_busy
        self.update = update
        self.busy = False
        self.events = queue.Queue()
        self.paths = []
        self.result = None
        self.columnconfigure(0, weight=1)
        self.rowconfigure(2, weight=2, minsize=54)
        self.rowconfigure(6, weight=1, minsize=35)
        self.intro = ttk.Label(self, text='加入或拖曳簡報至清單。合併、刪除隱藏頁及舊檔轉換需桌面版 PowerPoint。', style='Hint.TLabel')
        self.intro.grid(row=0, column=0, sticky='w', pady=(0, 8))
        toolbar = ttk.Frame(self)
        toolbar.grid(row=1, column=0, sticky='ew', pady=(0, 7))
        self.buttons = []
        actions = [('加入檔案', add), ('清除清單', clear), ('移除選取檔案', lambda: remove(self.selected())),
                   ('上移', lambda: self.move(-1)), ('下移', lambda: self.move(1)), ('依檔名排序', self.sort)]
        for label, action in actions:
            button = ttk.Button(toolbar, text=label, command=action)
            button.pack(side='left', padx=(0, 7))
            self.buttons.append(button)
        table = ttk.Frame(self)
        table.grid(row=2, column=0, sticky='nsew')
        self.files = ttk.Treeview(table, columns=('name', 'path'), show='tree headings', height=5, selectmode='extended')
        self.files.heading('#0', text='順序')
        self.files.column('#0', width=66, minwidth=50, stretch=False, anchor='center')
        self.files.heading('name', text='簡報')
        self.files.column('name', width=270, minwidth=120)
        self.files.heading('path', text='檔案位置')
        self.files.column('path', width=500, minwidth=100)
        self.files.pack(side='left', fill='both', expand=True)
        scroll = ttk.Scrollbar(table, orient='vertical', command=self.files.yview)
        scroll.pack(side='right', fill='y')
        self.files.configure(yscrollcommand=scroll.set)
        self.files.bind('<<TreeviewSelect>>', lambda event: self.update_controls())

        options = ttk.LabelFrame(self, text='這次要處理的項目', padding=(10, 5))
        options.grid(row=3, column=0, sticky='ew', pady=(9, 4))
        self.notes = tk.BooleanVar()
        self.hidden = tk.BooleanVar()
        self.comments = tk.BooleanVar()
        self.keywords = tk.BooleanVar()
        self.merge = tk.BooleanVar()
        self.checks = []
        for index, (label, var) in enumerate([
            ('刪除備忘稿', self.notes), ('刪除隱藏投影片', self.hidden), ('刪除註解', self.comments),
            ('套用已勾選的關鍵字項目', self.keywords), ('按清單順序合併為一份', self.merge),
        ]):
            check = ttk.Checkbutton(options, text=label, variable=var, command=self.update_controls)
            check.grid(row=index // 3, column=index % 3, sticky='w', padx=(0, 20), pady=2)
            self.checks.append(check)
        self.select_cleanup = ttk.Button(options, text='勾選三項清理', command=self.select_all_cleanup)
        self.select_cleanup.grid(row=0, column=3, sticky='e')
        options.columnconfigure(3, weight=1)
        self.keyword_hint = tk.StringVar()
        ttk.Label(self, textvariable=self.keyword_hint, style='Hint.TLabel').grid(row=4, column=0, sticky='w', pady=(0, 4))
        controls = ttk.Frame(self)
        controls.grid(row=5, column=0, sticky='ew', pady=(0, 6))
        self.run_button = ttk.Button(controls, text='執行並另存', style='Accent.TButton', command=self.run)
        self.run_button.pack(side='left', padx=(0, 8))
        self.open_button = ttk.Button(controls, text='開啟輸出資料夾', command=self.open_result)
        self.open_button.pack(side='left')
        self.status = tk.StringVar(value='清理會另存副本。刪除隱藏頁與合併需 Windows 桌面版 PowerPoint。')
        self.status_label = ttk.Label(controls, textvariable=self.status, style='Hint.TLabel')
        self.status_label.pack(side='left', padx=12)
        log_box = ttk.Frame(self)
        log_box.grid(row=6, column=0, sticky='nsew')
        self.log = tk.Text(log_box, height=4, wrap='word', state='disabled', background='#ffffff',
                           foreground='#193449', font=('Microsoft JhengHei', 10), relief='flat', padx=8, pady=5)
        self.log.pack(side='left', fill='both', expand=True)
        scroll = ttk.Scrollbar(log_box, orient='vertical', command=self.log.yview)
        scroll.pack(side='right', fill='y')
        self.log.configure(yscrollcommand=scroll.set)

    def selected(self):
        return [self.paths[int(i)] for i in self.files.selection()]

    def select_all_cleanup(self):
        for option in (self.notes, self.hidden, self.comments):
            option.set(True)
        self.update_controls()

    def sync(self, files):
        supported = [path for path in files if Path(path).suffix.lower() in ppt_batch.FORMATS]
        if set(supported) != set(self.paths):
            self.paths = [path for path in self.paths if path in supported]
            self.paths.extend(path for path in supported if path not in self.paths)
            self.draw()
        self.update_controls()

    def draw(self, selected=()):
        self.files.delete(*self.files.get_children())
        for index, path in enumerate(self.paths):
            label = 'even' if index % 2 == 0 else 'odd'
            self.files.insert('', 'end', iid=str(index), text=str(index + 1),
                              values=(Path(path).name, str(Path(path).parent)), tags=(label,))
            if path in selected:
                self.files.selection_add(str(index))
        self.files.tag_configure('odd', background='#edf4f7')
        self.update_controls()

    def move(self, direction):
        selected = set(self.selected())
        indices = range(1, len(self.paths)) if direction == -1 else range(len(self.paths) - 2, -1, -1)
        for index in indices:
            other = index + direction
            if self.paths[index] in selected and self.paths[other] not in selected:
                self.paths[index], self.paths[other] = self.paths[other], self.paths[index]
        self.draw(selected)

    def sort(self):
        selected = self.selected()
        self.paths.sort(key=lambda path: [int(word) if word.isdecimal() else word.lower()
                                         for word in re.split(r'(\d+)', Path(path).name)])
        self.draw(selected)

    def update_controls(self):
        blocked = self.busy or self.other_busy()
        selected = self.files.selection()
        available = self.keyword_jobs(self.paths)
        self.keyword_hint.set(f'目前有 {sum(len(job.ids) for job in available.values())} 個已勾選的關鍵字項目。'
                              '可到「關鍵字掃描」確認或調整勾選。')
        for index, button in enumerate(self.buttons):
            enabled = bool(not blocked and (index < 2 or (selected if index < 5 else len(self.paths) > 1)))
            button.configure(state='normal' if enabled else 'disabled')
        for check in self.checks:
            check.configure(state='disabled' if blocked else 'normal')
        self.select_cleanup.configure(state='disabled' if blocked else 'normal')
        enough = len(self.paths) >= (2 if self.merge.get() else 1)
        actions = any(var.get() for var in (self.notes, self.hidden, self.comments, self.merge))
        valid_keywords = not self.keywords.get() or bool(available)
        enabled = not blocked and enough and valid_keywords and (actions or self.keywords.get())
        self.run_button.configure(state='normal' if enabled else 'disabled')
        self.open_button.configure(state='normal' if self.result and not blocked else 'disabled')

    def write_log(self, value, clear=False):
        self.log.configure(state='normal')
        if clear:
            self.log.delete('1.0', 'end')
        self.log.insert('end', value)
        self.log.see('end')
        self.log.configure(state='disabled')

    def open_result(self):
        try:
            os.startfile(str(self.result))
        except OSError as error:
            messagebox.showerror('無法開啟輸出資料夾', str(error), parent=self)

    def run(self):
        if self.busy or self.other_busy():
            return
        target = self.destination('PowerPoint_Batch')
        if not target:
            return
        paths = self.paths.copy()
        options = ppt_batch.Options(self.notes.get(), self.hidden.get(), self.comments.get())
        merge = self.merge.get()
        keywords = self.keyword_jobs(paths) if self.keywords.get() else {}
        if not (options.notes or options.hidden or options.comments or merge):
            paths = [path for path in paths if path in keywords]
        self.result = None
        self.busy = True
        self.update()
        self.write_log('', clear=True)
        self.status.set('正在處理簡報……')

        def worker():
            results = []
            groups = [paths] if merge else [[path] for path in paths]
            for index, group in enumerate(groups, 1):
                try:
                    selection = {path: keywords[path] for path in group if path in keywords}
                    result = ppt_batch.process(group, target if merge else target / f'{index:03d}', options, merge, selection)
                    results.append(result)
                    after = result['after']
                    self.events.put(('log', f"完成 {Path(result['output']).name}，共 {after['slides']} 張。\n"))
                except Exception as error:
                    results.append({'source': group, 'error': str(error)})
                    self.events.put(('log', f"未輸出 {', '.join(Path(p).name for p in group)}：{error}\n"))
            try:
                target.mkdir(parents=True, exist_ok=True)
                with (target / 'batch.report.json').open('x', encoding='utf-8') as stream:
                    json.dump({'operations': {'notes': options.notes, 'hidden': options.hidden, 'comments': options.comments,
                                              'merge': merge}, 'results': results}, stream, ensure_ascii=False, indent=2)
            except OSError as error:
                self.events.put(('log', f'未能寫入批次報告：{error}\n'))
            success = sum('output' in item for item in results)
            self.events.put(('done', (success, len(groups) - success, target)))

        threading.Thread(target=worker, daemon=True).start()
        self.after(100, self.poll)

    def poll(self):
        try:
            while True:
                kind, value = self.events.get_nowait()
                if kind == 'log':
                    self.write_log(value)
                else:
                    success, errors, target = value
                    self.result = target if target.is_dir() else None
                    self.busy = False
                    self.status.set(f'處理完成。成功 {success} 份，未輸出 {errors} 份。')
                    self.update()
        except queue.Empty:
            if self.busy:
                self.after(100, self.poll)
