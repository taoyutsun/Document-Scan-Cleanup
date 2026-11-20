"""Merge checks and native Office integration for Word and Excel."""
import re

import office_annotations as annotations
import ppt_batch as ppt
import ppt_structure as opc

S, W = annotations.S, annotations.W


def workbook_sheets(data):
    return list(opc.parse(data['xl/workbook.xml']).findall(S+'sheets/'+S+'sheet'))


def workbook_names(data):
    root = opc.parse(data['xl/workbook.xml'])
    return {e.get('name', '').casefold() for e in root.findall(S+'definedNames/'+S+'definedName')
            if e.get('localSheetId') is None}


def excel_name_plan(decks):
    used, scoped_names, tables = set(), set(), set()
    result = []
    systems = []
    for data in decks:
        root = opc.parse(data['xl/workbook.xml'])
        options = root.find(S+'workbookPr')
        systems.append(options is not None and options.get('date1904', '0') in {'1', 'true'})
        if any(name.startswith(('xl/externalLinks/', 'xl/pivotTables/', 'xl/pivotCache/', 'xl/model/'))
               or name == 'xl/connections.xml' for name in data):
            raise ValueError('Excel 合併暫不處理外部連結、資料連線、資料模型或樞紐分析的活頁簿。')
        names = workbook_names(data)
        if scoped_names & names:
            raise ValueError('Excel 活頁簿含同名的全簿命名範圍，請先更名再合併。')
        scoped_names.update(names)
        current_tables = {opc.parse(payload).get('displayName', '').casefold() for name, payload in data.items()
                          if name.startswith('xl/tables/') and name.endswith('.xml')}
        if current_tables & tables:
            raise ValueError('Excel 活頁簿含同名的表格，請先更名再合併。')
        tables.update(current_tables)
        sheet_names = [sheet.get('name') for sheet in workbook_sheets(data)]
        reserved = used | {name.casefold() for name in sheet_names}
        planned = []
        for name in sheet_names:
            candidate = name
            if candidate.casefold() in used:
                index = 2
                while True:
                    suffix = f' ({index})'
                    candidate = name[:31-len(suffix)] + suffix
                    if candidate.casefold() not in reserved | used:
                        break
                    index += 1
            if candidate != name and any(re.search(rb'\b(?:INDIRECT|HYPERLINK)\s*\(', payload, re.I)
                                         for part, payload in data.items() if part.startswith('xl/') and part.endswith('.xml')):
                raise ValueError('Excel 同名工作表需要更名，且檔案含文字引用公式。請先調整工作表名稱再合併。')
            used.add(candidate.casefold())
            planned.append(candidate)
        result.append(planned)
    if len(set(systems)) > 1:
        raise ValueError('Excel 活頁簿使用不同日期系統，請統一後再合併。')
    return result


def excel_layout(data):
    workbook = 'xl/workbook.xml'
    targets = {rel.get('Id'): (rel.get('Type', ''), opc.internal_target(workbook, rel))
               for rel in opc.relations(data, workbook)}
    strings = []
    if 'xl/sharedStrings.xml' in data:
        strings = [''.join(t.text or '' for t in e.iter(S+'t'))
                   for e in opc.parse(data['xl/sharedStrings.xml']).findall(S+'si')]
    layout = []
    for sheet in workbook_sheets(data):
        kind, part = targets[sheet.get(ppt.RID)]
        if not kind.endswith('/worksheet'):
            raise ValueError('Excel 合併目前支援一般工作表。')
        cells = []
        for cell in opc.parse(data[part]).iter(S+'c'):
            formula = cell.find(S+'f')
            if formula is not None:
                cells.append((cell.get('r'), 'formula', '#REF!' in (formula.text or '')))
                continue
            kind = cell.get('t', 'n')
            value = cell.findtext(S+'v')
            if kind == 's' and value is not None:
                kind, value = 'text', strings[int(value)]
            elif kind == 'inlineStr':
                kind, value = 'text', ''.join(t.text or '' for t in cell.iter(S+'t'))
            if value is not None:
                cells.append((cell.get('r'), kind, value))
        layout.append((sheet.get('name'), sheet.get('state', 'visible'), cells))
    return layout


def verify_excel_merge(inputs, output, names):
    expected = [(name, state, cells) for data, group_names in zip(inputs, names)
                for (_, state, cells), name in zip(excel_layout(data), group_names)]
    if excel_layout(output) != expected:
        raise ValueError('Excel 合併後的工作表順序、儲存格或隱藏狀態驗證未通過。')
    if workbook_names(output) != set().union(*(workbook_names(data) for data in inputs)):
        raise ValueError('Excel 合併後的全簿命名範圍驗證未通過。')
    if any(name.startswith('xl/externalLinks/') for name in output):
        raise ValueError('Excel 合併後出現外部活頁簿引用，未輸出。')


def word_body_text(data):
    body = opc.parse(data['word/document.xml']).find(W+'body')
    return ' '.join(' '.join(''.join(node.text or '' for node in p.iter(W+'t'))
                             for p in body.iter(W+'p')).split())


def merge_word_documents(app, result, sources):
    track = bool(result.TrackRevisions)
    result.TrackRevisions = False
    for path in sources:
        incoming = app.Documents.OpenNoRepairDialog(path, ConfirmConversions=False, ReadOnly=True,
                                                     AddToRecentFiles=False, Visible=False, OpenAndRepair=False)
        try:
            if incoming.HasVBProject or incoming.ProtectionType != -1:
                raise ValueError('Word 合併來源含巨集或編輯保護。')
            track = track or bool(incoming.TrackRevisions)
            start = result.Sections.Count + 1
            end = result.Content.End - 1
            result.Range(end, end).InsertBreak(2)  # wdSectionBreakNextPage
            end = result.Content.End - 1
            result.Range(end, end).InsertFile(path, ConfirmConversions=False, Link=False, Attachment=False)
            if result.Sections.Count != start + incoming.Sections.Count - 1:
                raise ValueError('Word 合併後分節數量與來源不一致。')
            for i in range(1, incoming.Sections.Count+1):
                copy_word_section(incoming.Sections(i), result.Sections(start+i-1), first=i==1)
        finally:
            incoming.Close(False)
    result.TrackRevisions = track


def copy_word_section(source, destination, *, first):
    original = source.PageSetup
    updated = destination.PageSetup
    # Word's InsertFile imports internal section breaks, but the final section
    # settings and header/footer stories need separate handling.
    for prop in ('Orientation', 'PageWidth', 'PageHeight', 'TopMargin', 'BottomMargin',
                 'LeftMargin', 'RightMargin', 'Gutter', 'HeaderDistance', 'FooterDistance',
                 'DifferentFirstPageHeaderFooter', 'OddAndEvenPagesHeaderFooter'):
        setattr(updated, prop, getattr(original, prop))
    if first:
        updated.SectionStart = 2  # wdSectionNewPage
    for field in ('Headers', 'Footers'):
        source_items = getattr(source, field)
        dest_items = getattr(destination, field)
        for i in range(1, 4):
            previous = source_items(i)
            current = dest_items(i)
            if previous.Exists or current.Exists:
                current.LinkToPrevious = False
                if previous.Exists:
                    current.Range.FormattedText = previous.Range.FormattedText
                else:
                    current.Range.Text = ''


def merge_excel_workbooks(app, result, sources, names):
    for path, assigned in zip(sources, names[1:]):
        incoming = app.Workbooks.Open(path, UpdateLinks=0, ReadOnly=True, AddToMru=False,
                                      IgnoreReadOnlyRecommended=True)
        try:
            if incoming.HasVBProject or incoming.ProtectStructure or incoming.Connections.Count:
                raise ValueError('Excel 合併來源含巨集、保護或資料連線。')
            if incoming.Sheets.Count != len(assigned):
                raise ValueError('Excel 開啟後的工作表數量與來源不同。')
            source_visibility = []
            for i, name in enumerate(assigned, 1):
                sheet = incoming.Sheets(i)
                source_visibility.append(sheet.Visible)
                if sheet.Name != name:
                    sheet.Name = name
                if sheet.Visible != -1:
                    sheet.Visible = -1
            # Copy each workbook's sheets as one group so cross-sheet references
            # are remapped within the merged workbook by Excel.
            count = result.Sheets.Count
            last = result.Sheets(count)
            visibility = last.Visible
            if visibility != -1:
                last.Visible = -1
            try:
                # Passing Missing through this COM collection can drop After.
                incoming.Worksheets.Copy(None, last)
            finally:
                if visibility != -1:
                    last.Visible = visibility
            if result.Sheets.Count != count + len(assigned):
                raise ValueError('Excel 未將完整工作表群組插入指定的活頁簿。')
            for name, visibility in zip(assigned, source_visibility):
                result.Sheets(name).Visible = visibility
        finally:
            incoming.Close(False)
