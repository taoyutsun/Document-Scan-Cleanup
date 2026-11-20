"""Plan and execute a mixed Office batch while keeping the original files."""
from __future__ import annotations

from dataclasses import dataclass
from pathlib import Path
from tempfile import TemporaryDirectory

import document_cleanup_core as core
import document_legacy as legacy
import office_annotations as annotations
import office_merge
import ppt_batch as ppt
import ppt_structure as opc
import word_custom_xml

WORD_FORMATS = {'.docx', '.doc', '.rtf'}
EXCEL_FORMATS = {'.xlsx', '.xls'}
FORMATS = ppt.FORMATS | WORD_FORMATS | EXCEL_FORMATS
MAIN = {'.docx': 'word/document.xml', '.xlsx': 'xl/workbook.xml'}
CONTENT_TYPE = {'.docx': 'wordprocessingml.document.main', '.xlsx': 'spreadsheetml.sheet.main'}


@dataclass(frozen=True)
class Actions:
    ppt: ppt.Options = ppt.Options()
    merge: bool = False
    word_comments: bool = False
    word_accept: bool = False
    excel_comments: bool = False
    excel_notes: bool = False
    word_merge: bool = False
    excel_merge: bool = False


@dataclass(frozen=True)
class Task:
    paths: tuple
    kind: str
    merge: bool = False


def family(path):
    suffix = Path(path).suffix.lower()
    if suffix in ppt.FORMATS:
        return 'PowerPoint'
    if suffix in WORD_FORMATS:
        return 'Word'
    if suffix in EXCEL_FORMATS:
        return 'Excel'
    return ''


def plan(paths, actions, keywords):
    tasks, skipped = [], []
    groups = {}
    for kind, enabled in merge_options(actions).items():
        if enabled:
            group = tuple(path for path in paths if family(path) == kind)
            if len(group) >= 2:
                groups[kind] = group
    selected = {'PowerPoint': any((actions.ppt.notes, actions.ppt.hidden, actions.ppt.comments)),
                'Word': actions.word_comments or actions.word_accept,
                'Excel': actions.excel_comments or actions.excel_notes}
    for path in paths:
        kind = family(path)
        if kind in groups:
            if path == groups[kind][0]:
                tasks.append(Task(groups[kind], kind, True))
            continue
        if kind and (selected[kind] or path in keywords):
            tasks.append(Task((path,), kind))
        else:
            skipped.append({'source': [str(path)], 'status': 'skipped', 'reason': '沒有適用或勾選的處理項目。'})
    return tasks, skipped


def merge_options(actions):
    return {'PowerPoint': actions.merge, 'Word': actions.word_merge, 'Excel': actions.excel_merge}


def read_office(raw, extension, *, editable):
    if len(raw) > core.MAX_SIZE:
        raise ValueError('Office 文件超過 600 MB。')
    data, _ = core.package(raw)
    core.validate_package(data, extension[1:])
    if ppt.part_types(data).get(MAIN[extension]) != 'application/vnd.openxmlformats-officedocument.' + CONTENT_TYPE[extension] + '+xml':
        raise ValueError('請使用不含巨集的標準 DOCX 或 XLSX。')
    root = opc.parse(data[MAIN[extension]])
    if root.tag != {'.docx': annotations.W+'document', '.xlsx': annotations.S+'workbook'}[extension]:
        raise ValueError('Strict Office 格式請先另存為標準 DOCX 或 XLSX。')
    if editable:
        if any(p.lower().startswith('_xmlsignatures/') or 'vbaproject' in p.lower() for p in data):
            raise ValueError('文件含數位簽章或巨集，未輸出。')
        if extension == '.docx' and core.docx_protected(data):
            raise ValueError('Word 文件有編輯或追蹤保護，未輸出。')
        if extension == '.xlsx':
            if root.find(annotations.S+'workbookProtection') is not None or any(
                    opc.parse(payload).find(annotations.S+'sheetProtection') is not None
                    for p, payload in data.items() if p.startswith('xl/worksheets/') and p.endswith('.xml')):
                raise ValueError('Excel 活頁簿或工作表有編輯保護，未輸出。')
    return data


def load_source(path):
    path = Path(path).resolve(strict=True)
    extension = path.suffix.lower()
    limit = 100 * 1024**2 if extension in legacy.CONVERT else core.MAX_SIZE
    if path.stat().st_size > limit:
        raise ValueError('檔案超過處理大小上限。')
    raw = path.read_bytes()
    digest = core.digest(raw)
    if extension in legacy.CONVERT:
        raw = legacy.convert(raw, extension)
        extension = legacy.CONVERT[extension]
    return path, raw, extension, digest


def preview_excel(path):
    if family(path) != 'Excel':
        raise ValueError('請選擇 Excel 活頁簿。')
    source, raw, extension, _ = load_source(path)
    if extension != '.xlsx':
        raise ValueError('請選擇 Excel 活頁簿。')
    return {'source': str(source), 'sheets': annotations.hidden_sheets(read_office(raw, extension, editable=False))}


def verify_protected(before, after, *, word_stores=None):
    def label_parts(data):
        return {name: value for name, value in data.items() if 'labelinfo' in name.lower()}
    if label_parts(before) != label_parts(after):
        raise ValueError('處理變更了敏感度標籤資料，未輸出。')
    if 'word/document.xml' in before:
        if word_stores is None:
            word_custom_xml.verify_rewrite(before, after)
        else:
            word_custom_xml.verify_same(word_stores, word_custom_xml.stores(after))
    elif word_custom_xml.custom_parts(before) != word_custom_xml.custom_parts(after):
        raise ValueError('處理變更了自訂 XML，未輸出。')
    def label_links(data):
        return sorted(opc.internal_target('', node) for node in opc.relations(data, '')
                      if node.get('Type') == ppt.LABEL_REL)
    if label_links(before) != label_links(after):
        raise ValueError('處理變更了敏感度標籤關聯，未輸出。')
    custom = 'docProps/custom.xml'
    def marked_properties(data):
        if custom not in data:
            return []
        return sorted((node.get('name', ''), tuple(node.itertext())) for node in opc.parse(data[custom])
                      if isinstance(node.tag, str) and core.protected_node(custom, node))
    if marked_properties(before) != marked_properties(after):
        raise ValueError('處理變更了敏感度相關屬性，未輸出。')


def prepare_document(path, kind, actions, job):
    source, raw, extension, source_hash = load_source(path)
    original = read_office(raw, extension, editable=True)
    data = original.copy()
    if job:
        with TemporaryDirectory(prefix='document-office-keywords-') as temp:
            result = core.clean_file(source, temp, job.ids, job.words, job.mode, False, job.source_hash)
            data = read_office(Path(result['output']).read_bytes(), extension, editable=True)
    if kind == 'Word':
        before = annotations.word_info(original)
        if actions.word_accept and (before['revision_marks'] or before['tracking']):
            ppt.check_native_input(data)
            comments = annotations.word_comment_text(data)
            data = read_office(legacy.finalize_word(ppt.write_package(data)), extension, editable=True)
            # Word can discard a comment attached to a deleted passage when accepting revisions.
            if not actions.word_comments and annotations.word_comment_text(data) != comments:
                raise ValueError('接受修訂時 Word 變更了要保留的註解。可改用含註解的原檔檢查。')
        if actions.word_comments:
            annotations.remove_word_comments(data)
        after = annotations.word_info(data)
        if ((actions.word_comments and after['comments']) or
                (actions.word_accept and (after['revision_marks'] or after['tracking']))):
            raise ValueError('Word 註解或修訂驗證未通過，未輸出。')
    else:
        before = annotations.excel_info(original)
        if actions.excel_comments or actions.excel_notes:
            annotations.remove_excel_annotations(data, threaded=actions.excel_comments, notes=actions.excel_notes)
        after = annotations.excel_info(data)
        if (after['threaded_comments'] != (0 if actions.excel_comments else before['threaded_comments'])
                or after['notes'] != (0 if actions.excel_notes else before['notes'])
                or after['hidden_sheets'] != before['hidden_sheets']):
            raise ValueError('Excel 註解、附註或工作表驗證未通過，未輸出。')
    verify_protected(original, data)
    core.validate_package(data, extension[1:])
    return source, source_hash, data, before, after, extension


def combine_documents(inputs, kind, extension):
    decks = [item[2] for item in inputs]
    first_labels = None
    for index, (source, _, data, _, _, _) in enumerate(inputs):
        try:
            labels = ppt.check_native_input(data)
        except ValueError as error:
            raise ValueError(f'{source.name}：{error}') from error
        if index == 0:
            first_labels = labels
        elif labels and labels != first_labels:
            raise ValueError('合併來源的有效敏感度標籤與首份檔案不一致，未輸出。')
        if kind != 'Word' and index and word_custom_xml.custom_parts(data):
            raise ValueError(f'{source.name} 含自訂 XML，無法驗證合併後完整保留。')
    expected_stores = word_custom_xml.merge_plan(decks) if kind == 'Word' else None
    names = office_merge.excel_name_plan(decks) if kind == 'Excel' else []
    if names:
        for deck in decks:
            office_merge.excel_layout(deck)
    raw = legacy.merge_office([ppt.write_package(data) for data in decks], extension, sheet_names=names)
    result = read_office(raw, extension, editable=True)
    if expected_stores is not None:
        word_custom_xml.preserve_after_merge(result, expected_stores, decks)
    verify_protected(decks[0], result, word_stores=expected_stores)
    core.validate_package(result, extension[1:])
    if ppt.active_labels(result) != first_labels:
        raise ValueError('Office 另存時變更了有效敏感度標籤，未輸出。')
    if kind == 'Word':
        info = annotations.word_info(result)
        info['custom_xml_stores_retained'] = len(expected_stores)
        comments = sorted(comment for data in decks for comment in annotations.word_comment_text(data))
        if (office_merge.word_body_text(result) != ' '.join(filter(None, map(office_merge.word_body_text, decks))) or
                annotations.word_comment_text(result) != comments or
                info['revision_marks'] != sum(item[4]['revision_marks'] for item in inputs) or
                info['tracking'] != any(item[4]['tracking'] for item in inputs)):
            raise ValueError('Word 合併後的正文、註解或修訂驗證未通過。')
    else:
        office_merge.verify_excel_merge(decks, result, names)
        info = annotations.excel_info(result)
        for key in ('notes', 'threaded_comments'):
            if info[key] != sum(item[4][key] for item in inputs):
                raise ValueError('Excel 合併後的註解或附註數量與來源不一致。')
        info['sheet_groups'] = [{'source': str(item[0]), 'sheets': group} for item, group in zip(inputs, names)]
    return result, info


def process(task, directory, actions, keywords):
    selected = {path: keywords[path] for path in task.paths if path in keywords}
    if task.kind == 'PowerPoint':
        return ppt.process(task.paths, directory, actions.ppt, task.merge, selected)
    if len(task.paths) < 2 and task.merge or len(task.paths) != 1 and not task.merge:
        raise ValueError('合併至少需要同格式的兩份檔案。')
    inputs = [prepare_document(path, task.kind, actions, selected.get(path)) for path in task.paths]
    source, _, data, _, after, extension = inputs[0]
    if task.merge:
        data, after = combine_documents(inputs, task.kind, extension)
    raw_out = ppt.write_package(data)
    for path, source_hash, *_ in inputs:
        if core.digest(path.read_bytes()) != source_hash:
            raise ValueError(f'{path.name} 在處理時已變更，請重新執行。')
    directory.mkdir(parents=True, exist_ok=True)
    name = ({'Word': '合併文件', 'Excel': '合併活頁簿'}[task.kind] if task.merge else source.stem+'.批次清理')
    output = directory / (name+extension)
    with output.open('xb') as stream:
        try:
            stream.write(raw_out)
        except BaseException:
            stream.close()
            output.unlink(missing_ok=True)
            raise
    return {'source': [str(item[0]) for item in inputs], 'output': str(output), 'before': [item[3] for item in inputs],
            'after': after, 'keyword_items': sum(len(job.ids) for job in selected.values())}
