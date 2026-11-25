"""Batch operations for sharing copies of PowerPoint presentations."""
from __future__ import annotations

import json
import os
import subprocess
import sys
import tempfile
import time
from dataclasses import dataclass
from io import BytesIO
from pathlib import Path
from zipfile import ZIP_DEFLATED, ZipFile

from lxml import etree

import document_cleanup_core as core
import ppt_structure as opc

FORMATS = {'.pptx', '.ppt', '.pps'}
NOTES = {'application/vnd.openxmlformats-officedocument.presentationml.notesSlide+xml'}
COMMENTS = {
    'application/vnd.openxmlformats-officedocument.presentationml.comments+xml',
    'application/vnd.ms-powerpoint.comments+xml',
}
AUTHORS = {
    'application/vnd.openxmlformats-officedocument.presentationml.commentAuthors+xml',
    'application/vnd.ms-powerpoint.authors+xml',
}
PRESENTATION = 'ppt/presentation.xml'
ANIMATION_PART_TYPES = {
    'application/vnd.openxmlformats-officedocument.presentationml.slide+xml',
    'application/vnd.openxmlformats-officedocument.presentationml.slideLayout+xml',
    'application/vnd.openxmlformats-officedocument.presentationml.slideMaster+xml',
}
TIMING = '{' + opc.P + '}timing'
RID = '{' + opc.R + '}id'
MODERN_COMMENT = '{http://schemas.microsoft.com/office/powerpoint/2018/8/main}commentRel'
LABEL_NS = '{http://schemas.microsoft.com/office/2020/mipLabelMetadata}'
LABEL_REL = 'http://schemas.microsoft.com/office/2020/02/relationships/classificationlabels'


@dataclass(frozen=True)
class Options:
    notes: bool = False
    hidden: bool = False
    comments: bool = False
    animations: bool = False

    @property
    def has_cleanup(self):
        return self.notes or self.hidden or self.comments or self.animations


@dataclass(frozen=True)
class KeywordSelection:
    ids: tuple
    words: tuple
    source_hash: str
    mode: str
    prune: bool


def read_presentation(raw):
    if len(raw) > core.MAX_SIZE:
        raise ValueError('簡報超過 600 MB。')
    data, _ = core.package(raw)
    if any(name.lower().startswith('_xmlsignatures/') for name in data):
        raise ValueError('簡報含數位簽章，請使用已授權的可編輯副本。')
    core.validate_package(data, 'pptx')
    types = part_types(data)
    if types.get(PRESENTATION) != 'application/vnd.openxmlformats-officedocument.presentationml.presentation.main+xml':
        raise ValueError('請使用標準、未加密且不含巨集的 PPTX 簡報。')
    if any('vbaproject' in name.lower() or 'vbaproject' in kind.lower() for name, kind in types.items()):
        raise ValueError('簡報含有巨集。')
    return data


def part_types(data):
    return {node.get('PartName').lstrip('/'): node.get('ContentType', '')
            for node in opc.parse(data['[Content_Types].xml']) if node.get('PartName')}


def slide_parts(data):
    root = opc.parse(data[PRESENTATION])
    relations = {rel.get('Id'): opc.internal_target(PRESENTATION, rel)
                 for rel in opc.relations(data, PRESENTATION) if rel.get('Type', '').endswith('/slide')}
    if root.tag != '{' + opc.P + '}presentation':
        raise ValueError('請先在 PowerPoint 另存標準 PPTX。')
    slides = [relations.get(node.get(RID)) for node in root.findall('p:sldIdLst/p:sldId', opc.NS)]
    if not slides or None in slides or len(set(slides)) != len(slides):
        raise ValueError('找不到有效的投影片清單。')
    return slides


def inspect(data):
    slides = slide_parts(data)
    types = part_types(data)
    return {
        'slides': len(slides),
        'hidden': sum(opc.parse(data[part]).get('show', '1').lower() in {'0', 'false'} for part in slides),
        'notes': sum(kind in NOTES for kind in types.values()),
        'comment_parts': sum(kind in COMMENTS for kind in types.values()),
    }


def slide_annotations(data, options, skip_hidden=False):
    """Check that native import keeps unselected notes and comment text in order."""
    types = part_types(data)
    signatures = []
    for slide in slide_parts(data):
        if skip_hidden and opc.parse(data[slide]).get('show', '1').lower() in {'0', 'false'}:
            continue
        notes = []
        comments = []
        for rel in opc.relations(data, slide):
            part = opc.internal_target(slide, rel)
            kind = types.get(part)
            if kind in NOTES and not options.notes:
                root = opc.parse(data[part])
                for shape in root.findall('.//p:sp', opc.NS):
                    placeholder = shape.find('p:nvSpPr/p:nvPr/p:ph', opc.NS)
                    if placeholder is not None and placeholder.get('type') == 'body':
                        notes.extend(node.text for node in shape.iter('{' + core.A + '}t') if node.text)
            if kind in COMMENTS and not options.comments:
                root = opc.parse(data[part])
                comments.extend(node.text for node in root.iter() if core.local(node) in {'text', 't'} and node.text)
        signatures.append((tuple(notes), tuple(comments)))
    return signatures


def write_package(data):
    stream = BytesIO()
    with ZipFile(stream, 'w', ZIP_DEFLATED, allowZip64=True) as archive:
        for name, content in data.items():
            archive.writestr(name, content)
    return stream.getvalue()


def remove_annotations(data, options):
    kinds = (NOTES if options.notes else set()) | (COMMENTS | AUTHORS if options.comments else set())
    unwanted = {part for part, kind in part_types(data).items() if kind in kinds}
    if not unwanted:
        return
    candidates = opc.reachable(data, unwanted)
    for name in list(data):
        if not name.endswith('.rels'):
            continue
        owner = opc.relowner(name)
        if owner in unwanted:
            continue
        relationships = opc.parse(data[name])
        removed_ids = set()
        for relation in list(relationships):
            if isinstance(relation.tag, str) and opc.internal_target(owner, relation) in unwanted:
                removed_ids.add(relation.get('Id'))
                relationships.remove(relation)
        if not removed_ids:
            continue
        data[name] = opc.xml(relationships)
        if not owner or not owner.endswith('.xml'):
            continue
        root = opc.parse(data[owner])
        changed = False
        for element in list(root.iter()):
            if not isinstance(element.tag, str):
                continue
            if not any(etree.QName(attr).namespace == opc.R and value in removed_ids
                       for attr, value in element.attrib.items()):
                continue
            if element.tag != MODERN_COMMENT:
                raise ValueError('發現未支援的備忘稿或註解關聯。')
            parent = element.getparent()
            parent.remove(element)
            while not len(parent) and parent.tag in {'{' + opc.P + '}ext', '{' + opc.P + '}extLst'}:
                empty_extension = parent
                parent = parent.getparent()
                parent.remove(empty_extension)
            changed = True
        if changed:
            data[owner] = opc.xml(root)
    # Keep shared dependencies, including images also used on visible slides.
    # Follow references from all parts outside the candidate set, even if a part
    # was already orphaned in the input.
    owners = {opc.relowner(name) for name in data if name.endswith('.rels')}
    candidates -= opc.reachable(data, {''} | (owners - candidates))
    deleted = candidates | unwanted | {opc.relpath(part) for part in unwanted}
    if any(core.protected_part(name) for name in deleted if not name.endswith('.rels')):
        raise ValueError('移除此內容會影響受保護的零件。')
    for name in deleted:
        data.pop(name, None)
    types = opc.parse(data['[Content_Types].xml'])
    for node in list(types):
        if node.get('PartName', '').lstrip('/') in deleted:
            types.remove(node)
    data['[Content_Types].xml'] = opc.xml(types)
    if options.notes and 'docProps/app.xml' in data:
        props = opc.parse(data['docProps/app.xml'])
        count = props.find('ep:Notes', opc.NS)
        if count is not None:
            count.text = '0'
            data['docProps/app.xml'] = opc.xml(props)
    core.validate_package(data, 'pptx')


def prepare(path, options, keywords=None):
    path = Path(path).resolve(strict=True)
    extension = path.suffix.lower()
    if extension not in FORMATS or path.name.startswith('~$'):
        raise ValueError('請選擇 PPTX、PPT 或 PPS 簡報。')
    if path.stat().st_size > (core.MAX_SIZE if extension == '.pptx' else 100 * 1024 * 1024):
        raise ValueError('檔案超過大小限制。')
    raw = path.read_bytes()
    source_hash = core.digest(raw)
    if keywords is not None:
        with tempfile.TemporaryDirectory(prefix='document-ppt-keywords-') as temporary:
            result = core.clean_file(path, temporary, keywords.ids, keywords.words, keywords.mode,
                                     keywords.prune, keywords.source_hash)
            raw = Path(result['output']).read_bytes()
    elif extension != '.pptx':
        from document_legacy import convert
        raw = convert(raw, extension)
    data = read_presentation(raw)
    before = inspect(data)
    remove_annotations(data, options)
    return path, source_hash, data, before


def remove_animation_timelines(data):
    """Remove shared and per-slide timing without editing transitions or media parts."""
    removed = 0
    for part, kind in part_types(data).items():
        if kind not in ANIMATION_PART_TYPES:
            continue
        root = opc.parse(data[part])
        timelines = list(root.iter(TIMING))
        if not timelines:
            continue
        for timeline in timelines:
            timeline.getparent().remove(timeline)
        data[part] = opc.xml(root)
        removed += len(timelines)
    return removed


def active_labels(data):
    """Read the Office label state without changing the label or its history."""
    linked = [opc.internal_target('', relation) for relation in opc.relations(data, '')
              if relation.get('Type') == LABEL_REL]
    suspected = {name for name in data if name.lower().endswith('/labelinfo.xml')}
    if len(linked) > 1 or not suspected.issubset(linked) or (linked and linked[0] not in data):
        raise ValueError('敏感度標籤資料的關聯無法確認。')
    custom = 'docProps/custom.xml'
    if custom in data and any(core.protected_node(custom, node) for node in opc.parse(data[custom])
                              if isinstance(node.tag, str)):
        raise ValueError('含舊格式或未知的敏感度分類屬性，無法驗證其保存狀態。')
    if not linked:
        return frozenset()
    root = opc.parse(data[linked[0]])
    if root.tag != LABEL_NS + 'labelList':
        raise ValueError('敏感度標籤資料採用未支援的格式。')
    truth = {'1': '1', 'true': '1', '0': '0', 'false': '0'}
    signatures = set()
    for node in root:
        if not isinstance(node.tag, str):
            continue
        enabled = truth.get(node.get('enabled'))
        removed = truth.get(node.get('removed', '0'))
        if node.tag != LABEL_NS + 'label' or len(node) or enabled is None or removed is None:
            raise ValueError('無法確認敏感度標籤的狀態。')
        if enabled == '0' and removed == '1':
            continue
        if enabled != '1' or removed != '0' or not node.get('id') or not node.get('siteId'):
            raise ValueError('無法確認敏感度標籤的狀態。')
        fields = dict(node.attrib, enabled=enabled, removed=removed)
        signatures.add(tuple(sorted(fields.items())))
    return frozenset(signatures)


def check_native_input(data):
    for name, payload in data.items():
        if name.lower().startswith(('ppt/activex/', 'word/activex/', 'xl/activex/')):
            raise ValueError('簡報含 ActiveX 元件，無法使用 PowerPoint 自動處理。')
        if name.endswith('.rels'):
            for rel in opc.parse(payload):
                if rel.get('TargetMode') == 'External' and not rel.get('Type', '').endswith('/hyperlink'):
                    raise ValueError('簡報含外部連結物件，請先處理連結後再執行。')
    return active_labels(data)


def native_plan(decks, paths):
    expected = frozenset()
    base = 0
    for index, (data, path) in enumerate(zip(decks, paths)):
        try:
            labels = check_native_input(data)
        except (ValueError, etree.XMLSyntaxError) as error:
            raise ValueError(f'{path.name}：{error}') from error
        if not labels:
            continue
        if expected and expected != labels:
            raise ValueError(f'{paths[base].name} 與 {path.name} 的有效敏感度標籤不一致。請依公司規範確認合併檔應使用的分類。')
        if not expected:
            base, expected = index, labels
    return base, expected


def native_transform(decks, hidden, base):
    if sys.platform != 'win32':
        raise ValueError('合併或刪除隱藏投影片需要 Windows 桌面版 PowerPoint。')
    with tempfile.TemporaryDirectory(prefix='document-ppt-batch-') as temporary:
        folder = Path(temporary)
        inputs = []
        for i, data in enumerate(decks, 1):
            path = folder / f'input-{i}.pptx'
            path.write_bytes(write_package(data))
            inputs.append(str(path))
        request = folder / 'request.json'
        response = folder / 'response.json'
        output = folder / 'output.pptx'
        request.write_text(json.dumps({'inputs': inputs, 'base': base, 'response': str(response),
                                       'output': str(output), 'hidden': hidden}), encoding='utf-8')
        command = [sys.executable]
        if not getattr(sys, 'frozen', False):
            command.append(str(Path(__file__).with_name('document_scan_cleanup.py')))
        try:
            process = subprocess.run([*command, '--ppt-batch-worker', str(request)],
                                     creationflags=subprocess.CREATE_NO_WINDOW, timeout=180,
                                     stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL)
        except subprocess.TimeoutExpired as error:
            raise ValueError('PowerPoint 處理超過 180 秒。請檢查是否有 PowerPoint 對話框，再重試。') from error
        state = json.loads(response.read_text(encoding='utf-8')) if response.exists() else {}
        if process.returncode or not output.is_file():
            raise ValueError(state.get('error', 'PowerPoint 無法完成此次處理。'))
        return read_presentation(output.read_bytes())


def powerpoint_processes():
    import re
    result = subprocess.run(['tasklist.exe', '/FI', 'IMAGENAME eq POWERPNT.EXE', '/FO', 'CSV', '/NH'],
                            capture_output=True, creationflags=subprocess.CREATE_NO_WINDOW, timeout=10, check=True)
    return {int(pid) for pid in re.findall(rb'"POWERPNT\.EXE","(\d+)"', result.stdout, re.I)}


def native_worker(request_path):
    import pythoncom
    import win32com.client
    request = json.loads(Path(request_path).read_text(encoding='utf-8'))
    response = Path(request['response'])
    app = deck = incoming = slide = None
    owned_pid = None
    pythoncom.CoInitialize()
    try:
        if powerpoint_processes():
            raise ValueError('請先關閉桌面版 PowerPoint，再執行合併或刪除隱藏投影片。')
        app = win32com.client.DispatchEx('PowerPoint.Application')
        started = powerpoint_processes()
        if len(started) != 1 or app.Presentations.Count:
            raise ValueError('無法確認 PowerPoint 專用程序，請關閉 PowerPoint 再重試。')
        owned_pid = started.pop()
        app.AutomationSecurity = 3  # msoAutomationSecurityForceDisable
        app.DisplayAlerts = 1      # ppAlertsNone
        sources = request['inputs']
        base = request['base']
        deck = app.Presentations.Open(sources[base], 0, 0, 0)
        # Use the labelled file as the base while preserving the user's slide order.
        order = list(range(base - 1, -1, -1)) + list(range(base + 1, len(sources)))
        for index in order:
            source = sources[index]
            incoming = app.Presentations.Open(source, -1, 0, 0)
            try:
                size = (incoming.PageSetup.SlideWidth, incoming.PageSetup.SlideHeight)
                if any(abs(a - b) > 0.1 for a, b in zip(size, (deck.PageSetup.SlideWidth, deck.PageSetup.SlideHeight))):
                    raise ValueError('合併的投影片尺寸不同，請先統一尺寸。')
                count = incoming.Slides.Count
                layouts = [(slide.Design.Index, slide.CustomLayout.Index) for slide in incoming.Slides]
                start = 0 if index < base else deck.Slides.Count
                if deck.Slides.InsertFromFile(source, start) != count:
                    raise ValueError('PowerPoint 匯入的投影片張數不一致。')
                cloned = {index: deck.Designs.Clone(incoming.Designs(index))
                          for index in dict.fromkeys(design for design, _ in layouts)}
                for i, (design, layout) in enumerate(layouts, start + 1):
                    inserted = deck.Slides(i)
                    inserted.CustomLayout = cloned[design].SlideMaster.CustomLayouts(layout)
                inserted = None
                cloned.clear()
            finally:
                incoming.Close()
                incoming = None
        if request['hidden']:
            for index in range(deck.Slides.Count, 0, -1):
                slide = deck.Slides(index)
                if slide.SlideShowTransition.Hidden == -1:
                    slide.Delete()
                slide = None
        if not deck.Slides.Count:
            raise ValueError('移除後沒有投影片，未產生輸出檔。')
        deck.SaveCopyAs(request['output'], 24)
        response.write_text('{}', encoding='utf-8')
        return 0
    except Exception as error:
        response.write_text(json.dumps({'error': str(error)}, ensure_ascii=False), encoding='utf-8')
        return 1
    finally:
        if deck is not None:
            try:
                deck.Saved = -1  # Discard edits made to the temporary working copy.
                deck.Close()
            except Exception:
                pass
        if owned_pid is not None:
            try:
                app.Quit()
            except Exception:
                pass
        slide = incoming = deck = app = None
        pythoncom.CoUninitialize()
        if owned_pid is not None:
            # Consecutive batch files need the server to finish exiting before
            # the next worker tests whether PowerPoint is already in use.
            deadline = time.monotonic() + 5
            while time.monotonic() < deadline:
                try:
                    if owned_pid not in powerpoint_processes():
                        break
                except subprocess.SubprocessError:
                    break
                time.sleep(0.15)


def process(paths, directory, options, merge=False, keywords=None):
    if not paths or (merge and len(paths) < 2) or (not merge and len(paths) != 1):
        raise ValueError('合併至少需兩份簡報。個別清理每份簡報分開另存。')
    keywords = keywords or {}
    if not (merge or options.has_cleanup or keywords):
        raise ValueError('請勾選至少一個清理項目。')
    prepared = [prepare(path, options, keywords.get(path)) for path in paths]
    decks = [item[2] for item in prepared]
    before = [item[3] for item in prepared]
    hidden = sum(item['hidden'] for item in before)
    if options.hidden and hidden == sum(item['slides'] for item in before):
        raise ValueError('全部投影片皆為隱藏，未產生空白簡報。')
    use_powerpoint = merge or bool(options.hidden and hidden)
    if use_powerpoint:
        base, expected_labels = native_plan(decks, [item[0] for item in prepared])
        data = native_transform(decks, bool(options.hidden and hidden), base)
    else:
        data = decks[0]
    remove_annotations(data, options)
    # Clear timing on the final package, including timelines imported by PowerPoint during a merge.
    timelines_removed = remove_animation_timelines(data) if options.animations else 0
    after = inspect(data)
    if (after['slides'] != sum(item['slides'] for item in before) - (hidden if options.hidden else 0)
            or after['hidden'] != (0 if options.hidden else hidden)
            or options.notes and after['notes'] or options.comments and after['comment_parts']):
        raise ValueError('輸出驗證不符，未產生簡報。')
    if use_powerpoint:
        if active_labels(data) != expected_labels:
            raise ValueError('PowerPoint 另存時變更了有效敏感度標籤。為保留原分類，本次未輸出。')
        expected_annotations = [signature for deck in decks for signature in slide_annotations(deck, options, options.hidden)]
        if slide_annotations(data, options) != expected_annotations:
            raise ValueError('PowerPoint 變更了未勾選刪除的備忘稿或註解，未產生輸出檔。')
    for path, source_hash, _, _ in prepared:
        if core.digest(path.read_bytes()) != source_hash:
            raise ValueError(f'處理期間來源發生變更：{path.name}')
    directory = Path(directory)
    directory.mkdir(parents=True, exist_ok=True)
    name = '合併簡報.pptx' if merge else f'{prepared[0][0].stem}.cleaned.pptx'
    output = directory / name
    raw = write_package(data)
    # Exclusive output: existing documents are never overwritten.
    with output.open('xb') as stream:
        try:
            stream.write(raw)
        except BaseException:
            stream.close()
            output.unlink(missing_ok=True)
            raise
    return {'output': str(output), 'source': [str(item[0]) for item in prepared], 'before': before, 'after': after,
            'animation_timelines_removed': timelines_removed,
            'keyword_items': sum(len(keywords[path].ids) for path in paths if path in keywords)}
