"""Structural cleanup for Word comments/revisions and Excel annotations."""
import re

from lxml import etree

import document_cleanup_core as core
import ppt_batch
import ppt_structure as opc

W = '{' + core.W + '}'
S = '{' + core.S + '}'
X = '{urn:schemas-microsoft-com:office:excel}'
VML_SHAPE = '{urn:schemas-microsoft-com:vml}shape'
WORD_COMMENT_TYPES = {'comments', 'commentsextended', 'commentsids', 'commentsextensible'}
COMMENT_ANCHORS = {W + name for name in ('commentReference', 'commentRangeStart', 'commentRangeEnd')}
REVISION = re.compile(r'^(?:ins|del|delText|delInstrText|move(?:From|To)(?:Range(?:Start|End))?|'
                      r'conflict(?:Ins|Del)|customXml(?:Ins|Del|MoveFrom|MoveTo|ConflictIns|ConflictDel)Range(?:Start|End)|'
                      r'(?:pPr|rPr|sectPr|tblPr|tblPrEx|tblGrid|trPr|tcPr|numbering)Change|cell(?:Ins|Del|Merge))$')


def drop_parts(data, seeds):
    if not seeds:
        return
    candidates = opc.reachable(data, seeds)
    for part in list(data):
        if not part.endswith('.rels'):
            continue
        owner = opc.relowner(part)
        if owner in seeds:
            continue
        root = opc.parse(data[part])
        remove = [node for node in root if isinstance(node.tag, str) and opc.internal_target(owner, node) in seeds]
        if not remove:
            continue
        ids = {node.get('Id') for node in remove}
        if owner.endswith('.xml') and any(etree.QName(attr).namespace == opc.R and value in ids
                                          for node in opc.parse(data[owner]).iter() for attr, value in node.attrib.items()):
            raise ValueError('發現未支援的註解關聯，未輸出。')
        for node in remove:
            root.remove(node)
        data[part] = opc.xml(root)
    surviving = {opc.relowner(name) for name in data if name.endswith('.rels')}
    delete = (candidates - opc.reachable(data, {''} | (surviving - candidates))) | seeds
    delete.update(opc.relpath(name) for name in seeds)
    if any(core.protected_part(name) for name in delete if not name.endswith('.rels')):
        raise ValueError('清理會影響其他受保護的零件，未輸出。')
    for name in delete:
        data.pop(name, None)
    types = opc.parse(data['[Content_Types].xml'])
    for node in list(types):
        if node.get('PartName', '').lstrip('/') in delete:
            types.remove(node)
    data['[Content_Types].xml'] = opc.xml(types)


def word_part_kind(content_type):
    return content_type.rsplit('.', 1)[-1].partition('+')[0].lower()


def word_revisions(data):
    for name, payload in data.items():
        if not name.startswith('word/') or not name.endswith('.xml'):
            continue
        for node in opc.parse(payload).iter():
            if isinstance(node.tag, str) and REVISION.fullmatch(core.local(node)) and (
                    node.tag.startswith(W) or node.tag.startswith('{http://schemas.microsoft.com/office/word/')):
                yield node


def word_comment_text(data):
    return sorted((node.get(W+'author', ''), ''.join(node.itertext()))
                  for name, content_type in ppt_batch.part_types(data).items() if word_part_kind(content_type) == 'comments'
                  for node in opc.parse(data[name]).iter(W+'comment'))


def word_info(data):
    track = False
    if 'word/settings.xml' in data:
        flag = opc.parse(data['word/settings.xml']).find(W+'trackRevisions')
        track = flag is not None and flag.get(W+'val', '1').lower() not in {'0', 'false', 'off'}
    return {'comments': len(word_comment_text(data)), 'revision_marks': sum(1 for _ in word_revisions(data)),
            'tracking': track}


def remove_word_comments(data):
    types = ppt_batch.part_types(data)
    delete = {name for name, kind in types.items() if name.startswith('word/') and word_part_kind(kind) in WORD_COMMENT_TYPES}
    revision_authors = {node.get(W+'author') for node in word_revisions(data) if node.get(W+'author')}
    for name, kind in types.items():
        if not name.startswith('word/') or word_part_kind(kind) != 'people':
            continue
        root = opc.parse(data[name])
        for person in list(root):
            if core.local(person)=='person' and not any(etree.QName(key).localname == 'author' and value in revision_authors
                                                       for key, value in person.attrib.items()):
                root.remove(person)
        if not len(root):
            delete.add(name)
        else:
            data[name] = opc.xml(root)
    for name in list(data):
        if name in delete or not name.startswith('word/') or not name.endswith('.xml'):
            continue
        root = opc.parse(data[name])
        anchors = [node for node in root.iter() if node.tag in COMMENT_ANCHORS]
        if anchors:
            for node in anchors:
                node.getparent().remove(node)
            data[name] = opc.xml(root)
    drop_parts(data, delete)


def hidden_sheets(data):
    return [{'name': node.get('name', ''), 'state': node.get('state')}
            for node in opc.parse(data['xl/workbook.xml']).findall(S+'sheets/'+S+'sheet')
            if node.get('state') in {'hidden', 'veryHidden'}]


def sheet_annotations(data):
    types = ppt_batch.part_types(data)
    worksheets = [name for name, kind in types.items()
                  if kind == 'application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml']
    attached = set()
    for sheet in worksheets:
        targets = {}
        for rel in opc.relations(data, sheet):
            kind = rel.get('Type', '').rsplit('/', 1)[-1]
            targets.setdefault(kind, []).append(opc.internal_target(sheet, rel))
        threads = {part: opc.parse(data[part]) for part in targets.get('threadedComment', []) if part}
        thread_cells = {node.get('ref') for root in threads.values() for node in root
                        if core.local(node)=='threadedComment' and not node.get('parentId')}
        comments = {part: opc.parse(data[part]) for part in targets.get('comments', []) if part}
        for part in set(threads) | set(comments):
            if part in attached:
                raise ValueError('發現共用的 Excel 註解零件，無法安全區分工作表。')
            attached.add(part)
        yield threads, thread_cells, comments, targets.get('vmlDrawing', [])
    declared = {part for part, kind in types.items() if kind in {
        'application/vnd.ms-excel.threadedcomments+xml',
        'application/vnd.openxmlformats-officedocument.spreadsheetml.comments+xml',
    }}
    if declared - attached:
        raise ValueError('活頁簿有未連結到工作表的註解零件，未輸出。')


def excel_info(data):
    threaded = notes = 0
    for threads, cells, comments, _ in sheet_annotations(data):
        threaded += sum(sum(core.local(node)=='threadedComment' for node in root) for root in threads.values())
        notes += sum(sum(node.get('ref') not in cells for node in root.iter(S+'comment')) for root in comments.values())
    return {'threaded_comments': threaded, 'notes': notes, 'hidden_sheets': hidden_sheets(data)}


def vml_cell(client):
    row, col = client.findtext(X+'Row'), client.findtext(X+'Column')
    if row is None or col is None or not row.isdecimal() or not col.isdecimal():
        raise ValueError('Excel 附註包含無法識別的圖形位置。')
    row, col = int(row), int(col)+1
    if row >= 1048576 or col > 16384:
        raise ValueError('Excel 附註圖形位置超出範圍。')
    letters = ''
    while col:
        col, rem = divmod(col-1, 26)
        letters = chr(65+rem)+letters
    return letters+str(row+1)


def remove_excel_annotations(data, *, threaded, notes):
    delete = set()
    if threaded:
        delete.update(part for part, kind in ppt_batch.part_types(data).items()
                      if kind == 'application/vnd.ms-excel.person+xml')
    for threads, cells, comments, drawings in sheet_annotations(data):
        if threaded:
            delete.update(threads)
        removed_cells = set(cells) if threaded else set()
        for part, root in comments.items():
            records = list(root.iter(S+'comment'))
            remove = [node for node in records if (threaded if node.get('ref') in cells else notes)]
            if not remove:
                continue
            removed_cells.update(node.get('ref') for node in remove)
            if len(records) == len(remove):
                delete.add(part)
                continue
            for node in remove:
                node.getparent().remove(node)
            authors = root.find(S+'authors')
            kept_ids = sorted({int(node.get('authorId')) for node in root.iter(S+'comment')})
            if authors is None or any(index < 0 or index >= len(authors) for index in kept_ids):
                raise ValueError('Excel 附註作者索引無法驗證。')
            mapping = {old: new for new, old in enumerate(kept_ids)}
            for index, node in enumerate(list(authors)):
                if index not in mapping:
                    authors.remove(node)
            for node in root.iter(S+'comment'):
                node.set('authorId', str(mapping[int(node.get('authorId'))]))
            data[part] = opc.xml(root)
        if not removed_cells:
            continue
        for drawing in drawings:
            if not drawing:
                continue
            root = opc.parse(data[drawing])
            changed = False
            for shape in list(root.iter(VML_SHAPE)):
                client = shape.find(X+'ClientData')
                if client is not None and client.get('ObjectType') == 'Note' and vml_cell(client) in removed_cells:
                    shape.getparent().remove(shape)
                    changed = True
            if changed:
                data[drawing] = opc.xml(root)
    drop_parts(data, delete)
