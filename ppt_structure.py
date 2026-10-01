"""Offline PPTX template housekeeping. Originals and active markings are preserved."""
from __future__ import annotations

import copy
import hashlib
import json
import posixpath
import re
from collections import Counter
from datetime import datetime
from io import BytesIO
from pathlib import Path
from urllib.parse import unquote
from zipfile import ZipFile, BadZipFile

from lxml import etree as X

from app_metadata import VERSION
P = "http://schemas.openxmlformats.org/presentationml/2006/main"
R = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
REL = "http://schemas.openxmlformats.org/package/2006/relationships"
EP = "http://schemas.openxmlformats.org/officeDocument/2006/extended-properties"
VT = "http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes"
NS = {"p": P, "r": R, "ep": EP, "vt": VT}
PRUNABLE = ("ppt/slideMasters/", "ppt/slideLayouts/", "ppt/theme/", "ppt/media/")


class CleanupError(Exception):
    pass


def parse(data):
    if b"<!DOCTYPE" in data.upper():
        raise CleanupError("不支援含有 DTD 的 XML。")
    return X.fromstring(data, X.XMLParser(resolve_entities=False, no_network=True))


def xml(root):
    return X.tostring(root, encoding="UTF-8", xml_declaration=True, standalone=True)


def sha(data):
    return hashlib.sha256(data).hexdigest()


def relpath(part):
    if not part:
        return "_rels/.rels"
    folder, name = posixpath.split(part)
    return posixpath.join(folder, "_rels", name + ".rels")


def relowner(name):
    if name == "_rels/.rels":
        return ""
    folder, leaf = posixpath.split(name)
    return posixpath.join(posixpath.dirname(folder), leaf[:-5])


def target(part, value):
    value = unquote(value.split("#", 1)[0])
    resolved = posixpath.normpath(value.lstrip("/") if value.startswith("/")
                                  else posixpath.join(posixpath.dirname(part), value))
    if resolved.startswith("../") or "\\" in resolved:
        raise CleanupError("檔案關聯包含不支援的路徑。")
    return resolved


def relations(data, part):
    name = relpath(part)
    return [e for e in parse(data[name]) if isinstance(e.tag, str)] if name in data else []


def internal_target(part, relation):
    if relation.get("TargetMode") == "External":
        return None
    return target(part, relation.get("Target", ""))


def typed_targets(data, part, kind):
    return [internal_target(part, e) for e in relations(data, part)
            if e.get("Type", "").endswith("/" + kind) and internal_target(part, e)]


def reachable(data, roots):
    seen, pending = set(), list(roots)
    while pending:
        part = pending.pop()
        if part in seen:
            continue
        if part and part not in data:
            raise CleanupError("遺失關聯零件：" + part)
        seen.add(part)
        rp = relpath(part)
        if rp in data:
            seen.add(rp)
        pending.extend(t for e in relations(data, part)
                       if (t := internal_target(part, e)) and t not in seen)
    return seen


def validate(data):
    for name, b in data.items():
        if name.endswith((".xml", ".rels")):
            root = parse(b)
            if name.endswith(".rels"):
                root = [e for e in root if isinstance(e.tag, str)]
                owner = relowner(name)
                if owner and owner not in data:
                    raise CleanupError("孤立的關聯檔：" + name)
                ids = [e.get("Id") for e in root]
                if len(ids) != len(set(ids)):
                    raise CleanupError("重複的關聯 ID：" + name)
                for e in root:
                    t = internal_target(owner, e)
                    if t and t not in data:
                        raise CleanupError("遺失關聯零件：" + t)
    ct = parse(data["[Content_Types].xml"])
    for e in ct:
        if not isinstance(e.tag, str):
            continue
        if X.QName(e).localname == "Override" and e.get("PartName", "").lstrip("/") not in data:
            raise CleanupError("內容類型指向不存在的零件。")
    for part, listname, kind in [("ppt/presentation.xml", "sldMasterIdLst", "slideMaster")]:
        listed = parse(data[part]).find("p:" + listname, NS)
        ids = {e.get("Id") for e in relations(data, part) if e.get("Type", "").endswith("/" + kind)}
        if listed is None or not len(listed) or {e.get("{" + R + "}id") for e in listed} != ids:
            raise CleanupError("母片清單與檔案關聯不一致。")
    for part in data:
        if re.fullmatch(r"ppt/slideMasters/[^/]+\.xml", part):
            listed = parse(data[part]).find("p:sldLayoutIdLst", NS)
            ids = {e.get("Id") for e in relations(data, part) if e.get("Type", "").endswith("/slideLayout")}
            if listed is None or not len(listed) or {e.get("{" + R + "}id") for e in listed} != ids:
                raise CleanupError("版面配置清單與檔案關聯不一致。")


def counts(data):
    return {key: sum(bool(re.fullmatch(folder + r"/[^/]+\.xml", n)) for n in data)
            for key, folder in [("slides", "ppt/slides"), ("masters", "ppt/slideMasters"),
                                ("layouts", "ppt/slideLayouts"), ("themes", "ppt/theme")]}


def trim_ids(data, part, container, kind, kept):
    rp = relpath(part)
    rr = parse(data[rp])
    removed_ids = set()
    for e in list(rr):
        if e.get("Type", "").endswith("/" + kind) and internal_target(part, e) not in kept:
            removed_ids.add(e.get("Id"))
            rr.remove(e)
    if not removed_ids:
        return
    root = parse(data[part])
    ids = root.find("p:" + container, NS)
    if ids is None:
        raise CleanupError("缺少範本清單：" + part)
    for e in list(ids):
        if e.get("{" + R + "}id") in removed_ids:
            ids.remove(e)
    data[part], data[rp] = xml(root), xml(rr)


def update_properties(data, original, deleted, renamed):
    name = "docProps/app.xml"
    if name not in data:
        return []
    root = parse(data[name])
    notes = []
    old_themes = Counter(parse(b).get("name", "") for n, b in original.items()
                         if re.fullmatch(r"ppt/theme/[^/]+\.xml", n))
    kept_themes = Counter(parse(b).get("name", "") for n, b in data.items()
                          if re.fullmatch(r"ppt/theme/[^/]+\.xml", n))
    for item in renamed:
        kept_themes[item["new_name"]] -= 1
        kept_themes[item["old_name"]] += 1
    vanished = old_themes - kept_themes
    hp = root.find("ep:HeadingPairs/vt:vector", NS)
    tp = root.find("ep:TitlesOfParts/vt:vector", NS)
    if (vanished or renamed) and hp is not None and tp is not None:
        pairs = [e for e in hp if isinstance(e.tag, str)]
        if len(pairs) % 2:
            raise CleanupError("無法安全解析文件屬性的分類清單。")
        cursor, changed = 0, False
        for i in range(0, len(pairs), 2):
            label = "".join(pairs[i].itertext()).strip()
            num = pairs[i + 1].find("vt:i4", NS)
            if num is None:
                raise CleanupError("無法安全解析文件屬性的數量。")
            size = int(num.text)
            segment = [e for e in tp if isinstance(e.tag, str)][cursor:cursor + size]
            if len(segment) != size:
                raise CleanupError("文件屬性的分類數量不一致。")
            # Only adjust the known theme category, never arbitrary matching text.
            if label.casefold() in {"themes", "theme", "佈景主題", "布景主題", "主题", "主題", "テーマ"}:
                remaining = vanished.copy()
                replacements = {item["old_name"]: item["new_name"] for item in renamed}
                removed = 0
                for e in segment:
                    value = e.text or ""
                    if remaining[value] > 0:
                        tp.remove(e)
                        remaining[value] -= 1
                        removed += 1
                    elif value in replacements:
                        e.text = replacements[value]
                        changed = True
                num.text = str(size - removed)
                cursor += size - removed
                changed |= bool(removed)
            else:
                cursor += size
        if cursor != len([e for e in tp if isinstance(e.tag, str)]):
            raise CleanupError("文件屬性的分類數量不一致。")
        tp.set("size", str(len([e for e in tp if isinstance(e.tag, str)])))
        if changed:
            notes.append("同步 TitlesOfParts 的佈景主題顯示名稱與數量；其餘分類名稱保留。")
    elif vanished or renamed:
        notes.append("未找到標準佈景主題屬性分類，保留原屬性並由剩餘字樣檢查回報。")
    if notes:
        data[name] = xml(root)
    return notes


def analyze(path):
    path = Path(path).resolve()
    if path.suffix.lower() != ".pptx":
        raise CleanupError("目前僅支援 .pptx；.ppt、.pptm、加密檔案請先由 PowerPoint 另存適當副本。")
    if path.stat().st_size > 600 * 1024 * 1024:
        raise CleanupError("檔案超過工具的處理大小上限（600 MB）。")
    raw = path.read_bytes()
    try:
        with ZipFile(BytesIO(raw)) as z:
            infos = z.infolist()
            if len(infos) > 25000 or sum(e.file_size for e in infos) > 600 * 1024 * 1024:
                raise CleanupError("檔案超過工具的處理大小上限（解壓縮後 600 MB）。")
            if len({e.filename for e in infos}) != len(infos):
                raise CleanupError("ZIP 中有重複零件名稱。")
            for e in infos:
                if e.flag_bits & 1 or "\\" in e.filename or e.filename.startswith("/") or ".." in e.filename.split("/"):
                    raise CleanupError("不支援加密或含有不正常路徑的檔案。")
            data = {e.filename: z.read(e) for e in infos if not e.is_dir()}
    except BadZipFile as exc:
        raise CleanupError("這不是可讀取的 PPTX；可能已加密或損毀。") from exc
    if any(n.startswith("_xmlsignatures/") for n in data):
        raise CleanupError("檔案含有數位簽章，修改會使簽章失效，因此未處理。")
    required = {"ppt/presentation.xml", "[Content_Types].xml", "_rels/.rels"}
    if not required.issubset(data) or X.QName(parse(data["ppt/presentation.xml"])).namespace != P:
        raise CleanupError("不支援此 PPTX 結構；Strict Open XML 請先另存標準 PowerPoint 簡報。")
    validate(data)
    original = data.copy()
    pres = parse(data["ppt/presentation.xml"])
    slide_rels = {e.get("Id"): internal_target("ppt/presentation.xml", e)
                  for e in relations(data, "ppt/presentation.xml") if e.get("Type", "").endswith("/slide")}
    ids = pres.find("p:sldIdLst", NS)
    slides = [slide_rels.get(e.get("{" + R + "}id")) for e in ids] if ids is not None else []
    if not slides or any(s is None for s in slides):
        raise CleanupError("找不到完整的投影片清單。")
    layouts, masters = set(), set()
    for s in slides:
        found = typed_targets(data, s, "slideLayout")
        if len(found) != 1:
            raise CleanupError("投影片的版面配置關聯不唯一：" + s)
        layouts.update(found)
    for layout in layouts:
        found = typed_targets(data, layout, "slideMaster")
        if len(found) != 1:
            raise CleanupError("版面配置的母片關聯不唯一：" + layout)
        masters.update(found)
    all_master_layout = {n for n in data if re.fullmatch(r"ppt/(slideMasters|slideLayouts)/[^/]+\.xml", n)}
    unused = all_master_layout - layouts - masters
    candidates = reachable(data, unused)
    trim_ids(data, "ppt/presentation.xml", "sldMasterIdLst", "slideMaster", masters)
    for master in masters:
        trim_ids(data, master, "sldLayoutIdLst", "slideLayout", layouts)
    live = reachable(data, [""])
    deleted = sorted(n for n in candidates - live if n.startswith(PRUNABLE))
    for n in deleted:
        data.pop(n, None)
    ct = parse(data["[Content_Types].xml"])
    for e in list(ct):
        if e.get("PartName", "").lstrip("/") in deleted:
            ct.remove(e)
    if deleted:
        data["[Content_Types].xml"] = xml(ct)
    renamed = []  # Theme name edits are handled only by explicit scan selections.
    property_changes = update_properties(data, original, deleted, renamed)
    validate(data)
    modified = sorted(n for n in data if data[n] != original[n])
    allowed = {"ppt/presentation.xml", "ppt/_rels/presentation.xml.rels", "docProps/app.xml", "[Content_Types].xml"}
    allowed.update(masters)
    allowed.update(relpath(m) for m in masters)
    allowed.update(item["part"] for item in renamed)
    if set(modified) - allowed:
        raise CleanupError("有非預期的內容變更，已停止。")
    # Verify every active formatting dependency, including hidden slides, is preserved.
    for n in data:
        if n not in allowed and data[n] != original[n]:
            raise CleanupError("內容保留驗證失敗：" + n)
    for m in masters:
        a, b = parse(original[m]), parse(data[m])
        for root in [a, b]:
            group = root.find("p:sldLayoutIdLst", NS)
            if group is not None:
                root.remove(group)
        if X.tostring(a, method="c14n") != X.tostring(b, method="c14n"):
            raise CleanupError("使用中母片的樣式發生變更，已停止。")
    for item in renamed:
        a, b = parse(original[item["part"]]), parse(data[item["part"]])
        a.set("name", item["new_name"])
        if X.tostring(a, method="c14n") != X.tostring(b, method="c14n"):
            raise CleanupError("佈景主題除名稱外發生變更，已停止。")
    special = [n for n in data if n.startswith(("ppt/embeddings/", "ppt/activeX/"))]
    report = {"tool_version": VERSION, "source": str(path), "source_sha256": sha(raw),
              "time": datetime.now().isoformat(timespec="seconds"),
              "before": counts(original), "after": counts(data),
              "deleted_parts": deleted, "modified_parts": modified,
              "property_changes": property_changes,
              "renamed_themes": renamed,
              "retained_masters": sorted(masters), "retained_layouts": sorted(layouts),
              "embedded_objects_not_scanned": special,
              "verification": {"all_relationships_valid": True, "slide_parts_identical": True,
                               "active_master_style_identical": True,
                               "theme_style_identical": True,
                               "all_other_retained_parts_identical": True},
              "limitations": ["僅檢查 XML 文字與屬性；不做圖片 OCR 或內嵌物件內容掃描。",
                              "不移除使用中內容的機密字樣或正式敏感度標籤。",
                              "範本整理保留使用中的圖形與背景，不代表文件已完成所有內容檢查。",
                              "此報告僅描述範本結構變更，分享前請檢查內容與用途。"]}
    return data, infos, report

