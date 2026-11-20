"""Identify Word custom XML stores by their package properties, not their file names."""
from collections import Counter
from dataclasses import dataclass
from uuid import UUID

from lxml import etree

import ppt_structure as opc

DOCUMENT = 'word/document.xml'
DATASTORE = '{http://schemas.openxmlformats.org/officeDocument/2006/customXml}'
DATA_REL = opc.R + '/customXml'
PROPERTIES_REL = opc.R + '/customXmlProps'
CONTENT_TYPES = '{http://schemas.openxmlformats.org/package/2006/content-types}'
WORD_BINDINGS = {'{http://schemas.openxmlformats.org/wordprocessingml/2006/main}dataBinding',
                 '{http://schemas.microsoft.com/office/word/2012/wordml}dataBinding'}


def custom_parts(data):
    return {name: content for name, content in data.items() if name.startswith('customXml/')}


def incoming_links(data, parts):
    links = Counter()
    for name in data:
        if not name.endswith('.rels'):
            continue
        owner = opc.relowner(name)
        for relation in opc.relations(data,owner):
            target = opc.internal_target(owner,relation)
            if target in parts:
                links[owner, relation.get('Type'), target] += 1
    return links


@dataclass(frozen=True)
class Store:
    data: bytes
    properties: bytes
    data_type: str
    properties_type: str

    def matches(self, other):
        if (self.data_type,self.properties_type) != (other.data_type,other.properties_type):
            return False
        return all(left==right or etree.tostring(opc.parse(left), method='c14n') ==
                   etree.tostring(opc.parse(right), method='c14n')
                   for left,right in ((self.data,other.data),(self.properties,other.properties)))


def store_id(value):
    try:
        return UUID(value).hex
    except (ValueError, TypeError, AttributeError) as error:
        raise ValueError('無法識別 Word 自訂 XML 的資料識別碼。') from error


def stores(data):
    parts = custom_parts(data)
    if not parts:
        return {}
    content_types = opc.parse(data['[Content_Types].xml'])
    overrides = {node.get('PartName').lstrip('/'): node.get('ContentType')
                 for node in content_types.findall(CONTENT_TYPES+'Override')}
    defaults = {node.get('Extension').lower(): node.get('ContentType')
                for node in content_types.findall(CONTENT_TYPES+'Default')}

    def content_type(name):
        value = overrides.get(name) or defaults.get(name.rsplit('.',1)[-1].lower())
        if not value:
            raise ValueError('Word 自訂 XML 缺少內容類型。')
        return value

    result = {}
    covered = set()
    expected_links = Counter()
    for relation in opc.relations(data, DOCUMENT):
        if relation.get('Type') != DATA_REL:
            continue
        part = opc.internal_target(DOCUMENT, relation)
        if part not in parts:
            raise ValueError('Word 自訂 XML 的主文件關聯無法驗證。')
        subrels = opc.relations(data, part)
        if len(subrels)!=1 or subrels[0].get('Type')!=PROPERTIES_REL:
            raise ValueError('Word 自訂 XML 含未支援的附屬資料。')
        properties = opc.internal_target(part, subrels[0])
        if properties not in parts or {part, properties} & covered:
            raise ValueError('Word 自訂 XML 的資料屬性關聯不完整或重複。')
        root = opc.parse(parts[properties])
        if root.tag!=DATASTORE+'datastoreItem':
            raise ValueError('Word 自訂 XML 缺少標準資料屬性。')
        key = store_id(root.get(DATASTORE+'itemID'))
        if key in result:
            raise ValueError('Word 自訂 XML 有重複的資料識別碼。')
        result[key] = Store(parts[part], parts[properties], content_type(part), content_type(properties))
        covered.update((part, properties, opc.relpath(part)))
        expected_links.update(((DOCUMENT, DATA_REL, part),(part, PROPERTIES_REL, properties)))
    if covered != parts.keys() or incoming_links(data,parts) != expected_links:
        raise ValueError('Word 自訂 XML 含無法驗證的附件或關聯。')
    return result


def verify_same(expected, actual):
    if expected.keys() != actual.keys() or any(not store.matches(actual[key]) for key,store in expected.items()):
        raise ValueError('Word 自訂 XML 的資料內容、屬性或關聯無法確認完整保留，未輸出。')


def verify_rewrite(before, after):
    original = custom_parts(before)
    if original==custom_parts(after) and (not original or incoming_links(before,original)==incoming_links(after,original)):
        return
    verify_same(stores(before),stores(after))


def merge_plan(decks):
    merged = {}
    for deck in decks:
        for key, store in stores(deck).items():
            if key in merged and not merged[key].matches(store):
                raise ValueError('Word 合併來源共用同一個自訂 XML 識別碼，但資料內容不同，未輸出。')
            merged.setdefault(key, store)
    return merged


def bindings(data, keys):
    result = Counter()
    if not keys:
        return result
    for name, payload in data.items():
        if not name.startswith('word/') or not name.endswith('.xml'):
            continue
        for node in opc.parse(payload).iter():
            if node.tag not in WORD_BINDINGS:
                continue
            attributes = dict(node.attrib)
            for name in attributes:
                if etree.QName(name).localname == 'storeItemID':
                    key = store_id(attributes[name])
                    if key in keys:
                        attributes[name] = key
                        result[(node.tag,tuple(sorted(attributes.items())))] += 1
                    break
    return result


def preserve_after_merge(data, expected, sources):
    actual = stores(data)
    if actual.keys()-expected.keys() or any(not expected[key].matches(store) for key,store in actual.items()):
        raise ValueError('Word 合併時變更了原有自訂 XML，未輸出。')
    required_bindings = sum((bindings(deck,expected) for deck in sources), Counter())
    if bindings(data,expected) != required_bindings:
        raise ValueError('Word 合併時變更了繫結自訂 XML 的內容控制項，未輸出。')
    missing = [store for key,store in expected.items() if key not in actual]
    if missing:
        attach_stores(data, missing)
    verify_same(expected,stores(data))


def attach_stores(data, missing):
    main_rels = opc.relpath(DOCUMENT)
    relationships = opc.parse(data[main_rels]) if main_rels in data else etree.Element('{'+opc.REL+'}Relationships')
    used = {node.get('Id') for node in relationships}
    types = opc.parse(data['[Content_Types].xml'])
    index = 1
    for store in missing:
        while True:
            part = f'customXml/item{index}.xml'
            props = f'customXml/itemProps{index}.xml'
            rels = opc.relpath(part)
            rid = f'rIdCustomXml{index}'
            if rid not in used and all(name not in data for name in (part,props,rels)):
                break
            index += 1
        data[part] = store.data
        data[props] = store.properties
        subrels = etree.Element('{'+opc.REL+'}Relationships')
        etree.SubElement(subrels,'{'+opc.REL+'}Relationship', Id='rId1', Type=PROPERTIES_REL,
                         Target=props.rsplit('/',1)[-1])
        data[rels] = opc.xml(subrels)
        etree.SubElement(relationships,'{'+opc.REL+'}Relationship',Id=rid,Type=DATA_REL,Target='../'+part)
        for name, kind in ((part,store.data_type),(props,store.properties_type)):
            etree.SubElement(types,CONTENT_TYPES+'Override',PartName='/'+name,ContentType=kind)
        index += 1
    data[main_rels] = opc.xml(relationships)
    data['[Content_Types].xml'] = opc.xml(types)
