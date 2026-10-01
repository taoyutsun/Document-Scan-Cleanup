"""Generate minimal, non-company Office files for reproducible engine tests."""
from pathlib import Path
from zipfile import ZipFile,ZIP_DEFLATED
from email.message import EmailMessage
from email import policy
import base64

W='http://schemas.openxmlformats.org/wordprocessingml/2006/main'
P='http://schemas.openxmlformats.org/presentationml/2006/main'
A='http://schemas.openxmlformats.org/drawingml/2006/main'
R='http://schemas.openxmlformats.org/officeDocument/2006/relationships'
REL='http://schemas.openxmlformats.org/package/2006/relationships'
S='http://schemas.openxmlformats.org/spreadsheetml/2006/main'
CT='http://schemas.openxmlformats.org/package/2006/content-types'
PNG=base64.b64decode('iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mP8/x8AAwMCAO+j3ioAAAAASUVORK5CYII=')

def relationships(items):
    return ('<Relationships xmlns="'+REL+'">'+''.join(f'<Relationship Id="{rid}" Type="{R}/{kind}" Target="{dest}"/>' for rid,kind,dest in items)+'</Relationships>')

def content_types(parts):
    return '<Types xmlns="'+CT+'"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Default Extension="png" ContentType="image/png"/>'+''.join(f'<Override PartName="/{n}" ContentType="{typ}"/>' for n,typ in parts.items())+'</Types>'

def write(path,parts):
    with ZipFile(path,'w',ZIP_DEFLATED) as z:
        for n,b in parts.items():z.writestr(n,b if isinstance(b,bytes) else b.encode('utf-8'))

def make(folder):
    folder=Path(folder);folder.mkdir(parents=True,exist_ok=True)
    doc={
      '_rels/.rels':relationships([('r1','officeDocument','word/document.xml')]),
      'word/document.xml':f'<w:document xmlns:w="{W}" xmlns:r="{R}"><w:body><w:p><w:r><w:t>Ordinary content</w:t></w:r></w:p><w:sectPr><w:footerReference w:type="default" r:id="rFooter"/><w:pgSz w:w="11906" w:h="16838"/><w:pgMar w:top="1440" w:bottom="1440" w:left="1440" w:right="1440"/></w:sectPr></w:body></w:document>',
      'word/_rels/document.xml.rels':relationships([('rFooter','footer','footer1.xml'),('rStyle','styles','styles.xml'),('rSetting','settings','settings.xml'),('rNum','numbering','numbering.xml')]),
      'word/footer1.xml':f'<w:ftr xmlns:w="{W}"><w:p><w:pPr><w:pStyle w:val="ConfidentialityLabelSpec"/></w:pPr><w:r><w:t>Confidential</w:t></w:r></w:p><w:p><w:r><w:t>Keep footer</w:t></w:r></w:p></w:ftr>',
      'word/styles.xml':f'<w:styles xmlns:w="{W}"><w:style w:type="paragraph" w:styleId="ConfidentialityLabelSpec"><w:name w:val="Confidentiality Label Spec"/><w:rPr><w:b/></w:rPr></w:style></w:styles>',
      'word/settings.xml':f'<w:settings xmlns:w="{W}"><w:documentProtection w:edit="readOnly" w:enforcement="0"/></w:settings>',
      'word/numbering.xml':f'<w:numbering xmlns:w="{W}"><w:abstractNum w:abstractNumId="0"><w:lvl w:ilvl="0"><w:pStyle w:val="ConfidentialityLabelSpec"/></w:lvl></w:abstractNum></w:numbering>',
      'word/media/image1.png':PNG}
    doc['[Content_Types].xml']=content_types({'word/document.xml':'application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml','word/footer1.xml':'application/vnd.openxmlformats-officedocument.wordprocessingml.footer+xml','word/styles.xml':'application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml','word/settings.xml':'application/vnd.openxmlformats-officedocument.wordprocessingml.settings+xml','word/numbering.xml':'application/vnd.openxmlformats-officedocument.wordprocessingml.numbering+xml'})
    write(folder/'sample-template.docx',doc)
    def tree(text):return f'<p:cSld><p:spTree><p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr><p:grpSpPr/><p:sp><p:nvSpPr><p:cNvPr id="2" name="Highly Confidential"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr><p:spPr/><p:txBody><a:bodyPr/><a:lstStyle/><a:p><a:r><a:t>{text}</a:t></a:r></a:p></p:txBody></p:sp></p:spTree></p:cSld>'
    ppt={'_rels/.rels':relationships([('r1','officeDocument','ppt/presentation.xml')]),'ppt/media/image1.png':PNG}
    masterlist=''.join(f'<p:sldMasterId id="{2147483647+i}" r:id="m{i}"/>' for i in range(1,4))
    slidelist=''.join(f'<p:sldId id="{255+i}" r:id="s{i}"/>' for i in range(1,7))
    ppt['ppt/presentation.xml']=f'<p:presentation xmlns:p="{P}" xmlns:a="{A}" xmlns:r="{R}"><p:sldMasterIdLst>{masterlist}</p:sldMasterIdLst><p:sldIdLst>{slidelist}</p:sldIdLst><p:sldSz cx="12192000" cy="6858000"/></p:presentation>'
    ppt['ppt/_rels/presentation.xml.rels']=relationships([(f'm{i}','slideMaster',f'slideMasters/slideMaster{i}.xml') for i in range(1,4)]+[(f's{i}','slide',f'slides/slide{i}.xml') for i in range(1,7)]+[('t0','theme','theme/theme1.xml')])
    types={'ppt/presentation.xml':'application/vnd.openxmlformats-officedocument.presentationml.presentation.main+xml'}
    for i,text in enumerate(['Highly Confidential','Confidential (Restricted)','Confidential'],1):
        n=f'ppt/slideMasters/slideMaster{i}.xml';ppt[n]=f'<p:sldMaster xmlns:p="{P}" xmlns:a="{A}" xmlns:r="{R}">'+tree(text)+f'<p:sldLayoutIdLst><p:sldLayoutId id="{2147483650+i}" r:id="l1"/></p:sldLayoutIdLst><p:txStyles/></p:sldMaster>'
        types[n]='application/vnd.openxmlformats-officedocument.presentationml.slideMaster+xml'
        ppt[f'ppt/slideMasters/_rels/slideMaster{i}.xml.rels']=relationships([('l1','slideLayout',f'../slideLayouts/slideLayout{i}.xml'),('t1','theme',f'../theme/theme{i}.xml')])
        n=f'ppt/slideLayouts/slideLayout{i}.xml';ppt[n]=f'<p:sldLayout xmlns:p="{P}" xmlns:a="{A}" xmlns:r="{R}">'+tree('Layout')+'</p:sldLayout>'
        types[n]='application/vnd.openxmlformats-officedocument.presentationml.slideLayout+xml'
        ppt[f'ppt/slideLayouts/_rels/slideLayout{i}.xml.rels']=relationships([('m1','slideMaster',f'../slideMasters/slideMaster{i}.xml')])
        n=f'ppt/theme/theme{i}.xml';ppt[n]=f'<a:theme xmlns:a="{A}" name="{text}"><a:themeElements/></a:theme>'
        types[n]='application/vnd.openxmlformats-officedocument.theme+xml'
    for i in range(1,7):
        n=f'ppt/slides/slide{i}.xml';ppt[n]=f'<p:sld xmlns:p="{P}" xmlns:a="{A}" xmlns:r="{R}"'+(' show="0"' if i==6 else '')+'>'+tree('Public slide')+'</p:sld>'
        types[n]='application/vnd.openxmlformats-officedocument.presentationml.slide+xml'
        ppt[f'ppt/slides/_rels/slide{i}.xml.rels']=relationships([('l1','slideLayout','../slideLayouts/slideLayout3.xml')])
    ppt['[Content_Types].xml']=content_types(types);write(folder/'sample-template.pptx',ppt)
    simple=folder/'fixtures';simple.mkdir(exist_ok=True)
    (simple/'text.txt').write_bytes('一般內容\r\nConfidential review\r\n保留此行\r\n'.encode('utf-8-sig'))
    msg=EmailMessage();msg['Subject']='Test Confidential';msg.set_content('Confidential notice');msg.add_attachment((simple/'text.txt').read_bytes(),maintype='text',subtype='plain',filename='attachment.txt')
    (simple/'message.eml').write_bytes(msg.as_bytes(policy=policy.SMTP))
    sheet={'_rels/.rels':relationships([('r1','officeDocument','xl/workbook.xml')]),'xl/workbook.xml':f'<workbook xmlns="{S}" xmlns:r="{R}"><sheets><sheet name="Hidden" sheetId="1" state="hidden" r:id="r1"/></sheets></workbook>',
           'xl/_rels/workbook.xml.rels':relationships([('r1','worksheet','worksheets/sheet1.xml'),('r2','sharedStrings','sharedStrings.xml')]),
           'xl/sharedStrings.xml':f'<sst xmlns="{S}"><si><r><t>Confi</t></r><r><t>dential</t></r></si></sst>',
           'xl/worksheets/sheet1.xml':f'<worksheet xmlns="{S}"><sheetData><row r="1" hidden="1"><c r="B1" t="s"><v>0</v></c></row></sheetData></worksheet>'}
    sheet['[Content_Types].xml']=content_types({'xl/workbook.xml':'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml'})
    write(simple/'hidden.xlsx',sheet)
