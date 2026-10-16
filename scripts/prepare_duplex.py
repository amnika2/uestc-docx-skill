"""Materialize unnumbered blank backs without advancing Roman page numbers."""
import argparse
import copy
import json
from pathlib import Path
from zipfile import ZipFile, ZIP_DEFLATED
from lxml import etree as E

W='http://schemas.openxmlformats.org/wordprocessingml/2006/main'
R='http://schemas.openxmlformats.org/officeDocument/2006/relationships'
PKG='http://schemas.openxmlformats.org/package/2006/relationships'
CT='http://schemas.openxmlformats.org/package/2006/content-types'
NS={'w':W}


def prepare(path, stats):
    path=Path(path)
    with ZipFile(path) as z:
        members={name:z.read(name) for name in z.namelist()}
    doc=E.fromstring(members['word/document.xml'])
    body=doc.find('w:body',NS)
    sections=doc.findall('.//w:sectPr',NS)
    if len(sections)!=len(stats):
        raise ValueError('Section layout metadata does not match document')
    rels=E.fromstring(members['word/_rels/document.xml.rels'])
    types=E.fromstring(members['[Content_Types].xml'])
    used={r.get('Id') for r in rels}
    ids={}
    for kind,tag in [('header','hdr'),('footer','ftr')]:
        name=f'academic-blank-{kind}.xml'
        part=f'word/{name}'
        xml=E.Element('{'+W+'}'+tag,nsmap={'w':W})
        p=E.SubElement(xml,'{'+W+'}p')
        pr=E.SubElement(p,'{'+W+'}pPr')
        spacing=E.SubElement(pr,'{'+W+'}spacing')
        spacing.set('{'+W+'}before','0'); spacing.set('{'+W+'}after','0')
        borders=E.SubElement(pr,'{'+W+'}pBdr')
        bottom=E.SubElement(borders,'{'+W+'}bottom');bottom.set('{'+W+'}val','nil')
        members[part]=E.tostring(xml,xml_declaration=True,encoding='UTF-8',standalone=True)
        rid='rIdAcademicBlank'+kind.title()
        if rid not in used:
            E.SubElement(rels,'{'+PKG+'}Relationship',Id=rid,Type=R+'/'+kind,Target=name)
        ids[kind]=rid
        if not any(e.get('PartName')=='/'+part for e in types):
            E.SubElement(types,'{'+CT+'}Override',PartName='/'+part,ContentType='application/vnd.openxmlformats-officedocument.wordprocessingml.'+kind+'+xml')
    inserted=0
    for i,(section,stat) in enumerate(zip(sections,stats)):
        typ=section.find('w:type',NS)
        if typ is not None and typ.get('{'+W+'}val')=='oddPage':
            typ.set('{'+W+'}val','nextPage')
        role=stat['role']
        if role not in ['cover','front'] or stat['pages']%2==0:
            continue
        parent=section.getparent()
        terminal=parent.getparent()
        if terminal.tag!='{'+W+'}p':
            raise ValueError('Front section does not end in a paragraph')
        blank=E.Element('{'+W+'}p')
        pr=E.SubElement(blank,'{'+W+'}pPr')
        sp=E.SubElement(pr,'{'+W+'}sectPr')
        for kind in ['header','footer']:
            for subtype in ['default','even','first']:
                ref=E.SubElement(sp,'{'+W+'}'+kind+'Reference')
                ref.set('{'+W+'}type',subtype);ref.set('{'+R+'}id',ids[kind])
        kind=E.SubElement(sp,'{'+W+'}type');kind.set('{'+W+'}val','nextPage')
        for tag in ['pgSz','pgMar','cols','docGrid']:
            e=section.find('w:'+tag,NS)
            if e is not None:sp.append(copy.deepcopy(e))
        body.insert(body.index(terminal)+1,blank)
        inserted+=1
    for name,xml in [('word/document.xml',doc),('word/_rels/document.xml.rels',rels),('[Content_Types].xml',types)]:
        members[name]=E.tostring(xml,xml_declaration=True,encoding='UTF-8',standalone=True)
    temporary=path.with_suffix('.duplex.tmp.docx')
    with ZipFile(temporary,'w',ZIP_DEFLATED) as z:
        for name,data in members.items():z.writestr(name,data)
    temporary.replace(path)
    return inserted


def main():
    ap=argparse.ArgumentParser()
    ap.add_argument('docx');ap.add_argument('layout')
    args=ap.parse_args()
    stats=json.loads(Path(args.layout).read_text(encoding='utf-8-sig'))
    print('Unnumbered blank backs:',prepare(args.docx,stats))


if __name__=='__main__':main()
