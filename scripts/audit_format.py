"""Check observable academic DOCX formatting, without modifying the file."""
import argparse
import json
import re
from pathlib import Path
from zipfile import ZipFile
from lxml import etree as E

NS = {'w': 'http://schemas.openxmlformats.org/wordprocessingml/2006/main',
      'm': 'http://schemas.openxmlformats.org/officeDocument/2006/math'}
W = '{'+NS['w']+'}'


def audit(path, profile_path=None):
    profile = json.loads(Path(profile_path or Path(__file__).with_name('academic-profile.json')).read_text(encoding='utf-8'))
    failures = []
    with ZipFile(path) as z:
        for name in z.namelist():
            if name.endswith(('.xml','.rels')):
                E.fromstring(z.read(name))
        sx = E.fromstring(z.read('word/styles.xml'))
        dx = E.fromstring(z.read('word/document.xml'))
        nx = E.fromstring(z.read('word/numbering.xml'))
        styles = {s.get(W+'styleId'):s for s in sx.findall('w:style',NS)}
        canonical=['Normal','Heading1','Heading2','Heading3','Heading4','FigureCaption',
                   'TableCaption','Formula','FormulaNumber','Reference','Code','Footer','Header','TOC1','TOC2','TableText']
        names={re.sub(r'\s+','',c).lower():c for c in canonical}
        names.update({'正文':'Normal','标题1':'Heading1','标题2':'Heading2','标题3':'Heading3',
                      '标题4':'Heading4','页眉':'Header','页脚':'Footer'})
        aliases={}
        for sid,style in list(styles.items()):
            name=style.find('w:name',NS)
            key=re.sub(r'\s+','',name.get(W+'val','')).lower() if name is not None else sid.lower()
            if key in names:
                aliases[sid]=names[key]
                styles[names[key]]=style
        defaults={g:sx.find(f'w:docDefaults/w:{g}Default/w:{g}',NS) for g in ['pPr','rPr']}

        def check(label, actual, expected):
            if str(actual) != str(expected):
                failures.append(f'{label}: got {actual!r}, expected {expected!r}')

        def style_value(style, element, attr='val', group='pPr'):
            seen=set(); current=styles[style]
            while current is not None and id(current) not in seen:
                seen.add(id(current))
                node=current.find(f'w:{group}/w:{element}',NS)
                if node is not None:
                    value=node.get(W+attr)
                    if value is None and attr=='val' and element in ['keepNext','keepLines','pageBreakBefore']:
                        return 'true'
                    if value is not None:
                        return {'宋体':'SimSun','黑体':'SimHei'}.get(value,value)
                parent=current.find('w:basedOn',NS)
                current=styles.get(parent.get(W+'val')) if parent is not None else None
            default=defaults[group]
            node=None if default is None else default.find('w:'+element,NS)
            value=None if node is None else node.get(W+attr)
            return {'宋体':'SimSun','黑体':'SimHei'}.get(value,value)

        for sid in ['Normal','Heading1','Heading2','Heading3','Heading4','FigureCaption',
                    'TableCaption','Formula','FormulaNumber','Reference','Code','Footer','Header', 'TOC1','TOC2']:
            if sid not in styles:
                failures.append('Missing style '+sid)
        if failures:
            return failures
        for sid in ['Normal','Heading1','Heading2','Heading3','Heading4','Reference','TOC1','TOC2','TableText']:
            check(sid+' fixed line',style_value(sid,'spacing','line'),400)
            check(sid+' line rule',style_value(sid,'spacing','lineRule'),'exact')
        check('Normal size',style_value('Normal','sz',group='rPr'),24)
        check('Normal Chinese',style_value('Normal','rFonts','eastAsia','rPr'),'SimSun')
        check('Normal Latin',style_value('Normal','rFonts','ascii','rPr'),'Times New Roman')
        check('Normal first line',style_value('Normal','ind','firstLine'),480)
        check('Normal alignment',style_value('Normal','jc'),'both')
        for i,h in enumerate(profile['headings'],1):
            sid='Heading'+str(i)
            check(sid+' size',style_value(sid,'sz',group='rPr'),int(h['size']*2))
            check(sid+' before',style_value(sid,'spacing','before'),int(h['before']*20))
            check(sid+' after',style_value(sid,'spacing','after'),int(h['after']*20))
            check(sid+' Chinese',style_value(sid,'rFonts','eastAsia','rPr'),'SimHei')
            check(sid+' keepNext',style_value(sid,'keepNext'),'true')
        if style_value('Heading1','pageBreakBefore') in ['true','1','on']:
            failures.append('Heading1 adds pageBreakBefore on top of chapter section breaks')
        for sid,key in [('FigureCaption','figureCaption'),('TableCaption','tableCaption')]:
            spec=profile[key]
            check(sid+' size',style_value(sid,'sz',group='rPr'),21)
            check(sid+' before',style_value(sid,'spacing','before'),spec['before']*20)
            check(sid+' after',style_value(sid,'spacing','after'),spec['after']*20)
        check('Formula single line',style_value('Formula','spacing','line'),240)
        check('Formula automatic line',style_value('Formula','spacing','lineRule'),'auto')
        check('Formula number single line',style_value('FormulaNumber','spacing','line'),240)
        check('Formula number automatic line',style_value('FormulaNumber','spacing','lineRule'),'auto')
        check('Formula number Latin font',style_value('FormulaNumber','rFonts','ascii','rPr'),profile['formula']['numberFont'])
        check('Formula number size',style_value('FormulaNumber','sz',group='rPr'),int(profile['formula']['size']*2))
        check('Code Latin font',style_value('Code','rFonts','ascii','rPr'),'Consolas')
        check('Code size',style_value('Code','sz',group='rPr'),21)
        check('Footer size',style_value('Footer','sz',group='rPr'),18)
        for sp in dx.findall('.//w:sectPr',NS):
            m=sp.find('w:pgMar',NS)
            for side in ['top','bottom','left','right']:
                check('Page margin '+side,m.get(W+side),profile['page']['margin'])
            for side in ['header','footer']:
                check('Page distance '+side,m.get(W+side),profile['page'][side])
        for p in dx.findall('.//w:p',NS):
            sid=p.find('w:pPr/w:pStyle',NS)
            sid='' if sid is None else aliases.get(sid.get(W+'val'),sid.get(W+'val'))
            if sid in ['Heading2','Heading3','Heading4'] and p.find('w:pPr/w:pageBreakBefore',NS) is not None:
                failures.append('Subheading has a forced page break')
            if p.find('w:pPr/w:spacing',NS) is not None and sid in ['', 'Normal'] and p.find('w:pPr/w:sectPr',NS) is None:
                sp=p.find('w:pPr/w:spacing',NS)
                # Cover and unindented labels have deliberate, non-body spacing.
                ind=p.find('w:pPr/w:ind',NS)
                if ind is None or ind.get(W+'firstLine','480')!='0':
                    if sp.get(W+'line') and sp.get(W+'line')!='400':
                        failures.append('Body paragraph overrides fixed 20pt line')
            if sid=='Heading1':
                explicit=p.find('w:pPr/w:pageBreakBefore',NS)
                if explicit is not None and explicit.get(W+'val','true') not in ['0','false','off']:
                    failures.append('Chapter heading adds a duplicate page break')
        fields=' '.join(dx.xpath('//w:instrText/text() | //w:fldSimple/@w:instr',namespaces=NS))
        if dx.findall('.//m:oMath',NS) and 'equation' not in fields:
            failures.append('Equations found without automatic equation fields')
        for field in dx.xpath('//w:fldSimple[contains(@w:instr,"SEQ equation")]',namespaces=NS):
            paragraph=field.xpath('ancestor::w:p[1]',namespaces=NS)[0]
            node=paragraph.find('w:pPr/w:pStyle',NS)
            sid='' if node is None else aliases.get(node.get(W+'val'),node.get(W+'val'))
            if sid!='FormulaNumber':
                failures.append('Equation number does not use FormulaNumber style')
        header_fields=[]
        for name in z.namelist():
            if re.fullmatch(r'word/header\d+\.xml',name):
                hx=E.fromstring(z.read(name))
                header_fields.extend(hx.xpath('//w:instrText/text() | //w:fldSimple/@w:instr',namespaces=NS))
        header_instructions=' '.join(header_fields)
        if profile.get('layout',{}).get('bodyHeader')=='chapter':
            if not re.search(r'STYLEREF\s+1\s+\\n\b',header_instructions):
                failures.append('Body header lacks the full chapter-number STYLEREF field')
            if not re.search(r'STYLEREF\s+1(?:\s|$)',header_instructions):
                failures.append('Body header lacks the chapter-title STYLEREF field')
        if any(p.find('w:pPr/w:numPr',NS) is not None for p in dx.findall('.//w:p',NS)):
            levels=nx.findall('w:abstractNum/w:lvl',NS)
            if not any(l.find('w:isLgl',NS) is not None for l in levels):
                failures.append('No legal-numbering style for Arabic subheadings')
        fmt=[e.get(W+'fmt') for e in dx.findall('.//w:sectPr/w:pgNumType',NS)]
        if dx.xpath('//w:pStyle[@w:val="FrontHeading"]',namespaces=NS) and 'upperRoman' not in fmt:
            failures.append('Multi-section academic document lacks front Roman numbering')
    return failures


def main():
    ap=argparse.ArgumentParser()
    ap.add_argument('docx'); ap.add_argument('--profile')
    args=ap.parse_args()
    problems=audit(args.docx,args.profile)
    if problems:
        print('\n'.join(problems)); raise SystemExit(1)
    print('Academic formatting audit PASSED')


if __name__=='__main__':
    main()
