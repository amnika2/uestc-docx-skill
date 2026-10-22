'use strict';
const fs = require('fs');
const path = require('path');
const {loadDependency} = require('./runtime');
const d = loadDependency('docx');
const profile = require('./academic-profile.json');
const {Document, Paragraph, TextRun, Table, TableRow, TableCell, Header, Footer,
  ImageRun, Math:WordMath, MathRun, SimpleField, AlignmentType:A, LineRuleType:L,
  HeadingLevel:H, LevelFormat, WidthType, BorderStyle, VerticalAlign,
  TableOfContents, PageNumber, NumberFormat, SectionType, LevelSuffix} = d;
const W = profile.page.width - 2*profile.page.margin;
const FONT = {ascii:profile.body.latin,hAnsi:profile.body.latin,eastAsia:profile.body.font};
const HEAD = {...FONT,eastAsia:'SimHei'};
const FORMULA_NUMBER_FONT = {ascii:profile.formula.numberFont,hAnsi:profile.formula.numberFont,eastAsia:profile.formula.numberFont};
const NONE = {style:BorderStyle.NONE,size:0,color:'auto'};
const BORDERS_NONE = {top:NONE,bottom:NONE,left:NONE,right:NONE,insideHorizontal:NONE,insideVertical:NONE};
const fixed = (points,before=0,after=0)=>({line:points*20,lineRule:L.EXACT,before:before*20,after:after*20});
const single = (before=0,after=0)=>({line:240,lineRule:L.AUTO,before:before*20,after:after*20});
const resetIndent = {firstLine:0,left:0,right:0};

function styles() {
  const normal = {id:'Normal',name:'Normal',run:{font:FONT,size:24,bold:false,color:'000000',snapToGrid:false},
    paragraph:{alignment:A.JUSTIFIED,indent:{firstLine:480},spacing:fixed(20),widowControl:true,keepNext:false}};
  const headings = profile.headings.map((h,i)=>({id:'Heading'+(i+1),name:'Heading '+(i+1),basedOn:'Normal',next:'Normal',quickFormat:true,
    run:{font:HEAD,size:h.size*2,bold:false},paragraph:{alignment:h.alignment==='center'?A.CENTER:A.LEFT,
      indent:resetIndent,spacing:fixed(20,h.before,h.after),outlineLevel:i,keepNext:true,keepLines:true}}));
  const capStyle=(id,p,keepNext)=>({id,name:id,basedOn:'Normal',run:{font:FONT,size:p.size*2,bold:false},
    paragraph:{alignment:A.CENTER,indent:resetIndent,spacing:fixed(20,p.before,p.after),keepNext,keepLines:true}});
  const pp=[normal,...headings,
    {id:'Title',name:'Title',basedOn:'Normal',run:{font:FONT,size:32,bold:true},paragraph:{alignment:A.CENTER,indent:resetIndent,spacing:{line:360,lineRule:L.AUTO,before:0,after:0},keepNext:true}},
    {id:'FrontHeading',name:'FrontHeading',basedOn:'Normal',run:{font:HEAD,size:30,bold:false},paragraph:{alignment:A.CENTER,indent:resetIndent,spacing:fixed(20,24,18),keepNext:true,keepLines:true}},
    capStyle('FigureCaption',profile.figureCaption,false),capStyle('TableCaption',profile.tableCaption,true),
    {id:'TableText',name:'TableText',basedOn:'Normal',run:{font:FONT,size:21},paragraph:{alignment:A.CENTER,indent:resetIndent,spacing:fixed(20),keepLines:true}},
    {id:'Formula',name:'Formula',basedOn:'Normal',run:{font:{...FONT,ascii:'Cambria Math',hAnsi:'Cambria Math'},size:24},paragraph:{alignment:A.CENTER,indent:resetIndent,spacing:single(6,6),keepLines:true}},
    {id:'FormulaNumber',name:'FormulaNumber',basedOn:'Normal',run:{font:FORMULA_NUMBER_FONT,size:24},paragraph:{alignment:A.RIGHT,indent:resetIndent,spacing:single(6,6),keepLines:true}},
    {id:'Reference',name:'Reference',basedOn:'Normal',run:{font:FONT,size:21},paragraph:{alignment:A.JUSTIFIED,indent:{left:360,hanging:360,firstLine:0},spacing:fixed(20),keepLines:true}},
    {id:'Header',name:'Header',basedOn:'Normal',run:{font:FONT,size:21},paragraph:{alignment:A.CENTER,indent:resetIndent,spacing:fixed(20),border:{bottom:{style:BorderStyle.SINGLE,size:4,color:'000000'}}}},
    {id:'Footer',name:'Footer',basedOn:'Normal',run:{font:FONT,size:18},paragraph:{alignment:A.CENTER,indent:resetIndent,spacing:single()}},
    {id:'FootnoteText',name:'Footnote Text',basedOn:'Normal',run:{font:FONT,size:18},paragraph:{alignment:A.JUSTIFIED,indent:{left:270,hanging:270},spacing:single()}},
    {id:'Code',name:'Code',basedOn:'Normal',run:{font:{ascii:'Consolas',hAnsi:'Consolas',eastAsia:'SimSun'},size:21,snapToGrid:false},paragraph:{alignment:A.LEFT,indent:resetIndent,spacing:fixed(14),keepLines:true,keepNext:false}}
  ];
  for(let i=1;i<=4;i++) pp.push({id:'TOC'+i,name:'toc '+i,basedOn:'Normal',run:{font:i===1?HEAD:FONT,size:24,bold:false},paragraph:{alignment:A.LEFT,indent:{firstLine:0,left:(i-1)*240},spacing:fixed(20),keepLines:true}});
  return {default:{document:{run:normal.run,paragraph:normal.paragraph}},paragraphStyles:pp};
}

function numbering() {
  return {config:[{reference:'chapters',levels:[0,1,2,3].map(i=>({level:i,
    format:i===0?LevelFormat.CHINESE_COUNTING:LevelFormat.DECIMAL,
    text:i===0?'第%1章':Array.from({length:i+1},(_,j)=>'%'+(j+1)).join('.'),
    isLegalNumberingStyle:i>0,suffix:LevelSuffix.SPACE,alignment:A.LEFT,
    style:{paragraph:{indent:{left:0,hanging:0}},run:{font:HEAD}}}))},
    {reference:'references',levels:[{level:0,format:LevelFormat.DECIMAL,text:'[%1]',suffix:LevelSuffix.SPACE,
      alignment:A.LEFT,style:{paragraph:{indent:{left:360,hanging:360}},run:{font:FONT,size:21}}}]},
    {reference:'bullets',levels:[{level:0,format:LevelFormat.BULLET,text:'•',alignment:A.LEFT,style:{paragraph:{indent:{left:480,hanging:240}}}}]}]};
}

function inline(text) {
  const result=[];
  for(const part of String(text).split(/(\$[^$]+\$|\[\d+(?:[,，\-–]\d+)*\])/).filter(Boolean)){
    if(part.startsWith('$') && part.endsWith('$')) result.push(mathFromLatex(part.slice(1,-1)));
    else result.push(new TextRun({text:part,superScript:/^\[\d/.test(part)}));
  }
  return result;
}
function mathFromLatex(latex) {
  const temml=loadDependency('temml');
  const {mathmlToDocxChildren}=require('./mathml-to-docx');
  const ml=temml.renderToString(latex,{displayMode:true,throwOnError:true});
  const children=mathmlToDocxChildren(ml);
  if(!children.length) throw new Error('Empty native equation: '+latex);
  return new WordMath({children});
}

class AcademicBuilder {
  constructor(){this.chapter=0;this.figure=0;this.tableCount=0;this.equation=0;this.prefix=null;}
  body(t){return new Paragraph({style:'Normal',children:inline(t)});}
  text(t,opts={}){return new Paragraph({style:'Normal',indent:resetIndent,...opts,children:inline(t)});}
  heading(t,level=1){
    if(level<1 || level>4) throw new Error('Heading level must be 1..4');
    if(level===1){this.chapter++;this.figure=this.tableCount=this.equation=0;this.prefix=null;}
    const p=new Paragraph({heading:H['HEADING_'+level],numbering:{reference:'chapters',level:level-1},children:[new TextRun(t),...(level===1?[new SimpleField('SEQ chapter \\h','')]:[])]});
    p.academicHeading=true;p.academicHeadingLevel=level;p.academicTitle=t;p.academicNumbered=true;return p;
  }
  unnumbered(t,{appendix=null}={}){
    if(appendix){this.prefix=appendix;this.figure=this.tableCount=this.equation=0;}
    const p=new Paragraph({heading:H.HEADING_1,numbering:false,children:[new TextRun(t)]});
    p.academicHeading=true;p.academicHeadingLevel=1;p.academicTitle=t;p.academicNumbered=false;return p;
  }
  frontHeading(t){const p=new Paragraph({style:'FrontHeading',children:[new TextRun(t)]});p.academicHeading=true;return p;}
  code(t){return new Paragraph({style:'Code',children:[new TextRun({text:String(t)||' ',font:{ascii:'Consolas',hAnsi:'Consolas',eastAsia:'SimSun'},size:21})]});}
  ref(t){return new Paragraph({style:'Reference',numbering:{reference:'references',level:0},children:[new TextRun(t)]});}
  toc(){return new TableOfContents('',{hyperlink:true,headingStyleRange:'1-4'});}
  sequence(kind,count){
    const appendix=this.prefix;
    return [appendix?new TextRun(appendix):new SimpleField('SEQ chapter \\c',String(this.chapter)),
      new TextRun('-'),new SimpleField(appendix?'SEQ '+kind+'_'+appendix:'SEQ '+kind+' \\s 1',String(count))];
  }
  caption(t,kind='figure'){
    const isFigure=kind==='figure';const count=isFigure?++this.figure:++this.tableCount;
    return new Paragraph({style:isFigure?'FigureCaption':'TableCaption',children:[new TextRun(isFigure?'图 ':'表 '),...this.sequence(kind,count),new TextRun(' '+t)]});
  }
  image(file,width,height,description){
    const maximum=W/15;const ratio=width>maximum?maximum/width:1;
    return new Paragraph({alignment:A.CENTER,indent:resetIndent,spacing:single(6,0),keepNext:true,keepLines:true,
      children:[new ImageRun({type:path.extname(file).slice(1).replace('jpeg','jpg'),data:fs.readFileSync(file),
        transformation:{width:width*ratio,height:height*ratio},altText:{title:description,description,name:path.basename(file)}})]});
  }
  table(headers,rows,widths=null,{style=profile.table.style}={}){
    if(!headers.length || rows.some(row=>row.length!==headers.length))throw new Error('Table column count mismatch');
    if(!widths){widths=headers.map(()=>Math.floor(W/headers.length));widths[widths.length-1]+=W-widths.reduce((a,b)=>a+b,0);}
    else if(widths.reduce((a,b)=>a+b,0)!==W){const total=widths.reduce((a,b)=>a+b,0);widths=widths.map(v=>Math.round(v*W/total));widths[widths.length-1]+=W-widths.reduce((a,b)=>a+b,0);}
    const rule={style:BorderStyle.SINGLE,size:4,color:'000000'};
    const all=[headers,...rows];const short=all.length<=12;
    const tr=all.map((row,i)=>new TableRow({tableHeader:i===0,cantSplit:true,children:row.map((v,j)=>{
      const borders=style==='three-line'?{top:i===0?{...rule,size:12}:NONE,bottom:i===0?{...rule,size:6}:i===all.length-1?{...rule,size:12}:NONE,left:NONE,right:NONE}:{top:rule,bottom:rule,left:rule,right:rule};
      return new TableCell({width:{size:widths[j],type:WidthType.DXA},verticalAlign:VerticalAlign.CENTER,
        margins:{top:0,bottom:0,left:90,right:90},borders,children:[new Paragraph({style:'TableText',keepNext:short&&i<all.length-1,children:inline(String(v))})]});})}));
    return new Table({width:{size:W,type:WidthType.DXA},columnWidths:widths,borders:style==='three-line'?BORDERS_NONE:{top:rule,bottom:rule,left:rule,right:rule,insideHorizontal:rule,insideVertical:rule},rows:tr});
  }
  formula(children,number=null){
    const count=++this.equation;const side=profile.formula.sideWidth;
    const math=typeof children==='string'?mathFromLatex(children):new WordMath({children});
    const numberRun=text=>new TextRun({text,font:FORMULA_NUMBER_FONT,size:profile.formula.size*2});
    const label=number!==null?[numberRun('('+number+')')]:[numberRun('('),...this.sequence('equation',count),numberRun(')')];
    const widths=[side,W-2*side,side];
    return new Table({width:{size:W,type:WidthType.DXA},columnWidths:widths,borders:BORDERS_NONE,rows:[new TableRow({cantSplit:true,children:[[],[math],label].map((v,i)=>new TableCell({width:{size:widths[i],type:WidthType.DXA},verticalAlign:VerticalAlign.CENTER,margins:{top:0,bottom:0,left:0,right:0},borders:BORDERS_NONE,
      children:[new Paragraph({style:i===2?'FormulaNumber':'Formula',alignment:i===2?A.RIGHT:A.CENTER,spacing:single(6,6),children:v})]}))})]});
  }
}

function normalizeFlow(children){
  let previous=null;
  for(const child of children){
    if(child.academicHeading && previous?.academicHeading){
      const spacing=new Paragraph({spacing:{before:0}}).properties.root.find(x=>x.rootKey==='w:spacing');
      child.properties.push(spacing);
    }
    previous=child;
  }
  return children;
}
function header(t){return new Header({children:[new Paragraph({style:'Header',children:Array.isArray(t)?t:[new TextRun(t)]})]});}
const chapterHeader=title=>header([new SimpleField('STYLEREF 1 \\n','第一章'),new TextRun(' '),new SimpleField('STYLEREF 1',title)]);
const blankHeader=()=>new Header({children:[new Paragraph({indent:resetIndent,border:{bottom:NONE},spacing:single(),children:[]})]});
const footer=()=>new Footer({children:[new Paragraph({style:'Footer',children:[new TextRun({children:[PageNumber.CURRENT],font:FONT,size:18})]})]});
const blankFooter=()=>new Footer({children:[]});
const page=(number=null)=>({size:{width:profile.page.width,height:profile.page.height},margin:{top:profile.page.margin,bottom:profile.page.margin,left:profile.page.margin,right:profile.page.margin,header:profile.page.header,footer:profile.page.footer},...(number?{pageNumbers:number}:{})});

function createAcademicDocument({title,creator='付烁',cover,coverTemplate=profile.cover.default,coverFields={},front=[],body=[],bodyHeader=title,duplex=profile.layout.duplexDefault}){
  const defaultCover = cover===undefined && coverTemplate==='uestc-course';
  cover = cover ?? (defaultCover?[new Paragraph({children:[]})]:[]);
  const sections=[];
  if(cover.length)sections.push({properties:{page:page(),type:SectionType.NEXT_PAGE},headers:{default:blankHeader(),even:blankHeader(),first:blankHeader()},footers:{default:blankFooter(),even:blankFooter(),first:blankFooter()},children:normalizeFlow(cover)});
  front.forEach((f,i)=>sections.push({properties:{page:page({formatType:NumberFormat.UPPER_ROMAN,...(i===0?{start:1}:{})}),type:duplex?SectionType.ODD_PAGE:SectionType.NEXT_PAGE},headers:{default:header(f.title),even:header(f.title)},footers:{default:footer(),even:footer()},children:normalizeFlow(f.children)}));
  const chapters=[];let current=[];
  for(const item of body){if(item.academicHeadingLevel===1 && current.length){chapters.push(current);current=[];}current.push(item);}
  if(current.length)chapters.push(current);
  chapters.forEach((children,i)=>{
    const lead=children.find(item=>item.academicHeadingLevel===1);
    const bodySectionHeader=lead?.academicNumbered?chapterHeader(lead.academicTitle||bodyHeader):header(lead?.academicTitle||bodyHeader);
    sections.push({properties:{page:page({formatType:NumberFormat.DECIMAL,...(i===0?{start:1}:{})}),type:i===0&&duplex?SectionType.ODD_PAGE:SectionType.NEXT_PAGE},headers:{default:bodySectionHeader},footers:{default:footer()},children:normalizeFlow(children)});
  });
  const document=new Document({creator,lastModifiedBy:creator,title,description:title,features:{updateFields:true},compatibility:{suppressSpacingAtTopOfPage:false,suppressSpBfAfterPgBrk:false},evenAndOddHeaderAndFooters:false,styles:styles(),numbering:numbering(),sections});
  if(defaultCover)document.academicDefaultCover={title,...coverFields};
  return document;
}
async function packAcademicDocument(doc){
  const buffer=await d.Packer.toBuffer(doc);
  const JSZip=loadDependency('jszip');
  const zip=await JSZip.loadAsync(buffer);
  const file=zip.file('word/numbering.xml');
  if(file){
    const xml=await file.async('string');
    const fixedXml=xml.replace(/(<w:suff\b[^>]*\/>)(<w:isLgl\s*\/>)/g,'$2$1');
    if(fixedXml!==xml)zip.file('word/numbering.xml',fixedXml);
  }
  const documentFile=zip.file('word/document.xml');
  if(documentFile)zip.file('word/document.xml',normalizeSectionBoundaries(await documentFile.async('string')));
  if(doc.academicDefaultCover)await require('./default_cover').applyDefaultCover(zip,doc.academicDefaultCover);
  return zip.generateAsync({type:'nodebuffer',compression:'DEFLATE'});
}

// A next-page section break alone owns chapter pagination. In particular, a
// separate normal-height section-marker paragraph can overflow a full last page.
function normalizeSectionBoundaries(xml){
  const {XMLParser,XMLBuilder}=loadDependency('fast-xml-parser');
  const options={preserveOrder:true,ignoreAttributes:false,trimValues:false,parseTagValue:false,parseAttributeValue:false};
  const tree=new XMLParser(options).parse(xml);
  const document=tree.find(node=>node['w:document'])?.['w:document'];
  const body=document?.find(node=>node['w:body'])?.['w:body'];
  if(!body)throw new Error('Missing Word document body');
  const empty=node=>{
    if(node['w:pPr']||node['w:rPr'])return true;
    if(node['w:r'])return node['w:r'].every(empty);
    if(node['w:t'])return node['w:t'].every(part=>!String(part['#text']||'').trim());
    if(node['w:br'])return node[':@']?.['@_w:type']==='page';
    return false;
  };
  const clearTerminalBreaks=paragraph=>{
    for(let i=paragraph.length-1;i>=0;i--){
      const run=paragraph[i]['w:r'];
      if(!run){if(paragraph[i]['w:pPr'])continue;break;}
      while(run.length){
        const tail=run[run.length-1];
        if(tail['w:br']&&tail[':@']?.['@_w:type']==='page')run.pop();
        else break;
      }
      if(run.every(empty))paragraph.splice(i,1);else break;
    }
  };
  for(let i=0;i<body.length;i++){
    const paragraph=body[i]['w:p'];
    const properties=paragraph?.find(node=>node['w:pPr'])?.['w:pPr'];
    const sectionIndex=properties?.findIndex(node=>node['w:sectPr']);
    if(sectionIndex===undefined||sectionIndex<0||!paragraph.every(empty))continue;
    // Remove only empty/explicit-page-break boundary paragraphs, retaining all
    // content, bookmarks and other anchors.
    while(i>0&&body[i-1]['w:p']?.every(empty)){
      const previous=body[i-1]['w:p'];
      if(previous.some(node=>node['w:pPr']?.some(part=>part['w:sectPr'])))break;
      body.splice(i-1,1);i--;
    }
    const previous=body[i-1]?.['w:p'];
    if(previous){
      clearTerminalBreaks(previous);
      let previousProperties=previous.find(node=>node['w:pPr']);
      if(!previousProperties){previousProperties={'w:pPr':[]};previous.unshift(previousProperties);}
      if(!previousProperties['w:pPr'].some(node=>node['w:sectPr'])){
        previousProperties['w:pPr'].push(properties[sectionIndex]);
        body.splice(i,1);i--;
        continue;
      }
    }
    // Word requires a paragraph after a table. Keep its section marker compact,
    // with no inherited heading pagination, spacing or keep-with-next behavior.
    properties.splice(0,properties.length,
      {'w:keepNext':[],':@':{'@_w:val':'0'}},
      {'w:keepLines':[],':@':{'@_w:val':'0'}},
      {'w:pageBreakBefore':[],':@':{'@_w:val':'0'}},
      {'w:spacing':[],':@':{'@_w:before':'0','@_w:after':'0','@_w:line':'20','@_w:lineRule':'exact'}},
      {'w:rPr':[{'w:sz':[],':@':{'@_w:val':'2'}}]},properties[sectionIndex]);
  }
  // Section-start Heading1 paragraphs must not add a second page break.
  for(const node of body){
    const properties=node['w:p']?.find(part=>part['w:pPr'])?.['w:pPr'];
    if(properties?.some(part=>part['w:pStyle']&&part[':@']?.['@_w:val']==='Heading1')){
      for(let i=properties.length-1;i>=0;i--)if(properties[i]['w:pageBreakBefore'])properties.splice(i,1);
    }
  }
  return new XMLBuilder({...options,suppressEmptyNode:true}).build(tree);
}
module.exports={d,profile,W,FONT,HEAD,FORMULA_NUMBER_FONT,NONE,BORDERS_NONE,fixed,single,styles,numbering,inline,mathFromLatex,AcademicBuilder,createAcademicDocument,packAcademicDocument,normalizeFlow,normalizeSectionBoundaries};
