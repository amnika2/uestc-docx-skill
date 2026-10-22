#!/usr/bin/env node
'use strict';
const fs=require('fs');
const path=require('path');
const {AcademicBuilder,createAcademicDocument,packAcademicDocument,d}=require('./academic_doc');

const stripHeading=t=>t.replace(/^(?:第[一二三四五六七八九十百\d]+章\s*|[一二三四五六七八九十]+[、．.]\s*|\d+(?:\.\d+)*[.、]?\s+)/,'').trim();
const stripCaption=t=>t.replace(/^[图表]\s*[A-Z\d]+(?:[-.]\d+)?\s*/,'').trim();
function parseMarkdown(source,baseDir=process.cwd()){
  const b=new AcademicBuilder();const lines=source.replace(/\r/g,'').split('\n');
  const front=[],body=[];let target=body,title='中文学术报告',reference=false;
  let i=0;
  while(i<lines.length){
    const raw=lines[i],t=raw.trim();
    if(!t || t==='---'){i++;continue;}
    if(/^#\s+/.test(t)){title=t.replace(/^#\s+/,'');i++;continue;}
    const h=t.match(/^(#{2,5})\s+(.+)$/);
    if(h){
      const label=h[2].trim();
      if(h[1].length===2 && /^(?:摘要|摘 要|Abstract|目录|目 录)$/.test(label)){
        const key=label.replace(/ /g,'');const f={title:key,children:[b.frontHeading(label)]};front.push(f);target=f.children;reference=false;
        if(key==='目录')target.push(b.toc());i++;continue;
      }
      target=body;
      if(h[1].length===2 && label==='参考文献'){target.push(b.unnumbered(label));reference=true;}
      else if(h[1].length===2 && /^附录/.test(label)){const prefix=(label.match(/^附录\s*([A-Z])/)||[])[1]||'A';target.push(b.unnumbered(label,{appendix:prefix}));reference=false;}
      else{target.push(b.heading(stripHeading(label),h[1].length-1));reference=false;}
      i++;continue;
    }
    if(t.startsWith('```') || t.startsWith('~~~')){
      const marker=t.slice(0,3);i++;
      while(i<lines.length && !lines[i].trim().startsWith(marker)){target.push(b.code(lines[i]));i++;}
      i++;continue;
    }
    if(t.startsWith('$$')){
      let latex=t.slice(2),closed=latex.endsWith('$$');
      if(closed)latex=latex.slice(0,-2);i++;
      while(!closed && i<lines.length){const line=lines[i].trim();if(line.endsWith('$$')){latex+='\n'+line.slice(0,-2);closed=true;}else latex+='\n'+line;i++;}
      if(!closed)throw new Error('Unclosed display equation');
      target.push(b.formula(latex.trim()));continue;
    }
    const im=t.match(/^!\[([^\]]*)\]\(([^)]+)\)$/);
    if(im){
      const file=path.resolve(baseDir,im[2]);
      const meta=fs.readFileSync(file);let width=550,height=330;
      if(path.extname(file).toLowerCase()==='.png'){const w=meta.readUInt32BE(16),hh=meta.readUInt32BE(20);height=width*hh/w;}
      const description=im[1]||path.basename(file);target.push(b.image(file,width,height,description));i++;
      if(i<lines.length && /^图\s*(?:\d|[A-Z])/.test(lines[i].trim())){target.push(b.caption(stripCaption(lines[i].trim()),'figure'));i++;}
      else target.push(b.caption(description,'figure'));
      continue;
    }
    if(/^表\s*(?:\d|[A-Z])/.test(t)){target.push(b.caption(stripCaption(t),'table'));i++;continue;}
    if(t.startsWith('|') && i+1<lines.length && /^\s*\|?[\s:|\-]+\|?\s*$/.test(lines[i+1])){
      const parseRow=x=>x.trim().replace(/^\|/,'').replace(/\|$/,'').split('|').map(v=>v.trim());
      const heads=parseRow(t);i+=2;const rows=[];
      while(i<lines.length && lines[i].trim().startsWith('|')){rows.push(parseRow(lines[i]));i++;}
      target.push(b.table(heads,rows));continue;
    }
    if(t.startsWith('<table')){
      let html=raw;i++;
      while(!html.includes('</table>') && i<lines.length)html+='\n'+lines[i++];
      const rows=[...html.matchAll(/<tr\b[^>]*>([\s\S]*?)<\/tr>/gi)].map(row=>[...row[1].matchAll(/<t[hd]\b[^>]*>([\s\S]*?)<\/t[hd]>/gi)].map(cell=>cell[1].replace(/<[^>]*>/g,'').replace(/&amp;/g,'&').replace(/&lt;/g,'<').replace(/&gt;/g,'>').trim()));
      if(!rows.length)throw new Error('Empty HTML table');target.push(b.table(rows[0],rows.slice(1)));continue;
    }
    if(reference && /^\[\d+\]/.test(t)){target.push(b.ref(t.replace(/^\[\d+\]\s*/,'')));i++;continue;}
    if(/^[-*]\s+/.test(t)){target.push(new d.Paragraph({numbering:{reference:'bullets',level:0},indent:{firstLine:0,left:480,hanging:240},children:[new d.TextRun(t.replace(/^[-*]\s+/,''))]}));i++;continue;}
    const pieces=[t];i++;
    while(i<lines.length && lines[i].trim() && !/^(?:#{1,5}\s|```|~~~|\$\$|!\[|\||<table|表\s*(?:\d|[A-Z])|[-*]\s+)/.test(lines[i].trim()))pieces.push(lines[i++].trim());
    target.push(/^(?:关键词|Keywords)[:：]/.test(t)?b.text(pieces.join(' ')):b.body(pieces.join(' ')));
  }
  if(!front.some(f=>f.title==='目录') && body.filter(x=>x.academicHeadingLevel===1).length>1)front.push({title:'目录',children:[b.frontHeading('目 录'),b.toc()]});
  return {title,front,body};
}

async function main(){
  const input=process.argv[2];
  if(!input){console.error('Usage: node new_doc.js input.md [output.docx] [cover-fields.json]');process.exitCode=2;return;}
  const file=path.resolve(input),output=path.resolve(process.argv[3]||file.replace(/\.[^.]+$/,'.docx'));
  const options=parseMarkdown(fs.readFileSync(file,'utf8'),path.dirname(file));
  if(process.argv[4])options.coverFields=JSON.parse(fs.readFileSync(path.resolve(process.argv[4]),'utf8'));
  fs.writeFileSync(output,await packAcademicDocument(createAcademicDocument(options)));
  console.log(output);
}
if(require.main===module)main().catch(e=>{console.error(e.stack);process.exitCode=1;});
module.exports={parseMarkdown,stripHeading,stripCaption};
