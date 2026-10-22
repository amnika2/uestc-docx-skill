'use strict';
const fs = require('fs');
const path = require('path');
const {loadDependency} = require('./runtime');
const options = {preserveOrder:true,ignoreAttributes:false,trimValues:false,
  parseTagValue:false,parseAttributeValue:false};
const clone = value => JSON.parse(JSON.stringify(value));
const node = (tree, key) => tree.find(part => part[key]);

async function applyDefaultCover(zip, fields) {
  const JSZip = loadDependency('jszip');
  const {XMLParser,XMLBuilder} = loadDependency('fast-xml-parser');
  const parser = new XMLParser(options);
  const builder = new XMLBuilder({...options,suppressEmptyNode:true});
  const profile = require('./academic-profile.json');
  const templatePath = path.resolve(__dirname,'..',profile.cover.asset);
  const assetRoot = path.dirname(templatePath);
  const template = await JSZip.loadAsync(fs.readFileSync(templatePath));
  const widths = JSON.parse(fs.readFileSync(path.join(assetRoot,'cover-fields.json'),'utf8'));
  const read = async (archive, name) => parser.parse(await archive.file(name).async('string'));
  const source = await read(template,'word/document.xml');
  const destination = await read(zip,'word/document.xml');
  const sourceRoot = node(source,'w:document');
  const destinationRoot = node(destination,'w:document');
  const sourceBody = node(sourceRoot['w:document'],'w:body')['w:body'];
  const body = node(destinationRoot['w:document'],'w:body')['w:body'];
  const end = body.findIndex(p => p['w:p']?.some(part =>
    part['w:pPr']?.some(prop => prop['w:sectPr'])));
  if(end < 0) throw new Error('Default cover requires a separate cover section');
  const section = node(node(body[end]['w:p'],'w:pPr')['w:pPr'],'w:sectPr');
  const paragraphs = clone(sourceBody.filter(p => p['w:p']));

  // Keep the inherited blank header/footer references, with the source page
  // geometry and grid that position the cover artwork and information rows.
  const sourceSection = clone(node(sourceBody,'w:sectPr')['w:sectPr']);
  section['w:sectPr'] = [...section['w:sectPr'].filter(p =>
    p['w:headerReference'] || p['w:footerReference'] || p['w:type']),
    ...sourceSection.filter(p => !p['w:headerReference'] && !p['w:footerReference'] && !p['w:type'])];

  const visit = (value, callback) => {
    if(!value || typeof value !== 'object') return;
    callback(value);
    for(const child of Object.values(value)) visit(child,callback);
  };
  let drawingId=0;
  visit(destination,part=>{
    if(part['wp:docPr'])drawingId=Math.max(drawingId,Number(part[':@']?.['@_id']||0));
  });
  visit(paragraphs,part=>{if(part['wp:docPr'])part[':@']['@_id']=String(++drawingId);});
  const width = text => [...text].reduce((total,c) => total+(c.codePointAt(0)>255?2:1),0);
  visit(paragraphs, part => {
    if(typeof part['#text'] !== 'string') return;
    part['#text'] = part['#text'].replace(/\{\{(\w+)\}\}/g, (_, key) => {
      const text = String(fields[key] ?? '');
      const spare = Math.max(0,widths[key].width-width(text));
      return ' '.repeat(Math.floor(spare/2))+text+' '.repeat(Math.ceil(spare/2));
    });
  });
  const sourceRels = node(await read(template,'word/_rels/document.xml.rels'),'Relationships')['Relationships'];
  const relTree = await read(zip,'word/_rels/document.xml.rels');
  const rels = node(relTree,'Relationships')['Relationships'];
  const used = new Set();
  visit(paragraphs, part => {
    for(const key of ['@_r:embed','@_r:link','@_r:id'])
      if(part[key]) used.add(part[key]);
  });
  const contentTree = await read(zip,'[Content_Types].xml');
  const content = node(contentTree,'Types')['Types'];
  let counter=0;
  for(const id of used) {
    const rel = sourceRels.find(r => r[':@']?.['@_Id']===id);
    const attrs = rel[':@'];
    if(!attrs['@_Type'].endsWith('/image')) throw new Error('Unexpected cover relationship: '+id);
    const extension = path.extname(attrs['@_Target']);
    let newId, target;
    do {counter++;newId='rIdUESTCCover'+counter;target='media/uestc-cover-'+counter+extension;}
    while(rels.some(r=>r[':@']?.['@_Id']===newId) || zip.file('word/'+target));
    zip.file('word/'+target,await template.file('word/'+attrs['@_Target']).async('nodebuffer'));
    rels.push({Relationship:[],':@':{...attrs,'@_Id':newId,'@_Target':target}});
    visit(paragraphs, part => {
      for(const key of ['@_r:embed','@_r:link','@_r:id']) if(part[key]===id)part[key]=newId;
    });
    if(!content.some(p=>p.Default && p[':@']?.['@_Extension']===extension.slice(1)))
      content.push({Default:[],':@':{'@_Extension':extension.slice(1),'@_ContentType':extension==='.png'?'image/png':'image/jpeg'}});
  }
  const last = paragraphs[paragraphs.length-1]['w:p'];
  node(last,'w:pPr')['w:pPr'].push(section);
  body.splice(0,end+1,...paragraphs);
  for(const [key,value] of Object.entries(sourceRoot[':@']))
    if(key.startsWith('@_xmlns:') && !(key in destinationRoot[':@'])) destinationRoot[':@'][key]=value;
  const sourceStyles = node(await read(template,'word/styles.xml'),'w:styles')['w:styles'];
  const styleTree = await read(zip,'word/styles.xml');
  const targetStyles = node(styleTree,'w:styles')['w:styles'];
  targetStyles.push(clone(sourceStyles.find(s=>s[':@']?.['@_w:styleId']==='UESTCCoverNormal')));
  // Retain the source cover's visible underlines when information fields are
  // empty. Keep ordinary spaces so the original font metrics remain unchanged.
  const settingsTree=await read(zip,'word/settings.xml');
  const settings=node(settingsTree,'w:settings')['w:settings'];
  let compat=node(settings,'w:compat');
  if(!compat){compat={'w:compat':[]};settings.push(compat);}
  if(!node(compat['w:compat'],'w:ulTrailSpace'))compat['w:compat'].unshift({'w:ulTrailSpace':[]});
  zip.file('word/document.xml',builder.build(destination));
  zip.file('word/styles.xml',builder.build(styleTree));
  zip.file('word/_rels/document.xml.rels',builder.build(relTree));
  zip.file('[Content_Types].xml',builder.build(contentTree));
  zip.file('word/settings.xml',builder.build(settingsTree));
}
module.exports = {applyDefaultCover};
