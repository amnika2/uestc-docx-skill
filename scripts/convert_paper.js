#!/usr/bin/env node
'use strict';
// Compatibility entrypoint; all formatting lives in the shared academic profile.
const fs=require('fs');
const path=require('path');
const {parseMarkdown}=require('./new_doc');
const {createAcademicDocument,packAcademicDocument}=require('./academic_doc');
async function main(){
  if(!process.argv[2]){console.error('Usage: node convert_paper.js input.md [output.docx]');process.exitCode=2;return;}
  const input=path.resolve(process.argv[2]);const output=path.resolve(process.argv[3]||input.replace(/\.[^.]+$/,'.docx'));
  fs.writeFileSync(output,await packAcademicDocument(createAcademicDocument(parseMarkdown(fs.readFileSync(input,'utf8'),path.dirname(input)))));
  console.log(output);
}
main().catch(e=>{console.error(e.stack);process.exitCode=1;});
