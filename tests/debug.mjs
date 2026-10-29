import fs from 'node:fs';
import JSZip from 'jszip';
import { XMLParser } from 'fast-xml-parser';
const p=new XMLParser({ignoreAttributes:false,attributeNamePrefix:'@_',removeNSPrefix:true,allowBooleanAttributes:true,parseTagValue:false,trimValues:false});
const asArray=v=>v==null?[]:Array.isArray(v)?v:[v];
function each(node,key,fn){
  if(!node||typeof node!=='object')return;
  if(Array.isArray(node)){node.forEach(i=>each(i,key,fn));return;}
  for(const [k,v] of Object.entries(node)){
    if(k===key)asArray(v).forEach(fn);
    each(v,key,fn);
  }
}
function collectText(node){
  let out='';
  if(node==null)return out;
  if(typeof node==='string')return node;
  if(typeof node==='number'||typeof node==='boolean')return String(node);
  if(Array.isArray(node)){node.forEach(i=>out+=collectText(i));return out;}
  if(typeof node!=='object')return out;
  for(const [k,v] of Object.entries(node)){
    if(k==='t'){
      asArray(v).forEach(i=>{if(typeof i==='string')out+=i;else if(i&&typeof i==='object'&&i['#text'])out+=i['#text'];});
    } else if(k!=='rPr'&&k!=='pPr'&&k!=='endParaRPr') out+=collectText(v);
  }
  return out;
}
const z=await JSZip.loadAsync(fs.readFileSync('../mvp/tests/fixtures/bad.docx'));
const d=p.parse(await z.file('word/document.xml').async('string'));
each(d.document.body,'hyperlink',h=>console.log('FOUND',JSON.stringify(collectText(h)),h['@_id']));
