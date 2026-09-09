import PizZip from 'pizzip';
import {DOMParser,XMLSerializer} from '@xmldom/xmldom';
const W='http://schemas.openxmlformats.org/wordprocessingml/2006/main';
const V='urn:schemas-microsoft-com:vml';
const parts=zip=>Object.keys(zip.files).filter(n=>/^word\/(document|header\d+|footer\d+)\.xml$/.test(n));
const parse=xml=>new DOMParser().parseFromString(xml,'text/xml');
const all=(node,ns,tag)=>Array.from(node.getElementsByTagNameNS(ns,tag));
export function wordArtStrings(bytes){const zip=new PizZip(bytes);return [...new Set(parts(zip).flatMap(n=>all(parse(zip.file(n).asText()),V,'textpath').map(t=>t.getAttribute('string')||'')).filter(Boolean))];}
export function replaceWordArtFields(zip,data){for(const n of parts(zip)){const d=parse(zip.file(n).asText());let changed=false;for(const t of all(d,V,'textpath')){const text=t.getAttribute('string')||'';const next=text.replace(/\{([^{}]+)\}/g,(_,key)=>String(data[key.trim()]??''));if(text!==next){t.setAttribute('string',next);changed=true;}}if(changed)zip.file(n,new XMLSerializer().serializeToString(d));}return zip;}
export function prepareNameCardTemplate(bytes,originalName=''){
 const zip=new PizZip(bytes);let firstPairOnly=false,replacements=0;
 for(const n of parts(zip)){
  const d=parse(zip.file(n).asText());
  if(n==='word/document.xml'){
   const tables=all(d,W,'tbl');
   if(tables.length===1){const table=tables[0],rows=Array.from(table.childNodes).filter(c=>c.nodeType===1&&c.localName==='tr');
    const values=rows.map(row=>all(row,V,'textpath').map(t=>t.getAttribute('string')));
    if(rows.length>2&&rows.length%2===0&&values.every((v,i)=>v.length===1&&v[0]===values[i-i%2][0])){
     rows.slice(2).forEach(row=>table.removeChild(row));firstPairOnly=true;
    }
   }
  }
  if(originalName){
   for(const t of all(d,V,'textpath')){const text=t.getAttribute('string')||'';if(text.includes(originalName)){replacements+=text.split(originalName).length-1;t.setAttribute('string',text.split(originalName).join('{姓名}'));}}
   for(const paragraph of all(d,W,'p')){
    if(all(paragraph,W,'p').length)continue;
    const nodes=all(paragraph,W,'t');let text=nodes.map(t=>t.textContent||'').join('');let index=text.lastIndexOf(originalName);
    // Replace from right to left to preserve run formatting and split-run names.
    while(index>=0){let offset=0;const end=index+originalName.length;for(const node of nodes){const value=node.textContent||'',start=offset;offset+=value.length;if(offset<=index||start>=end)continue;const left=Math.max(0,index-start),right=Math.min(value.length,end-start);node.textContent=value.slice(0,left)+(start<=index?'{姓名}':'')+value.slice(right);}replacements++;text=nodes.map(t=>t.textContent||'').join('');index=text.lastIndexOf(originalName,index-1);}
   }
  }
  zip.file(n,new XMLSerializer().serializeToString(d));
 }
 return {bytes:zip.generate({type:'uint8array',compression:'DEFLATE'}),firstPairOnly,replacements};
}
