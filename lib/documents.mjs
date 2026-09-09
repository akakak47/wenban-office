import {wordArtStrings,replaceWordArtFields} from './namecard-templates.mjs';
import PizZip from 'pizzip';
import Docxtemplater from 'docxtemplater';
import {Document,Packer,Paragraph,TextRun,AlignmentType,PageOrientation,BorderStyle,Table,TableRow,TableCell,WidthType,VerticalAlign,HeightRule} from 'docx';
export const DOCX_TYPE='application/vnd.openxmlformats-officedocument.wordprocessingml.document';
export function inspectTemplate(bytes) {
 const zip=new PizZip(bytes);const main=zip.file('word/document.xml');if(!main)throw new Error('请使用有效的 .docx 文件；旧版 .doc 请先另存为 .docx。');
 const fields=new Set();let headerFields=false;
 for(const value of wordArtStrings(bytes))for(const match of value.matchAll(/\{([^{}]+)\}/g))fields.add(match[1].trim());
 for(const name of Object.keys(zip.files).filter(n=>/^word\/(document|header\d+|footer\d+)\.xml$/.test(n))){
 const xml=zip.file(name).asText();if(name!=='word/document.xml'&&/<v:textpath[^>]+string="[^"]*\{/.test(xml))headerFields=true;const plain=xml.replace(/<\/w:p>/g,'\n').replace(/<[^>]+>/g,'');
 for(const match of plain.matchAll(/\{([^{}]+)\}/g)){const key=match[1].trim();if(!/^[#/@^]/.test(key))fields.add(key);if(name!=='word/document.xml')headerFields=true;}
 }
 const xml=main.asText();
 return {fields:[...fields],text:wordArtStrings(bytes).join('\n')+'\n'+xml.replace(/<\/w:p>/g,'\n').replace(/<[^>]+>/g,'').replace(/&amp;/g,'&').replace(/&lt;/g,'<').replace(/&gt;/g,'>'),canMerge:!headerFields&&(xml.match(/<w:sectPr[ >]/g)||[]).length<=1};
}
export function renderTemplate(bytes,data) {
 try {const doc=new Docxtemplater(new PizZip(bytes),{paragraphLoop:true,linebreaks:true,nullGetter:()=>'',errorLogging:false});doc.render(data);return replaceWordArtFields(doc.getZip(),data).generate({type:'uint8array',compression:'DEFLATE'});}catch(e){throw new Error('模板无法填充。请使用单层花括号占位符（如 {姓名}），并检查是否有未闭合的花括号。'+(e.properties?.errors?.[0]?.properties?.explanation||''));}
}
export function mergeTemplateDocuments(documents) {
 if(!documents.length)throw new Error('没有可生成的记录');
 const zip=new PizZip(documents[0]);const xml=zip.file('word/document.xml').asText();
 const match=xml.match(/<w:body>([\s\S]*?)<\/w:body>/);if(!match)throw new Error('无法识别 Word 正文');
 const section=match[1].match(/<w:sectPr[\s\S]*?<\/w:sectPr>\s*$/)?.[0]||'';
 let id=1;
 const bodies=documents.map((bytes,i)=>{
 let body=new PizZip(bytes).file('word/document.xml').asText().match(/<w:body>([\s\S]*?)<\/w:body>/)[1].replace(/<w:sectPr[\s\S]*?<\/w:sectPr>\s*$/,'');
 // Inline drawings must have unique identifiers after repeating a template.
 body=body.replace(/(<v:shape\b[^>]*\bid=")[^"]+/g,(_,prefix)=>prefix+'wenban_merge_'+id++).replace(/(<wp:docPr\b[^>]*\bid=")[^"]+/g,(_,prefix)=>prefix+id++).replace(/<w:bookmarkStart\b[^>]*\/>|<w:bookmarkEnd\b[^>]*\/>/g,'');
 return (i?'<w:p><w:pPr><w:spacing w:before="0" w:after="0" w:line="1" w:lineRule="exact"/></w:pPr><w:r><w:br w:type="page"/></w:r></w:p>':'')+body;
 }).join('');
 zip.file('word/document.xml',xml.replace(match[0],'<w:body>'+bodies+section+'</w:body>'));
 return zip.generate({type:'uint8array',compression:'DEFLATE'});
}
const font='Microsoft YaHei';
const p=(text,size=28,extra={})=>new Paragraph({alignment:AlignmentType.CENTER,spacing:{after:220},children:[new TextRun({text,font,size,...extra})]});
export async function certificateBytes(records,{title='荣誉证书',award='优秀学生',issuer='学校教务处',date='2026年9月',template=false}={}) {
 const section=record=>({properties:{page:{size:{orientation:PageOrientation.LANDSCAPE,width:11906,height:16838},margin:{top:950,bottom:900,left:1250,right:1250},borders:{pageBorderTop:{style:BorderStyle.DOUBLE,size:12,color:'A72D3C'},pageBorderBottom:{style:BorderStyle.DOUBLE,size:12,color:'A72D3C'},pageBorderLeft:{style:BorderStyle.DOUBLE,size:12,color:'A72D3C'},pageBorderRight:{style:BorderStyle.DOUBLE,size:12,color:'A72D3C'}}}},children:[p(title,72,{bold:true,color:'9E2839'}),p('CERTIFICATE OF HONOR',20,{color:'9E2839',characterSpacing:70}),p(record['姓名']+(template?'':' 同学'),48,{bold:true}),p('在本学期学习与实践中表现突出，被评为',28),p(record['奖项']||award,46,{bold:true,color:'9E2839'}),p('特发此证，以资鼓励。',28),p(issuer,24),p(date,22)]});
 return new Uint8Array(await Packer.toArrayBuffer(new Document({sections:records.map(section)})));
}
export async function nameCardBytes(names,{size=96,subtitle='',color='222222'}={}) {
 if(!names.length)throw new Error('请先导入姓名');
 const esc=value=>String(value).replace(/[&<>"']/g,ch=>({'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;',"'":'&apos;'}[ch]));
 const safeColor=/^[0-9a-f]{6}$/i.test(color)?color:'222222';
 const pages=names.map((name,i)=>({properties:{page:{size:{width:11906,height:16838},margin:{top:720,bottom:720,left:720,right:720}}},children:[new Paragraph({children:[new TextRun(`WENBAN_FOLD_${i}`)]})]}));
 const zip=new PizZip(await Packer.toArrayBuffer(new Document({sections:pages})));
 let xml=zip.file('word/document.xml').asText();
 for(const [prefix,uri] of [['v','urn:schemas-microsoft-com:vml'],['o','urn:schemas-microsoft-com:office:office'],['w10','urn:schemas-microsoft-com:office:word'],['wps','http://schemas.microsoft.com/office/word/2010/wordprocessingShape'],['a','http://schemas.openxmlformats.org/drawingml/2006/main']])if(!xml.includes(`xmlns:${prefix}=`))xml=xml.replace('<w:document ',`<w:document xmlns:${prefix}="${uri}" `);
 names.forEach((name,i)=>{
  const units=[...name].reduce((n,ch)=>n+(ch.charCodeAt(0)>255?1:.6),0);
  const pts=Math.max(12,Math.min(Number(size)||96,Math.floor(480/Math.max(1,units))));
  const anchor=(top,height,index,shape)=>`<w:r><w:drawing><wp:anchor distT="0" distB="0" distL="0" distR="0" simplePos="0" relativeHeight="${i*3+index+1}" behindDoc="0" locked="0" layoutInCell="1" allowOverlap="1"><wp:simplePos x="0" y="0"/><wp:positionH relativeFrom="page"><wp:posOffset>457200</wp:posOffset></wp:positionH><wp:positionV relativeFrom="page"><wp:posOffset>${Math.round(top*12700)}</wp:posOffset></wp:positionV><wp:extent cx="6645910" cy="${Math.round(height*12700)}"/><wp:effectExtent l="0" t="0" r="0" b="0"/><wp:wrapNone/><wp:docPr id="${i*3+index+1}" name="Name card ${i+1} panel ${index+1}"/><wp:cNvGraphicFramePr/><a:graphic><a:graphicData uri="http://schemas.microsoft.com/office/word/2010/wordprocessingShape"><wps:wsp>${shape}</wps:wsp></a:graphicData></a:graphic></wp:anchor></w:drawing></w:r>`;
  // Use editable WordArt, as in the supplied reference. Text boxes can keep
  // their text upright despite shape rotation in some Word-compatible readers.
  const art=(text,fs,center,rotation,index)=>{
   const u=[...text].reduce((n,ch)=>n+(ch.charCodeAt(0)>255?1:.6),0);
   const width=Math.min(513.3,Math.max(fs,fs*u));const height=fs*1.12;
   return `<w:r><w:rPr><w:rFonts w:ascii="华文行楷" w:hAnsi="华文行楷" w:eastAsia="华文行楷"/></w:rPr><w:pict><v:shape id="wenban_art_${i}_${index}" o:spt="136" type="#_x0000_t136" style="position:absolute;left:0pt;${rotation?'flip:x y;':''}margin-left:${(595.3-width)/2}pt;margin-top:${center-height/2}pt;height:${height}pt;width:${width}pt;mso-position-horizontal-relative:page;mso-position-vertical-relative:page;" fillcolor="#${safeColor}" filled="t" stroked="f" coordsize="21600,21600" adj="10800"><v:path textpathok="t"/><v:textpath on="t" fitshape="t" fitpath="t" trim="t" xscale="f" string="${esc(text)}" style="font-family:华文行楷;font-size:${fs}pt;v-text-align:center;"/><w10:wrap type="none"/></v:shape></w:pict></w:r>`;
  };
  const box=(top,rotation,index)=>{
   const center=top+324.95/2;
   return art(name,pts,center+(subtitle?(rotation?16:-16):0),rotation,index)+(subtitle?art(subtitle,Math.min(18,480/Math.max(1,subtitle.length)),center+(rotation?-1:1)*(pts*.56+26),rotation,index+4):'');
  };
  const line=anchor(420.95,0,2,`<wps:cNvSpPr/><wps:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="6645910" cy="0"/></a:xfrm><a:prstGeom prst="line"><a:avLst/></a:prstGeom><a:ln w="6350"><a:solidFill><a:srgbClr val="C8CDD4"/></a:solidFill><a:prstDash val="dash"/></a:ln></wps:spPr><wps:bodyPr/>`);
  const marker=new RegExp('<w:p>(?:(?!<w:p>).)*?WENBAN_FOLD_'+i+'(?:(?!<w:p>).)*?</w:p>','s');
  xml=xml.replace(marker,'<w:p><w:pPr><w:spacing w:before="0" w:after="0"/></w:pPr>'+box(48.0,180,0)+box(468.95,0,1)+line+'</w:p>');
 });
 zip.file('word/document.xml',xml);
 return zip.generate({type:'uint8array',compression:'DEFLATE'});
}
export function zipFiles(files){const zip=new PizZip();files.forEach(({name,bytes})=>zip.file(name,bytes));return zip.generate({type:'uint8array',compression:'DEFLATE'});}
export function download(bytes,name,type=DOCX_TYPE){const blob=new Blob([bytes],{type});const url=URL.createObjectURL(blob);const a=document.createElement('a');a.href=url;a.download=name;document.body.appendChild(a);a.click();a.remove();setTimeout(()=>URL.revokeObjectURL(url),30000);}
