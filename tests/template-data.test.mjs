import test from 'node:test';
import assert from 'node:assert/strict';
import * as data from '../lib/template-data.mjs';
test('columns preserve blanks, duplicates and string identifiers',()=>{
 assert.equal(typeof data.columnsToRecords,'function');
 const r=data.columnsToRecords([{field:'姓名',text:'张三\n\n张三\n'},{field:'学号',text:'001\n002\n003'}],['姓名','学号']);
 assert.equal(r.errors.length,0);assert.deepEqual(r.records,[{'姓名':'张三','学号':'001'},{'姓名':'','学号':'002'},{'姓名':'张三','学号':'003'}]);
});
test('unequal columns require explicit padding and missing fields block export',()=>{
 const c=[{field:'姓名',text:'甲\n乙'},{field:'电话',text:'0123'}];
 assert.ok(data.columnsToRecords(c,['姓名','电话']).errors.length);
 assert.equal(data.columnsToRecords(c,['姓名','电话'],true).records[1]['电话'],'');
 assert.ok(data.columnsToRecords(c,['姓名','电话','学号'],true).errors.some(x=>x.includes('学号')));
 assert.ok(data.columnsToRecords([...c,c[0]],['姓名'],true).errors.some(x=>x.includes('重复')));
});

test('multiple fields fill repeated text, tables and headers without rebuilding styles',async()=>{
 const {Document,Packer,Paragraph,TextRun,Table,TableRow,TableCell,Header}=await import('docx');
 const {default:PizZip}=await import('pizzip');
 const {inspectTemplate,renderTemplate}=await import('../lib/documents.mjs');
 const bytes=await Packer.toArrayBuffer(new Document({sections:[{headers:{default:new Header({children:[new Paragraph('联系方式：{联系方式}')]})},children:[new Paragraph({children:[new TextRun({text:'{姓名}',bold:true,color:'AD1234',size:36}),new TextRun(' / {姓名}')]}),new Table({rows:[new TableRow({children:[new TableCell({children:[new Paragraph('{学号}')]})]})]})]}]}));
 const info=inspectTemplate(bytes);assert.deepEqual([...info.fields].sort(),['姓名','学号','联系方式'].sort());assert.equal(info.canMerge,false);
 const records=data.columnsToRecords([{field:'姓名',text:'张 & 李'},{field:'学号',text:'00123'},{field:'联系方式',text:'0123456789'}],info.fields).records;
 const rendered=new PizZip(renderTemplate(bytes,records[0]));const xml=rendered.file('word/document.xml').asText();
 assert.equal((xml.match(/张 &amp; 李/g)||[]).length,2);assert.ok(xml.includes('00123'));assert.ok(xml.includes('AD1234'));assert.ok(xml.includes('<w:b'));assert.ok(xml.includes('<w:tbl>'));
 const header=Object.keys(rendered.files).find(n=>/^word\/header\d+\.xml$/.test(n));assert.ok(rendered.file(header).asText().includes('0123456789'));
});
