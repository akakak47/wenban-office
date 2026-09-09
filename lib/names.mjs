export function parseNames(text, dedupe=false) {
 const names=String(text).split(/[\r\n\t,，;；]+/).map(s=>s.trim()).filter(Boolean);
 return dedupe?[...new Set(names)]:names;
}
export function makeGrid(names,columns=4,order='row',fixedRows=0) {
 if(!Number.isInteger(columns)||columns<1||columns>50)throw new Error('列数应为 1–50 的整数');
 const rows=Math.max(fixedRows,Math.ceil(names.length/columns));
 return Array.from({length:rows},(_,r)=>Array.from({length:columns},(_,c)=>names[order==='column'?c*rows+r:r*columns+c]??''));
}
export const toTsv=grid=>grid.map(row=>row.join('\t')).join('\n');
export function extractRows(matrix,headerRow=0) {
 const source=matrix[headerRow]||[];const used=new Set();
 const headers=Array.from({length:source.length},(_,i)=>{const base=String(source[i]??'').trim()||`第${i+1}列`;let h=base,n=2;while(used.has(h))h=base+'_'+n++;used.add(h);return h;});
 const rows=matrix.slice(headerRow+1).filter(r=>r.some(v=>String(v??'').trim())).map(r=>Object.fromEntries(headers.map((h,i)=>[h,String(r[i]??'').trim()])));
 return {headers,rows};
}
export function guessName(headers) {return headers.find(h=>/^(姓名|学生姓名|嘉宾姓名|名字|name|full.?name)$/i.test(h))||headers.find(h=>/姓名|名字|name/i.test(h))||headers[0]||'';}
