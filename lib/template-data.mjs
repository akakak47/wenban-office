export function splitColumn(text) {
 if(!text)return [];
 // A single terminal newline is the clipboard's row separator, not a record.
 return text.replace(/\r\n?/g,'\n').replace(/\n$/,'').split('\n');
}
export function columnsToRecords(columns,fields,pad=false) {
 const errors=[];const seen=new Set();
 const parsed=columns.map(c=>({field:c.field.trim(),values:splitColumn(c.text)}));
 for(const c of parsed){if(!c.field)errors.push('请填写每列的字段名称');else if(seen.has(c.field))errors.push(`字段「${c.field}」重复`);seen.add(c.field);if(/[{}\n\r]/.test(c.field))errors.push('字段名称不应包含花括号或换行');if(c.values.some(v=>v.includes('\t')))errors.push(`「${c.field}」粘贴了多列，请每个输入框只粘贴一列`);}
 for(const f of fields)if(!seen.has(f))errors.push(`缺少模板字段「${f}」的信息列`);
 const count=Math.max(0,...parsed.map(c=>c.values.length));
 if(count>500)errors.push('单次最多生成 500 条记录，请分批处理');
 if(!pad&&parsed.some(c=>c.values.length!==count))errors.push('各列行数不一致，请补齐数据或勾选缺失值留空');
 const records=Array.from({length:Math.min(count,500)},(_,i)=>Object.fromEntries(parsed.map(c=>[c.field,c.values[i]??''])));
 return {records,count,errors,counts:parsed.map(c=>c.values.length)};
}
