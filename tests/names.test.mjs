import test from 'node:test';
import assert from 'node:assert/strict';
import {parseNames,makeGrid,toTsv,extractRows} from '../lib/names.mjs';
test('keeps duplicates and names with spaces, discards empty lines',()=>assert.deepEqual(parseNames('张三\r\n李四\n\n张三\nMary Jane'),['张三','李四','张三','Mary Jane']));
test('pads incomplete last row and preserves TSV blanks',()=>assert.equal(toTsv(makeGrid(['一','二','三','四'],3,'row')),'一\t二\t三\n四\t\t'));
test('column-first order is deterministic',()=>assert.deepEqual(makeGrid(['一','二','三','四','五'],3,'column'),[['一','三','五'],['二','四','']]));
test('sheet extraction handles duplicated headers and blank records',()=>assert.deepEqual(extractRows([['姓名','姓名'],['甲','乙'],['','']]),{headers:['姓名','姓名_2'],rows:[{'姓名':'甲','姓名_2':'乙'}]}));
test('fixed row count is preserved when names are fewer than rows',()=>assert.deepEqual(makeGrid(['甲','乙'],1,'row',4),[['甲'],['乙'],[''],['']]));
