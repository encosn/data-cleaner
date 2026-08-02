import * as XLSX from 'xlsx';
import fs from 'node:fs';
import { readGrid, unmergeCells, trimEmpty, detectHeaderRows, parseSheet } from '../lib/parse.js';
XLSX.set_fs(fs);
const M=(a,b,c,d)=>({s:{r:a,c:b},e:{r:c,c:d}});
const aoa=[['시도','구군',...Array.from({length:10},(_,i)=>'항목'+i)]]; const merges=[];
for(let g=0;g<4000;g++){const s=aoa.length;for(let k=0;k<5;k++)aoa.push([k===0?'시도'+g:null,'구'+g+'_'+k,...Array.from({length:10},(_,i)=>g*10+i)]);merges.push(M(s,0,aoa.length-1,0));}
const ws=XLSX.utils.aoa_to_sheet(aoa); ws['!merges']=merges;
let t=Date.now(); const rg=readGrid(ws); console.log('readGrid', Date.now()-t);
t=Date.now(); const un=unmergeCells(rg.grid,{merges:rg.merges}); console.log('unmerge', Date.now()-t);
t=Date.now(); const tr=trimEmpty(un.grid); console.log('trimEmpty', Date.now()-t);
t=Date.now(); detectHeaderRows(tr.rows,{merges:rg.merges}); console.log('detectHeader', Date.now()-t);
t=Date.now(); parseSheet(ws); console.log('전체', Date.now()-t);
