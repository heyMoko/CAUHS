import { test } from 'node:test';
import assert from 'node:assert/strict';
import { readFile } from 'node:fs/promises';
import ExcelJS from 'exceljs';
import { csvText, demoDataset, hasMissingDocument, normalizeId, parseWorkbook, spentFor } from '../src/data';

test('서류 빈칸 규칙은 금액·예정일과 독립적이며 예외 원문을 보존한다',()=>{
  const base=demoDataset().records[0];
  assert.equal(hasMissingDocument({...base,receipt:' ',certificate:'O',fee:null,end:'2027-03-31'}),true);
  assert.equal(hasMissingDocument({...base,receipt:'△',certificate:''}),true);
  assert.equal(hasMissingDocument({...base,receipt:'해당없음',certificate:'O'}),false);
  assert.equal(hasMissingDocument({...base,receipt:'X',certificate:'X'}),false);
  assert.equal(normalizeId(700),'00700');assert.equal(normalizeId('R0221'),'R0221');
});
test('CSV는 따옴표·개행을 보존하고 수식 실행 문자열을 방어한다',()=>{
  const csv=csvText(['이름','교육명'],[['=1+1','교육 "A"\n다음 줄']]);
  assert.ok(csv.startsWith('\uFEFF'));
  assert.ok(csv.includes('"\'=1+1"'));assert.ok(csv.includes('"교육 ""A""\n다음 줄"'));
});
test('누락 시트와 잘못된 양식을 설명하며 원본 수식 준비행은 교육으로 세지 않는다',async()=>{
  const workbook=new ExcelJS.Workbook();
  const sheet=workbook.addWorksheet('보수교육 2026');
  sheet.addRow(['번호','사번','이름','교육명','위탁예산부서','교육비','영수증','이수증']);
  sheet.addRow([1,700,'가상직원','교육','간호본부',40000,'O',null]);
  sheet.addRow([2,null,{formula:'NA()',result:{error:'#N/A'}},null,null,null,null,null]);
  const buffer=await workbook.xlsx.writeBuffer();
  const a=await parseWorkbook(buffer as unknown as ArrayBuffer,'test.xlsx');
  const b=await parseWorkbook(buffer as unknown as ArrayBuffer,'test.xlsx');
  assert.equal(a.records.length,1);assert.equal(a.records[0].employeeId,'00700');
  assert.equal(a.records[0].matched,false);assert.equal(a.records[0].certificate,'');
  assert.equal(b.records.length,a.records.length);assert.equal(a.budgets.length,0);
  assert.ok(a.issues.some(i=>i.message.includes('연락처')));
  const invalid=new ExcelJS.Workbook();invalid.addWorksheet('Sheet1');
  const invalidBuffer = await invalid.xlsx.writeBuffer();
  await assert.rejects(()=>parseWorkbook(invalidBuffer as unknown as ArrayBuffer,'bad.xlsx'),/시트를 찾지 못했습니다/);
});
test('실제 원본의 교육·예산·서류 빈칸을 독립 기준값과 대조한다',{skip:!process.env.CAUHS_TEST_XLSX},async()=>{
  const buffer=await readFile(process.env.CAUHS_TEST_XLSX!);
  const data=await parseWorkbook(buffer.buffer.slice(buffer.byteOffset,buffer.byteOffset+buffer.byteLength) as ArrayBuffer,'원본.xlsx');
  assert.equal(data.records.length,1834);
  assert.equal(data.employeeCount,1983);assert.equal(data.budgets.length,39);
  assert.equal(data.records.reduce((s,r)=>s+(r.fee??0),0),80018102);
  assert.equal(data.budgets.filter(b=>!b.separate).reduce((s,b)=>s+b.amount,0),154000000);
  assert.equal(data.budgets.filter(b=>!b.separate).reduce((s,b)=>s+spentFor(data.records,b.department),0),75158102);
  assert.equal(data.budgets.find(b=>b.department==='연구행정관리팀')?.amount,350000);
  assert.equal(data.budgets.find(b=>b.department==='연구윤리팀')?.amount,50000);
  assert.equal(data.records.filter(hasMissingDocument).length,100);
  assert.equal(data.records.filter(r=>!r.matched).length,49);
  assert.equal(data.issues.filter(i=>i.message.includes('원본 사용금액')).length,0);
});
