import ExcelJS from 'exceljs';

export type RecordItem = {
  id: string; employeeId: string; name: string; department: string; budgetDepartment: string;
  category: '보수교육' | '보수교육 외'; title: string; start: string; end: string;
  fee: number | null; feeText: string; receipt: string; certificate: string;
  phone: string; mobile: string; extension: string; matched: boolean;
  resolution: string; resolutionStatus: string; year: string; source: string; row: number;
};
export type Budget = { department: string; amount: number; source: string; separate: boolean };
export type Issue = { message: string; location: string };
export type Dataset = { records: RecordItem[]; budgets: Budget[]; issues: Issue[]; employeeCount: number; filename: string; loadedAt: string; demo: boolean };

function cellValue(cell: ExcelJS.Cell): unknown {
  const value = cell.value;
  if (value && typeof value === 'object' && !(value instanceof Date)) {
    if ('result' in value) return value.result;
    if ('richText' in value) return value.richText.map(x => x.text).join('');
    if ('error' in value) return '';
    if ('text' in value) return value.text;
    return '';
  }
  return value;
}
const txt = (value: unknown) => value == null ? '' : String(value).trim();
export const normalizeId = (value: unknown) => {
  const text = txt(value);
  return /^\d+$/.test(text) ? text.padStart(5, '0') : text;
};
export const isBlank = (value: string | null | undefined) => !value?.trim();
export const hasMissingDocument = (r: RecordItem) => isBlank(r.receipt) || isBlank(r.certificate);
export const missingType = (r: RecordItem) => isBlank(r.receipt) && isBlank(r.certificate) ? 'both' : isBlank(r.receipt) ? 'receipt' : isBlank(r.certificate) ? 'certificate' : 'none';
export const phoneOf = (r: RecordItem) => r.mobile || r.phone || (r.extension ? `내선 ${r.extension}` : '');
export const spentFor = (records: RecordItem[], dept: string) => records.filter(r => r.budgetDepartment === dept).reduce((sum, r) => sum + (r.fee ?? 0), 0);
export const categorySpent = (records: RecordItem[], dept: string, category: string) => records.filter(r => r.budgetDepartment === dept && r.category === category).reduce((sum, r) => sum + (r.fee ?? 0), 0);

function dateString(value: unknown) {
  if (value instanceof Date) return value.toISOString().slice(0, 10);
  const s = txt(value);
  if (/^\d{8}$/.test(s)) {
    const formatted = `${s.slice(0,4)}-${s.slice(4,6)}-${s.slice(6,8)}`;
    const d = new Date(`${formatted}T00:00:00Z`);
    if (!Number.isNaN(d.getTime()) && d.toISOString().slice(0,10) === formatted) return formatted;
  }
  if (/^\d{4}-\d{2}-\d{2}$/.test(s)) return s;
  return s;
}

export async function parseWorkbook(buffer: ArrayBuffer, filename: string): Promise<Dataset> {
  const workbook = new ExcelJS.Workbook();
  await workbook.xlsx.load(buffer);
  const issues: Issue[] = [];
  const findSheet = (name: string) => workbook.worksheets.find(s => s.name.replace(/[\s_]/g, '') === name.replace(/[\s_]/g, ''));
  const staffSheet = findSheet('전직원');
  const budgetSheet = findSheet('교육행정팀(부서별_위탁예산)');
  const educationSheets = [findSheet('보수교육_2026'), findSheet('보수교육외_2026')];
  if (!educationSheets.some(Boolean)) throw new Error('보수교육_2026 또는 보수교육외_2026 시트를 찾지 못했습니다. 교육 내역이 포함된 .xlsx 파일을 선택해 주세요.');
  const staff = new Map<string, { name: string; phone: string; mobile: string; extension: string }>();
  if (staffSheet) {
    const headers = new Map<string, number>();
    staffSheet.getRow(1).eachCell((c, col) => headers.set(txt(cellValue(c)), col));
    for (let i=2; i<=staffSheet.rowCount; i++) {
      const row = staffSheet.getRow(i);
      const get = (key: string) => headers.has(key) ? cellValue(row.getCell(headers.get(key)!)) : '';
      const id = normalizeId(get('사원번호'));
      if (!id) continue;
      if (staff.has(id)) issues.push({ message: `사원번호 ${id}가 직원 명부에 중복되어 있습니다.`, location: `${staffSheet.name}!A${i}` });
      else staff.set(id, { name: txt(get('사원명')), phone: txt(get('전화번호')), mobile: txt(get('휴대폰번호')), extension: txt(get('내선번호')) });
    }
  } else issues.push({ message: '전직원 시트가 없어 연락처를 연결하지 못했습니다.', location: '전직원' });
  const records: RecordItem[] = [];
  educationSheets.forEach((sheet, index) => {
    if (!sheet) { issues.push({ message: `${index===0 ? '보수교육' : '보수교육외'} 시트가 없습니다. 나머지 시트만 집계했습니다.`, location: '통합문서' }); return; }
    const headers = new Map<string, number>();
    sheet.getRow(1).eachCell((c, col) => headers.set(txt(cellValue(c)), col));
    const required = ['사번','이름','교육명','위탁예산부서','교육비','영수증','이수증'];
    const absent = required.filter(x => !headers.has(x));
    if (absent.length) throw new Error(`${sheet.name} 시트에 필수 열이 없습니다: ${absent.join(', ')}`);
    for (let i=2; i<=sheet.rowCount; i++) {
      const row = sheet.getRow(i);
      const get = (key: string) => headers.has(key) ? cellValue(row.getCell(headers.get(key)!)) : '';
      const employeeId = normalizeId(get('사번'));
      if (!employeeId) continue;
      const name = txt(get('이름'));
      const title = txt(get('교육명'));
      if (!name || !title) { issues.push({message:'사번은 있지만 이름 또는 교육명이 없어 제외한 행입니다.',location:`${sheet.name}!${i}행`}); continue; }
      const person = staff.get(employeeId);
      if (!person) issues.push({message:`직원 명부에 없는 사번 ${employeeId}: 연락처를 확인해 주세요.`,location:`${sheet.name}!C${i}`});
      else if (person.name && person.name !== name) issues.push({message:`사번 ${employeeId}의 교육 내역 이름과 명부 이름이 다릅니다.`,location:`${sheet.name}!D${i}`});
      const feeValue = get('교육비');
      const end = dateString(get('교육종료'));
      if (end && !/^\d{4}-\d{2}-\d{2}$/.test(end)) issues.push({message:`교육종료일 형식을 확인해 주세요: ${end}`,location:`${sheet.name}!${i}행`});
      const year = txt(get('예산연도'));
      if (year && year !== '2026') issues.push({message:`예산연도 확인: ${year}. 엑셀 호환 집계에는 포함했습니다.`,location:`${sheet.name}!B${i}`});
      records.push({id:`${sheet.name}:${i}`,employeeId,name,department:txt(get('부서')),budgetDepartment:txt(get('위탁예산부서')),category:index===0?'보수교육':'보수교육 외',title,start:dateString(get('교육시작')),end,fee:typeof feeValue==='number'&&Number.isFinite(feeValue)?feeValue:null,feeText:txt(feeValue),receipt:txt(get('영수증')),certificate:txt(get('이수증')),phone:person?.phone??'',mobile:person?.mobile??'',extension:person?.extension??'',matched:!!person,resolution:dateString(get('지출결의')),resolutionStatus:dateString(get('지출결의여부')),year,source:sheet.name,row:i});
    }
  });
  if (!records.length) throw new Error('읽을 수 있는 교육 기록이 없습니다. 사번·이름·교육명이 입력되어 있는지 확인해 주세요.');
  const budgets: Budget[] = [];
  if (budgetSheet) {
    let headerRow = 0, departmentColumn = 0, budgetColumn = 0, spentColumn = 0;
    for (let i=1;i<=Math.min(12,budgetSheet.rowCount);i++) {
      const row = budgetSheet.getRow(i);
      row.eachCell((c,col) => { if(txt(cellValue(c))==='확정예산2') { headerRow=i;budgetColumn=col;departmentColumn=col-2;spentColumn=col+1; } });
    }
    if (!headerRow) issues.push({message:'부서 예산 집계표의 확정예산2 열을 찾지 못했습니다.',location:budgetSheet.name});
    else {
      for (let i=headerRow+1;i<=budgetSheet.rowCount;i++) {
        const row = budgetSheet.getRow(i), department = txt(cellValue(row.getCell(departmentColumn))), amount = cellValue(row.getCell(budgetColumn));
        if (department==='합계') break;
        if (!department || typeof amount !== 'number' || !Number.isFinite(amount)) continue;
        budgets.push({department,amount,source:`${budgetSheet.name}!${row.getCell(budgetColumn).address}`,separate:department==='교육행정팀'});
        const cachedSpent = cellValue(row.getCell(spentColumn));
        if (typeof cachedSpent==='number' && Math.abs(cachedSpent-spentFor(records,department))>0.5) issues.push({message:`${department}: 원본 사용금액과 교육 내역 재합산 금액이 다릅니다.`,location:`${budgetSheet.name}!${row.getCell(spentColumn).address}`});
      }
    }
    if (budgets.some(b=>b.separate)) issues.push({message:'교육행정팀은 별도 예산입니다. 원본 O7과 아래 산출표 G270의 기준이 달라 위탁예산 합계에서 분리했습니다.',location:`${budgetSheet.name}!O7 / G270`});
  } else issues.push({message:'부서별 위탁예산 시트가 없어 확정예산과 잔액을 표시할 수 없습니다.',location:'교육행정팀(부서별_위탁예산)'});
  const known = new Set(budgets.map(b=>b.department));
  for (const dept of new Set(records.filter(r=>r.fee!==null).map(r=>r.budgetDepartment))) {
    if (!known.has(dept)) issues.push({message:`${dept||'부서 미입력'}: 숫자 교육비가 있지만 확정예산을 찾지 못했습니다.`,location:'부서 예산 연결'});
  }
  return {records,budgets,issues,employeeCount:staff.size,filename,loadedAt:new Date().toISOString(),demo:false};
}

export function csvText(headers: string[], rows: (string | number)[][]) {
  const escape = (value: string | number) => {
    let text = String(value);
    if (/^[=+\-@\t\r]/.test(text)) text = `'${text}`;
    return `"${text.replace(/"/g,'""')}"`;
  };
  return '\uFEFF' + [headers,...rows].map(row=>row.map(escape).join(',')).join('\r\n');
}

export function demoDataset(): Dataset {
  const budgets: Budget[] = [
    {department:'간호본부',amount:62000000,source:'예시',separate:false},
    {department:'영상의학과',amount:6368000,source:'예시',separate:false},
    {department:'QPS팀',amount:6511000,source:'예시',separate:false},
    {department:'약제팀',amount:6710000,source:'예시',separate:false},
    {department:'진단검사의학과',amount:2940000,source:'예시',separate:false},
    {department:'교육행정팀',amount:120168000,source:'예시',separate:true},
  ];
  const names = ['김하늘','이서윤','박지우','최도윤','정다은','윤서준','한지민','오수빈','임하준','강서현','신예린','송민준'];
  const titles = ['환자안전과 간호 실무 보수교육','의료영상 최신 기술 연수교육','환자안전 전담자 역량 강화 교육','임상약학 실무 역량 교육','진단검사 최신 동향 학술대회','직원 교육 운영 실무 워크숍'];
  const records = Array.from({length:30},(_,i): RecordItem => ({
    id:`demo:${i}`,employeeId:`D${String(i%12+1).padStart(4,'0')}`,name:names[i%12],department:i%6===0?'142병동':budgets[i%6].department,budgetDepartment:budgets[i%6].department,
    category:i%3===0?'보수교육':'보수교육 외',title:titles[i%6],start:`2026-10-${String(i%20+1).padStart(2,'0')}`,end:`2026-10-${String(i%20+1).padStart(2,'0')}`,
    fee:[40000,180000,120000,90000,250000,60000][i%6],feeText:String([40000,180000,120000,90000,250000,60000][i%6]),
    receipt:i%5===0||i%5===2?'':'O',certificate:i%5===0||i%5===1?'':'O',phone:'',mobile:`010-0000-${String(i%12+1).padStart(4,'0')}`,extension:'',matched:true,resolution:'2026-10-31',resolutionStatus:i%5>=3?'2026-10-20':'',year:'2026',source:'가상 예시 데이터',row:i+2,
  }));
  return {records,budgets,issues:[],employeeCount:12,filename:'샘플 데이터 · 실제 직원 정보가 아닙니다',loadedAt:new Date().toISOString(),demo:true};
}
