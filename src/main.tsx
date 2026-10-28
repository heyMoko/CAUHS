import React, { useEffect, useMemo, useRef, useState } from 'react';
import { createRoot } from 'react-dom/client';
import { ArrowDown, ArrowDownToLine, ArrowUp, ArrowUpDown, ArrowUpRight, Building2, Check, ChevronLeft, ChevronRight, CircleHelp, ClipboardCheck, FileCheck2, FileSpreadsheet, FolderOpen, LayoutDashboard, Search, ShieldCheck, Upload, Users, Wallet, X, AlertCircle, LoaderCircle, RotateCcw, Menu } from 'lucide-react';
import { categorySpent, csvText, demoDataset, hasMissingDocument, isBlank, missingType, parseWorkbook, phoneOf, spentFor, type Dataset, type RecordItem } from './data';
import { clearLastUpload, readLastUpload, saveLastUpload } from './persistence';
import './styles.css';
import cauhsLogo from './cauhs-logo.png';

type View = 'overview' | 'documents' | 'budgets' | 'upload';
const money = (v: number) => v.toLocaleString('ko-KR');
const collator = new Intl.Collator('ko', {numeric:true});
function App() {
  const [data,setData] = useState<Dataset>(demoDataset);
  const [savedLocally,setSavedLocally] = useState(false);
  const [view,setView] = useState<View>('overview');
  const [search,setSearch] = useState('');
  const [department,setDepartment] = useState('all');
  const [category,setCategory] = useState('all');
  const [docFilter,setDocFilter] = useState('missing');
  const [sort,setSort] = useState({key:'name',asc:true});
  const [page,setPage] = useState(1);
  const [busy,setBusy] = useState(false);
  const [error,setError] = useState('');
  const [toast,setToast] = useState('');
  const [modal,setModal] = useState<'help'|'issues'|null>(null);
  const [detail,setDetail] = useState<RecordItem|null>(null);
  const [drag,setDrag] = useState(false);
  const [mobileMenu,setMobileMenu] = useState(false);
  const [budgetSearch,setBudgetSearch] = useState('');
  const [includeSeparate,setIncludeSeparate] = useState(false);
  const [budgetSort,setBudgetSort] = useState({key:'department',asc:true});
  const input = useRef<HTMLInputElement>(null);
  useEffect(()=>{
    let active=true;
    void readLastUpload().then(saved=>{
      if(active&&saved&&!saved.demo){setData(saved);setSavedLocally(true);}
    }).catch(()=>{
      if(active)notify('이 브라우저에 저장된 엑셀 데이터를 불러오지 못했습니다.');
    });
    return ()=>{active=false;};
  },[]);
  const notify = (text: string) => { setToast(text); window.setTimeout(()=>setToast(''),3500); };
  const navigate = (next: View) => {setView(next);setMobileMenu(false);setPage(1);};
  const resetFilters = () => {setSearch('');setDepartment('all');setCategory('all');setDocFilter('missing');setPage(1);};
  const missing = data.records.filter(hasMissingDocument);
  const both = missing.filter(r=>missingType(r)==='both');
  const onlyReceipt = missing.filter(r=>missingType(r)==='receipt');
  const onlyCertificate = missing.filter(r=>missingType(r)==='certificate');
  const people = new Set(missing.map(r=>r.employeeId)).size;
  const departments = Array.from(new Set(data.records.map(r=>r.budgetDepartment))).filter(Boolean).sort(collator.compare);
  const filtered = useMemo(()=>{
    const q = search.trim().toLowerCase(), digits=q.replace(/[-\s]/g,'');
    return data.records.filter(r=>{
      if (department!=='all'&&r.budgetDepartment!==department) return false;
      if (category!=='all'&&r.category!==category) return false;
      if (docFilter==='missing'&&!hasMissingDocument(r)) return false;
      if (['both','receipt','certificate'].includes(docFilter)&&missingType(r)!==docFilter) return false;
      if (docFilter==='complete'&&!(r.receipt.toUpperCase()==='O'&&r.certificate.toUpperCase()==='O')) return false;
      if (q && ![r.name,r.employeeId,r.title,r.department,r.budgetDepartment,phoneOf(r)].some(v=>v.toLowerCase().includes(q)) && !(digits&&phoneOf(r).replace(/[-\s]/g,'').includes(digits))) return false;
      return true;
    }).sort((a,b)=>{
      const value = (r:RecordItem) => sort.key==='phone'?phoneOf(r):sort.key==='fee'?r.fee??-1:r[sort.key as keyof RecordItem];
      const av=value(a),bv=value(b);
      const compare = typeof av==='number'&&typeof bv==='number'?av-bv:collator.compare(String(av??''),String(bv??''));
      return (sort.asc?compare:-compare)||collator.compare(a.id,b.id);
    });
  },[data,search,department,category,docFilter,sort]);
  const maxPage = Math.max(1,Math.ceil(filtered.length/12));
  const currentPage = Math.min(page,maxPage);
  const pageRows = filtered.slice((currentPage-1)*12,currentPage*12);
  const budgetRows = useMemo(()=>{
    const known = new Set(data.budgets.map(b=>b.department));
    const extras = Array.from(new Set(data.records.filter(r=>r.fee!==null).map(r=>r.budgetDepartment))).filter(d=>!known.has(d)).map(d=>({department:d||'부서 미입력',amount:null,source:'미설정',separate:false}));
    return [...data.budgets,...extras].filter(b=>(includeSeparate||!b.separate)&&b.department.includes(budgetSearch)).map(b=>({...b,spent:spentFor(data.records,b.department),missing:data.records.filter(r=>r.budgetDepartment===b.department&&hasMissingDocument(r)).length})).sort((a,b)=>{
      const val=(r:typeof a)=>budgetSort.key==='balance'?(r.amount??0)-r.spent:budgetSort.key==='amount'?r.amount??-1:budgetSort.key==='spent'?r.spent:r.department;
      const av=val(a),bv=val(b);return (typeof av==='number'&&typeof bv==='number'?av-bv:collator.compare(String(av),String(bv)))*(budgetSort.asc?1:-1);
    });
  },[data,includeSeparate,budgetSearch,budgetSort]);
  const scopedBudgets = data.budgets.filter(b=>includeSeparate||!b.separate);
  const budgetTotal = scopedBudgets.reduce((s,b)=>s+b.amount,0);
  const spentTotal = scopedBudgets.reduce((s,b)=>s+spentFor(data.records,b.department),0);
  const utilization = budgetTotal ? spentTotal/budgetTotal*100 : 0;
  const missingPhone = missing.filter(r=>!phoneOf(r)).length;
  async function load(file?: File) {
    if(!file||busy) return;
    setError('');
    if(!file.name.toLowerCase().endsWith('.xlsx')) {setError('.xlsx 형식의 엑셀 파일을 선택해 주세요.');return;}
    if(file.size>30*1024*1024) {setError('30MB 이하의 엑셀 파일을 선택해 주세요.');return;}
    setBusy(true);
    try {
      const next=await parseWorkbook(await file.arrayBuffer(),file.name);
      setData(next);resetFilters();setIncludeSeparate(false);setBudgetSearch('');setView('overview');
      try {
        await saveLastUpload(next);
        setSavedLocally(true);
        notify(`교육 ${money(next.records.length)}건을 불러와 이 브라우저에 저장했습니다.`);
      } catch {
        setSavedLocally(false);
        notify('엑셀은 불러왔지만 브라우저 저장에 실패했습니다. 브라우저 저장공간을 확인해 주세요.');
      }
    } catch(e) {setError(e instanceof Error?e.message:'파일을 읽지 못했습니다. 올바른 엑셀 파일인지 확인해 주세요.');}
    finally {setBusy(false);if(input.current)input.current.value='';}
  }
  async function forgetUpload() {
    if(!window.confirm('이 브라우저에 저장된 엑셀 데이터를 삭제할까요?')) return;
    try {
      await clearLastUpload();
      setData(demoDataset);setSavedLocally(false);resetFilters();setView('overview');
      notify('저장된 엑셀 데이터를 삭제했습니다.');
    } catch {
      notify('저장된 데이터를 삭제하지 못했습니다.');
    }
  }
  function download(kind:'documents'|'budgets') {
    const text=kind==='documents'?csvText(['이름','사번','교육명','전화번호','근무부서','위탁예산부서','교육구분','교육종료','교육비/처리','영수증','이수증','원본위치'],filtered.map(r=>[r.name,r.employeeId,r.title,phoneOf(r)||'연락처 없음',r.department,r.budgetDepartment,r.category,r.end,r.fee??r.feeText,r.receipt||'누락',r.certificate||'누락',`${r.source}!${r.row}행`])):csvText(['예산부서','확정예산(원)','등록교육비(원)','잔액(원)','증빙서류누락건수'],budgetRows.map(b=>[b.department,b.amount??'미설정',b.spent,b.amount===null?'미설정':b.amount-b.spent,b.missing]));
    const url=URL.createObjectURL(new Blob([text],{type:'text/csv;charset=utf-8;'}));const a=document.createElement('a');a.href=url;a.download=`CAUHS_${kind==='documents'?'서류확인':'부서예산'}_${new Date().toISOString().slice(0,10)}.csv`;a.click();URL.revokeObjectURL(url);notify('현재 조회 결과를 내려받았습니다.');
  }
  function showDocs(filter='missing',dept='all') {setDocFilter(filter);setDepartment(dept);setSearch('');setCategory('all');navigate('documents');}
  function column(label:string,key:string) {
    return <button className="sort-button" onClick={()=>{setSort({key,asc:sort.key===key?!sort.asc:true});setPage(1);}}>{label}{sort.key===key?(sort.asc?<ArrowUp size={13}/>:<ArrowDown size={13}/>):<ArrowUpDown size={12}/>}</button>;
  }
  function documentStatus(value:string,kind:string) {
    if(isBlank(value))return <span className="status missing">{kind} 누락</span>;
    if(value.toUpperCase()==='O')return <span className="status complete"><Check size={11}/>제출</span>;
    return <span className="status other" title={value}>{value}</span>;
  }
  const titles:Record<View,string>={overview:'전체 현황',documents:'증빙서류 누락현황',budgets:'부서별 예산',upload:'엑셀 업로드'};
  return <div className="app-shell">
    <input ref={input} type="file" accept=".xlsx" className="file-input" aria-label="교육 예산 엑셀 파일 선택" onChange={e=>void load(e.target.files?.[0])}/>

    {mobileMenu&&<button className="sidebar-scrim" aria-label="메뉴 닫기" onClick={()=>setMobileMenu(false)}/>}
    <header className="site-header">
      <a className="site-brand" href="#" onClick={e=>{e.preventDefault();navigate('overview');}}><img className="brand-logo" src={cauhsLogo} alt="CAUHS"/><span>CAUHS<small>직원출장 및 지출결의 내역</small></span></a>
      <nav className="main-nav" aria-label="주 메뉴">
        <button className={view==='overview'?'active':''} onClick={()=>navigate('overview')}><LayoutDashboard size={16}/>전체 현황</button>
        <button className={view==='documents'?'active':''} onClick={()=>navigate('documents')}><FileCheck2 size={16}/>증빙서류 누락현황<span className="nav-count">{missing.length}</span></button>
        <button className={view==='budgets'?'active':''} onClick={()=>navigate('budgets')}><Wallet size={16}/>부서별 예산</button>
        <button className={view==='upload'?'active':''} onClick={()=>navigate('upload')}><Upload size={16}/>엑셀 업로드</button>
      </nav>
    </header>    <main>
      <header className="topbar"><div className="breadcrumb"><button className="mobile-toggle" onClick={()=>setMobileMenu(!mobileMenu)} aria-label="메뉴 열기"><Menu size={20}/></button><span>직원출장 및 지출결의 내역</span><ChevronRight size={13}/><strong>{titles[view]}</strong></div></header>
      <div className="content">
        <div className="page-heading"><div><div className="eyebrow">출장 · 지출결의 관리</div><h1>{view==='overview'?'직원출장 및 지출결의 내역':titles[view]}</h1><p>{view==='overview'?'부서별 확정예산과 교육비 사용 현황, 환급 서류 제출 여부를 한눈에 확인합니다.':view==='documents'?'영수증·이수증이 비어 있는 교육과 직원 연락처를 확인하세요.':view==='budgets'?'예산부서를 기준으로 확정예산과 등록 교육비, 잔액을 확인하세요.':'관리 중인 엑셀을 올리면 예산과 서류 현황을 바로 정리합니다.'}</p></div><button className="primary" onClick={()=>input.current?.click()} disabled={busy}>{busy?<LoaderCircle size={16} className="spin"/>:<Upload size={16}/>} {busy?'엑셀 읽는 중…':'엑셀 업로드'}</button></div>
        {error&&<div className="error-banner" role="alert"><AlertCircle size={18}/><span>{error}</span><button onClick={()=>setError('')} aria-label="오류 닫기"><X size={16}/></button></div>}
        <div className={`source-bar ${data.demo?'demo':''}`}><span className="source-icon"><FileSpreadsheet size={17}/></span><span className="source-filename">{data.filename}</span><span className="source-tag">{data.demo?'체험용 예시':savedLocally?'이 브라우저에 저장됨':'업로드 완료 · 저장 안 됨'}</span><span className="source-meta">{data.demo?'엑셀을 올려 실제 현황을 확인하세요':`${new Date(data.loadedAt).toLocaleString('ko-KR')} 반영`}</span>{!data.demo&&<><button className="text-link" onClick={()=>setModal('issues')}>검토 {data.issues.length}건<ChevronRight size={14}/></button><button className="text-link" onClick={()=>void forgetUpload()}>저장 데이터 삭제</button></>}</div>

        {view==='overview'&&<>
          <section className="summary-grid" aria-label="요약 지표">
            <div className="metric"><div className="metric-top"><span>위탁 확정예산</span><span className="metric-icon mint"><Wallet size={18}/></span></div><strong>{data.budgets.length?money(data.budgets.filter(b=>!b.separate).reduce((s,b)=>s+b.amount,0)):'—'}<small>원</small></strong><p>{data.budgets.filter(b=>!b.separate).length}개 예산부서 기준</p></div>
            <div className="metric"><div className="metric-top"><span>등록 교육비</span><span className="metric-icon blue"><ClipboardCheck size={18}/></span></div><strong>{money(data.budgets.filter(b=>!b.separate).reduce((s,b)=>s+spentFor(data.records,b.department),0))}<small>원</small></strong><p>엑셀 기준 교육비 합계 <span className="subtle">· 실제 지급액과 다름</span></p></div>
            <div className="metric"><div className="metric-top"><span>위탁예산 잔액</span><span className="metric-icon violet"><Building2 size={18}/></span></div><strong>{data.budgets.length?money(data.budgets.filter(b=>!b.separate).reduce((s,b)=>s+b.amount-spentFor(data.records,b.department),0)):'—'}<small>원</small></strong><p>확정예산 − 등록 교육비</p></div>
            <button className="metric action-metric" onClick={()=>showDocs()}><div className="metric-top"><span>영수증·이수증 누락</span><span className="metric-icon orange"><FileCheck2 size={18}/></span></div><strong>{money(missing.length)}<small>건</small></strong><p>확인할 직원 <b>{money(people)}명</b><ArrowUpRight size={15}/></p></button>
          </section>
          <div className="overview-grid">
            <section className="panel budget-panel"><div className="panel-heading"><div><h2>부서별 예산 현황</h2><p>주요 부서의 등록 교육비와 잔액</p></div><button className="text-link" onClick={()=>navigate('budgets')}>전체 보기<ChevronRight size={14}/></button></div><div className="chart-legend"><span><i className="legend-dot used"/>등록 교육비</span><span><i className="legend-dot remaining"/>잔액</span><span className="chart-unit">단위: 원</span></div><div className="budget-bars">{data.budgets.filter(b=>!b.separate).sort((a,b)=>b.amount-a.amount).slice(0,5).map(b=>{const spent=spentFor(data.records,b.department);const percent=b.amount?spent/b.amount*100:0;return <button className="budget-bar-row" key={b.department} onClick={()=>showDocs('all',b.department)}><div className="bar-label"><strong>{b.department}</strong><span>{money(spent)} <small>/ {money(b.amount)}</small></span></div><div className="bar-track"><span style={{width:`${Math.min(100,Math.max(0,percent))}%`}}/></div><span className={`bar-percent ${percent>100?'over':''}`}>{percent.toFixed(1)}%</span></button>;})}{!data.budgets.length&&<div className="small-empty">예산 시트가 없어 부서별 차트를 표시할 수 없습니다.</div>}</div><div className="panel-footnote"><span className="tiny-dot"/>교육행정팀 별도 예산은 위탁 합계에서 제외합니다.</div></section>
            <section className="panel docs-panel"><div className="panel-heading"><div><h2>영수증·이수증 누락 현황</h2><p>영수증 또는 이수증이 제출되지 않은 교육을 종류별로 확인하세요</p></div><span className="soft-icon"><FileCheck2 size={19}/></span></div><div className="donut-area"><div className="donut" style={{background:missing.length?`conic-gradient(#d77a4c 0 ${both.length/missing.length*100}%, #e9ad6e ${both.length/missing.length*100}% ${(both.length+onlyCertificate.length)/missing.length*100}%, #8aaba0 ${(both.length+onlyCertificate.length)/missing.length*100}% 100%)`:'#e9edea'}}><div><small>증빙 누락</small><strong>{missing.length}<span>건</span></strong></div></div><div className="donut-caption"><Users size={14}/><b>{people}명</b>의 직원에게 해당해요</div></div><div className="doc-types">{[['both','영수증·이수증 모두 누락',both.length,'#d77a4c'],['certificate','이수증 누락',onlyCertificate.length,'#e9ad6e'],['receipt','영수증 누락',onlyReceipt.length,'#8aaba0']].map(([key,label,count,color])=><button key={key} onClick={()=>showDocs(String(key))}><span><i style={{background:String(color)}}/>{label}</span><strong>{count}<small>건</small><ChevronRight size={14}/></strong></button>)}</div></section>
          </div>
          <section className="panel"><div className="panel-heading"><div><h2>영수증·이수증 누락 교육 목록 <span className="count-chip">{missing.length}</span></h2><p>영수증 또는 이수증이 제출되지 않은 교육입니다.</p></div><button className="text-link" onClick={()=>showDocs()}>전체 확인하기<ArrowUpRight size={15}/></button></div><div className="table-scroll"><table className="preview-table"><thead><tr><th>이름 / 사번</th><th>교육명</th><th>전화번호</th><th>영수증</th><th>이수증</th><th/></tr></thead><tbody>{missing.slice(0,5).map(r=><tr key={r.id}><td><div className="person"><span className="person-avatar">{r.name.slice(0,1)}</span><div><strong>{r.name}</strong><small>{r.employeeId}</small></div></div></td><td><div className="course-title">{r.title}</div><small className="secondary-text">{r.department} · {r.category}</small></td><td className="phone-cell">{phoneOf(r)||<span className="no-phone">연락처 없음</span>}</td><td>{documentStatus(r.receipt,'영수증')}</td><td>{documentStatus(r.certificate,'이수증')}</td><td><button className="icon-button" aria-label={`${r.name} 교육 상세`} onClick={()=>setDetail(r)}><ChevronRight size={17}/></button></td></tr>)}</tbody></table>{!missing.length&&<div className="empty-state"><Check size={30}/><h3>누락된 영수증·이수증이 없습니다</h3><p>현재 파일에서 영수증·이수증 누락 항목을 찾지 못했습니다.</p></div>}</div><div className="table-bottom"><span>교육 {money(data.records.length)}건 중 누락 교육 {money(missing.length)}건</span><button className="text-link" onClick={()=>showDocs()}>목록에서 검색·정렬하기<ChevronRight size={14}/></button></div></section>
        </>}

        {view==='documents'&&<>
          <div className="document-summary"><div><span className="metric-icon orange"><FileCheck2 size={20}/></span><p>확인할 교육<strong>{missing.length}<small>건</small></strong></p></div><div><span className="metric-icon mint"><Users size={20}/></span><p>해당 직원<strong>{people}<small>명</small></strong></p></div><div><span className="metric-icon blue"><FolderOpen size={20}/></span><p>전체 교육<strong>{money(data.records.length)}<small>건</small></strong></p></div><p className="document-summary-note">교육 예정·금액·결의 상태와 관계없이<br/>서류 칸이 비어 있는 교육을 표시합니다.</p></div>
          <section className="panel document-list"><div className="document-tabs" role="group" aria-label="서류 상태 필터">{[['missing','누락 전체',missing.length],['both','영수증·이수증 모두 누락',both.length],['receipt','영수증만',onlyReceipt.length],['certificate','이수증만',onlyCertificate.length],['complete','모두 제출',data.records.filter(r=>r.receipt.toUpperCase()==='O'&&r.certificate.toUpperCase()==='O').length],['all','전체 교육',data.records.length]].map(([key,label,count])=><button key={key} aria-pressed={docFilter===key} className={docFilter===key?'selected':''} onClick={()=>{setDocFilter(String(key));setPage(1);}}>{label}<span>{count}</span></button>)}</div>
          <div className="filter-row"><label className="search-field"><Search size={17}/><input placeholder="이름, 사번, 교육명, 전화번호 검색" aria-label="이름 사번 교육명 전화번호 검색" value={search} onChange={e=>{setSearch(e.target.value);setPage(1);}}/>{search&&<button aria-label="검색 지우기" onClick={()=>{setSearch('');setPage(1);}}><X size={14}/></button>}</label><select aria-label="예산부서 필터" value={department} onChange={e=>{setDepartment(e.target.value);setPage(1);}}><option value="all">모든 예산부서</option>{departments.map(d=><option key={d}>{d}</option>)}</select><select aria-label="교육구분 필터" value={category} onChange={e=>{setCategory(e.target.value);setPage(1);}}><option value="all">모든 교육구분</option><option>보수교육</option><option>보수교육 외</option></select><button className="icon-button reset" title="필터 초기화" aria-label="필터 초기화" onClick={resetFilters}><RotateCcw size={16}/></button><button className="secondary export" onClick={()=>download('documents')}><ArrowDownToLine size={15}/>CSV 내려받기</button></div>
          <div className="results-line"><span>조회 결과 <strong>{money(filtered.length)}건</strong><span className="subtle"> · 직원 {new Set(filtered.map(r=>r.employeeId)).size}명</span></span><span>열 제목을 눌러 정렬할 수 있어요</span></div>
          <div className="table-scroll"><table className="records-table"><thead><tr><th>{column('이름','name')}</th><th>{column('사번','employeeId')}</th><th>{column('교육명','title')}</th><th>{column('전화번호','phone')}</th><th>영수증</th><th>이수증</th><th>{column('교육종료','end')}</th><th/></tr></thead><tbody>{pageRows.map(r=><tr key={r.id}><td><strong>{r.name}</strong><small className="secondary-text block">{r.department}</small></td><td><span className="employee-id">{r.employeeId}</span></td><td><button className="course-button" onClick={()=>setDetail(r)}>{r.title}</button><small className="secondary-text block">{r.category} · {r.fee===null?r.feeText||'금액 미입력':`${money(r.fee)}원`}</small></td><td className="phone-cell">{phoneOf(r)?<a href={`tel:${phoneOf(r).replace(/[^\d+]/g,'')}`}>{phoneOf(r)}</a>:<span className="no-phone">연락처 없음</span>}</td><td>{documentStatus(r.receipt,'영수증')}</td><td>{documentStatus(r.certificate,'이수증')}</td><td className="date-cell">{r.end||'—'}</td><td><button className="icon-button" aria-label={`${r.name} 교육 상세`} onClick={()=>setDetail(r)}><ChevronRight size={16}/></button></td></tr>)}</tbody></table></div>
          {!filtered.length&&<div className="empty-state"><Search size={30}/><h3>조건에 맞는 교육이 없습니다</h3><p>검색어나 필터를 변경해 보세요.</p><button className="secondary" onClick={resetFilters}>필터 초기화</button></div>}
          <div className="table-bottom"><span>{filtered.length?`${(currentPage-1)*12+1}–${Math.min(currentPage*12,filtered.length)} / ${money(filtered.length)}건`:'0건'}</span><div className="pagination"><button disabled={currentPage===1} onClick={()=>setPage(currentPage-1)} aria-label="이전 페이지"><ChevronLeft size={16}/></button><span><b>{currentPage}</b> / {maxPage}</span><button disabled={currentPage===maxPage} onClick={()=>setPage(currentPage+1)} aria-label="다음 페이지"><ChevronRight size={16}/></button></div></div></section>
          <div className="inline-note"><CircleHelp size={15}/><span>엑셀에서 공란은 제출 여부가 기록되지 않은 상태입니다. ‘해당없음’, ‘예산 이관’, ‘△’ 등 원문 값은 누락으로 처리하지 않습니다.{missingPhone>0&&` 연락처가 없는 누락 교육은 ${missingPhone}건입니다.`}</span></div>
        </>}

        {view==='budgets'&&<>
          <div className="budget-intro"><div><small>{includeSeparate?'교육행정팀 포함 · 참고 합계':'위탁예산 합계'}</small><strong>{data.budgets.length?money(budgetTotal):'—'}<span>원</span></strong></div><div><small>등록 교육비</small><strong>{money(spentTotal)}<span>원</span></strong></div><div><small>잔액</small><strong className="green-text">{data.budgets.length?money(budgetTotal-spentTotal):'—'}<span>원</span></strong></div><div><small>예산 소진율</small><strong>{utilization.toFixed(1)}<span>%</span></strong></div></div>
          <section className="panel"><div className="panel-heading"><div><h2>예산부서별 상세 <span className="count-chip">{budgetRows.length}</span></h2><p>숫자로 입력된 교육비를 원본과 같은 방식으로 합산합니다.</p></div><button className="secondary" onClick={()=>download('budgets')}><ArrowDownToLine size={15}/>CSV 내려받기</button></div><div className="filter-row budget-filter"><label className="search-field"><Search size={17}/><input placeholder="예산부서 검색" aria-label="예산부서 검색" value={budgetSearch} onChange={e=>setBudgetSearch(e.target.value)}/></label><label className="checkbox"><input type="checkbox" checked={includeSeparate} onChange={e=>setIncludeSeparate(e.target.checked)}/>교육행정팀 별도 예산 포함</label></div><div className="table-scroll"><table className="budget-table"><thead><tr>{[['예산부서','department'],['확정예산','amount'],['등록 교육비','spent'],['잔액','balance']].map(([label,key])=><th key={key}><button className="sort-button" onClick={()=>setBudgetSort({key,asc:budgetSort.key===key?!budgetSort.asc:true})}>{label}{budgetSort.key===key?(budgetSort.asc?<ArrowUp size={12}/>:<ArrowDown size={12}/>):<ArrowUpDown size={12}/>}</button></th>)}<th>소진율</th><th>증빙 누락</th><th/></tr></thead><tbody>{budgetRows.map(b=>{const pct=b.amount?b.spent/b.amount*100:null;return <tr key={b.department}><td><strong>{b.department}</strong>{b.separate&&<span className="separate-label">별도</span>}</td><td>{b.amount===null?<span className="no-phone">미설정</span>:money(b.amount)}</td><td>{money(b.spent)}</td><td className={b.amount!==null&&b.amount-b.spent<0?'negative':'balance'}>{b.amount===null?'—':money(b.amount-b.spent)}</td><td><div className="mini-progress"><span><i style={{width:`${Math.min(100,pct??0)}%`}}/></span><small>{pct===null?'—':`${pct.toFixed(1)}%`}</small></div></td><td><button className={`count-link ${b.missing?'has-missing':''}`} onClick={()=>showDocs('missing',b.department)}>{b.missing}건</button></td><td><button className="icon-button" aria-label={`${b.department} 교육 내역`} onClick={()=>showDocs('all',b.department)}><ChevronRight size={16}/></button></td></tr>;})}</tbody></table></div>{!budgetRows.length&&<div className="empty-state"><Wallet size={30}/><h3>표시할 예산부서가 없습니다</h3><p>검색어와 업로드 파일의 예산 시트를 확인해 주세요.</p></div>}<div className="table-bottom"><span>단위: 원</span><span>교육 내역의 ‘위탁예산부서’ 기준</span></div></section><div className="inline-note"><CircleHelp size={15}/><span>등록 교육비에는 증빙 누락·교육 예정 기록도 포함됩니다. 실제 환급 완료 금액이 아닙니다.{includeSeparate&&' 교육행정팀 예산은 원본 내 기준 차이가 있어 합계 해석 시 확인이 필요합니다.'}</span></div>
        </>}

        {view==='upload'&&<div className="upload-layout"><section className="panel upload-panel"><div className={`dropzone ${drag?'dragging':''} ${busy?'loading':''}`} onDragOver={e=>{e.preventDefault();setDrag(true);}} onDragLeave={()=>setDrag(false)} onDrop={e=>{e.preventDefault();setDrag(false);void load(e.dataTransfer.files[0]);}}><span className="upload-illustration">{busy?<LoaderCircle className="spin" size={34}/>:<FileSpreadsheet size={34}/>}</span><h2>{busy?'엑셀을 정리하고 있어요':'엑셀 파일을 여기에 놓아주세요'}</h2><p>또는 아래 버튼을 눌러 파일을 선택하세요.</p><button className="primary" disabled={busy} onClick={()=>input.current?.click()}><Upload size={16}/>파일 선택</button><small>.xlsx 형식 · 최대 30MB</small></div><div className="upload-privacy"><ShieldCheck size={18}/><p><strong>업로드 데이터는 이 브라우저에만 저장돼요.</strong><br/>같은 브라우저와 기기에서 다시 열면 마지막 엑셀이 복원됩니다. 상단의 ‘저장 데이터 삭제’로 지울 수 있습니다.</p></div></section><section className="upload-guide"><span className="eyebrow">HOW IT WORKS</span><h2>지금 쓰는 엑셀,<br/>그대로 올리면 됩니다.</h2>{[['01','교육 내역 불러오기','보수교육_2026, 보수교육외_2026 시트에서 교육비와 서류 상태를 읽습니다.'],['02','예산과 연락처 연결','부서별 위탁예산과 전직원 명부를 사번으로 연결합니다.'],['03','필요한 정보 바로 찾기','예산 현황을 보고, 영수증·이수증 누락 직원의 이름·사번·교육명·전화번호를 검색하세요.']].map(([n,title,desc])=><div className="guide-step" key={n}><span>{n}</span><div><h3>{title}</h3><p>{desc}</p></div></div>)}<div className="upload-tip"><CircleHelp size={16}/><p>같은 파일을 다시 올리면 현재 데이터를 교체합니다. 교육비가 중복해서 더해지지 않습니다.</p></div></section></div>}
        <footer className="content-footer"><span>CAUHS <i/> 교육예산 관리</span><span>2026 예산연도</span></footer>
      </div>
    </main>
    {toast&&<div className="toast" role="status"><Check size={17}/>{toast}</div>}
    {(modal||detail)&&<div className="modal-overlay" onClick={()=>{setModal(null);setDetail(null);}}><section className="modal" role="dialog" aria-modal="true" aria-label={detail?'교육 상세':modal==='help'?'사용 가이드':'업로드 검토'} onClick={e=>e.stopPropagation()}><button className="modal-close icon-button" aria-label="닫기" onClick={()=>{setModal(null);setDetail(null);}}><X size={20}/></button>
      {detail?<><span className="eyebrow">EDUCATION DETAIL</span><h2>{detail.name}님의 교육</h2><h3 className="detail-title">{detail.title}</h3><dl className="detail-grid">{[['사번',detail.employeeId],['전화번호',phoneOf(detail)||'연락처 없음'],['근무부서',detail.department],['예산부서',detail.budgetDepartment],['교육구분',detail.category],['예산연도',detail.year],['교육시작',detail.start||'미입력'],['교육종료',detail.end||'미입력'],['교육비',detail.fee===null?detail.feeText||'미입력':`${money(detail.fee)}원`],['지출결의',detail.resolution||'미입력'],['지출결의여부',detail.resolutionStatus||'미입력'],['직원 연결',detail.matched?'직원 명부 연결됨':'직원 명부 미연결']].map(([label,value])=><div key={label}><dt>{label}</dt><dd>{value}</dd></div>)}</dl><div className="detail-status"><span>영수증 {documentStatus(detail.receipt,'영수증')}</span><span>이수증 {documentStatus(detail.certificate,'이수증')}</span></div><div className="detail-source">원본 위치: {detail.source} · {detail.row}행</div></>:modal==='help'?<><span className="eyebrow">QUICK GUIDE</span><h2>교육예산 관리 사용 가이드</h2><div className="help-content"><h3>1. 파일 올리기</h3><p>상단의 엑셀 업로드 버튼으로 원본 .xlsx 파일을 선택하세요. 처음에는 가상 예시가 표시됩니다.</p><h3>2. 증빙 누락 확인하기</h3><p>증빙서류 누락현황에서 영수증 또는 이수증이 누락된 교육을 확인할 수 있습니다. 이름·사번·교육명·전화번호로 검색하고 열 제목으로 정렬하세요.</p><h3>3. 예산 확인하기</h3><p>부서별 예산은 위탁예산부서 기준입니다. 사용금액은 숫자 교육비 합계이며 실제 환급 완료액과 다릅니다.</p><h3>4. 결과 내려받기</h3><p>현재 필터와 정렬이 적용된 목록을 CSV로 내려받을 수 있습니다.</p><div className="help-note">업로드 데이터는 서버에 전송되지 않고 이 브라우저 저장소에만 보관됩니다. 브라우저 저장 데이터를 지우거나 다른 기기·브라우저를 사용하면 다시 업로드해야 합니다. 실제 영수증·이수증 파일 검증과 지급 처리는 이번 버전에 포함되지 않습니다.</div></div></>:<><span className="eyebrow">IMPORT REVIEW</span><h2>업로드 검토 <span className="count-chip">{data.issues.length}</span></h2><p className="modal-description">원본 기록은 유지하고 확인이 필요한 항목을 표시했습니다.</p><div className="issue-list">{data.issues.map((issue,i)=><div className="issue" key={i}><AlertCircle size={17}/><div><p>{issue.message}</p><small>{issue.location}</small></div></div>)}{!data.issues.length&&<p>추가로 확인할 항목이 없습니다.</p>}</div></>}
    </section></div>}
  </div>;
}
createRoot(document.getElementById('root')!).render(<React.StrictMode><App/></React.StrictMode>);








