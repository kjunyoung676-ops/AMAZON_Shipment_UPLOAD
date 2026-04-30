"use client"

import { useState, useRef, useEffect } from "react"
import * as XLSX from "xlsx"

const SKU_ORDER = [
  "HS124018Z-5W","HS124018Z-5B","HS124518W-5B","HS124518W-5W",
  "HS604012-4B","HS604012-4W","HS604015HP-5W","HS604015HPW-5W",
  "HS604015W-5B","HS604015W-5W","HS804016.5HPW-5W","HS804018HP-5W",
  "HS904016.5W-5W","HS904016.5W-5B","HS904018-5B","HS904018-5W",
]

const MASTER_INIT: Record<string, MasterItem> = {
  "HS604012-4B":      {sku:"HS604012-4B-1Pack-JP",           asin:"B0CTLLCPF8", to:"H8",  loc:"J-07-B", cpp:24, fba:1020, price:12300, kg:15.50, bx:65,   by:41.5, bz:12  },
  "HS604012-4W":      {sku:"HS604012-4W-1Pack",              asin:"B0CQZJCCCB", to:"H6",  loc:"J-07-W", cpp:24, fba:1020, price:12300, kg:15.50, bx:65,   by:41.5, bz:12  },
  "HS904018-5B":      {sku:"HS904018-5B-1Pack-JP",           asin:"B0CTLDTN5V", to:"H9",  loc:"J-09-B", cpp:16, fba:1100, price:16400, kg:24.38, bx:94.5, by:41.5, bz:12.5},
  "HS904018-5W":      {sku:"US-HSBWB6-090-040-180-5S-1Pack", asin:"B09NLJMGDZ", to:"H2",  loc:"J-09-W", cpp:16, fba:1100, price:16400, kg:22.00, bx:94.5, by:41.5, bz:12.5},
  "HS124018Z-5B":     {sku:"HS124018Z-5B-1Pack-JP",          asin:"B0CTM7TKRC", to:"HG1", loc:"J-10-B", cpp:16, fba:1756, price:20300, kg:31.44, bx:125,  by:41.5, bz:12.5},
  "HS124018Z-5W":     {sku:"HK124018Z-5W-1Pack",             asin:"B0BYJ6C1PG", to:"H3",  loc:"J-10-W", cpp:16, fba:1756, price:20300, kg:31.44, bx:125,  by:41.5, bz:12.5},
  "HS604015W-5B":     {sku:"HS604015W-5B-JP",                asin:"B0CTKFC5LR", to:"H4",  loc:"J-12-B", cpp:24, fba:1020, price:14800, kg:12.90, bx:80,   by:41.5, bz:13  },
  "HS604015W-5W":     {sku:"HS604015W-5W-JAPAN",             asin:"B0CWQ8QX1Y", to:"H11", loc:"J-12-W", cpp:24, fba:1100, price:14800, kg:20.00, bx:80,   by:41.5, bz:13  },
  "HS904016.5W-5B":   {sku:"HS9040165W-5B-JP",               asin:"B0CTKDFQ47", to:"H5",  loc:"J-13-B", cpp:16, fba:1532, price:18800, kg:27.20, bx:42,   by:96,   bz:13  },
  "HS904016.5W-5W":   {sku:"HS904016.5W-5W-JAPAN",           asin:"B0CWQF2XC9", to:"H12", loc:"J-13-W", cpp:16, fba:1532, price:18800, kg:27.20, bx:96,   by:42,   bz:13  },
  "HS124518W-5B":     {sku:"HS124518W-5B-JAPAN",             asin:"B0CWQ9X3QB", to:"H10", loc:"J-14-B", cpp:16, fba:1756, price:22300, kg:35.30, bx:125,  by:47,   bz:13.5},
  "HS124518W-5W":     {sku:"HS124518W-5W-JAPAN",             asin:"B0CWQGKPMG", to:"H13", loc:"J-14-W", cpp:16, fba:1756, price:22300, kg:35.30, bx:125,  by:47,   bz:13.5},
  "HS604015HP-5W":    {sku:"HS604015HP-5W",                  asin:"B0GVSSFYRZ", to:"",    loc:"J-15-W", cpp:24, fba:1100, price:16300, kg:22.00, bx:79.8, by:41.8, bz:13.7},
  "HS804018HP-5W":    {sku:"HS804018HP-5W",                  asin:"B0GVSWH65N", to:"",    loc:"J-16-W", cpp:16, fba:1100, price:19400, kg:27.00, bx:94.5, by:42.5, bz:13.7},
  "HS604015HPW-5W":   {sku:"HS604015HPW-5W",                 asin:"B0GVT1R5HY", to:"",    loc:"J-17-W", cpp:18, fba:1532, price:18800, kg:24.00, bx:79.8, by:41.5, bz:16.8},
  "HS804016.5HPW-5W": {sku:"HS804016.5HPW-5W",               asin:"B0GVSPLDG3", to:"",    loc:"J-18-W", cpp:12, fba:1532, price:21800, kg:29.10, bx:94.5, by:42.0, bz:16.8},
}

const ALIASES: Record<string, string[]> = {
  shipment_date: ["출하 계획(공장)","출하 계획"],
  shipment_time: ["출하 시간"],
  destination:   ["국가","목적지","도착지"],
  model_code:    ["BOM 품번","품번"],
  sku:           ["품목"],
  color:         ["색상"],
  quantity:      ["수량","qty"],
  qty_total:     ["수량 합계"],
  ctn_count:     ["컨 수량"],
  ft:            ["FT"],
  type_old:      ["TYPE (구)"],
  location:      ["TYPE (신)","로케이션"],
  etd:           ["출항(ETD)"],
  eta:           ["도착(ETA)"],
  forwarding:    ["포워딩"],
  carrier:       ["선사"],
}
const FF_COLS = ["shipment_date","shipment_time","destination","etd","eta","qty_total","ctn_count","ft","sku","model_code","color"]
const CB = ["#dbeafe","#d1fae5","#fef3c7","#fce7f3","#ede9fe","#ffedd5","#e0f2fe","#dcfce7"]
const CT = ["#1e40af","#065f46","#92400e","#9d174d","#4c1d95","#9a3412","#075985","#166534"]
const TH: React.CSSProperties = {padding:"7px 10px",textAlign:"left",fontWeight:500,fontSize:11,color:"var(--color-text-primary)",borderBottom:"1.5px solid var(--color-border-secondary)",whiteSpace:"nowrap",background:"var(--color-background-secondary)",position:"sticky",top:0}
const TD: React.CSSProperties = {padding:"5px 10px",borderBottom:"0.5px solid var(--color-border-tertiary)",fontSize:12,overflow:"hidden",textOverflow:"ellipsis",whiteSpace:"nowrap",color:"var(--color-text-primary)"}
const INP: React.CSSProperties = {border:"none",background:"transparent",fontSize:12,padding:"1px 0",outline:"none",color:"var(--color-text-primary)",fontFamily:"inherit",width:"100%"}
const IBLU: React.CSSProperties = {background:"rgba(219,234,254,0.35)"}
const EMPTY_SKU = {sku:"",realSku:"",asin:"",to:"",loc:"",cpp:16,fba:0,price:0,kg:0,bx:0,by:0,bz:0}

const LS_KEY = "shipment_app_v1"
function lsLoad<T>(key: string, fallback: T): T {
  try { const r = localStorage.getItem(LS_KEY+"_"+key); if (r) return JSON.parse(r) as T } catch { /**/ }
  return fallback
}
function lsSave(key: string, value: unknown) {
  try { localStorage.setItem(LS_KEY+"_"+key, JSON.stringify(value)) } catch { /**/ }
}

interface MasterItem { sku:string; asin:string; to:string; loc:string; cpp:number; fba:number; price:number; kg:number; bx:number; by:number; bz:number }
interface RowData { [key:string]:unknown; __groupStart?:boolean; quantity:number; container_no?:number }
interface LabelDebugEntry {
  file: string
  page: number
  skusOnPage: string[]       // PDF에서 읽힌 원문 SKU 텍스트
  skusNorm: string[]         // 정규화된 SKU (매칭 비교용)
  matchedLoc: string | null  // 매칭된 LOC
  labelCount: number
  note?: string
}

const calcGW  = (qty:number, nw:number) => Math.trunc(qty*nw*1.01*10)/10
const calcCBM = (qty:number, bx:number, by:number, bz:number) => Math.round(bx*by*bz/1000000*qty*100)/100

function fmtDate(v:unknown):unknown {
  if (v instanceof Date && !isNaN(v.getTime())) return String(v.getMonth()+1).padStart(2,'0')+"월 "+String(v.getDate()).padStart(2,'0')+"일"
  if (typeof v==='number' && v>40000 && v<60000) { const d=new Date(Date.UTC(1899,11,30)+v*86400000); if (!isNaN(d.getTime())) return String(d.getUTCMonth()+1).padStart(2,'0')+"월 "+String(d.getUTCDate()).padStart(2,'0')+"일" }
  return v
}
function normHeaders(rows:RowData[]) {
  if (!rows.length) return {data:[]}
  const lm:Record<string,string>={};Object.keys(rows[0]).forEach(k=>{lm[k.trim().toLowerCase()]=k})
  const rn:Record<string,string>={}
  for (const [can,als] of Object.entries(ALIASES)) for (const a of als) if (a.toLowerCase() in lm){rn[lm[a.toLowerCase()]]=can;break}
  return {data:rows.map(r=>{const nr:RowData={quantity:0};for (const [k,v] of Object.entries(r)) nr[rn[k]||k]=v;nr.__groupStart=r.__groupStart||false;nr.quantity=parseFloat(String(nr.quantity))||0;return nr})}
}
function parseSheet(ws:XLSX.WorkSheet):RowData[] {
  if (!ws['!ref']) return []
  const merges=ws['!merges']||[];const range=XLSX.utils.decode_range(ws['!ref']);const hr=range.s.r;const gsr=new Set<number>()
  for (const m of merges) if (m.s.c<=2&&m.e.r>m.s.r&&m.s.r>hr) gsr.add(m.s.r)
  const raw=XLSX.utils.sheet_to_json(ws,{defval:"",cellDates:true}) as Record<string,unknown>[]
  return raw.map((r,i)=>{const nr:RowData={quantity:0};for(const [k,v] of Object.entries(r)) nr[k]=fmtDate(v);nr.__groupStart=gsr.has(i+hr+1);return nr})
}
function forwardFill(rows:RowData[]):RowData[] {
  const res=rows.map(r=>({...r}));const last:Record<string,unknown>={};let ctnNo=0
  for (const row of res) {
    if (row.__groupStart) ctnNo++;if (ctnNo===0) ctnNo=1
    for (const col of FF_COLS){const v=row[col];if(v!==undefined&&v!==''&&v!==null)last[col]=v;else if(col in last)row[col]=last[col]}
    row.container_no=ctnNo
  }
  return res
}
function xlsDl(data:unknown[],sn:string,fn:string){const wb=XLSX.utils.book_new();const ws=Array.isArray(data[0])?XLSX.utils.aoa_to_sheet(data as unknown[][]):XLSX.utils.json_to_sheet(data as Record<string,unknown>[]);XLSX.utils.book_append_sheet(wb,ws,sn||"data");XLSX.writeFile(wb,fn)}

export default function ShipmentApp() {
  const [sheets,setSh]=useState<Record<string,RowData[]>>({})
  const [sheetNames,setSn]=useState<string[]>([])
  const [activeSheet,setAs]=useState<string|null>(null)
  const [file,setFile]=useState<File|null>(null)
  const [mode,setMode]=useState('1')
  const [master,setMaster]=useState<Record<string,MasterItem>>(()=>lsLoad("master",Object.fromEntries(Object.entries(MASTER_INIT).map(([k,v])=>[k,{...v}]))))
  const [s2meta,setS2meta]=useState<Record<string,Record<string,string>>>(()=>lsLoad("s2meta",{}))
  const [ctnMeta,setCtnMeta]=useState<Record<number,Record<string,string>>>(()=>lsLoad("ctnMeta",{}))
  useEffect(()=>{lsSave("master",master)},[master])
  useEffect(()=>{lsSave("s2meta",s2meta)},[s2meta])
  useEffect(()=>{lsSave("ctnMeta",ctnMeta)},[ctnMeta])
  const [coll,setColl]=useState<Record<string,boolean>>({})
  const [newSku,setNewSku]=useState({...EMPTY_SKU})
  const [fbaLoading,setFbaLoading]=useState(false)
  const [labelFiles,setLabelFiles]=useState<File[]>([])
  const [labelGroups,setLabelGroups]=useState<Record<string,Uint8Array>>({})
  const [labelCounts,setLabelCounts]=useState<Record<string,number>>({})
  const [labelLoading,setLabelLoading]=useState(false)
  const [labelStatus,setLabelStatus]=useState("")
  const [labelDebug,setLabelDebug]=useState<LabelDebugEntry[]>([])
  const [showDebug,setShowDebug]=useState(false)
  const fileRef=useRef<HTMLInputElement>(null)
  const labelRef=useRef<HTMLInputElement>(null)

  const raw=(activeSheet&&sheets[activeSheet])||[]
  const {data:nd}=raw.length?normHeaders(raw):{data:[] as RowData[]}
  const fd=nd.length?forwardFill(nd):[]
  const ctnNums=[...new Set(fd.map(r=>r.container_no as number))].sort((a,b)=>a-b)

  function loadFile(f:File){setFile(f);const rd=new FileReader();rd.onload=e=>{const wb=XLSX.read((e.target as FileReader).result,{type:'array'});const ns:Record<string,RowData[]>={};for(const sn of wb.SheetNames)ns[sn]=parseSheet(wb.Sheets[sn]);setSh(ns);setSn(wb.SheetNames);setAs(wb.SheetNames[0]);setMode('1')};rd.readAsArrayBuffer(f)}

  function buildS2rows(){
    const agg:Record<string,{sku:string,total:number}>={};for(const r of fd){const sk=String(r.sku||"");if(!sk)continue;if(!agg[sk])agg[sk]={sku:sk,total:0};agg[sk].total+=r.quantity}
    return Object.values(agg).sort((a,b)=>{const ai=SKU_ORDER.indexOf(a.sku),bi=SKU_ORDER.indexOf(b.sku);if(ai===-1&&bi===-1)return a.sku.localeCompare(b.sku);if(ai===-1)return 1;if(bi===-1)return -1;return ai-bi}).map(({sku,total})=>{
      const m=master[sku]||({} as MasterItem);const cpp=parseFloat(String(m.cpp))||16;const meta=s2meta[sku]||{}
      return {sku,total,cpp,asin:m.asin||"",to:m.to||"",loc:m.loc||"",pallets:Math.ceil(total/cpp),pages:parseFloat((total/6).toFixed(2)),gw:calcGW(total,parseFloat(String(m.kg))||0),cbm:calcCBM(total,parseFloat(String(m.bx))||0,parseFloat(String(m.by))||0,parseFloat(String(m.bz))||0),fc:meta.fc||"",address:meta.address||"",fbaId:meta.fbaId||"",amazonId:meta.amazonId||""}
    })
  }
  function buildS3groups(){return ctnNums.map(no=>{const rows=fd.filter(r=>r.container_no===no);const f0=rows[0]||{};const cm=ctnMeta[no]||{};return {no,container:cm.container||"",sealNo:cm.sealNo||"",shipment_date:String(f0.shipment_date||""),shipment_time:String(f0.shipment_time||""),destination:String(f0.destination||""),etd:String(f0.etd||""),eta:String(f0.eta||""),rows}})}
  const updM=(sku:string,f:string,v:string|number)=>setMaster(p=>({...p,[sku]:{...(p[sku]||{}),[f]:v} as MasterItem}))
  const updMeta=(sku:string,f:string,v:string)=>setS2meta(p=>({...p,[sku]:{...(p[sku]||{}),[f]:v}}))
  const updCtnMeta=(no:number,f:string,v:string)=>setCtnMeta(p=>({...p,[no]:{...(p[no]||{}),[f]:v}}))
  const togCtn=(k:string)=>setColl(c=>({...c,[k]:!c[k]}));const isOpen=(k:string)=>!coll[k]
  function tag(l:string,v:unknown){if(!v||v===""||v===0||v==="0")return null;return <span style={{fontSize:12,opacity:0.85}}><span style={{opacity:0.6,marginRight:2}}>{l}</span>{String(v)}</span>}

  function expS1(){xlsDl(fd.map(r=>({시트:activeSheet,컨:"컨"+r.container_no,출항일:r.shipment_date||"",목적지:r.destination||"",BOM품번:r.model_code||"",품목:r.sku||"",색상:r.color||"",수량:r.quantity,TYPE구:r.type_old||"",TYPE신:r.location||"",ETD:r.etd||"",ETA:r.eta||""})),activeSheet||"S1","1단계_"+(activeSheet||"data")+".xlsx")}
  function expS2(){const rows=buildS2rows();const hdr=[["약호","ASIN","구박스","신박스","카톤","PLT당카톤","팔레트","G.W(kg)","CBM","FC CENTER","주소","FBA ID","아마존 ID"]];const body=rows.map(r=>[r.sku,r.asin,r.to,r.loc,r.total,r.cpp,r.pallets,r.gw,r.cbm,r.fc,r.address,r.fbaId,r.amazonId]);const tC=rows.reduce((s,r)=>s+r.total,0),tP=rows.reduce((s,r)=>s+r.pallets,0),tW=rows.reduce((s,r)=>s+r.gw,0),tB=Math.round(rows.reduce((s,r)=>s+r.cbm,0)*100)/100;xlsDl([...hdr,...body,["합계","","","",tC,"",tP,tW,tB]],activeSheet||"S2","1차가공_"+(activeSheet||"data")+".xlsx")}
  function expS3(){const hdr=[["CONTAINER","SEAL NO.","약호","BOX CODE","ASIN","FBA ID","아마존 ID","FC CENTER","주소","PLT","CT","CTN/PLT","G.W(kg)","CBM"]];const body:unknown[][]=[];for(const g of buildS3groups()){let first=true;for(const r of g.rows){const m=master[String(r.sku)]||({} as MasterItem);const cpp=parseFloat(String(m.cpp))||16;const sk=s2meta[String(r.sku)]||{};body.push([first?g.container:"",first?g.sealNo:"",r.sku,r.location||m.loc||"",m.asin||"",sk.fbaId||"",sk.amazonId||"",sk.fc||"",first?(sk.address||""):"",Math.ceil(r.quantity/cpp),r.quantity,cpp+"CTN/PLT",calcGW(r.quantity,parseFloat(String(m.kg))||0),calcCBM(r.quantity,parseFloat(String(m.bx))||0,parseFloat(String(m.by))||0,parseFloat(String(m.bz))||0)]);first=false}};xlsDl([...hdr,...body],activeSheet||"S3","2차가공_"+(activeSheet||"data")+".xlsx")}

  async function expFbaUpload(){
    setFbaLoading(true)
    try{const res=await fetch("/UPLOAD_FORMAT.xlsx");if(!res.ok)throw new Error("파일 로드 실패 ("+res.status+")");const buf=await res.arrayBuffer();const wb=XLSX.read(buf,{type:"array",cellStyles:true,cellNF:true,cellDates:true,sheetStubs:true});const tplName=wb.SheetNames.find(n=>n.toLowerCase().includes("template"));if(!tplName)throw new Error("template 시트를 찾을 수 없습니다: "+wb.SheetNames.join(", "));const ws=wb.Sheets[tplName];const rows=buildS2rows();rows.forEach((r,i)=>{const row=9+i;const m=master[r.sku]||({} as MasterItem);const realSku=m.sku||r.sku;const set=(col:string,t:"s"|"n",v:string|number)=>{ws[col+row]={...(ws[col+row]||{}),t,v,w:String(v)}};set("A","s",realSku);set("B","n",r.total);set("F","n",1);set("G","n",r.total);set("H","n",m.bx||0);set("I","n",m.by||0);set("J","n",m.bz||0);set("K","n",m.kg||0)});if(rows.length>0){const decoded=XLSX.utils.decode_range(ws["!ref"]||"A1:K8");decoded.e.r=Math.max(decoded.e.r,8+rows.length-1);decoded.e.c=Math.max(decoded.e.c,10);ws["!ref"]=XLSX.utils.encode_range(decoded)};XLSX.writeFile(wb,"UPLOAD_FORMAT_filled.xlsx",{cellStyles:true,compression:true})}catch(e){alert("오류: "+(e as Error).message)}finally{setFbaLoading(false)}
  }

  // eslint-disable-next-line @typescript-eslint/no-explicit-any
  async function loadPdfjs():Promise<any>{
    // eslint-disable-next-line @typescript-eslint/no-explicit-any
    if((window as any).pdfjsLib)return(window as any).pdfjsLib
    // eslint-disable-next-line @typescript-eslint/no-explicit-any
    return new Promise<any>((resolve,reject)=>{const s=document.createElement('script');s.src='https://cdnjs.cloudflare.com/ajax/libs/pdf.js/3.11.174/pdf.min.js';s.onload=()=>{const lib=(window as any).pdfjsLib as any;lib.GlobalWorkerOptions.workerSrc='https://cdnjs.cloudflare.com/ajax/libs/pdf.js/3.11.174/pdf.worker.min.js';resolve(lib)};s.onerror=reject;document.head.appendChild(s)})
  }

  // ── SKU 정규화: 하이픈·공백·suffix 제거 후 비교용 문자열 반환 ──
  function normalizeSku(sku: string): string {
    return sku
      .replace(/\s+/g, '')           // 공백 전체 제거
      .replace(/-/g, '')             // 하이픈 전체 제거
      .replace(/1Pack(JP)?$/i, '')   // 1Pack / 1PackJP suffix 제거
      .toUpperCase()
  }

  // ── realSku → loc 매핑 헬퍼 ──────────────────────────────
  // 매칭 우선순위:
  // 1. 완전 일치
  // 2. 공백 제거 후 일치
  // 3. 하이픈+공백 제거 후 일치
  // 4. suffix(1Pack, 1Pack-JP 등) + 하이픈 제거 후 일치
  function skuToLoc(sku: string, realSkuToLoc: Record<string, string>): string | null {
    if (!sku) return null

    // 1. 완전 일치
    if (realSkuToLoc[sku]) return realSkuToLoc[sku]

    // PDF SKU 정규화
    const skuNorm = normalizeSku(sku)

    for (const [rsku, loc] of Object.entries(realSkuToLoc)) {
      // 2. 공백 제거 일치
      if (rsku.replace(/\s+/g, '') === sku.replace(/\s+/g, '')) return loc

      // 3~4. 하이픈+suffix 제거 후 일치
      const rskuNorm = normalizeSku(rsku)
      if (rskuNorm === skuNorm && rskuNorm.length > 5) return loc

      // 5. PDF SKU가 마스터 SKU를 포함하거나 그 반대 (부분 매칭)
      //    단, 너무 짧은 건 제외 (오탐 방지)
      if (skuNorm.length > 8 && rskuNorm.length > 8) {
        if (skuNorm.includes(rskuNorm) || rskuNorm.includes(skuNorm)) return loc
      }
    }
    return null
  }

  // ── 라벨 PDF 분류 ─────────────────────────────────────────
  // 로직:
  // 1. 각 페이지에서 pdfjs로 텍스트 추출 → 「単一のSKU」 다음 토큰이 라벨의 SKU
  // 2. 페이지 내 라벨 수 = 「単一のSKU」 출현 횟수
  // 3. SKU별로 LOC 결정 → 해당 SKU에 속하는 라벨 인덱스만 남기고
  //    나머지 라벨 위치에 흰색 사각형으로 마스킹
  // 4. 라벨 그리드: 2열 레이아웃, 위→아래·왼→오른 순서
  async function processLabels(files: File[]) {
    setLabelLoading(true)
    setLabelGroups({})
    setLabelCounts({})
    setLabelDebug([])
    setLabelStatus("라이브러리 로딩 중...")

    try {
      // realSku → loc
      const realSkuToLocMap: Record<string, string> = {}
      for (const m of Object.values(master)) {
        if (m.sku && m.loc) realSkuToLocMap[m.sku] = m.loc
      }

      const pdfjsLib = await loadPdfjs()
      const { PDFDocument, rgb } = await import('pdf-lib')

      type PageEntry = {
        fileBytes: Uint8Array      // 원본 PDF 바이트 (Uint8Array, 매번 slice로 복사)
        pageIdx: number            // 0-based
        labelsToKeep: number[]     // 유지할 라벨 인덱스 (빈 배열 = 전체 유지)
        totalLabels: number
      }

      const locMap: Record<string, PageEntry[]> = {}
      const debugLog: LabelDebugEntry[] = []
      let totalPages = 0
      let processed = 0

      // 전체 페이지 수
      for (const f of files) {
        const bytes = new Uint8Array(await f.arrayBuffer())
        const pdf = await pdfjsLib.getDocument({ data: bytes.slice() }).promise
        totalPages += pdf.numPages
        pdf.destroy()
      }

      for (const f of files) {
        const fileBytes = new Uint8Array(await f.arrayBuffer())
        const pdf = await pdfjsLib.getDocument({ data: fileBytes.slice() }).promise

        for (let i = 0; i < pdf.numPages; i++) {
          const page = await pdf.getPage(i + 1)
          const content = await page.getTextContent()
          // eslint-disable-next-line @typescript-eslint/no-explicit-any
          const items = content.items as any[]
          const tokenTexts = items.map((it)=>String(it.str || ""))
          const fullTextSpaced = tokenTexts.join(" ")
          const fullTextNoSpace = tokenTexts.join("").replace(/\s+/g, "")

          // 각 라벨의 SKU 추출
          // 1) 토큰 기반 추출
          // 2) 실패 시 전체 텍스트 정규식 fallback
          // 라벨 위치: transform[5](Y), transform[4](X)로 정렬
          type LabelInfo = { sku: string; y: number; x: number }
          const labelInfos: LabelInfo[] = []

          for (let j = 0; j < items.length; j++) {
            const str = items[j].str as string
            if (str && (str.includes('単一のSKU') || str === '単一のSKU')) {
              const y = items[j].transform?.[5] ?? 0
              const x = items[j].transform?.[4] ?? 0
              let sku = ''
              // 바로 뒤 최대 8토큰에서 SKU 패턴 찾기
              for (let k = j + 1; k < Math.min(j + 9, items.length); k++) {
                const s = (items[k].str as string)?.trim()
                if (s && s.length > 4
                    && !s.includes('数量') && !s.includes('JAN')
                    && !s.match(/^[\d\s]+$/) && !s.includes('のSKU')
                    && !s.includes('FBA') && !s.includes('Pack数')
                ) {
                  sku = s
                  break
                }
              }
              if (sku) labelInfos.push({ sku, y, x })
            }
          }

          // fallback: 텍스트 전체에서 「単一のSKU <SKU> 数量」 패턴 추출
          if (labelInfos.length === 0) {
            const re = /単一のSKU\s*([A-Za-z0-9.\-_/]+?)\s*(?:数量|JAN|FBA|$)/g
            const fallbackSkus: string[] = []
            let m: RegExpExecArray | null
            while ((m = re.exec(fullTextSpaced)) !== null) {
              const s = (m[1] || "").trim()
              if (s.length > 4) fallbackSkus.push(s)
            }
            fallbackSkus.forEach((sku, idx) => labelInfos.push({ sku, y: 10000 - idx * 40, x: idx % 2 === 0 ? 0 : 1000 }))
          }

          // fallback2: marker가 전혀 없어도 페이지 텍스트에서 마스터 realSku를 직접 탐색
          if (labelInfos.length === 0) {
            const matchedMasterSkus: string[] = []
            for (const m of Object.values(master)) {
              if (!m.sku) continue
              const norm = normalizeSku(m.sku)
              if (norm.length < 6) continue
              if (fullTextNoSpace.toUpperCase().includes(norm)) matchedMasterSkus.push(m.sku)
            }
            matchedMasterSkus.forEach((sku, idx) => labelInfos.push({ sku, y: 10000 - idx * 40, x: idx % 2 === 0 ? 0 : 1000 }))
          }

          if (labelInfos.length === 0) {
            debugLog.push({
              file: f.name,
              page: i + 1,
              skusOnPage: [],
              skusNorm: [],
              matchedLoc: null,
              labelCount: 0,
              note: `SKU marker not found on page (textLen=${fullTextNoSpace.length})`,
            })
            processed++
            continue
          }

          // Y 내림차순(상→하), 같은 행이면 X 오름차순(좌→우) 정렬
          labelInfos.sort((a, b) => {
            if (Math.abs(a.y - b.y) > 30) return b.y - a.y
            return a.x - b.x
          })

          const totalLabels = labelInfos.length
          const uniqueSkus = [...new Set(labelInfos.map(l => l.sku))]

          // SKU별 처리
          for (const targetSku of uniqueSkus) {
            const loc = skuToLoc(targetSku, realSkuToLocMap)
            if (!loc) continue

            // 해당 SKU의 라벨 인덱스
            const keptIndices = labelInfos
              .map((l, idx) => ({ l, idx }))
              .filter(({ l }) => {
                const ls = l.sku.replace(/\s+/g, '')
                const ts = targetSku.replace(/\s+/g, '')
                return ls === ts
              })
              .map(({ idx }) => idx)

            if (!locMap[loc]) locMap[loc] = []
            locMap[loc].push({
              fileBytes,
              pageIdx: i,
              labelsToKeep: keptIndices.length === totalLabels ? [] : keptIndices,
              totalLabels,
            })
          }

          const matchedLocs = uniqueSkus.map(s => skuToLoc(s, realSkuToLocMap)).filter(Boolean)
          debugLog.push({
            file: f.name,
            page: i + 1,
            skusOnPage: uniqueSkus,
            skusNorm: uniqueSkus.map(s => normalizeSku(s)),
            matchedLoc: matchedLocs.join(', ') || null,
            labelCount: totalLabels,
            note: matchedLocs.length ? undefined : "SKU extracted but no master mapping match",
          })

          processed++
          if (processed % 30 === 0 || processed === totalPages) {
            setLabelStatus(`페이지 분석 중... ${processed}/${totalPages}`)
          }
        }
        pdf.destroy()
      }

      setLabelDebug(debugLog)
      setShowDebug(true)  // 항상 디버그 패널 열기 (성공/실패 무관)

      if (!Object.keys(locMap).length) {
        setLabelStatus("매칭 실패 — 아래 디버그 패널에서 PDF 내 SKU 텍스트 확인")
        return
      }

      setLabelStatus("PDF 생성 중...")
      const result: Record<string, Uint8Array> = {}
      const counts: Record<string, number> = {}

      for (const [loc, entries] of Object.entries(locMap)) {
        let totalKept = 0
        const outDoc = await PDFDocument.create()

        for (const entry of entries) {
          // 매번 새로운 slice로 로드 → ArrayBuffer detach 방지
          const srcDoc = await PDFDocument.load(entry.fileBytes.slice())
          const [copiedPage] = await outDoc.copyPages(srcDoc, [entry.pageIdx])

          if (entry.labelsToKeep.length > 0) {
            // 혼합 페이지: 다른 SKU 라벨 위치를 흰색으로 마스킹
            const { width, height } = copiedPage.getSize()
            const cols = 2
            const rows = Math.ceil(entry.totalLabels / cols)
            const cellW = width / cols
            const cellH = height / rows

            for (let idx = 0; idx < entry.totalLabels; idx++) {
              if (!entry.labelsToKeep.includes(idx)) {
                const col = idx % cols
                const row = Math.floor(idx / cols)
                // PDF 좌표: 원점=좌하단, Y는 아래서 위로
                const x = col * cellW
                const y = height - (row + 1) * cellH
                copiedPage.drawRectangle({ x, y, width: cellW, height: cellH, color: rgb(1, 1, 1), opacity: 1 })
              }
            }
            totalKept += entry.labelsToKeep.length
          } else {
            totalKept += entry.totalLabels
          }

          outDoc.addPage(copiedPage)
        }

        counts[loc] = totalKept
        result[loc] = await outDoc.save()
      }

      setLabelGroups(result)
      setLabelCounts(counts)
      const unmatchedPages = debugLog.filter(d => !d.matchedLoc).length
      setLabelStatus(`완료 — ${Object.keys(result).length}개 BOX CODE 분류됨${unmatchedPages > 0 ? ` (미매칭 ${unmatchedPages}페이지)` : ""}`)
      if (unmatchedPages > 0) setShowDebug(true)

    } catch (e) {
      alert('처리 오류: ' + (e as Error).message)
      setLabelStatus("")
    } finally {
      setLabelLoading(false)
    }
  }

  function dlLabel(loc:string,bytes:Uint8Array){const a=document.createElement('a');a.href=URL.createObjectURL(new Blob([bytes],{type:'application/pdf'}));a.download=`${loc}.pdf`;a.click()}
  function dlAllLabels(){Object.entries(labelGroups).forEach(([l,b])=>dlLabel(l,b))}

  function SheetTabs(){
    if(!sheetNames.length)return null
    return(<div style={{display:"flex",gap:2,borderBottom:"0.5px solid var(--color-border-tertiary)",marginBottom:16}}>{sheetNames.map(sn=>{const {data:d}=(sheets[sn]||[]).length?normHeaders(sheets[sn]):{data:[] as RowData[]};const ff2=d.length?forwardFill(d):[];const cc=new Set(ff2.map(r=>r.container_no)).size;return(<button key={sn} onClick={()=>setAs(sn)} style={{padding:"6px 14px",fontSize:12,fontWeight:activeSheet===sn?500:400,cursor:"pointer",background:"transparent",border:"none",color:activeSheet===sn?"var(--color-text-primary)":"var(--color-text-secondary)",borderBottom:activeSheet===sn?"2px solid var(--color-text-primary)":"2px solid transparent",marginBottom:-1,display:"flex",alignItems:"center",gap:5}}>{sn}{cc>0&&<span style={{fontSize:10,padding:"1px 5px",borderRadius:10,background:"var(--color-background-secondary)",border:"0.5px solid var(--color-border-tertiary)",color:"var(--color-text-tertiary)"}}>{cc}컨</span>}</button>)})}</div>)
  }

  const s2rows=buildS2rows()
  const totC=s2rows.reduce((s,r)=>s+r.total,0),totP=s2rows.reduce((s,r)=>s+r.pallets,0),totW=s2rows.reduce((s,r)=>s+r.gw,0),totCBM=Math.round(s2rows.reduce((s,r)=>s+r.cbm,0)*100)/100

  function NavBtn({m,label}:{m:string,label:string}){
    const active=mode===m,avail=m==='1'||m==='map'||m==='label'||!!file
    return(<button onClick={()=>avail&&setMode(m)} style={{padding:"8px 16px",fontSize:12,fontWeight:active?500:400,cursor:avail?"pointer":"default",border:"none",background:active?"var(--color-text-primary)":"transparent",color:active?"var(--color-background-primary)":avail?"var(--color-text-secondary)":"var(--color-text-tertiary)"}}>{label}</button>)
  }

  return (
    <div style={{padding:"0.75rem 0",fontFamily:"var(--font-sans)"}}>
      {/* 내비 */}
      <div style={{display:"flex",alignItems:"center",gap:8,marginBottom:16}}>
        <div style={{border:"0.5px solid var(--color-border-tertiary)",borderRadius:"var(--border-radius-md)",overflow:"hidden",display:"flex"}}>
          <NavBtn m="1" label="① 업로드"/><div style={{width:"0.5px",background:"var(--color-border-tertiary)"}}/>
          <NavBtn m="2" label="② 1차 가공"/><div style={{width:"0.5px",background:"var(--color-border-tertiary)"}}/>
          <NavBtn m="label" label="② 라벨 분류"/><div style={{width:"0.5px",background:"var(--color-border-tertiary)"}}/>
          <NavBtn m="3" label="③ 2차 가공"/>
        </div>
        <button onClick={()=>setMode('map')} style={{marginLeft:"auto",fontSize:11,padding:"6px 14px",background:mode==='map'?"var(--color-background-secondary)":"transparent",border:"0.5px solid var(--color-border-tertiary)",borderRadius:"var(--border-radius-md)",cursor:"pointer",color:mode==='map'?"var(--color-text-primary)":"var(--color-text-secondary)"}}>매핑 관리</button>
      </div>

      {/* ══ STEP 1 ══ */}
      {mode==='1'&&(
        <div>
          <div style={{display:"flex",alignItems:"center",gap:8,marginBottom:14,flexWrap:"wrap"}}>
            <div onClick={()=>fileRef.current?.click()} onDragOver={e=>e.preventDefault()} onDrop={e=>{e.preventDefault();const f=e.dataTransfer.files[0];if(f)loadFile(f)}} style={{display:"flex",alignItems:"center",gap:8,padding:"6px 14px",border:"1px dashed var(--color-border-secondary)",borderRadius:"var(--border-radius-md)",cursor:"pointer",fontSize:13}}>
              <span>+</span><span style={{color:file?"var(--color-text-primary)":"var(--color-text-tertiary)"}}>{file?file.name:"파일 업로드 (xlsx · csv)"}</span>
            </div>
            <input ref={fileRef} type="file" accept=".xlsx,.xls,.csv" style={{display:"none"}} onChange={e=>e.target.files?.[0]&&loadFile(e.target.files[0])}/>
            {file&&<><span style={{fontSize:11,padding:"3px 9px",borderRadius:20,background:"var(--color-background-secondary)",border:"0.5px solid var(--color-border-tertiary)",color:"var(--color-text-secondary)"}}>{raw.length}행 · {ctnNums.length}컨 · {fd.reduce((s,r)=>s+r.quantity,0).toLocaleString()}개</span><button onClick={expS1} style={{fontSize:11,padding:"3px 10px"}}>xlsx 저장</button><button onClick={()=>setMode('2')} style={{marginLeft:"auto",fontSize:12,padding:"5px 14px",background:"var(--color-background-info)",border:"0.5px solid var(--color-border-info)",color:"var(--color-text-info)",borderRadius:"var(--border-radius-md)",cursor:"pointer",fontWeight:500}}>1차 가공으로 →</button></>}
          </div>
          <SheetTabs/>
          {!file?(
            <div style={{textAlign:"center",padding:"5rem 0",color:"var(--color-text-tertiary)"}}><div style={{fontSize:48,marginBottom:12}}>📦</div><p style={{fontSize:15}}>쉽먼트 파일을 업로드해주세요</p><p style={{fontSize:12,marginTop:4}}>시트가 여러 개면 탭으로 구분됩니다</p></div>
          ):(
            <div style={{display:"flex",flexDirection:"column",gap:10}}>
              {ctnNums.map(no=>{const rows=fd.filter(r=>r.container_no===no);const f0=rows[0],ci=(no-1)%CB.length;const sumQ=rows.reduce((s,r)=>s+r.quantity,0);const dQ=parseFloat(String(f0.qty_total))||sumQ;const key="s1_"+no,open=isOpen(key);return(
                <div key={no} style={{border:"0.5px solid var(--color-border-tertiary)",borderRadius:"var(--border-radius-md)",overflow:"hidden"}}>
                  <div onClick={()=>togCtn(key)} style={{background:CB[ci],color:CT[ci],padding:"9px 14px",display:"flex",alignItems:"center",gap:12,flexWrap:"wrap",cursor:"pointer",userSelect:"none"}}>
                    <span style={{fontWeight:500,fontSize:13}}>컨{no}</span><span style={{fontWeight:500}}>{String(f0.shipment_date||"")}</span>{tag("",f0.shipment_time)}<span style={{fontWeight:500}}>{String(f0.destination||"")}</span>{tag("ETD ",f0.etd)}{tag("ETA ",f0.eta)}{tag("선사 ",f0.carrier)}
                    <span style={{marginLeft:"auto",display:"flex",alignItems:"center",gap:8}}><span style={{fontWeight:500,fontSize:13}}>{dQ.toLocaleString()}개{f0.ctn_count&&f0.ctn_count!=="0"?" · "+f0.ctn_count+"컨":""}{f0.ft&&f0.ft!=="0"?" · "+f0.ft+"FT":""}</span><span style={{fontSize:11,opacity:0.6}}>{open?"▲":"▼"}</span></span>
                  </div>
                  {open&&(<div style={{overflowX:"auto"}}><table style={{width:"100%",borderCollapse:"collapse",fontSize:12,minWidth:600}}><thead><tr>{["BOM품번","약호(SKU)","색상","수량","TYPE(구)","신박스코드","ETD","ETA"].map(h=><th key={h} style={TH}>{h}</th>)}</tr></thead><tbody>{rows.map((r,i)=>(<tr key={i} style={{background:i%2===0?"transparent":"var(--color-background-secondary)"}}><td style={{...TD,fontFamily:"var(--font-mono)",fontSize:11,color:"var(--color-text-secondary)",minWidth:130}}>{String(r.model_code||"")}</td><td style={{...TD,fontWeight:500,minWidth:120}}>{String(r.sku||"")}</td><td style={{...TD,minWidth:50}}>{String(r.color||"")}</td><td style={{...TD,textAlign:"right",fontWeight:500,minWidth:50}}>{r.quantity.toLocaleString()}</td><td style={{...TD,color:"var(--color-text-secondary)"}}>{String(r.type_old||"")}</td><td style={{...TD,color:"var(--color-text-info)",fontWeight:500}}>{String(r.location||"")}</td><td style={TD}>{String(r.etd||"")}</td><td style={TD}>{String(r.eta||"")}</td></tr>))}</tbody><tfoot><tr style={{background:"var(--color-background-secondary)",borderTop:"1px solid var(--color-border-secondary)"}}><td colSpan={3} style={{...TD,textAlign:"right",fontSize:11,color:"var(--color-text-secondary)"}}>소계</td><td style={{...TD,textAlign:"right",fontWeight:500}}>{sumQ.toLocaleString()}</td><td colSpan={4} style={TD}></td></tr></tfoot></table></div>)}
                </div>
              )})}
            </div>
          )}
        </div>
      )}

      {/* ══ STEP 2 ══ */}
      {mode==='2'&&(
        <div>
          <div style={{display:"flex",alignItems:"center",gap:8,marginBottom:8,flexWrap:"wrap"}}>
            <span style={{fontSize:12,color:"var(--color-text-secondary)"}}>전체 합산 · <span style={{color:"var(--color-text-info)"}}>파란 셀</span> = 직접 입력</span>
            <button onClick={expS2} style={{marginLeft:"auto",fontSize:11,padding:"3px 10px"}}>xlsx 저장</button>
            <button onClick={expFbaUpload} disabled={fbaLoading} style={{fontSize:11,padding:"3px 10px",background:"var(--color-background-success)",border:"0.5px solid var(--color-border-success)",color:"var(--color-text-success)",borderRadius:"var(--border-radius-md)",cursor:fbaLoading?"default":"pointer",opacity:fbaLoading?0.6:1}}>{fbaLoading?"처리 중...":"FBA 업로드 양식"}</button>
            <button onClick={()=>setMode('3')} style={{fontSize:12,padding:"5px 14px",background:"var(--color-background-success)",border:"0.5px solid var(--color-border-success)",color:"var(--color-text-success)",borderRadius:"var(--border-radius-md)",cursor:"pointer",fontWeight:500}}>2차 가공으로 →</button>
          </div>
          <SheetTabs/>
          <div style={{overflowX:"auto",border:"0.5px solid var(--color-border-tertiary)",borderRadius:"var(--border-radius-md)"}}>
            <table style={{width:"100%",borderCollapse:"collapse",fontSize:12,minWidth:960}}>
              <thead>
                <tr><th colSpan={9} style={{...TH,borderRight:"2px solid var(--color-border-secondary)"}}>자동 계산</th><th colSpan={4} style={{...TH,...IBLU,color:"var(--color-text-info)"}}>직접 입력 (변동값)</th></tr>
                <tr>{["약호","ASIN","구박스","신박스","카톤","PLT당카톤","팔레트","G.W(kg)","CBM"].map(h=><th key={h} style={TH}>{h}</th>)}{["FC CENTER","주소","FBA ID","아마존 ID"].map(h=><th key={h} style={{...TH,...IBLU}}>{h}</th>)}</tr>
              </thead>
              <tbody>
                {s2rows.map((r,i)=>(<tr key={r.sku} style={{background:i%2===0?"transparent":"var(--color-background-secondary)"}}><td style={{...TD,fontWeight:500,minWidth:130}}>{r.sku}</td><td style={{...TD,fontFamily:"var(--font-mono)",fontSize:11,minWidth:100}}>{r.asin}</td><td style={{...TD,minWidth:45}}>{r.to}</td><td style={{...TD,color:"var(--color-text-info)",fontWeight:500,minWidth:70}}>{r.loc}</td><td style={{...TD,textAlign:"right",fontWeight:500,minWidth:55}}>{r.total.toLocaleString()}</td><td style={{...TD,textAlign:"center",minWidth:55}}>{r.cpp}</td><td style={{...TD,textAlign:"right",fontWeight:500,minWidth:50}}>{r.pallets}</td><td style={{...TD,textAlign:"right",minWidth:70,borderRight:"1px solid var(--color-border-tertiary)"}}>{r.gw.toFixed(1)}</td><td style={{...TD,textAlign:"right",minWidth:60,borderRight:"2px solid var(--color-border-secondary)"}}>{r.cbm}</td>{(["fc","address","fbaId","amazonId"] as const).map((f,idx)=>(<td key={f} style={{...TD,minWidth:[80,150,130,90][idx],...IBLU}}><input value={r[f]} onChange={e=>updMeta(r.sku,f,e.target.value)} style={INP} placeholder={["FC CENTER","주소","FBA ID","아마존 ID"][idx]}/></td>))}</tr>))}
              </tbody>
              <tfoot><tr style={{background:"var(--color-background-secondary)",borderTop:"2px solid var(--color-border-secondary)"}}><td colSpan={4} style={{...TD,fontWeight:500,textAlign:"right"}}>합계</td><td style={{...TD,fontWeight:500,textAlign:"right"}}>{totC.toLocaleString()}</td><td style={TD}></td><td style={{...TD,fontWeight:500,textAlign:"right"}}>{totP}</td><td style={{...TD,fontWeight:500,textAlign:"right"}}>{totW.toFixed(1)}</td><td style={{...TD,fontWeight:500,textAlign:"right"}}>{totCBM}</td><td colSpan={4} style={TD}></td></tr></tfoot>
            </table>
          </div>
          <div style={{marginTop:8,padding:"6px 10px",background:"var(--color-background-secondary)",borderRadius:"var(--border-radius-md)",fontSize:11,color:"var(--color-text-secondary)",border:"0.5px solid var(--color-border-tertiary)"}}>
            💡 라벨 분류는 PDF 내부의 Merchant SKU 텍스트를 직접 인식합니다. FBA ID는 선택사항이며, 마스터의 <strong>SKU(realSku)</strong> 값이 PDF 라벨 텍스트와 일치해야 합니다.
          </div>
        </div>
      )}

      {/* ══ 라벨 분류 ══ */}
      {mode==='label'&&(
        <div>
          <div style={{display:"flex",alignItems:"center",gap:8,marginBottom:14,flexWrap:"wrap"}}>
            <span style={{fontSize:12,color:"var(--color-text-secondary)"}}>FBA 카톤 라벨 PDF → SKU 자동 인식 → BOX CODE별 분류 (혼합 페이지 자동 마스킹)</span>
            {Object.keys(labelGroups).length>0&&(<button onClick={dlAllLabels} style={{marginLeft:"auto",fontSize:11,padding:"3px 12px",background:"var(--color-background-success)",border:"0.5px solid var(--color-border-success)",color:"var(--color-text-success)",borderRadius:"var(--border-radius-md)",cursor:"pointer",fontWeight:500}}>전체 다운로드</button>)}
          </div>

          {/* 업로드 드롭존 */}
          <div onClick={()=>!labelLoading&&labelRef.current?.click()} onDragOver={e=>e.preventDefault()} onDrop={e=>{e.preventDefault();const fs=Array.from(e.dataTransfer.files).filter(f=>f.type==='application/pdf');if(fs.length&&!labelLoading){setLabelFiles(fs);processLabels(fs)}}} style={{display:"flex",flexDirection:"column",alignItems:"center",justifyContent:"center",padding:"32px",border:"1px dashed var(--color-border-secondary)",borderRadius:"var(--border-radius-md)",cursor:labelLoading?"default":"pointer",marginBottom:16,gap:8,background:labelLoading?"var(--color-background-secondary)":"transparent"}}>
            <span style={{fontSize:32}}>{labelLoading?"⏳":"📄"}</span>
            <span style={{fontSize:13,color:"var(--color-text-secondary)",textAlign:"center"}}>{labelLoading?labelStatus:labelFiles.length?`${labelFiles.length}개 파일 처리 완료 (${labelStatus}) — 다시 업로드하려면 클릭`:"FBA 카톤 라벨 PDF 업로드 (여러 파일 동시 가능)"}</span>
            {!labelLoading&&(<span style={{fontSize:11,color:"var(--color-text-tertiary)"}}>페이지 내 SKU 텍스트 인식 · 혼합 페이지는 흰색 마스킹으로 분리</span>)}
          </div>
          <input ref={labelRef} type="file" accept="application/pdf" multiple style={{display:"none"}} onChange={e=>{const fs=Array.from(e.target.files||[]);if(fs.length){setLabelFiles(fs);processLabels(fs)}}}/>

          {/* 결과 테이블 */}
          {Object.keys(labelGroups).length>0&&!labelLoading&&(
            <div style={{border:"0.5px solid var(--color-border-tertiary)",borderRadius:"var(--border-radius-md)",overflow:"hidden",marginBottom:12}}>
              <table style={{width:"100%",borderCollapse:"collapse",fontSize:12}}>
                <thead><tr><th style={TH}>BOX CODE</th><th style={{...TH,textAlign:"right"}}>라벨 수</th><th style={{...TH,textAlign:"right"}}>파일 크기</th><th style={TH}></th></tr></thead>
                <tbody>{Object.entries(labelGroups).sort(([a],[b])=>a.localeCompare(b)).map(([loc,bytes],i)=>(<tr key={loc} style={{background:i%2===0?"transparent":"var(--color-background-secondary)"}}><td style={{...TD,fontWeight:500}}>{loc}</td><td style={{...TD,textAlign:"right"}}>{(labelCounts[loc]||0).toLocaleString()}장</td><td style={{...TD,textAlign:"right",color:"var(--color-text-secondary)"}}>{(bytes.length/1024).toFixed(0)}KB</td><td style={TD}><button onClick={()=>dlLabel(loc,bytes)} style={{fontSize:11,padding:"2px 10px",cursor:"pointer",borderRadius:4,border:"0.5px solid var(--color-border-secondary)",background:"transparent"}}>↓ PDF</button></td></tr>))}</tbody>
                <tfoot><tr style={{background:"var(--color-background-secondary)",borderTop:"1.5px solid var(--color-border-secondary)"}}><td style={{...TD,fontWeight:500}}>합계</td><td style={{...TD,textAlign:"right",fontWeight:500}}>{Object.values(labelCounts).reduce((s,c)=>s+c,0).toLocaleString()}장</td><td style={{...TD,textAlign:"right",fontWeight:500}}>{(Object.values(labelGroups).reduce((s,b)=>s+b.length,0)/1024).toFixed(0)}KB</td><td style={TD}></td></tr></tfoot>
              </table>
            </div>
          )}

          {/* 디버그 패널 */}
          {labelDebug.length>0&&(
            <div>
              <button onClick={()=>setShowDebug(v=>!v)} style={{fontSize:11,padding:"4px 12px",cursor:"pointer",border:"0.5px solid var(--color-border-tertiary)",borderRadius:"var(--border-radius-md)",background:"transparent",color:"var(--color-text-secondary)",marginBottom:8,display:"flex",alignItems:"center",gap:6}}>
                <span>{showDebug?"▲":"▼"}</span><span>디버그 패널 (PDF 내 SKU 원문 확인용)</span>
                <span style={{padding:"1px 6px",borderRadius:10,background:"rgba(16,185,129,0.15)",color:"var(--color-text-success)",fontSize:10}}>✅ {labelDebug.filter(d=>d.matchedLoc).length}페이지 매칭</span>
                <span style={{padding:"1px 6px",borderRadius:10,background:"rgba(239,68,68,0.1)",color:"var(--color-text-danger)",fontSize:10}}>❌ {labelDebug.filter(d=>!d.matchedLoc).length}페이지 미매칭</span>
              </button>
              {showDebug&&(
                <div style={{border:"0.5px solid var(--color-border-tertiary)",borderRadius:"var(--border-radius-md)",overflow:"hidden",fontSize:11}}>
                  {/* 마스터 SKU 비교 참고표 */}
                  <div style={{padding:"8px 12px",background:"rgba(219,234,254,0.3)",borderBottom:"1px solid var(--color-border-tertiary)",fontSize:10}}>
                    <span style={{fontWeight:500,marginRight:8}}>마스터 realSku 정규화 값:</span>
                    {Object.values(master).filter(m=>m.sku&&m.loc).map(m=>(
                      <span key={m.sku} style={{marginRight:6,padding:"1px 5px",borderRadius:4,background:"rgba(219,234,254,0.5)",fontFamily:"var(--font-mono)"}}>
                        {normalizeSku(m.sku)}
                      </span>
                    ))}
                  </div>
                  <div style={{maxHeight:400,overflowY:"auto"}}>
                    <table style={{width:"100%",borderCollapse:"collapse"}}>
                      <thead>
                        <tr style={{position:"sticky",top:0}}>
                          <th style={{...TH,fontSize:10,padding:"5px 8px"}}>파일</th>
                          <th style={{...TH,fontSize:10,padding:"5px 8px",textAlign:"right"}}>P</th>
                          <th style={{...TH,fontSize:10,padding:"5px 8px"}}>PDF 원문 SKU</th>
                          <th style={{...TH,fontSize:10,padding:"5px 8px"}}>정규화 (매칭용)</th>
                          <th style={{...TH,fontSize:10,padding:"5px 8px",textAlign:"center"}}>라벨수</th>
                          <th style={{...TH,fontSize:10,padding:"5px 8px"}}>BOX CODE</th>
                        </tr>
                      </thead>
                      <tbody>
                        {labelDebug.map((d,i)=>(<tr key={i} style={{background:d.matchedLoc?(i%2===0?"transparent":"var(--color-background-secondary)"):"rgba(254,202,202,0.3)"}}>
                          <td style={{...TD,fontSize:10,color:"var(--color-text-tertiary)",maxWidth:130,padding:"3px 8px"}}>{d.file}</td>
                          <td style={{...TD,fontSize:10,textAlign:"right",padding:"3px 8px"}}>{d.page}</td>
                          <td style={{...TD,fontSize:10,fontFamily:"var(--font-mono)",padding:"3px 8px",maxWidth:220,color:d.matchedLoc?"var(--color-text-primary)":"var(--color-text-danger)"}}>
                            {d.skusOnPage.join(' / ')||"(SKU 없음)"}
                            {d.skusOnPage.length>1&&<span style={{marginLeft:4,padding:"1px 4px",borderRadius:4,background:"rgba(234,179,8,0.2)",color:"#92400e",fontSize:9}}>혼합</span>}
                          </td>
                          <td style={{...TD,fontSize:10,fontFamily:"var(--font-mono)",padding:"3px 8px",maxWidth:200,color:"var(--color-text-secondary)"}}>
                            {d.skusNorm.join(' / ')||"—"}
                          </td>
                          <td style={{...TD,fontSize:10,textAlign:"center",padding:"3px 8px"}}>{d.labelCount}</td>
                          <td style={{...TD,fontSize:10,padding:"3px 8px",color:d.matchedLoc?"var(--color-text-info)":"var(--color-text-danger)",fontWeight:d.matchedLoc?500:400}}>
                            {d.matchedLoc||"❌ 미매칭"}
                          </td>
                        </tr>))}
                      </tbody>
                    </table>
                  </div>
                </div>
              )}
            </div>
          )}
        </div>
      )}

      {/* ══ STEP 3 ══ */}
      {mode==='3'&&(
        <div>
          <div style={{display:"flex",alignItems:"center",gap:8,marginBottom:14}}><span style={{fontSize:12,color:"var(--color-text-secondary)"}}>업로드 컨테이너별 분류 · CONTAINER/SEAL 직접 입력</span><button onClick={expS3} style={{marginLeft:"auto",fontSize:11,padding:"3px 10px"}}>xlsx 저장</button></div>
          <SheetTabs/>
          {!file?(<p style={{color:"var(--color-text-tertiary)",fontSize:13,textAlign:"center",padding:"3rem 0"}}>파일을 먼저 업로드해주세요</p>):(
            <div style={{overflowX:"auto",border:"0.5px solid var(--color-border-tertiary)",borderRadius:"var(--border-radius-md)"}}>
              <table style={{width:"100%",borderCollapse:"collapse",fontSize:12,minWidth:950}}>
                <thead><tr><th style={{...TH,minWidth:100}}>CONTAINER</th><th style={{...TH,minWidth:90,borderRight:"2px solid var(--color-border-secondary)"}}>SEAL NO.</th>{["약호","BOX CODE","ASIN","FBA ID","아마존 ID","FC CENTER","PLT","CT","CTN/PLT","G.W(kg)","CBM"].map(h=>(<th key={h} style={TH}>{h}</th>))}</tr></thead>
                <tbody>
                  {buildS3groups().map(g=>{
                    const rows=g.rows,ci=(g.no-1)%CB.length
                    const sumQ=rows.reduce((s,r)=>s+r.quantity,0)
                    const sumP=rows.reduce((s,r)=>{const cpp=parseFloat(String((master[String(r.sku)]||{}).cpp))||16;return s+Math.ceil(r.quantity/cpp)},0)
                    const sumW=rows.reduce((s,r)=>s+calcGW(r.quantity,parseFloat(String((master[String(r.sku)]||{}).kg))||0),0)
                    const sumCBM=Math.round(rows.reduce((s,r)=>{const m=master[String(r.sku)]||({} as MasterItem);return s+calcCBM(r.quantity,parseFloat(String(m.bx))||0,parseFloat(String(m.by))||0,parseFloat(String(m.bz))||0)},0)*100)/100
                    return[...rows.map((r,i)=>{const m=master[String(r.sku)]||({} as MasterItem);const cpp=parseFloat(String(m.cpp))||16;const sk=s2meta[String(r.sku)]||{};const gw=calcGW(r.quantity,parseFloat(String(m.kg))||0);const cbm=calcCBM(r.quantity,parseFloat(String(m.bx))||0,parseFloat(String(m.by))||0,parseFloat(String(m.bz))||0);return(<tr key={g.no+"_"+i} style={{background:i%2===0?"transparent":"var(--color-background-secondary)"}}>
                      {i===0&&(<><td rowSpan={rows.length+1} style={{...TD,verticalAlign:"middle",textAlign:"center",background:CB[ci],padding:"6px 8px",borderRight:"0.5px solid var(--color-border-tertiary)"}}><div style={{fontSize:10,color:CT[ci],opacity:0.7,marginBottom:3}}>컨{g.no}</div><input value={ctnMeta[g.no]?.container||""} onChange={e=>updCtnMeta(g.no,"container",e.target.value)} placeholder="CONTAINER" style={{...INP,textAlign:"center",fontWeight:500,fontSize:11,color:CT[ci],borderBottom:"1px solid "+CT[ci]+"66",width:88}}/></td><td rowSpan={rows.length+1} style={{...TD,verticalAlign:"middle",textAlign:"center",background:CB[ci],padding:"6px 8px",borderRight:"2px solid var(--color-border-secondary)"}}><input value={ctnMeta[g.no]?.sealNo||""} onChange={e=>updCtnMeta(g.no,"sealNo",e.target.value)} placeholder="SEAL NO." style={{...INP,textAlign:"center",fontSize:11,color:CT[ci],borderBottom:"1px solid "+CT[ci]+"66",width:78}}/></td></>)}
                      <td style={{...TD,fontWeight:500,minWidth:130}}>{String(r.sku)}</td><td style={{...TD,color:"var(--color-text-info)",fontWeight:500,minWidth:70}}>{String(r.location||m.loc||"")}</td><td style={{...TD,fontFamily:"var(--font-mono)",fontSize:11,minWidth:100}}>{m.asin||""}</td><td style={{...TD,fontFamily:"var(--font-mono)",fontSize:11,minWidth:110}}>{sk.fbaId||""}</td><td style={{...TD,fontFamily:"var(--font-mono)",fontSize:11,minWidth:90}}>{sk.amazonId||""}</td><td style={{...TD,fontWeight:500,minWidth:70}}>{sk.fc||""}</td><td style={{...TD,textAlign:"right",minWidth:40}}>{Math.ceil(r.quantity/cpp)}</td><td style={{...TD,textAlign:"right",fontWeight:500,minWidth:40}}>{r.quantity.toLocaleString()}</td><td style={{...TD,textAlign:"center",color:"var(--color-text-secondary)"}}>{cpp}CTN/PLT</td><td style={{...TD,textAlign:"right",color:"var(--color-text-secondary)"}}>{gw.toFixed(1)}</td><td style={{...TD,textAlign:"right",color:"var(--color-text-secondary)"}}>{cbm}</td>
                    </tr>)}),
                    <tr key={"sub_"+g.no} style={{background:CB[ci],borderTop:"1px solid "+CT[ci]+"44"}}><td colSpan={6} style={{...TD,color:CT[ci],fontWeight:500,textAlign:"right",fontSize:11}}>{(g.container||"컨"+g.no)+" 소계"}{g.shipment_date?" · "+g.shipment_date:""}{g.destination?"  "+g.destination:""}{g.etd?"  ETD "+g.etd:""}</td><td style={{...TD,color:CT[ci],fontWeight:500,textAlign:"right"}}>{sumP}</td><td style={{...TD,color:CT[ci],fontWeight:500,textAlign:"right"}}>{sumQ.toLocaleString()}</td><td style={TD}></td><td style={{...TD,color:CT[ci],fontWeight:500,textAlign:"right"}}>{sumW.toFixed(1)}</td><td style={{...TD,color:CT[ci],fontWeight:500,textAlign:"right"}}>{sumCBM}</td></tr>]
                  })}
                  {(()=>{const gC=fd.reduce((s,r)=>s+r.quantity,0);const gP=fd.reduce((s,r)=>{const cpp=parseFloat(String((master[String(r.sku)]||{}).cpp))||16;return s+Math.ceil(r.quantity/cpp)},0);const gW=fd.reduce((s,r)=>s+calcGW(r.quantity,parseFloat(String((master[String(r.sku)]||{}).kg))||0),0);const gCBM=Math.round(fd.reduce((s,r)=>{const m=master[String(r.sku)]||({} as MasterItem);return s+calcCBM(r.quantity,parseFloat(String(m.bx))||0,parseFloat(String(m.by))||0,parseFloat(String(m.bz))||0)},0)*100)/100;return(<tr style={{background:"var(--color-background-secondary)",borderTop:"2px solid var(--color-border-secondary)"}}><td colSpan={8} style={{...TD,fontWeight:500,textAlign:"right"}}>전체 합계</td><td style={{...TD,fontWeight:500,textAlign:"right"}}>{gP}</td><td style={{...TD,fontWeight:500,textAlign:"right"}}>{gC.toLocaleString()}</td><td style={TD}></td><td style={{...TD,fontWeight:500,textAlign:"right"}}>{gW.toFixed(1)}</td><td style={{...TD,fontWeight:500,textAlign:"right"}}>{gCBM}</td></tr>)})()}
                </tbody>
              </table>
            </div>
          )}
        </div>
      )}

      {/* ══ 매핑 관리 ══ */}
      {mode==='map'&&(
        <div>
          <div style={{display:"flex",alignItems:"center",gap:8,marginBottom:14}}>
            <span style={{fontSize:13,fontWeight:500}}>SKU 마스터</span><span style={{fontSize:12,color:"var(--color-text-tertiary)"}}>{Object.keys(master).length}개</span>
            <button onClick={()=>{if(!newSku.sku)return;setMaster(p=>({...p,[newSku.sku]:{sku:newSku.realSku||newSku.sku,asin:newSku.asin,to:newSku.to,loc:newSku.loc,cpp:+newSku.cpp,fba:+newSku.fba,price:+newSku.price,kg:+newSku.kg,bx:+newSku.bx,by:+newSku.by,bz:+newSku.bz}}));setNewSku({...EMPTY_SKU})}} style={{marginLeft:"auto",fontSize:11,padding:"3px 10px"}}>+ 추가</button>
            <button onClick={()=>{if(confirm("초기화하시겠습니까?"))setMaster(Object.fromEntries(Object.entries(MASTER_INIT).map(([k,v])=>[k,{...v}])))}} style={{fontSize:11,padding:"3px 10px",color:"var(--color-text-danger)",border:"0.5px solid var(--color-border-danger)",borderRadius:"var(--border-radius-md)",background:"transparent",cursor:"pointer"}}>초기화</button>
            <button onClick={()=>xlsDl(Object.entries(master).map(([sku,m])=>({약호:sku,SKU:m.sku,ASIN:m.asin,구박스:m.to,신박스:m.loc,PLT당카톤:m.cpp,FBA비용:m.fba,판매가:m.price,무게kg:m.kg,박스가로:m.bx,박스세로:m.by,박스높이:m.bz})),"마스터","SKU마스터.xlsx")} style={{fontSize:11,padding:"3px 10px"}}>xlsx 저장</button>
          </div>
          <div style={{marginBottom:8,padding:"5px 10px",background:"var(--color-background-secondary)",borderRadius:"var(--border-radius-md)",fontSize:11,color:"var(--color-text-secondary)"}}>💾 마스터 데이터는 브라우저에 자동 저장됩니다 (재배포 후에도 유지)</div>
          <div style={{overflowX:"auto",border:"0.5px solid var(--color-border-tertiary)",borderRadius:"var(--border-radius-md)"}}>
            <table style={{width:"100%",borderCollapse:"collapse",fontSize:12}}>
              <thead><tr>{["약호","SKU","ASIN","구박스","신박스","PLT당카톤","FBA비용(¥)","판매가(¥)","무게(kg)","박스가로","박스세로","박스높이",""].map(h=><th key={h} style={TH}>{h}</th>)}</tr></thead>
              <tbody>
                {Object.entries(master).map(([sku,m],i)=>(<tr key={sku} style={{background:i%2===0?"transparent":"var(--color-background-secondary)"}}><td style={{...TD,fontWeight:500,minWidth:130}}>{sku}</td><td style={{...TD,...IBLU,minWidth:160}}><input value={m.sku??""} onChange={e=>updM(sku,"sku",e.target.value)} style={{...INP,color:m.sku?"var(--color-text-primary)":"var(--color-text-secondary)"}} placeholder="Merchant SKU"/></td>{(["asin","to","loc"] as const).map(f=>(<td key={f} style={{...TD,...IBLU}}><input value={m[f]||""} onChange={e=>updM(sku,f,e.target.value)} style={INP}/></td>))}{(["cpp","fba","price","kg","bx","by","bz"] as const).map(f=>(<td key={f} style={{...TD,...IBLU}}><input type="number" value={m[f]||0} onChange={e=>updM(sku,f,Number(e.target.value))} style={{...INP,textAlign:"right"}}/></td>))}<td style={TD}><button onClick={()=>setMaster(p=>{const n={...p};delete n[sku];return n})} style={{fontSize:11,padding:"1px 6px",cursor:"pointer",borderRadius:4,background:"transparent",color:"var(--color-text-danger)",border:"0.5px solid var(--color-border-danger)"}}>삭제</button></td></tr>))}
              </tbody>
            </table>
          </div>
        </div>
      )}
    </div>
  )
}
