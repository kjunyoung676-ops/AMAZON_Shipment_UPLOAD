"use client"

import { useState, useRef, useEffect } from "react"
import * as XLSX from "xlsx"

const SKU_ORDER = [
  "HS604012-4B","HS604012-4W",
  "HS904018-5B","HS904018-5W",
  "HS124018Z-5B","HS124018Z-5W",
  "HS604015W-5B","HS604015W-5W",
  "HS904016.5W-5B","HS904016.5W-5W",
  "HS124518W-5B","HS124518W-5W",
  "HS604015HP-5W","HS804018HP-5W",
  "HS604015HPW-5W","HS804016.5HPW-5W",
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
  shipment_date: ["출하 계획(공장)","출하 계획","출하계획"],
  shipment_time: ["출하 시간","출하시간"],
  destination:   ["국가","목적지","도착지"],
  model_code:    ["BOM 품번","품번"],
  sku:           ["품목"],
  color:         ["색상"],
  quantity:      ["수량","qty"],
  qty_total:     ["수량 합계","합계"],
  ctn_count:     ["컨 수량"],
  ft:            ["FT","컨테이너"],
  type_old:      ["TYPE (구)"],
  location:      ["TYPE (신)","로케이션","신박스코드"],
  etd:           ["출항(ETD)"],
  eta:           ["도착(ETA)"],
  forwarding:    ["포워딩"],
  carrier:       ["선사"],
  booking_ref:   ["부킹여부","부킹","Booking","booking"],
}
const FF_COLS = ["shipment_date","shipment_time","destination","etd","eta","qty_total","ctn_count","ft","sku","model_code","color","booking_ref"]
const CB = ["#dbeafe","#d1fae5","#fef3c7","#fce7f3","#ede9fe","#ffedd5","#e0f2fe","#dcfce7"]
const CT = ["#1e40af","#065f46","#92400e","#9d174d","#4c1d95","#9a3412","#075985","#166534"]
const TH: React.CSSProperties = {padding:"7px 10px",textAlign:"left",fontWeight:500,fontSize:11,color:"var(--color-text-primary)",borderBottom:"1.5px solid var(--color-border-secondary)",whiteSpace:"nowrap",background:"var(--color-background-secondary)",position:"sticky",top:0}
const TD: React.CSSProperties = {padding:"5px 10px",borderBottom:"0.5px solid var(--color-border-tertiary)",fontSize:12,overflow:"hidden",textOverflow:"ellipsis",whiteSpace:"nowrap",color:"var(--color-text-primary)"}
const INP: React.CSSProperties = {border:"none",background:"transparent",fontSize:12,padding:"1px 0",outline:"none",color:"var(--color-text-primary)",fontFamily:"inherit",width:"100%"}
const IBLU: React.CSSProperties = {background:"rgba(219,234,254,0.35)"}
const EMPTY_SKU = {sku:"",realSku:"",asin:"",to:"",loc:"",cpp:16,fba:0,price:0,kg:0,bx:0,by:0,bz:0}

const LS_KEY = "shipment_app_v2"
// 시트별 저장: key = LS_KEY + "_" + sheetName + "_" + dataKey
function lsLoad<T>(key: string, fallback: T): T {
  try { const r = localStorage.getItem(key); if (r) return JSON.parse(r) as T } catch { /**/ }
  return fallback
}
function lsSave(key: string, value: unknown) {
  try { localStorage.setItem(key, JSON.stringify(value)) } catch { /**/ }
}
// 시트별 키 생성
function sheetKey(sheetName: string, dataKey: string) {
  return `${LS_KEY}_${sheetName}_${dataKey}`
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
  // 기존: 병합 셀 기반 그룹 감지
  for (const m of merges) if (m.s.c<=2&&m.e.r>m.s.r&&m.s.r>hr) gsr.add(m.s.r)
  const raw=XLSX.utils.sheet_to_json(ws,{defval:"",cellDates:true}) as Record<string,unknown>[]
  if (!raw.length) return []

  // 간소화 양식 감지: '컨테이너' 컬럼이 있으면 간소화 양식
  const headers = Object.keys(raw[0])
  const ctnCol = headers.find(h=>h.includes('컨테이너'))
  const isSimplified = !!ctnCol

  return raw.map((r,i)=>{
    const nr:RowData={quantity:0}
    for(const [k,v] of Object.entries(r)) nr[k]=fmtDate(v)

    if (isSimplified && ctnCol) {
      // 컨테이너 컬럼에 값(40FT 등)이 있으면 새 컨테이너 시작
      const ctnVal = String(r[ctnCol]||'').trim()
      nr.__groupStart = ctnVal !== '' && ctnVal !== 'null'
    } else {
      nr.__groupStart = gsr.has(i+hr+1)
    }
    return nr
  })
}
function forwardFill(rows:RowData[], masterRef?: Record<string,{loc:string}>):RowData[] {
  const res=rows.map(r=>({...r}));const last:Record<string,unknown>={};let ctnNo=0
  for (const row of res) {
    if (row.__groupStart) ctnNo++;if (ctnNo===0) ctnNo=1
    for (const col of FF_COLS){const v=row[col];if(v!==undefined&&v!==''&&v!==null)last[col]=v;else if(col in last)row[col]=last[col]}
    row.container_no=ctnNo
    // 신박스코드(location)가 없으면 마스터 loc에서 자동 채우기
    if ((!row.location||row.location==='') && masterRef) {
      const sku=String(row.sku||'')
      if (sku && masterRef[sku]?.loc) row.location=masterRef[sku].loc
    }
    // TYPE(구)(type_old)가 없으면 마스터 to에서 자동 채우기
    if ((!row.type_old||row.type_old==='') && masterRef) {
      const sku=String(row.sku||'')
      if (sku && (masterRef[sku] as {to?:string})?.to) row.type_old=(masterRef[sku] as {to?:string}).to
    }
  }
  return res
}
function xlsDl(data:unknown[],sn:string,fn:string){const wb=XLSX.utils.book_new();const ws=Array.isArray(data[0])?XLSX.utils.aoa_to_sheet(data as unknown[][]):XLSX.utils.json_to_sheet(data as Record<string,unknown>[]);XLSX.utils.book_append_sheet(wb,ws,sn||"data");XLSX.writeFile(wb,fn)}

// ── JSZip 외과적 XLSX 패치 헬퍼 (모듈 레벨) ─────────────────────────────
// 특정 셀을 inline string으로 교체 (이미지/서식/드로잉 손대지 않음)
function xmlSetCell(xml:string, cellRef:string, value:string):string{
  const esc=value.replace(/&/g,'&amp;').replace(/</g,'&lt;').replace(/>/g,'&gt;')
  const re=new RegExp(`<c\\s+r="${cellRef}"([^>]*)(?:/>|>[\\s\\S]*?</c>)`)
  if(re.test(xml)){
    return xml.replace(re,(_m,attrs)=>{
      const s=(attrs.match(/\bs="(\d+)"/)??[])[1]
      return `<c r="${cellRef}"${s?` s="${s}"`:''}  t="inlineStr"><is><t>${esc}</t></is></c>`
    })
  }
  // 셀이 없으면 해당 행에 삽입
  const rowNum=(cellRef.match(/\d+/)??[])[0]
  if(!rowNum) return xml
  const rowRe=new RegExp(`(<row[^>]+\\br="${rowNum}"[^>]*>)`)
  if(rowRe.test(xml)) return xml.replace(rowRe,`$1<c r="${cellRef}" t="inlineStr"><is><t>${esc}</t></is></c>`)
  return xml
}
function xmlClearCell(xml:string, cellRef:string):string{ return xmlSetCell(xml,cellRef,'') }

// workbook.xml + _rels → { 시트이름: 'xl/worksheets/sheetN.xml' }
async function getSheetFileMap(zip:{file:(p:string)=>({async:(t:'text')=>Promise<string>})|null}):Promise<Map<string,string>>{
  const wbXml=await zip.file('xl/workbook.xml')!.async('text')
  const relsXml=await zip.file('xl/_rels/workbook.xml.rels')!.async('text')
  const nameToRid=new Map<string,string>()
  for(const m of wbXml.matchAll(/<sheet\s[^>]*name="([^"]+)"[^>]*r:id="([^"]+)"/g)) nameToRid.set(m[1],m[2])
  const ridToPath=new Map<string,string>()
  for(const m of relsXml.matchAll(/<Relationship\s[^>]*Id="([^"]+)"[^>]*Target="([^"]+)"/g)) ridToPath.set(m[1],m[2])
  const result=new Map<string,string>()
  for(const [name,rid] of nameToRid){
    const rel=ridToPath.get(rid)
    if(rel) result.set(name, rel.startsWith('/')?rel.slice(1):`xl/${rel}`)
  }
  return result
}

export default function ShipmentApp() {
  const [sheets,setSh]=useState<Record<string,RowData[]>>({})
  const [sheetNames,setSn]=useState<string[]>([])
  const [activeSheet,setAs]=useState<string|null>(null)
  const [file,setFile]=useState<File|null>(null)
  const [mode,setMode]=useState('1')
  const [master,setMaster]=useState<Record<string,MasterItem>>(()=>lsLoad(LS_KEY+"_master",Object.fromEntries(Object.entries(MASTER_INIT).map(([k,v])=>[k,{...v}]))))
  // s2meta/ctnMeta/pltNotes: 시트별로 저장 — activeSheet 변경 시 로드
  const [s2meta,setS2meta]=useState<Record<string,Record<string,string>>>({})
  const [ctnMeta,setCtnMeta]=useState<Record<number,Record<string,string>>>({})
  const [pltNotes,setPltNotes]=useState<Record<string,string>>({})
  // BL별 입항지 선택: key=blId, value='kansai'|'kanto'
  const [blPortMode,setBlPortMode]=useState<Record<string,string>>(()=>lsLoad(LS_KEY+"_blPortMode",{}))
  useEffect(()=>{if(activeSheet)lsSave(sheetKey(activeSheet,"blPortMode"),blPortMode)},[blPortMode,activeSheet])

  // 시트 변경 시 해당 시트의 저장 데이터 로드
  useEffect(()=>{
    if(!activeSheet)return
    setS2meta(lsLoad(sheetKey(activeSheet,"s2meta"),{}))
    setCtnMeta(lsLoad(sheetKey(activeSheet,"ctnMeta"),{}))
    setPltNotes(lsLoad(sheetKey(activeSheet,"pltNotes"),{}))
    setBlPortMode(lsLoad(sheetKey(activeSheet,"blPortMode"),{}))
  },[activeSheet])

  useEffect(()=>{if(activeSheet)lsSave(sheetKey(activeSheet,"s2meta"),s2meta)},[s2meta,activeSheet])
  useEffect(()=>{if(activeSheet)lsSave(sheetKey(activeSheet,"ctnMeta"),ctnMeta)},[ctnMeta,activeSheet])
  useEffect(()=>{if(activeSheet)lsSave(sheetKey(activeSheet,"pltNotes"),pltNotes)},[pltNotes,activeSheet])
  const [coll,setColl]=useState<Record<string,boolean>>({})
  const [newSku,setNewSku]=useState({...EMPTY_SKU})
  const [fbaLoading,setFbaLoading]=useState(false)
  const [fbaUploadCtnFilter,setFbaUploadCtnFilter]=useState<number[]>([])
  const [labelFiles,setLabelFiles]=useState<File[]>([])
  const [pltLabelFiles,setPltLabelFiles]=useState<File[]>([])
  const [labelGroups,setLabelGroups]=useState<Record<string,Uint8Array>>({})
  const [labelCounts,setLabelCounts]=useState<Record<string,number>>({})
  const [labelLoading,setLabelLoading]=useState(false)
  const [labelStatus,setLabelStatus]=useState("")
  const [labelDebug,setLabelDebug]=useState<LabelDebugEntry[]>([])
  const [showDebug,setShowDebug]=useState(false)
  const fileRef=useRef<HTMLInputElement>(null)
  const labelRef=useRef<HTMLInputElement>(null)
  const pltLabelRef=useRef<HTMLInputElement>(null)
  const metaJsonRef=useRef<HTMLInputElement>(null)

  // ── FBA Inbound SP-API 상태 ───────────────────────────────────
  interface FbaStep { status:'idle'|'running'|'done'|'error'|'waiting'; msg?:string; data?:unknown }
  interface PlacementOption { placementOptionId:string; status:string; fees?:unknown; shipmentIds?:string[] }
  interface TransportOption { transportOptionId:string; carrierId?:string; carrierName?:string; shippingMode?:string; quote?:unknown }
  const FBA_STEPS=['createPlan','packingOptions','setPackingInfo','placementOptions','transport','labels','bol'] as const
  type FbaStepKey=typeof FBA_STEPS[number]
  const initSteps=():Record<FbaStepKey,FbaStep>=>
    Object.fromEntries(FBA_STEPS.map(s=>[s,{status:'idle'}])) as Record<FbaStepKey,FbaStep>
  const [fbaSteps,setFbaSteps]=useState<Record<FbaStepKey,FbaStep>>(initSteps)
  const [fbaPlanId,setFbaPlanId]=useState('')
  const [fbaShipmentIds,setFbaShipmentIds]=useState<string[]>([])
  const [fbaPlacementOptions,setFbaPlacementOptions]=useState<PlacementOption[]>([])
  const [fbaSelectedPlacement,setFbaSelectedPlacement]=useState('')
  const [fbaTransportOptions,setFbaTransportOptions]=useState<Record<string,TransportOption[]>>({})
  const [fbaLabelUrls,setFbaLabelUrls]=useState<Record<string,string>>({})
  const [fbaBolUrls,setFbaBolUrls]=useState<Record<string,string>>({})
  const [fbaRunning,setFbaRunning]=useState(false)
  const [fbaWarehouse,setFbaWarehouse]=useState<'saitama'|'osaka'>('saitama')
  const FBA_WAREHOUSES={saitama:{label:'사이타마 창고',city:'Saitama'},osaka:{label:'오사카 창고',city:'Osaka'}} as const
  const [fbaShippingMode,setFbaShippingMode]=useState<'FTL'|'LTL'|'SP'>('FTL')
  // 출하 예정일 — readyToShipWindow.start 로 사용 (기본: 오늘)
  const [fbaShipDate,setFbaShipDate]=useState<string>(()=>new Date().toISOString().slice(0,10))
  const FBA_SHIPPING_MODES=[
    {key:'FTL' as const, label:'FTL', desc:'Full Truck Load'},
    {key:'LTL' as const, label:'LTL', desc:'Less Than Truck Load'},
    {key:'SP'  as const, label:'SP',  desc:'Small Parcel'},
  ]

  // ── 총합 탭 상태 ──────────────────────────────────────────────────────────
  const [totalExcludedSheets,setTotalExcludedSheets]=useState<Set<string>>(()=>new Set(lsLoad(LS_KEY+'_totalExcl',[] as string[])))
  const [totalAdjustments,setTotalAdjustments]=useState<Record<string,number>>(()=>lsLoad(LS_KEY+'_totalAdj',{} as Record<string,number>))
  useEffect(()=>{lsSave(LS_KEY+'_totalExcl',[...totalExcludedSheets])},[totalExcludedSheets])
  useEffect(()=>{lsSave(LS_KEY+'_totalAdj',totalAdjustments)},[totalAdjustments])

  // ── CI/PL 변환기 상태 ────────────────────────────────────────────────────
  const [ciplFile,setCiplFile]=useState<File|null>(null)
  const [ciplWB,setCiplWB]=useState<XLSX.WorkBook|null>(null)
  const [ciplPort,setCiplPort]=useState<'kanto'|'kansai'>('kanto')
  const ciplRef=useRef<HTMLInputElement>(null)
  const ciplRawRef=useRef<ArrayBuffer|null>(null)

  // ── 납품처 정보 상태 ─────────────────────────────────────────────────────
  const [dlBL_Saitama,setDlBL_Saitama]=useState('')
  const [dlBL_Osaka,setDlBL_Osaka]=useState('')

  // ── CI/PL 변환기 ────────────────────────────────────────────────────────
  // 일본용 CI/PL 고정 상수
  const JP_CONSIGNEE_LINES=['Homedant Co., Ltd. (CJ LOGISTICS JAPAN ON BEHALF OF)','367-26 Daegotbuk-ro, Daegot-myeon, ','Gimpo-si, Gyeonggi-do','SOUTH KOREA','+82 70-4157-3393','global@speedrack.kr']
  const JP_ACP_LINES=['CJ LOGISTICS JAPAN CORP.','TOKYO TO MINATO KU NISHISHINBASHI 2-7-4 CJ BLDG.,7F','TEL : 03-3500-5842','ACP No. : 1000-25-0891']
  const JP_NOTIFY_KANTO=['CJ LOGISTICS JAPAN CORP. HONSHA HM.BYUN','5F, 1662 Shimohayami, Kuki-shi, Saitama,','346-0022, Japan','TEL: 0335005842','keunje.park@cj.net']
  const JP_NOTIFY_KANSAI=['CJ LOGISTICS JAPAN OSAKA BRANCH','4F Prologis park 5, 8-4-47','Nankohigashi Suminoe-ku Osaka-city Osaka 559-0031 Japan','+81-6-6690-0217','Keunje Park, kj_park@jhss-jp.com']

  function setCiplCell(ws:XLSX.WorkSheet, addr:string, v:string){
    if(!ws[addr]) ws[addr]={t:'s',v,w:v}
    else{ ws[addr].v=v; ws[addr].w=v; delete ws[addr].f }
  }
  function clearCiplCell(ws:XLSX.WorkSheet, addr:string){
    if(ws[addr]){ ws[addr].v=''; ws[addr].w=''; delete ws[addr].f }
  }

  function loadCiplFile(f:File){
    setCiplFile(f)
    const reader=new FileReader()
    reader.onload=e=>{
      const buf=e.target!.result as ArrayBuffer
      ciplRawRef.current=buf
      const wb=XLSX.read(new Uint8Array(buf),{type:'array',cellStyles:true,cellFormula:true})
      setCiplWB(wb)
    }
    reader.readAsArrayBuffer(f)
  }

  async function downloadCipl(){
    if(!ciplWB||!ciplRawRef.current) return
    // SheetJS로 파싱된 값에서 미리 읽기 (쓰기는 JSZip이 담당)
    const inv=ciplWB.Sheets['Invoice ']
    const pack=ciplWB.Sheets['Packing ']
    if(!inv||!pack){ alert('Invoice 또는 Packing 시트를 찾을 수 없습니다'); return }
    const cntrCount=String(pack['H18']?.v||'')  // H18 덮기 전에 먼저 저장
    const invNo=String(inv['H3']?.v||ciplFile?.name.replace('.xlsx','')||'JP')
    const notify=ciplPort==='kanto'?JP_NOTIFY_KANTO:JP_NOTIFY_KANSAI

    // JSZip으로 원본 파일 열기 (이미지/서식/드로잉 100% 보존)
    const {default:JSZip}=await import('jszip')
    const zip=await JSZip.loadAsync(ciplRawRef.current)
    const sheetMap=await getSheetFileMap(zip)

    async function patchSheet(name:string, patches:[string,string][], clears:string[]){
      const path=sheetMap.get(name)
      if(!path){ console.warn(`시트 없음: "${name}"`); return }
      const f=zip.file(path); if(!f) return
      let xml=await f.async('text')
      for(const [cell,val] of patches) xml=xmlSetCell(xml,cell,val)
      for(const cell of clears) xml=xmlClearCell(xml,cell)
      zip.file(path,xml)
    }

    // Invoice 시트
    await patchSheet('Invoice ',[
      ['A8','global@speedrack.kr'],
      ['A10',JP_CONSIGNEE_LINES[0]],['A11',JP_CONSIGNEE_LINES[1]],
      ['A12',JP_CONSIGNEE_LINES[2]],['A13',JP_CONSIGNEE_LINES[3]],
      ['A14',JP_CONSIGNEE_LINES[4]],['A15',JP_CONSIGNEE_LINES[5]],
      ['H15','(10)ACP information'],
      ['H16',JP_ACP_LINES[0]],['H17',JP_ACP_LINES[1]],
      ['H18',JP_ACP_LINES[2]],['H19',JP_ACP_LINES[3]],
      ['H20','(11)Country of Origin'],['H21','SOUTH KOREA'],
      ['H22','(12)Terms of delivery and payment'],['H23','Shipping condition : CFR'],
    ],[])

    // Packing 시트
    const packPatches:[string,string][]=[
      ['A8','global@speedrack.kr'],
      ['A10',JP_CONSIGNEE_LINES[0]],['A11',JP_CONSIGNEE_LINES[1]],
      ['A12',JP_CONSIGNEE_LINES[2]],['A13',JP_CONSIGNEE_LINES[3]],
      ['A14',JP_CONSIGNEE_LINES[4]],['A15',JP_CONSIGNEE_LINES[5]],
      ['H10',notify[0]],['H11',notify[1]],['H12',notify[2]],
      ['H13',notify[3]],['H14',notify[4]],
      ['H16','(9)ACP information'],
      ['H17',JP_ACP_LINES[0]],['H18',JP_ACP_LINES[1]],
      ['H19',JP_ACP_LINES[2]],['H20',JP_ACP_LINES[3]],
      ['H22','(10)Other references'],
    ]
    if(cntrCount) packPatches.push(['H23',cntrCount])
    await patchSheet('Packing ',packPatches,['H15','H21',...(!cntrCount?['H23']:[])])

    // 다운로드 (이미지/드로잉은 zip 내에서 원본 그대로 유지됨)
    const out=await zip.generateAsync({type:'arraybuffer'})
    const blob=new Blob([out],{type:'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'})
    const a=document.createElement('a'); a.href=URL.createObjectURL(blob)
    a.download=`CI_PL_${invNo}_일본용.xlsx`; a.click()
  }

  // ── 납품처 세관 양식 다운로드 ──────────────────────────────────────────
  async function downloadDelivery(type:'saitama'|'osaka', blNo:string){
    if(!blNo.trim()){ alert('B/L 번호를 입력해주세요'); return }
    const {default:JSZip}=await import('jszip')
    const url=`/templates/delivery_${type==='saitama'?'SAITAMA':'OSAKA'}.xlsx`
    const res=await fetch(url)
    if(!res.ok){ alert('템플릿 파일을 불러올 수 없습니다'); return }
    const buf=await res.arrayBuffer()
    // JSZip으로 열기 → B6만 외과적으로 교체 → 정부 세관 양식 100% 원형 유지
    const zip=await JSZip.loadAsync(buf)
    const sheetMap=await getSheetFileMap(zip)
    for(const path of sheetMap.values()){
      const f=zip.file(path); if(!f) continue
      let xml=await f.async('text')
      xml=xmlSetCell(xml,'B6',blNo.trim())
      zip.file(path,xml)
    }
    const out=await zip.generateAsync({type:'arraybuffer'})
    const blob=new Blob([out],{type:'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'})
    const a=document.createElement('a'); a.href=URL.createObjectURL(blob)
    a.download=`運送先一覧様式_${type.toUpperCase()}_${blNo.trim()}.xlsx`; a.click()
  }

  // ── 변동값 JSON 내보내기 — 모든 시트 전체 포함 ──────────────
  function exportMetaJson() {
    // localStorage에 저장된 모든 시트 키를 스캔해서 전부 백업
    const allSheets: Record<string, {s2meta: unknown, ctnMeta: unknown, pltNotes: unknown}> = {}
    // 현재 열린 시트 포함
    const knownSheets = [...sheetNames, activeSheet].filter(Boolean) as string[]
    for (const sn of [...new Set(knownSheets)]) {
      // 현재 activeSheet는 state에서, 나머지는 localStorage에서
      if (sn === activeSheet) {
        allSheets[sn] = { s2meta, ctnMeta, pltNotes }
      } else {
        const sm = lsLoad(sheetKey(sn, "s2meta"), null)
        const cm = lsLoad(sheetKey(sn, "ctnMeta"), null)
        const pn = lsLoad(sheetKey(sn, "pltNotes"), null)
        if (sm !== null || cm !== null || pn !== null) {
          allSheets[sn] = { s2meta: sm || {}, ctnMeta: cm || {}, pltNotes: pn || {} }
        }
      }
    }
    const data = { sheets: allSheets, master, _v: 3, _date: new Date().toISOString() }
    const blob = new Blob([JSON.stringify(data, null, 2)], { type: 'application/json' })
    const a = document.createElement('a')
    a.href = URL.createObjectURL(blob)
    a.download = `shipment_meta_${new Date().toISOString().slice(0,10)}.json`
    a.click()
  }

  // ── 변동값 JSON 불러오기 — localStorage에 직접 쓰고 state도 갱신 ──
  function importMetaJson(file: File) {
    const rd = new FileReader()
    rd.onload = ev => {
      try {
        const d = JSON.parse(ev.target?.result as string)

        // v3 형식: { sheets: { 시트명: { s2meta, ctnMeta, pltNotes } } }
        if (d._v === 3 && d.sheets) {
          let restoredCount = 0
          for (const [sn, data] of Object.entries(d.sheets as Record<string, {s2meta: unknown, ctnMeta: unknown, pltNotes: unknown}>)) {
            // localStorage에 직접 저장
            lsSave(sheetKey(sn, "s2meta"), data.s2meta)
            lsSave(sheetKey(sn, "ctnMeta"), data.ctnMeta)
            lsSave(sheetKey(sn, "pltNotes"), data.pltNotes)
            restoredCount++
            // 현재 activeSheet라면 state도 즉시 갱신
            if (sn === activeSheet) {
              setS2meta(data.s2meta as Record<string,Record<string,string>>)
              setCtnMeta(data.ctnMeta as Record<number,Record<string,string>>)
              setPltNotes(data.pltNotes as Record<string,string>)
            }
          }
          if (d.master) {
            lsSave(LS_KEY+"_master", d.master)
            setMaster(d.master)
          }
          alert(`복원 완료! ${restoredCount}개 시트 데이터 복원됨\n시트 탭을 클릭하면 해당 시트 데이터가 로드됩니다.`)

        // v2 구버전 호환: { sheetName, s2meta, ctnMeta, pltNotes }
        } else if (d.s2meta || d.ctnMeta) {
          const sn = d.sheetName || activeSheet || 'default'
          if (d.s2meta) lsSave(sheetKey(sn, "s2meta"), d.s2meta)
          if (d.ctnMeta) lsSave(sheetKey(sn, "ctnMeta"), d.ctnMeta)
          if (d.pltNotes) lsSave(sheetKey(sn, "pltNotes"), d.pltNotes)
          if (sn === activeSheet) {
            if (d.s2meta) setS2meta(d.s2meta)
            if (d.ctnMeta) setCtnMeta(d.ctnMeta)
            if (d.pltNotes) setPltNotes(d.pltNotes)
          }
          if (d.master) { lsSave(LS_KEY+"_master", d.master); setMaster(d.master) }
          alert(`복원 완료 (구버전 형식)! [${sn}] SKU메타 ${Object.keys(d.s2meta||{}).length}개\n다른 시트는 수동으로 복원이 필요할 수 있습니다.`)
        } else {
          alert('알 수 없는 JSON 형식입니다')
        }
      } catch (e) { alert('JSON 파일 오류: ' + (e as Error).message) }
    }
    rd.readAsText(file)
  }

  // 시트 변경 시 해당 시트의 저장 데이터 로드 (importMetaJson과 충돌 방지)
  // → importMetaJson이 localStorage에 직접 쓰기 때문에 여기서 읽으면 올바른 값이 나옴

  // ── FBA Inbound SP-API 헬퍼 ────────────────────────────────────
  async function apiFba(action:string, params:Record<string,unknown>={}){
    const res=await fetch('/api/fba-inbound',{method:'POST',headers:{'Content-Type':'application/json'},body:JSON.stringify({action,...params})})
    const json=await res.json()
    if(!json.ok)throw new Error(json.error||'알 수 없는 오류')
    return json
  }
  function updFbaStep(key:FbaStepKey,status:FbaStep['status'],msg?:string,data?:unknown){
    setFbaSteps(prev=>({...prev,[key]:{status,msg,data}}))
  }
  async function pollOp(operationId:string,stepKey:FbaStepKey,label:string):Promise<void>{
    for(let i=0;i<60;i++){
      await new Promise(r=>setTimeout(r,3000))
      const res=await apiFba('pollOperation',{operationId})
      if(res.operationStatus==='SUCCESS')return
      if(res.operationStatus==='FAILED')throw new Error(`${label} 실패: ${JSON.stringify(res.operationProblems)}`)
      updFbaStep(stepKey,'running',`${label} 대기 중... ${(i+1)*3}초`)
    }
    throw new Error(`${label} 시간 초과`)
  }

  async function runFbaInbound(){
    if(fbaRunning||!s2rows.length)return
    setFbaRunning(true)
    setFbaSteps(initSteps())
    setFbaPlacementOptions([])
    setFbaSelectedPlacement('')
    setFbaTransportOptions({})
    setFbaLabelUrls({})
    setFbaBolUrls({})
    try{
      const items=s2rows.map(r=>({
        msku:master[r.sku]?.sku||r.sku,
        quantity:r.total,
        cpp:r.cpp,
        bx:master[r.sku]?.bx||0,
        by:master[r.sku]?.by||0,
        bz:master[r.sku]?.bz||0,
        kg:master[r.sku]?.kg||0,
      }))
      // ① 출하 플랜 생성
      updFbaStep('createPlan','running','출하 플랜 생성 중...')
      const planName=`${activeSheet||'Shipment'}_${new Date().toISOString().slice(0,10)}`
      const createRes=await apiFba('createPlan',{
        planName,
        warehouseKey:fbaWarehouse,
        items:items.map(i=>({msku:i.msku,quantity:i.quantity})),
      })
      await pollOp(createRes.operationId,'createPlan','플랜 생성')
      const planDetailRes=await apiFba('getPlan',{planId:createRes.inboundPlanId})
      const planId:string=planDetailRes.inboundPlanId||createRes.inboundPlanId
      setFbaPlanId(planId)
      updFbaStep('createPlan','done',`플랜 생성 완료 — ${planId}`)

      // ② 패킹 옵션 생성 + 확정
      // packingGroups 배열(ID)을 여기서 추출 — setPackingInfo에서 사용
      updFbaStep('packingOptions','running','패킹 옵션 생성 중...')
      const pkgOptRes=await apiFba('generatePackingOptions',{planId})
      await pollOp(pkgOptRes.operationId,'packingOptions','패킹 옵션')
      const pkgListRes=await apiFba('listPackingOptions',{planId})
      const packingOptions=(pkgListRes.packingOptions||[]) as Array<{packingOptionId:string; packingGroups?:string[]}>
      if(!packingOptions.length)throw new Error('패킹 옵션이 없습니다')
      const packingOptionId=packingOptions[0].packingOptionId
      const packingGroupIds:string[]=packingOptions[0].packingGroups||[]
      const cfmPkgRes=await apiFba('confirmPackingOption',{planId,packingOptionId})
      await pollOp(cfmPkgRes.operationId,'packingOptions','패킹 옵션 확정')
      updFbaStep('packingOptions','done',`패킹 옵션 확정 — ${packingOptionId}`)

      // ③ 박스 내용물 입력 (모든 packing group을 단일 요청으로 전송)
      // Shipment는 generatePlacementOptions 후에 생성됨
      updFbaStep('setPackingInfo','running','박스 내용물 입력 중...')
      if(!packingGroupIds.length)throw new Error('Packing Group ID를 찾을 수 없습니다')
      const packRes=await apiFba('setPackingInfo',{planId,packingGroupIds,items})
      if(packRes.operationId)await pollOp(packRes.operationId,'setPackingInfo','패킹 정보 저장')
      updFbaStep('setPackingInfo','done',`패킹 정보 저장 완료 (${packingGroupIds.length}개 그룹)`)

      // ④ FC 배치 옵션 생성 → 사용자 선택 대기
      // 이 시점에 shipment가 생성되며, 각 placementOption 안에 shipmentIds가 포함됨
      updFbaStep('placementOptions','running','FC 배치 옵션 생성 중...')
      const plcOptRes=await apiFba('generatePlacementOptions',{planId})
      await pollOp(plcOptRes.operationId,'placementOptions','FC 배치 옵션')
      const plcListRes=await apiFba('listPlacementOptions',{planId})
      const placements=(plcListRes.placementOptions||[]) as PlacementOption[]
      setFbaPlacementOptions(placements)
      // placement options 안의 shipmentIds를 미리 수집 (phase 2에서 덮어씀)
      const allShipIds=[...new Set(placements.flatMap(p=>p.shipmentIds||[]))]
      if(allShipIds.length)setFbaShipmentIds(allShipIds)
      updFbaStep('placementOptions','waiting',`${placements.length}개 FC 옵션 — 아래서 선택 후 [계속] 버튼을 눌러주세요`,placements)
    }catch(err){
      const msg=err instanceof Error?err.message:String(err)
      setFbaSteps(prev=>{const n={...prev};for(const k of FBA_STEPS)if(n[k].status==='running')n[k]={status:'error',msg};return n})
      alert(`FBA 오류: ${msg}`)
    }finally{setFbaRunning(false)}
  }

  async function runFbaPhase2(){
    if(!fbaPlanId||!fbaSelectedPlacement||fbaRunning)return
    setFbaRunning(true)
    try{
      // 선택된 placement option에서 shipmentIds 취득 (placement마다 다를 수 있음)
      const selectedPl=fbaPlacementOptions.find(p=>p.placementOptionId===fbaSelectedPlacement)
      const shipmentIds:string[]=selectedPl?.shipmentIds?.length
        ? selectedPl.shipmentIds
        : fbaShipmentIds
      if(!shipmentIds.length)throw new Error('Shipment ID를 찾을 수 없습니다 (placement 재확인 필요)')
      setFbaShipmentIds(shipmentIds)

      // ⑤ FC 배치 확정
      updFbaStep('placementOptions','running','FC 배치 확정 중...')
      const cfmRes=await apiFba('confirmPlacement',{planId:fbaPlanId,placementOptionId:fbaSelectedPlacement})
      if(cfmRes.operationId)await pollOp(cfmRes.operationId,'placementOptions','FC 배치 확정')
      updFbaStep('placementOptions','done',`FC 배치 확정 — ${fbaSelectedPlacement}`)

      // ⑥ 운송 옵션 생성 + 확정 + 팔레트 정보 등록
      updFbaStep('transport','running','운송 옵션 설정 중...')
      // 출하 예정일 기준: readyDate = 사용자 입력, endDate = +30일
      const readyDate=new Date(fbaShipDate)
      const endDate=new Date(readyDate);endDate.setDate(endDate.getDate()+30)
      // MSKU → cpp 역매핑 (팔레트 수 계산용)
      const mskuToCpp:Record<string,number>={}
      for(const m of Object.values(master)){if(m.sku)mskuToCpp[m.sku]=m.cpp}

      for(const shipmentId of shipmentIds){
        // 운송 옵션 생성
        const genRes=await apiFba('generateTransportOptions',{planId:fbaPlanId,shipmentId,readyToShipWindow:{start:readyDate.toISOString(),end:endDate.toISOString()}})
        if(genRes.operationId)await pollOp(genRes.operationId,'transport','운송 옵션')
        const tListRes=await apiFba('listTransportOptions',{planId:fbaPlanId,shipmentId})
        const tOpts=(tListRes.transportationOptions||[]) as TransportOption[]
        setFbaTransportOptions(prev=>({...prev,[shipmentId]:tOpts}))
        // 선택된 운송 모드(FTL/LTL/SP) 우선, 없으면 첫 번째
        const chosen=tOpts.find(t=>t.shippingMode===fbaShippingMode)||tOpts[0]
        if(!chosen)throw new Error(`${shipmentId}: 운송 옵션 없음`)
        const cfmTRes=await apiFba('confirmTransport',{planId:fbaPlanId,shipmentId,transportOptionId:chosen.transportOptionId})
        if(cfmTRes.operationId)await pollOp(cfmTRes.operationId,'transport','운송 확정')

        // 팔레트 수 자동 계산: listShipmentItems → 각 MSKU qty ÷ cpp 합산
        const shipItemsRes=await apiFba('listShipmentItems',{planId:fbaPlanId,shipmentId})
        const shipItems=(shipItemsRes.items||[]) as Array<{msku:string;quantity:number}>
        const palletCount=shipItems.reduce((sum,it)=>{
          const cpp=mskuToCpp[it.msku]||16
          return sum+Math.ceil(it.quantity/cpp)
        },0)

        // 팔레트 수 등록 (운송사는 env var SP_API_CARRIER_NAME 에서 서버가 읽음)
        await apiFba('updateTracking',{planId:fbaPlanId,shipmentId,palletCount})
      }
      updFbaStep('transport','done','운송 확정 + 팔레트 정보 등록 완료')

      // ⑦ 라벨 + BOL URL 취득
      updFbaStep('labels','running','라벨 URL 취득 중...')
      const labelUrls:Record<string,string>={}
      const bolUrls:Record<string,string>={}
      for(const shipmentId of shipmentIds){
        const cRes=await apiFba('getLabels',{planId:fbaPlanId,shipmentId})
        if(cRes.downloadURL)labelUrls[`carton_${shipmentId}`]=cRes.downloadURL
        const pRes=await apiFba('getPalletLabels',{planId:fbaPlanId,shipmentId})
        if(pRes.downloadURL)labelUrls[`pallet_${shipmentId}`]=pRes.downloadURL
        const bRes=await apiFba('getBol',{planId:fbaPlanId,shipmentId})
        if(bRes.downloadURL)bolUrls[shipmentId]=bRes.downloadURL
      }
      setFbaLabelUrls(labelUrls)
      setFbaBolUrls(bolUrls)
      updFbaStep('labels','done','카톤 + 파렛트 라벨 URL 취득 완료')
      updFbaStep('bol','done','BOL URL 취득 완료')
    }catch(err){
      const msg=err instanceof Error?err.message:String(err)
      setFbaSteps(prev=>{const n={...prev};for(const k of FBA_STEPS)if(n[k].status==='running')n[k]={status:'error',msg};return n})
      alert(`FBA 오류: ${msg}`)
    }finally{setFbaRunning(false)}
  }

  const raw=(activeSheet&&sheets[activeSheet])||[]
  const {data:nd}=raw.length?normHeaders(raw):{data:[] as RowData[]}
  const fd=nd.length?forwardFill(nd, master):[]
  const ctnNums=[...new Set(fd.map(r=>r.container_no as number))].sort((a,b)=>a-b)

  function loadFile(f:File){setFile(f);const rd=new FileReader();rd.onload=e=>{const wb=XLSX.read((e.target as FileReader).result,{type:'array'});const ns:Record<string,RowData[]>={};for(const sn of wb.SheetNames)ns[sn]=parseSheet(wb.Sheets[sn]);setSh(ns);setSn(wb.SheetNames);setAs(wb.SheetNames[0]);setMode('1')};rd.readAsArrayBuffer(f)}

  function buildS2rows(){
    // GW는 행별로 TRUNC 후 합산 (BL과 동일한 방식)
    const agg:Record<string,{sku:string,total:number,gw:number,cbm:number}>={};
    for(const r of fd){
      const sk=String(r.sku||"");if(!sk)continue;
      if(!agg[sk])agg[sk]={sku:sk,total:0,gw:0,cbm:0};
      agg[sk].total+=r.quantity;
      const m=master[sk]||({} as MasterItem);
      // 행별로 calcGW/calcCBM 계산 후 누적
      agg[sk].gw+=calcGW(r.quantity,parseFloat(String(m.kg))||0);
      agg[sk].cbm+=calcCBM(r.quantity,parseFloat(String(m.bx))||0,parseFloat(String(m.by))||0,parseFloat(String(m.bz))||0);
    }
    return Object.values(agg).sort((a,b)=>{const ai=SKU_ORDER.indexOf(a.sku),bi=SKU_ORDER.indexOf(b.sku);if(ai===-1&&bi===-1)return a.sku.localeCompare(b.sku);if(ai===-1)return 1;if(bi===-1)return -1;return ai-bi}).map(({sku,total,gw,cbm})=>{
      const m=master[sku]||({} as MasterItem);const cpp=parseFloat(String(m.cpp))||16;const meta=s2meta[sku]||{}
      return {sku,total,cpp,asin:m.asin||"",to:m.to||"",loc:m.loc||"",pallets:Math.ceil(total/cpp),pages:parseFloat((total/6).toFixed(2)),gw:Math.round(gw*10)/10,cbm:Math.round(cbm*100)/100,fc:meta.fc||"",address:meta.address||"",fbaId:meta.fbaId||"",amazonId:meta.amazonId||""}
    })
  }
  function buildTotalRows(){
    const agg:Record<string,number>={}
    for(const sn of sheetNames){
      if(totalExcludedSheets.has(sn))continue
      const raw=sheets[sn]||[];if(!raw.length)continue
      const{data:nd}=normHeaders(raw)
      const fd2=forwardFill(nd,master)
      for(const r of fd2){const sk=String(r.sku||'');if(!sk)continue;agg[sk]=(agg[sk]||0)+(r.quantity||0)}
    }
    const ordered=[...SKU_ORDER.filter(s=>s in agg),...Object.keys(agg).filter(s=>!SKU_ORDER.includes(s))]
    return ordered.map(sku=>({sku,total:agg[sku]||0,adj:totalAdjustments[sku]||0,final:(agg[sku]||0)+(totalAdjustments[sku]||0)})).filter(r=>r.total>0||r.adj!==0)
  }
  function buildS3groups(){return ctnNums.map(no=>{const rows=fd.filter(r=>r.container_no===no);const f0=rows[0]||{};const cm=ctnMeta[no]||{};return {no,container:cm.container||"",sealNo:cm.sealNo||"",shipment_date:String(f0.shipment_date||""),shipment_time:String(f0.shipment_time||""),destination:String(f0.destination||""),etd:String(f0.etd||""),eta:String(f0.eta||""),rows}})}
  const updM=(sku:string,f:string,v:string|number)=>setMaster(p=>({...p,[sku]:{...(p[sku]||{}),[f]:v} as MasterItem}))
  const updMeta=(sku:string,f:string,v:string)=>setS2meta(p=>({...p,[sku]:{...(p[sku]||{}),[f]:v}}))
  const updCtnMeta=(no:number,f:string,v:string)=>setCtnMeta(p=>({...p,[no]:{...(p[no]||{}),[f]:v}}))
  const togCtn=(k:string)=>setColl(c=>({...c,[k]:!c[k]}));const isOpen=(k:string)=>!coll[k]
  function tag(l:string,v:unknown){if(!v||v===""||v===0||v==="0")return null;return <span style={{fontSize:12,opacity:0.85}}><span style={{opacity:0.6,marginRight:2}}>{l}</span>{String(v)}</span>}

  function expTotal(){
    const rows=buildTotalRows()
    const selSheets=sheetNames.filter(sn=>!totalExcludedSheets.has(sn))
    const tR=rows.reduce((s,r)=>s+r.total,0),tA=rows.reduce((s,r)=>s+r.adj,0),tF=rows.reduce((s,r)=>s+r.final,0)
    const hdr=[["선택 시트",selSheets.join(', ')],["",""],["약호","합계","조정","최종합계"]]
    const body=rows.map(r=>[r.sku,r.total,r.adj,r.final])
    const ts=new Date().toISOString().slice(0,10)
    xlsDl([...hdr,...body,["합계",tR,tA,tF]],"총합","총합_"+ts+".xlsx")
  }
  function expS1(){xlsDl(fd.map(r=>({시트:activeSheet,컨:"컨"+r.container_no,출항일:r.shipment_date||"",목적지:r.destination||"",BOM품번:r.model_code||"",품목:r.sku||"",색상:r.color||"",수량:r.quantity,TYPE구:r.type_old||"",TYPE신:r.location||"",ETD:r.etd||"",ETA:r.eta||""})),activeSheet||"S1","1단계_"+(activeSheet||"data")+".xlsx")}
  function expS2(){const rows=buildS2rows();const hdr=[["약호","ASIN","구박스","신박스","카톤","PLT당카톤","팔레트","G.W(kg)","CBM","FC CENTER","주소","FBA ID","아마존 ID"]];const body=rows.map(r=>[r.sku,r.asin,r.to,r.loc,r.total,r.cpp,r.pallets,r.gw,r.cbm,r.fc,r.address,r.fbaId,r.amazonId]);const tC=rows.reduce((s,r)=>s+r.total,0),tP=rows.reduce((s,r)=>s+r.pallets,0),tW=rows.reduce((s,r)=>s+r.gw,0),tB=Math.round(rows.reduce((s,r)=>s+r.cbm,0)*100)/100;xlsDl([...hdr,...body,["합계","","","",tC,"",tP,tW,tB]],activeSheet||"S2","1차가공_"+(activeSheet||"data")+".xlsx")}
  function expS3(){const hdr=[["CONTAINER","SEAL NO.","약호","BOX CODE","ASIN","FBA ID","아마존 ID","FC CENTER","주소","PLT","CT","CTN/PLT","G.W(kg)","CBM"]];const body:unknown[][]=[];for(const g of buildS3groups()){let first=true;for(const r of g.rows){const m=master[String(r.sku)]||({} as MasterItem);const cpp=parseFloat(String(m.cpp))||16;const sk=s2meta[String(r.sku)]||{};body.push([first?g.container:"",first?g.sealNo:"",r.sku,r.location||m.loc||"",m.asin||"",sk.fbaId||"",sk.amazonId||"",sk.fc||"",first?(sk.address||""):"",Math.ceil(r.quantity/cpp),r.quantity,cpp+"CTN/PLT",calcGW(r.quantity,parseFloat(String(m.kg))||0),calcCBM(r.quantity,parseFloat(String(m.bx))||0,parseFloat(String(m.by))||0,parseFloat(String(m.bz))||0)]);first=false}};xlsDl([...hdr,...body],activeSheet||"S3","2차가공_"+(activeSheet||"data")+".xlsx")}

  async function expFbaUpload(ctnFilter?: number[]){
    setFbaLoading(true)
    try{
      const res=await fetch("/UPLOAD_FORMAT.xlsx");if(!res.ok)throw new Error("파일 로드 실패 ("+res.status+")");
      const buf=await res.arrayBuffer();
      const wb=XLSX.read(buf,{type:"array",cellStyles:true,cellNF:true,cellDates:true,sheetStubs:true});
      const tplName=wb.SheetNames.find(n=>n.toLowerCase().includes("template"));
      if(!tplName)throw new Error("template 시트를 찾을 수 없습니다: "+wb.SheetNames.join(", "));
      const ws=wb.Sheets[tplName];

      // 컨테이너 필터 적용: fd에서 해당 컨테이너 행만 추출 후 SKU별 합산
      const filteredFd = (ctnFilter&&ctnFilter.length>0)
        ? fd.filter(r=>ctnFilter.includes(r.container_no as number))
        : fd

      // buildS2rows와 동일 방식으로 필터된 fd 기준 집계
      const agg: Record<string,{sku:string,total:number,gw:number}> = {}
      for (const r of filteredFd) {
        const sk=String(r.sku||''); if(!sk) continue
        if (!agg[sk]) agg[sk]={sku:sk,total:0,gw:0}
        agg[sk].total+=r.quantity
        const m=master[sk]||({} as MasterItem)
        agg[sk].gw+=calcGW(r.quantity,parseFloat(String(m.kg))||0)
      }
      const rows=Object.values(agg)

      rows.forEach((r,i)=>{
        const row=9+i;const m=master[r.sku]||({} as MasterItem);const realSku=m.sku||r.sku;
        const set=(col:string,t:"s"|"n",v:string|number)=>{ws[col+row]={...(ws[col+row]||{}),t,v,w:String(v)}};
        set("A","s",realSku);set("B","n",r.total);set("F","n",1);set("G","n",r.total);
        set("H","n",m.bx||0);set("I","n",m.by||0);set("J","n",m.bz||0);set("K","n",m.kg||0)
      });
      if(rows.length>0){
        const decoded=XLSX.utils.decode_range(ws["!ref"]||"A1:K8");
        decoded.e.r=Math.max(decoded.e.r,8+rows.length-1);decoded.e.c=Math.max(decoded.e.c,10);
        ws["!ref"]=XLSX.utils.encode_range(decoded)
      }
      // 파일명에 컨테이너 범위 포함
      const suffix = (ctnFilter&&ctnFilter.length>0) ? `_컨${ctnFilter.join('_')}` : '_전체'
      XLSX.writeFile(wb,`UPLOAD_FORMAT_filled${suffix}.xlsx`,{cellStyles:true,compression:true})
    }catch(e){alert("오류: "+(e as Error).message)}finally{setFbaLoading(false)}
  }

  // ── 표지 페이지 생성 (카톤 라벨) ─────────────────────────────
  async function makeCoverPage(doc: {addPage:(s:[number,number])=>ReturnType<typeof doc.addPage>, embedFont:(f:string)=>Promise<ReturnType<typeof doc.embedFont>>}, loc: string) {
    const { StandardFonts, rgb: pdfRgb } = await import('pdf-lib')
    // eslint-disable-next-line @typescript-eslint/no-explicit-any
    const coverPage = (doc as any).addPage([595, 842])
    // eslint-disable-next-line @typescript-eslint/no-explicit-any
    let font: any
    try { font = await (doc as any).embedFont(StandardFonts.HelveticaBold) }
    catch { font = await (doc as any).embedFont(StandardFonts.Helvetica) }

    const realSku = locToRealSku(loc)
    const sku = locToSku(loc)

    // 외곽 테두리만 (잉크 최소화 — 배경/헤더/푸터 검정 바 없음)
    coverPage.drawRectangle({ x: 20, y: 20, width: 555, height: 802, borderColor: pdfRgb(0.3, 0.3, 0.3), borderWidth: 1.5 })

    // CARTON LABEL 헤더 텍스트
    const title = 'CARTON LABEL'
    const tw = font.widthOfTextAtSize(title, 26)
    coverPage.drawText(title, { x: (595 - tw) / 2, y: 775, size: 26, font, color: pdfRgb(0.15, 0.15, 0.15) })

    // 상단 구분선
    coverPage.drawLine({ start: {x: 40, y: 755}, end: {x: 555, y: 755}, thickness: 1, color: pdfRgb(0.5, 0.5, 0.5) })

    // BOX CODE (크게)
    const locSize = loc.length <= 7 ? 100 : 72
    const locW = font.widthOfTextAtSize(loc, locSize)
    coverPage.drawText(loc, { x: (595 - locW) / 2, y: 530, size: locSize, font, color: pdfRgb(0.05, 0.05, 0.05) })

    // 중간 구분선
    coverPage.drawLine({ start: {x: 40, y: 505}, end: {x: 555, y: 505}, thickness: 1, color: pdfRgb(0.5, 0.5, 0.5) })

    // realSku (중간 크기)
    const label2 = realSku || sku || ''
    if (label2) {
      const fs2 = label2.length > 25 ? 24 : 32
      const lw2 = font.widthOfTextAtSize(label2, fs2)
      coverPage.drawText(label2, { x: Math.max(40, (595 - lw2) / 2), y: 430, size: fs2, font, color: pdfRgb(0.2, 0.2, 0.2) })
    }

    // 하단 구분선 + 서브타이틀
    coverPage.drawLine({ start: {x: 40, y: 55}, end: {x: 555, y: 55}, thickness: 1, color: pdfRgb(0.5, 0.5, 0.5) })
    const sub2 = 'FBA CARTON SHIPPING LABEL'
    const sw2 = font.widthOfTextAtSize(sub2, 13)
    coverPage.drawText(sub2, { x: (595 - sw2) / 2, y: 30, size: 13, font, color: pdfRgb(0.3, 0.3, 0.3) })
  }

  // ── 파렛트 표지 페이지 생성 ───────────────────────────────────
  async function makePalletCoverPage(doc: unknown, fbaId: string, fc: string, pltRange: string) {
    const { StandardFonts, rgb: pdfRgb } = await import('pdf-lib')
    // eslint-disable-next-line @typescript-eslint/no-explicit-any
    const d = doc as any
    const coverPage = d.addPage([595, 842])
    // eslint-disable-next-line @typescript-eslint/no-explicit-any
    let font: any
    try { font = await d.embedFont(StandardFonts.HelveticaBold) }
    catch { font = await d.embedFont(StandardFonts.Helvetica) }

    // 파란 배경 상단 바
    coverPage.drawRectangle({ x: 0, y: 742, width: 595, height: 100, color: pdfRgb(0.07, 0.22, 0.55) })
    // 본문 배경
    coverPage.drawRectangle({ x: 0, y: 0, width: 595, height: 742, color: pdfRgb(0.97, 0.98, 1.0) })

    // 상단: PALLET LABEL
    const title = 'PALLET LABEL'
    const tw = font.widthOfTextAtSize(title, 28)
    coverPage.drawText(title, { x: (595 - tw) / 2, y: 790, size: 28, font, color: pdfRgb(1, 1, 1) })

    // FC센터 (가장 크게)
    const fcText = fc || 'FC'
    const fcSize = fcText.length <= 4 ? 120 : 80
    const fcW = font.widthOfTextAtSize(fcText, fcSize)
    coverPage.drawText(fcText, { x: (595 - fcW) / 2, y: 560, size: fcSize, font, color: pdfRgb(0.07, 0.22, 0.55) })

    // 구분선
    coverPage.drawLine({ start: {x:60, y:540}, end: {x:535, y:540}, thickness: 2, color: pdfRgb(0.7, 0.75, 0.85) })

    // 파렛트 범위
    const pltW = font.widthOfTextAtSize(pltRange, 52)
    coverPage.drawText(pltRange, { x: (595 - pltW) / 2, y: 460, size: 52, font, color: pdfRgb(0.1, 0.1, 0.1) })

    // FBA ID
    const fidW = font.widthOfTextAtSize(fbaId, 30)
    coverPage.drawText(fbaId, { x: (595 - fidW) / 2, y: 390, size: 30, font, color: pdfRgb(0.3, 0.35, 0.45) })

    // 하단 라벨
    const sub = 'FBA PALLET SHIPPING LABEL'
    const subW = font.widthOfTextAtSize(sub, 14)
    coverPage.drawRectangle({ x: 0, y: 0, width: 595, height: 40, color: pdfRgb(0.07, 0.22, 0.55) })
    coverPage.drawText(sub, { x: (595 - subW) / 2, y: 13, size: 14, font, color: pdfRgb(0.7, 0.8, 1.0) })
  }

  // eslint-disable-next-line @typescript-eslint/no-explicit-any
  async function loadPdfjs():Promise<any>{
    // eslint-disable-next-line @typescript-eslint/no-explicit-any
    if((window as any).pdfjsLib)return(window as any).pdfjsLib
    // eslint-disable-next-line @typescript-eslint/no-explicit-any
    return new Promise<any>((resolve,reject)=>{const s=document.createElement('script');s.src='https://cdnjs.cloudflare.com/ajax/libs/pdf.js/3.11.174/pdf.min.js';s.onload=()=>{const lib=(window as any).pdfjsLib as any;lib.GlobalWorkerOptions.workerSrc='https://cdnjs.cloudflare.com/ajax/libs/pdf.js/3.11.174/pdf.worker.min.js';resolve(lib)};s.onerror=reject;document.head.appendChild(s)})
  }

  // ── pdfjs 옵션 헬퍼 ─────────────────────────────────────────
  // eslint-disable-next-line @typescript-eslint/no-explicit-any
  function getPdfjsOpts(data: Uint8Array): any {
    return {
      data: data.slice(),
      cMapUrl: 'https://cdn.jsdelivr.net/npm/pdfjs-dist@3.11.174/cmaps/',
      cMapPacked: true,
      useWorkerFetch: false,
      isEvalSupported: false,
    }
  }

  // ── 파렛트 라벨만 따로 처리 (카톤 없이도 pltNotes 업데이트) ──
  async function processPalletOnly(pltFiles: File[]) {
    if (!pltFiles.length) return
    setLabelLoading(true)
    setLabelStatus("파렛트 라벨 처리 중...")
    try {
      const pdfjsLib = await loadPdfjs()
      const { PDFDocument } = await import('pdf-lib')
      const newPltNotes: Record<string, string> = { ...pltNotes }
      const newGroups: Record<string, Uint8Array> = { ...labelGroups }
      const newCounts: Record<string, number> = { ...labelCounts }

      for (const f of pltFiles) {
        const bytes = new Uint8Array(await f.arrayBuffer())
        const pdf2 = await pdfjsLib.getDocument(getPdfjsOpts(bytes)).promise
        let maxPlt = 0, detectedFbaId = '', detectedFc = ''

        for (let pi = 0; pi < pdf2.numPages; pi++) {
          const page = await pdf2.getPage(pi + 1)
          const content = await page.getTextContent()
          // eslint-disable-next-line @typescript-eslint/no-explicit-any
          const tokens = (content.items as any[]).map((it) => (it.str || '').trim()).filter(Boolean)
          const pageText = tokens.join(' ')
          const fbaMatch = pageText.match(/FBA[A-Z0-9]{8,12}/)
          if (fbaMatch && !detectedFbaId) detectedFbaId = fbaMatch[0]
          const pltMatch = pageText.match(/#[：:]\s*(\d+)\s*\/\s*(\d+)/)
          if (pltMatch) { const total = parseInt(pltMatch[2]); if (total > maxPlt) maxPlt = total }
          // FC센터 추출: 납품선 코드는 보통 3~5자 대문자 (HIY1, TPB5, VJNB 등)
          // 파렛트 라벨에서 "納品先:\nXXXX\n" 또는 "FBA STA (...)-XXXX" 패턴
          if (!detectedFc) {
            const staMatch = pageText.match(/FBA\s+STA\s+[^-]+-([A-Z]{3,5})\b/)
            if (staMatch) detectedFc = staMatch[1]
          }
          if (!detectedFc) {
            // "HIY1 - AMXL" 또는 "TPB5\n1980024" 패턴: 숫자 포함 4~5자 코드
            const codeMatch = pageText.match(/\b([A-Z]{2,4}\d{1,2})\b/)
            if (codeMatch && !['FBA'].includes(codeMatch[1])) detectedFc = codeMatch[1]
          }
          if (!detectedFc) {
            for (const tok of tokens) {
              if (/^[A-Z]{3,5}\d?$/.test(tok) && !['FBA','SKU','IXD','STA','IXD','AMXL'].includes(tok) && tok !== detectedFbaId) {
                detectedFc = tok; break
              }
            }
          }
        }
        pdf2.destroy()

        // 파일명에서 FBA ID + 파렛트 수 추출 (텍스트 추출 실패 시 fallback)
        if (!detectedFbaId) {
          const fnMatch = f.name.match(/^(FBA[A-Z0-9]+)/)
          if (fnMatch) detectedFbaId = fnMatch[1]
        }
        // 파일명에서 파렛트 수 추출: "1~18", "1~142" 패턴
        if (maxPlt === 0) {
          const pltNumMatch = f.name.match(/1[~\-](\d+)/)
          if (pltNumMatch) maxPlt = parseInt(pltNumMatch[1])
        }
        // 파일명에서 FC센터 추출 (앞부분 알파벳+숫자 코드)
        if (!detectedFc) {
          const fcMatch = f.name.match(/^([A-Z]{3,5}\d?)_/)
          if (fcMatch && !['FBA'].includes(fcMatch[1])) detectedFc = fcMatch[1]
        }

        if (detectedFbaId && maxPlt > 0) {
          const pltRange = `1~${maxPlt}`
          const pltLabel = `${detectedFc ? detectedFc+'_' : ''}${detectedFbaId}_${pltRange}`

          // FBA ID 매칭: 여러 전략으로 시도
          // s2meta.fbaId 예시: "FBA15G97HNPC", "FBA15G97HNPCU000001" 등
          // detectedFbaId 예시: "FBA15G97HNPC", "FBA15G97HNPCU000001" 등
          const detectedBase = detectedFbaId.replace(/U\d+$/, '').toUpperCase()

          for (const [sku, meta] of Object.entries(s2meta)) {
            const rawFbaId = meta.fbaId?.trim() || ''
            if (!rawFbaId) continue
            const metaBase = rawFbaId.replace(/U\d+$/, '').toUpperCase()

            const matched =
              metaBase === detectedBase ||                          // 정확히 일치
              rawFbaId.toUpperCase().startsWith(detectedBase) ||   // 저장된게 더 길 경우
              detectedBase.startsWith(metaBase) ||                 // PDF에서 읽힌게 더 길 경우
              metaBase.includes(detectedBase) ||
              detectedBase.includes(metaBase)

            if (matched) {
              const ctnNums2 = [...new Set(fd.map(r=>r.container_no as number))]
              for (const no of ctnNums2) { newPltNotes[`${no}_${sku}`] = pltRange }
            }
          }

          const pltKey = `__PLT__${pltLabel}`
          const pltDoc = await PDFDocument.create()
          await makePalletCoverPage(pltDoc, detectedFbaId, detectedFc, `Pallet ${pltRange}`)
          const srcPlt = await PDFDocument.load(bytes.slice())
          const pltPages = await pltDoc.copyPages(srcPlt, Array.from({length: srcPlt.getPageCount()}, (_,i)=>i))
          pltPages.forEach(p => pltDoc.addPage(p))
          await makePalletCoverPage(pltDoc, detectedFbaId, detectedFc, `Pallet ${pltRange}`)
          newGroups[pltKey] = await pltDoc.save()
          newCounts[pltKey] = maxPlt
        } else {
          // 매칭 실패 시 상태 메시지
          setLabelStatus(`파렛트 감지 실패: FBA ID="${detectedFbaId}" 파렛트수=${maxPlt} — 파일명: ${f.name}`)
        }
      }
      const matchedCount = Object.keys(newPltNotes).length - Object.keys(pltNotes).length
      setPltNotes(newPltNotes)
      setLabelGroups(newGroups)
      setLabelCounts(newCounts)
      setLabelStatus(`파렛트 ${pltFiles.length}개 처리 완료 — ${matchedCount > 0 ? matchedCount+'개 SKU 파렛트 번호 반영됨' : '⚠️ 매칭된 SKU 없음 (1차 가공에서 FBA ID 입력 확인)'}`)
    } catch(e) {
      alert('파렛트 처리 오류: ' + (e as Error).message)
      setLabelStatus("")
    } finally {
      setLabelLoading(false)
    }
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
  async function processLabels(files: File[], pltFiles: File[] = []) {
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

      // ── 파일명 prefix → LOC 매칭 맵 구성 ──────────────────
      // 파일명: "FBA15G97HNPC-xxxxxxxxxx.pdf" → prefix "FBA15G97HNPC"
      // s2meta[약호].fbaId = "FBA15G97HNPCU000001" → prefix "FBA15G97HNPC"
      // prefix가 일치하면 해당 약호의 loc를 사용
      const fbaIdPrefixToLoc: Record<string, string> = {}
      for (const [sku, meta] of Object.entries(s2meta)) {
        const fbaId = meta.fbaId?.trim()
        if (!fbaId) continue
        // FBA ID에서 U+숫자 suffix 제거 → prefix 추출
        // 예: FBA15G97HNPCU000001 → FBA15G97HNPC
        const prefix = fbaId.replace(/U\d+$/, '').toUpperCase()
        if (prefix.length > 5) {
          const loc = master[sku]?.loc
          if (loc) fbaIdPrefixToLoc[prefix] = loc
        }
      }

      // 파일명에서 prefix 추출하는 함수
      // 예: "FBA15G97HNPC-1777530666990.pdf" → "FBA15G97HNPC"
      function getPrefixFromFilename(filename: string): string {
        const base = filename.replace(/\.pdf$/i, '')
        const dashIdx = base.indexOf('-')
        return dashIdx > 0 ? base.slice(0, dashIdx).toUpperCase() : base.toUpperCase()
      }

      const pdfjsLib = await loadPdfjs()
      const { PDFDocument, rgb, StandardFonts } = await import('pdf-lib')

      type PageEntry = {
        fileBytes: Uint8Array
        pageIdx: number
        labelsToKeep: number[]
        totalLabels: number
        barcodeFound: {num: string; y: number; x: number}[]  // FBA unit IDs with raw PDF coordinates
      }

      const locMap: Record<string, PageEntry[]> = {}
      const debugLog: LabelDebugEntry[] = []
      let totalPages = 0
      let processed = 0

      // 전체 페이지 수
      for (const f of files) {
        const bytes = new Uint8Array(await f.arrayBuffer())
        const pdf = await pdfjsLib.getDocument(getPdfjsOpts(bytes)).promise
        totalPages += pdf.numPages
        pdf.destroy()
      }

      for (const f of files) {
        const fileBytes = new Uint8Array(await f.arrayBuffer())
        const pdf = await pdfjsLib.getDocument(getPdfjsOpts(fileBytes)).promise

        // ── 파일 단위 FBA ID prefix 매칭 ─────────────────────
        // pdfjs가 폰트를 못 읽어 SKU가 추출 안 되는 경우의 폴백
        const filenamePrefix = getPrefixFromFilename(f.name)
        const locFromFilename = fbaIdPrefixToLoc[filenamePrefix] ?? null

        for (let i = 0; i < pdf.numPages; i++) {
          const page = await pdf.getPage(i + 1)
          const content = await page.getTextContent()
          // eslint-disable-next-line @typescript-eslint/no-explicit-any
          const items = content.items as any[]

          // 각 라벨의 SKU 추출
          // 구조: 「単一のSKU」 → 다음 줄에 SKU (하이픈으로 끝나면 그 다음 줄과 합침)
          // 예: 'US-HSBWB6-090-040-180-5S-' + '1Pack' → 'US-HSBWB6-090-040-180-5S-1Pack'
          type LabelInfo = { sku: string; y: number; x: number }
          const labelInfos: LabelInfo[] = []

          for (let j = 0; j < items.length; j++) {
            const str = items[j].str as string
            if (!str) continue

            // 「単一のSKU」 마커 감지 (토큰이 쪼개질 경우도 대응)
            const isSingleSku = str === '単一のSKU' || str.includes('単一のSKU') ||
              (str === '単一の' && items[j+1]?.str === 'SKU') ||
              (str === '単' && items[j+1]?.str === '一のSKU')

            if (!isSingleSku) continue

            const y = items[j].transform?.[5] ?? 0
            const x = items[j].transform?.[4] ?? 0
            let sku = ''

            // 「単一のSKU」 다음 최대 10토큰에서 SKU 조합
            for (let k = j + 1; k < Math.min(j + 11, items.length); k++) {
              const s = (items[k].str as string)?.trim() ?? ''
              if (!s) continue

              // 건너뛸 토큰
              if (s.includes('数量') || s.includes('JAN') || s.includes('のSKU') ||
                  s.includes('Pack数') || /^[\d\s]+$/.test(s)) continue

              // FBA ID (FBA15... 형태)는 건너뜀
              if (/^FBA\d/.test(s)) continue

              if (s.length > 2) {
                // 첫 SKU 토큰 발견
                sku = s

                // 다음 토큰이 '1Pack' 계열이면 합침 (라인 줄바꿈으로 쪼개진 경우)
                for (let m = k + 1; m < Math.min(k + 4, items.length); m++) {
                  const next = (items[m].str as string)?.trim() ?? ''
                  if (!next) continue
                  // sku가 하이픈으로 끝나거나, 다음 토큰이 Pack/JAPAN/JP 계열이면 합침
                  if (sku.endsWith('-') || /^1?Pack|JAPAN$|^JP$|-JP$/i.test(next)) {
                    if (!next.includes('数量') && !next.includes('JAN') && !/^[\d\s]+$/.test(next)) {
                      sku += next
                      k = m
                    }
                  }
                  break
                }
                break
              }
            }

            if (sku) labelInfos.push({ sku, y, x })
          } // end for j


          // fallback 1: 전체 텍스트에서 마스터 realSku 직접 탐색
          if (labelInfos.length === 0) {
            const fullNoSp = items.map((it: {str:string}) => (it.str||'')).join('').replace(/\s+/g,'').toUpperCase()
            const matchedSkus: string[] = []
            for (const m of Object.values(master)) {
              if (!m.sku || !m.loc) continue
              const norm = normalizeSku(m.sku)
              if (norm.length < 6) continue
              if (fullNoSp.includes(norm)) matchedSkus.push(m.sku)
            }
            const unique = [...new Set(matchedSkus)]
            unique.forEach((sku, idx) => labelInfos.push({ sku, y: 10000 - idx * 40, x: idx % 2 === 0 ? 0 : 1000 }))
          }

          // fallback 2: 파일명 FBA ID prefix → LOC 직접 매칭
          // (pdfjs가 일본어 폰트를 못 읽어 SKU가 전혀 안 보이는 경우)
          if (labelInfos.length === 0 && locFromFilename) {
            // 페이지 내 FBA 라벨 ID 개수로 라벨 수 추정
            const rawText = items.map((it: {str:string}) => (it.str||'')).join('')
            const fbaMatches = rawText.match(/FBA\d{2}[A-Z0-9]+U\d+/g) || []
            const labelCount = fbaMatches.length || 6 // 기본 6개 (3×2 그리드)
            // 가상 라벨로 추가 (혼합 페이지 마스킹 없이 전체 페이지 포함)
            labelInfos.push({ sku: `__fbaPrefix__${locFromFilename}`, y: 10000, x: 0 })
          }

          if (labelInfos.length === 0) {
            const rawText = items.map((it: {str:string}) => (it.str||'')).join('').slice(0, 80)
            debugLog.push({
              file: f.name, page: i + 1, skusOnPage: [], skusNorm: [],
              matchedLoc: null, labelCount: 0,
              note: `SKU not found. fbaPrefix=${filenamePrefix}(no fbaId match). rawText: ${rawText}`
            })
            processed++
            continue
          }

          // Y 내림차순(상→하), 같은 행이면 X 오름차순(좌→우) 정렬
          labelInfos.sort((a, b) => {
            if (Math.abs(a.y - b.y) > 30) return b.y - a.y
            return a.x - b.x
          })

          // FBA unit ID 추출 (뒷면 페이지 숫자 표시용) — 좌표를 그대로 보존해 render 시 정확한 그리드 위치 계산
          const barcodeFound: {num: string; y: number; x: number}[] = (() => {
            const found: {num: string; y: number; x: number}[] = []
            let pending = '', pendingY = 0, pendingX = 0
            for (let j2 = 0; j2 < items.length; j2++) {
              const s = ((items[j2].str as string) || '').trim()
              if (!s) continue
              const y2 = items[j2].transform?.[5] ?? 0
              const x2 = items[j2].transform?.[4] ?? 0
              if (/^FBA[A-Z0-9]{6,}U\d{4,8}$/.test(s)) {
                found.push({num: s, y: y2, x: x2}); pending = ''; continue
              }
              if (/^FBA[A-Z0-9]{6,}$/.test(s) && !s.match(/U\d/)) {
                pending = s; pendingY = y2; pendingX = x2; continue
              }
              if (pending && /^U\d{4,8}$/.test(s)) {
                found.push({num: pending + s, y: pendingY, x: pendingX}); pending = ''; continue
              }
              if (s && !s.startsWith('FBA')) pending = ''
            }
            return found  // 정렬 안 함 — render 시 페이지 크기 기준으로 그리드 위치 계산
          })()

          const totalLabels = labelInfos.length
          const uniqueSkus = [...new Set(labelInfos.map(l => l.sku))]

          // SKU별 처리
          for (const targetSku of uniqueSkus) {
            // __fbaPrefix__ 가상 SKU 처리 (파일명 매칭)
            const loc = targetSku.startsWith('__fbaPrefix__')
              ? targetSku.replace('__fbaPrefix__', '')
              : skuToLoc(targetSku, realSkuToLocMap)
            if (!loc) continue

            const keptIndices = labelInfos
              .map((l, idx) => ({ l, idx }))
              .filter(({ l }) => l.sku === targetSku)
              .map(({ idx }) => idx)

            if (!locMap[loc]) locMap[loc] = []
            locMap[loc].push({
              fileBytes,
              pageIdx: i,
              labelsToKeep: keptIndices.length === totalLabels ? [] : keptIndices,
              totalLabels,
              barcodeFound,
            })
          }

          const matchedLocs = uniqueSkus.map(s =>
            s.startsWith('__fbaPrefix__') ? s.replace('__fbaPrefix__', '') : skuToLoc(s, realSkuToLocMap)
          ).filter(Boolean)
          const displaySkus = uniqueSkus.map(s =>
            s.startsWith('__fbaPrefix__') ? `[파일명매칭:${filenamePrefix}]` : s
          )
          debugLog.push({
            file: f.name,
            page: i + 1,
            skusOnPage: displaySkus,
            skusNorm: displaySkus.map(s => s.startsWith('[') ? s : normalizeSku(s)),
            matchedLoc: matchedLocs.join(', ') || null,
            labelCount: totalLabels,
            note: uniqueSkus.some(s => s.startsWith('__fbaPrefix__')) ? `파일명 prefix 매칭: ${filenamePrefix} → ${locFromFilename}` : undefined,
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
        const backFont = await outDoc.embedFont(StandardFonts.HelveticaBold)

        // 앞 표지 페이지
        await makeCoverPage(outDoc, loc, false)
        // 양면 인쇄용 빈 페이지 (표지 뒷면 = blank)
        outDoc.addPage([595, 842])

        for (const entry of entries) {
          const srcDoc = await PDFDocument.load(entry.fileBytes.slice())
          const [copiedPage] = await outDoc.copyPages(srcDoc, [entry.pageIdx])

          if (entry.labelsToKeep.length > 0) {
            const { width, height } = copiedPage.getSize()
            const cols = 2
            const rows = Math.max(3, Math.ceil(entry.totalLabels / cols))
            const cellW = width / cols
            const cellH = height / rows
            for (let idx = 0; idx < entry.totalLabels; idx++) {
              if (!entry.labelsToKeep.includes(idx)) {
                const col = idx % cols
                const row = Math.floor(idx / cols)
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

          // ── 뒷면 페이지: 바코드 숫자를 최대 크기로 (양면 인쇄용) ──
          if (entry.barcodeFound.length > 0) {
            const { width, height } = copiedPage.getSize()
            const backPage = outDoc.addPage([width, height])
            const cols2 = 2, rows2 = 3  // FBA 라벨 항상 2×3 고정 그리드
            const cW = width / cols2
            const cH = height / rows2

            // 바코드 y,x 좌표로 실제 그리드 위치 결정 (페이지 크기 기준)
            const barcodeGrid: Record<number, string> = {}
            for (const b of entry.barcodeFound) {
              const col = b.x < width / 2 ? 0 : 1
              const row = Math.min(rows2 - 1, Math.max(0, Math.floor((height - b.y) / cH)))
              const gridIdx = row * cols2 + col
              if (!(gridIdx in barcodeGrid)) barcodeGrid[gridIdx] = b.num
            }

            for (let idx = 0; idx < rows2 * cols2; idx++) {
              if (entry.labelsToKeep.length > 0 && !entry.labelsToKeep.includes(idx)) continue
              const num = barcodeGrid[idx]
              if (!num) continue

              const col = idx % cols2
              const row = Math.floor(idx / cols2)
              // 양면 인쇄 시 좌↔우 반전 (long-edge flip 기준)
              const backCol = (cols2 - 1) - col
              const cx = backCol * cW
              const cy = height - (row + 1) * cH

              // 셀 테두리
              backPage.drawRectangle({x: cx, y: cy, width: cW, height: cH, borderColor: rgb(0.75, 0.75, 0.75), borderWidth: 0.5})

              // U-number 파싱: 1~999 → 3자리 (001), 1000+ → 4자리 (1000)
              const uMatch = num.match(/U(\d+)$/)
              const unitNum = uMatch ? parseInt(uMatch[1], 10) : 0
              const digits = unitNum >= 1000 ? unitNum.toString() : unitNum.toString().padStart(3, '0')
              const maxW = cW - 32
              const fs = Math.min(160, maxW / (digits.length * 0.58))
              const tw = backFont.widthOfTextAtSize(digits, fs)
              backPage.drawText(digits, {x: cx + (cW - tw) / 2, y: cy + cH / 2 - fs / 2, size: fs, font: backFont, color: rgb(0, 0, 0)})
            }
          }
        }

        // 뒤 표지 페이지
        await makeCoverPage(outDoc, loc, false)

        counts[loc] = totalKept
        result[loc] = await outDoc.save()
      }

      // ── 파렛트 라벨 PDF 처리 ──────────────────────────────────
      // pltLabelFiles에서 FBA ID, FC센터, 파렛트범위 추출 → pltNotes 자동 업데이트
      if (pltFiles.length > 0) {
        setLabelStatus("파렛트 라벨 처리 중...")
        const newPltNotes: Record<string, string> = { ...pltNotes }

        for (const f of pltFiles) {
          const bytes = new Uint8Array(await f.arrayBuffer())
          const pdf2 = await pdfjsLib.getDocument(getPdfjsOpts(bytes)).promise
          let maxPlt = 0, detectedFbaId = '', detectedFc = ''

          for (let pi = 0; pi < pdf2.numPages; pi++) {
            const page = await pdf2.getPage(pi + 1)
            const content = await page.getTextContent()
            // eslint-disable-next-line @typescript-eslint/no-explicit-any
            const tokens = (content.items as any[]).map((it) => (it.str || '').trim()).filter(Boolean)
            const pageText = tokens.join(' ')

            // FBA ID 추출
            const fbaMatch = pageText.match(/FBA[A-Z0-9]{8,12}/)
            if (fbaMatch && !detectedFbaId) detectedFbaId = fbaMatch[0]

            // 파렛트 번호 추출: パレット #： N/TOTAL または #: N/TOTAL
            const pltMatch = pageText.match(/[Pp]al(?:let|レット)\s*#[：:]\s*(\d+)\s*\/\s*(\d+)/i) ||
              pageText.match(/#[：:]\s*(\d+)\s*\/\s*(\d+)/)
            if (pltMatch) {
              const total = parseInt(pltMatch[2])
              if (total > maxPlt) maxPlt = total
            }

            // 파일명에서 FBA ID 추출 (폰트 못 읽을 경우 fallback)
            if (!detectedFbaId) {
              const fnMatch = f.name.match(/^(FBA[A-Z0-9]+)/)
              if (fnMatch) detectedFbaId = fnMatch[1]
            }
            // FC센터 추출: "FBA STA (...)-HIY1" 패턴이 가장 확실
            if (!detectedFc) {
              const staMatch = pageText.match(/FBA\s+STA\s+[^-]+-([A-Z]{3,5})\b/)
              if (staMatch) detectedFc = staMatch[1]
            }
            if (!detectedFc) {
              const codeMatch = pageText.match(/\b([A-Z]{2,4}\d{1,2})\b/)
              if (codeMatch && !['FBA'].includes(codeMatch[1])) detectedFc = codeMatch[1]
            }
            if (!detectedFc) {
              for (const tok of tokens) {
                if (/^[A-Z]{3,5}\d?$/.test(tok) && !['FBA','SKU','IXD','STA','AMXL'].includes(tok) && tok !== detectedFbaId) {
                  detectedFc = tok; break
                }
              }
            }
          }
          pdf2.destroy()

          if (detectedFbaId && maxPlt > 0) {
            const pltRange = `1~${maxPlt}`
            const pltLabel = `${detectedFc ? detectedFc+'_' : ''}${detectedFbaId}_${pltRange}`
            // 해당 FBA ID를 가진 약호 찾아 pltNotes 업데이트
            for (const [sku, meta] of Object.entries(s2meta)) {
              const metaFbaId = meta.fbaId?.replace(/U\d+$/, '').toUpperCase()
              const targetPrefix = detectedFbaId.replace(/U\d+$/, '').toUpperCase()
              if (metaFbaId && metaFbaId === targetPrefix) {
                const ctnNums2 = [...new Set(fd.map(r=>r.container_no as number))]
                for (const no of ctnNums2) {
                  const key = `${no}_${sku}`
                  newPltNotes[key] = pltRange
                }
              }
            }
            // 파렛트 라벨 PDF 표지 달아서 결과에 포함 — key를 FC_FBAID_range 형태로
            const pltKey = `__PLT__${pltLabel}`
            if (!result[pltKey]) {
              const pltDoc = await PDFDocument.create()
              await makePalletCoverPage(pltDoc, detectedFbaId, detectedFc, `Pallet ${pltRange}`)
              const srcPlt = await PDFDocument.load(bytes.slice())
              const pltPages = await pltDoc.copyPages(srcPlt, Array.from({length: srcPlt.getPageCount()}, (_,i)=>i))
              pltPages.forEach(p => pltDoc.addPage(p))
              await makePalletCoverPage(pltDoc, detectedFbaId, detectedFc, `Pallet ${pltRange}`)
              result[pltKey] = await pltDoc.save()
              counts[pltKey] = maxPlt
            }
          }
        }
        setPltNotes(newPltNotes)
      } else if (Object.keys(newPltNotes ?? {}).length === 0 && pltFiles.length === 0) {
        // 파렛트 파일 없이 카톤만 처리한 경우 기존 pltNotes 유지
      }

      setLabelGroups(result)
      setLabelCounts(counts)
      const unmatchedPages = debugLog.filter(d => !d.matchedLoc).length
      setLabelStatus(`완료 — ${Object.keys(result).filter(k=>!k.startsWith('__PLT__')).length}개 BOX CODE + ${Object.keys(result).filter(k=>k.startsWith('__PLT__')).length}개 파렛트 분류됨${unmatchedPages > 0 ? ` (미매칭 ${unmatchedPages}페이지)` : ""}`)
      if (unmatchedPages > 0) setShowDebug(true)

    } catch (e) {
      alert('처리 오류: ' + (e as Error).message)
      setLabelStatus("")
    } finally {
      setLabelLoading(false)
    }
  }

  // loc → 약호/realSku 헬퍼
  function locToSku(loc: string): string {
    for (const [sku, m] of Object.entries(master)) { if (m.loc === loc) return sku }
    return ''
  }
  function locToRealSku(loc: string): string {
    for (const m of Object.values(master)) { if (m.loc === loc) return m.sku || '' }
    return ''
  }
  function locLabel(loc: string): string {
    const realSku = locToRealSku(loc)
    const sku = locToSku(loc)
    return (realSku||sku) ? `${loc} (${realSku||sku})` : loc
  }
  function dlLabel(loc: string, bytes: Uint8Array) {
    const isPlt = loc.startsWith('__PLT__')
    const ts = new Date().toISOString().slice(0,16).replace('T','_').replace(':','')
    let filename: string
    if (isPlt) {
      filename = loc.replace('__PLT__', '') + `_${ts}.pdf`
    } else {
      const realSku = locToRealSku(loc)
      const sku = locToSku(loc)
      filename = `${loc} (${realSku||sku||loc})_${ts}.pdf`
    }
    const a = document.createElement('a')
    a.href = URL.createObjectURL(new Blob([bytes], { type: 'application/pdf' }))
    a.download = filename
    a.click()
  }
  function dlAllLabels(){
    const entries = Object.entries(labelGroups)
    entries.forEach(([l,b], i) => {
      setTimeout(() => dlLabel(l, b), i * 400)
    })
  }

  function SheetTabs(){
    if(!sheetNames.length)return null
    return(<div style={{display:"flex",gap:2,borderBottom:"0.5px solid var(--color-border-tertiary)",marginBottom:16}}>{sheetNames.map(sn=>{const {data:d}=(sheets[sn]||[]).length?normHeaders(sheets[sn]):{data:[] as RowData[]};const ff2=d.length?forwardFill(d):[];const cc=new Set(ff2.map(r=>r.container_no)).size;return(<button key={sn} onClick={()=>setAs(sn)} style={{padding:"6px 14px",fontSize:12,fontWeight:activeSheet===sn?500:400,cursor:"pointer",background:"transparent",border:"none",color:activeSheet===sn?"var(--color-text-primary)":"var(--color-text-secondary)",borderBottom:activeSheet===sn?"2px solid var(--color-text-primary)":"2px solid transparent",marginBottom:-1,display:"flex",alignItems:"center",gap:5}}>{sn}{cc>0&&<span style={{fontSize:10,padding:"1px 5px",borderRadius:10,background:"var(--color-background-secondary)",border:"0.5px solid var(--color-border-tertiary)",color:"var(--color-text-tertiary)"}}>{cc}컨</span>}</button>)})}</div>)
  }

  const s2rows=buildS2rows()
  const totC=s2rows.reduce((s,r)=>s+r.total,0),totP=s2rows.reduce((s,r)=>s+r.pallets,0),totW=s2rows.reduce((s,r)=>s+r.gw,0),totCBM=Math.round(s2rows.reduce((s,r)=>s+r.cbm,0)*100)/100

  function NavBtn({m,label}:{m:string,label:string}){
    const active=mode===m,avail=m==='1'||m==='map'||m==='label'||m==='logistics'||m==='s3b'||m==='cj'||m==='fba-api'||m==='total'||m==='cipl'||m==='delivery'||!!file
    return(<button onClick={()=>avail&&setMode(m)} style={{padding:"8px 16px",fontSize:12,fontWeight:active?500:400,cursor:avail?"pointer":"default",border:"none",background:active?"var(--color-text-primary)":"transparent",color:active?"var(--color-background-primary)":avail?"var(--color-text-secondary)":"var(--color-text-tertiary)"}}>{label}</button>)
  }

  return (
    <div style={{padding:"0.75rem 0",fontFamily:"var(--font-sans)"}}>
      {/* 내비 */}
      <div style={{display:"flex",flexDirection:"column",gap:8,marginBottom:16}}>
        {/* 메인 워크플로우 */}
        <div style={{display:"flex",alignItems:"center",gap:8}}>
          <div style={{border:"0.5px solid var(--color-border-tertiary)",borderRadius:"var(--border-radius-md)",overflow:"hidden",display:"flex"}}>
            <NavBtn m="1" label="1. 업로드"/><div style={{width:"0.5px",background:"var(--color-border-tertiary)"}}/>
            <NavBtn m="2" label="1차 가공 (전품목 합산 정리)"/><div style={{width:"0.5px",background:"var(--color-border-tertiary)"}}/>
            <NavBtn m="cj" label="1.5. CJ 출고"/><div style={{width:"0.5px",background:"var(--color-border-tertiary)"}}/>
            <NavBtn m="label" label="2. 라벨 분류"/><div style={{width:"0.5px",background:"var(--color-border-tertiary)"}}/>
            <NavBtn m="logistics" label="2-5. 물류 전달"/><div style={{width:"0.5px",background:"var(--color-border-tertiary)"}}/>
            <NavBtn m="3" label="2차 가공 (CJ컨정보 전달)"/><div style={{width:"0.5px",background:"var(--color-border-tertiary)"}}/>
            <NavBtn m="fba-api" label="🚀 FBA 자동 등록"/><div style={{width:"0.5px",background:"var(--color-border-tertiary)"}}/>
            <NavBtn m="total" label="📊 총합"/>
          </div>
          <button onClick={()=>setMode('map')} style={{marginLeft:"auto",fontSize:11,padding:"6px 14px",background:mode==='map'?"var(--color-background-secondary)":"transparent",border:"0.5px solid var(--color-border-tertiary)",borderRadius:"var(--border-radius-md)",cursor:"pointer",color:mode==='map'?"var(--color-text-primary)":"var(--color-text-secondary)"}}>매핑 관리</button>
        </div>
        {/* 선적서류 섹션 */}
        <div style={{display:"flex",alignItems:"center",gap:8}}>
          <span style={{fontSize:11,fontWeight:600,color:"var(--color-text-tertiary)",whiteSpace:"nowrap",letterSpacing:"0.04em"}}>선적서류</span>
          <div style={{border:"0.5px solid var(--color-border-tertiary)",borderRadius:"var(--border-radius-md)",overflow:"hidden",display:"flex"}}>
            <NavBtn m="s3b" label="역산시트"/><div style={{width:"0.5px",background:"var(--color-border-tertiary)"}}/>
            <NavBtn m="cipl" label="CI/PL 변환기"/><div style={{width:"0.5px",background:"var(--color-border-tertiary)"}}/>
            <NavBtn m="delivery" label="납품처 정보"/>
          </div>
        </div>
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
              {ctnNums.map(no=>{const rows=fd.filter(r=>r.container_no===no);const f0=rows[0],ci=(no-1)%CB.length;const sumQ=rows.reduce((s,r)=>s+r.quantity,0);const dQ=parseFloat(String(f0.qty_total))||sumQ;const key="s1_"+no,open=isOpen(key)
                // 데이터 있는 컬럼만 표시
                const hasDate=rows.some(r=>r.shipment_date&&String(r.shipment_date).trim()!=='')
                const hasDest=rows.some(r=>r.destination&&String(r.destination).trim()!=='')
                const hasModel=rows.some(r=>r.model_code&&String(r.model_code).trim()!=='')
                const hasColor=rows.some(r=>r.color&&String(r.color).trim()!=='')
                const hasTypeOld=rows.some(r=>r.type_old&&String(r.type_old).trim()!=='')
                const hasEtd=rows.some(r=>r.etd&&String(r.etd).trim()!=='')
                const hasEta=rows.some(r=>r.eta&&String(r.eta).trim()!=='')
                return(
                <div key={no} style={{border:"0.5px solid var(--color-border-tertiary)",borderRadius:"var(--border-radius-md)",overflow:"hidden"}}>
                  <div onClick={()=>togCtn(key)} style={{background:CB[ci],color:CT[ci],padding:"9px 14px",display:"flex",alignItems:"center",gap:12,flexWrap:"wrap",cursor:"pointer",userSelect:"none"}}>
                    <span style={{fontWeight:500,fontSize:13}}>컨{no}</span>
                    {hasDate&&<span style={{fontWeight:500}}>{String(f0.shipment_date||"")}</span>}
                    {tag("",f0.shipment_time)}
                    {hasDest&&<span style={{fontWeight:500}}>{String(f0.destination||"")}</span>}
                    {tag("ETD ",f0.etd)}{tag("ETA ",f0.eta)}{tag("선사 ",f0.carrier)}
                    <span style={{marginLeft:"auto",display:"flex",alignItems:"center",gap:8}}><span style={{fontWeight:500,fontSize:13}}>{dQ.toLocaleString()}개{f0.ctn_count&&f0.ctn_count!=="0"?" · "+f0.ctn_count+"컨":""}{f0.ft&&f0.ft!=="0"?" · "+f0.ft:""}</span><span style={{fontSize:11,opacity:0.6}}>{open?"▲":"▼"}</span></span>
                  </div>
                  {open&&(<div style={{overflowX:"auto"}}><table style={{width:"100%",borderCollapse:"collapse",fontSize:12,minWidth:400}}>
                    <thead><tr>
                      {hasModel&&<th style={TH}>BOM품번</th>}
                      <th style={TH}>약호(SKU)</th>
                      {hasColor&&<th style={TH}>색상</th>}
                      <th style={{...TH,textAlign:"right"}}>수량</th>
                      {hasTypeOld&&<th style={TH}>TYPE(구)</th>}
                      <th style={TH}>신박스코드</th>
                      {hasEtd&&<th style={TH}>ETD</th>}
                      {hasEta&&<th style={TH}>ETA</th>}
                    </tr></thead>
                    <tbody>{rows.map((r,i)=>{
                      // TYPE(구) 없으면 마스터에서 자동
                      const typeOld = String(r.type_old||'')||String((master[String(r.sku||'')] as {to?:string})?.to||'')
                      const loc = String(r.location||'')||String((master[String(r.sku||'')] as {loc?:string})?.loc||'')
                      return(<tr key={i} style={{background:i%2===0?"transparent":"var(--color-background-secondary)"}}>
                        {hasModel&&<td style={{...TD,fontFamily:"var(--font-mono)",fontSize:11,color:"var(--color-text-secondary)",minWidth:120}}>{String(r.model_code||"")}</td>}
                        <td style={{...TD,fontWeight:500,minWidth:120}}>{String(r.sku||"")}</td>
                        {hasColor&&<td style={{...TD,minWidth:50}}>{String(r.color||"")}</td>}
                        <td style={{...TD,textAlign:"right",fontWeight:500,minWidth:50}}>{r.quantity.toLocaleString()}</td>
                        {hasTypeOld&&<td style={{...TD,color:"var(--color-text-secondary)"}}>{typeOld}</td>}
                        <td style={{...TD,color:"var(--color-text-info)",fontWeight:500}}>{loc}</td>
                        {hasEtd&&<td style={TD}>{String(r.etd||"")}</td>}
                        {hasEta&&<td style={TD}>{String(r.eta||"")}</td>}
                      </tr>)
                    })}</tbody>
                    <tfoot><tr style={{background:"var(--color-background-secondary)",borderTop:"1px solid var(--color-border-secondary)"}}>
                      <td colSpan={[hasModel,true,hasColor,true,hasTypeOld,true,hasEtd,hasEta].filter(Boolean).length-1} style={{...TD,textAlign:"right",fontSize:11,color:"var(--color-text-secondary)"}}>소계</td>
                      <td style={{...TD,textAlign:"right",fontWeight:500}}>{sumQ.toLocaleString()}</td>
                      <td colSpan={[hasTypeOld,true,hasEtd,hasEta].filter(Boolean).length} style={TD}></td>
                    </tr></tfoot>
                  </table></div>)}
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
            {/* 변동값 JSON 백업/복원 */}
            <button onClick={exportMetaJson} style={{fontSize:11,padding:"3px 10px",cursor:"pointer",border:"0.5px solid var(--color-border-secondary)",borderRadius:"var(--border-radius-md)",background:"transparent",color:"var(--color-text-secondary)"}}>⬇ 메타 JSON 백업</button>
            <label style={{fontSize:11,padding:"3px 10px",cursor:"pointer",border:"0.5px solid var(--color-border-info)",borderRadius:"var(--border-radius-md)",background:"rgba(219,234,254,0.25)",color:"var(--color-text-info)"}}>
              ⬆ 메타 JSON 복원
              <input ref={metaJsonRef} type="file" accept=".json" style={{display:"none"}} onChange={e=>{const f=e.target.files?.[0];if(f){importMetaJson(f);e.target.value=""}}}/>
            </label>
            <button onClick={expS2} style={{marginLeft:"auto",fontSize:11,padding:"3px 10px"}}>xlsx 저장</button>
            {/* FBA 업로드 양식: 컨테이너 선택 후 다운로드 */}
            <div style={{display:"flex",alignItems:"center",gap:4,flexWrap:"wrap"}}>
              <span style={{fontSize:11,color:"var(--color-text-tertiary)"}}>FBA 업로드:</span>
              {ctnNums.map(no=>(
                <button key={no} onClick={()=>setFbaUploadCtnFilter(prev=>prev.includes(no)?prev.filter(n=>n!==no):[...prev,no])} style={{fontSize:11,padding:"2px 8px",borderRadius:4,border:`0.5px solid ${fbaUploadCtnFilter.includes(no)?"var(--color-border-info)":"var(--color-border-tertiary)"}`,background:fbaUploadCtnFilter.includes(no)?"rgba(219,234,254,0.5)":"transparent",color:fbaUploadCtnFilter.includes(no)?"var(--color-text-info)":"var(--color-text-secondary)",cursor:"pointer",fontWeight:fbaUploadCtnFilter.includes(no)?600:400}}>컨{no}</button>
              ))}
              <button onClick={()=>setFbaUploadCtnFilter([])} style={{fontSize:10,padding:"2px 6px",borderRadius:4,border:"0.5px solid var(--color-border-tertiary)",background:"transparent",color:"var(--color-text-tertiary)",cursor:"pointer"}}>전체</button>
              <button onClick={()=>expFbaUpload(fbaUploadCtnFilter.length>0?fbaUploadCtnFilter:undefined)} disabled={fbaLoading} style={{fontSize:11,padding:"3px 10px",background:"var(--color-background-success)",border:"0.5px solid var(--color-border-success)",color:"var(--color-text-success)",borderRadius:"var(--border-radius-md)",cursor:fbaLoading?"default":"pointer",opacity:fbaLoading?0.6:1,fontWeight:500}}>
                {fbaLoading?"처리 중...":`↓ ${fbaUploadCtnFilter.length>0?`컨${fbaUploadCtnFilter.join('·')} 분리`:"전체"}`}
              </button>
            </div>
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
                {s2rows.map((r,i)=>(<tr key={r.sku} style={{background:i%2===0?"transparent":"var(--color-background-secondary)"}}><td style={{...TD,fontWeight:500,minWidth:130}}>{r.sku}</td><td style={{...TD,fontFamily:"var(--font-mono)",fontSize:11,minWidth:100}}>{r.asin}</td><td style={{...TD,minWidth:45}}>{r.to}</td><td style={{...TD,color:"var(--color-text-info)",fontWeight:500,minWidth:70}}>{r.loc}</td><td style={{...TD,textAlign:"right",fontWeight:500,minWidth:55}}>{r.total.toLocaleString()}</td><td style={{...TD,textAlign:"center",minWidth:55}}>{r.cpp}</td><td style={{...TD,textAlign:"right",fontWeight:500,minWidth:50}}>{r.pallets}</td><td style={{...TD,textAlign:"right",minWidth:70,borderRight:"1px solid var(--color-border-tertiary)"}}>{r.gw.toFixed(1)}</td><td style={{...TD,textAlign:"right",minWidth:60,borderRight:"2px solid var(--color-border-secondary)"}}>{r.cbm}</td>{(["fc","address","fbaId","amazonId"] as const).map((f,idx)=>(<td key={f} style={{...TD,minWidth:[80,150,130,90][idx],...IBLU}}><input
                  value={r[f]}
                  onChange={e=>updMeta(r.sku,f,e.target.value)}
                  onPaste={e=>{
                    const text = e.clipboardData.getData('text')
                    // 줄바꿈이 있으면 멀티행 붙여넣기
                    const lines = text.split(/\r?\n/).map(l=>l.trim()).filter((_,li)=>li===0||_!=='')
                    if (lines.length <= 1) return // 단일행이면 기본 동작
                    e.preventDefault()
                    // 현재 행(i)부터 아래로 순서대로 적용
                    lines.forEach((val, offset) => {
                      const targetRow = s2rows[i + offset]
                      if (targetRow) updMeta(targetRow.sku, f, val)
                    })
                  }}
                  style={INP}
                  placeholder={["FC CENTER","주소","FBA ID","아마존 ID"][idx]}
                /></td>))}</tr>))}
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
            <span style={{fontSize:12,color:"var(--color-text-secondary)"}}>카톤 + 파렛트 라벨 PDF → BOX CODE별 분류 · 표지 자동 삽입</span>
            {activeSheet&&<span style={{fontSize:11,padding:"2px 8px",borderRadius:10,background:"rgba(219,234,254,0.4)",border:"0.5px solid var(--color-border-info)",color:"var(--color-text-info)"}}>현재 시트: {activeSheet}</span>}
            {Object.keys(labelGroups).length>0&&(<button onClick={dlAllLabels} style={{marginLeft:"auto",fontSize:11,padding:"3px 12px",background:"var(--color-background-success)",border:"0.5px solid var(--color-border-success)",color:"var(--color-text-success)",borderRadius:"var(--border-radius-md)",cursor:"pointer",fontWeight:500}}>전체 다운로드</button>)}
          </div>
          <SheetTabs/>

          <div style={{display:"grid",gridTemplateColumns:"1fr 1fr",gap:12,marginBottom:16}}>
            {/* 카톤 라벨 드롭존 */}
            <div>
              <div style={{fontSize:11,color:"var(--color-text-tertiary)",marginBottom:6,fontWeight:500}}>카톤 라벨 PDF</div>
              <div onClick={()=>!labelLoading&&labelRef.current?.click()} onDragOver={e=>e.preventDefault()} onDrop={e=>{e.preventDefault();const fs=Array.from(e.dataTransfer.files).filter(f=>f.type==='application/pdf');if(fs.length&&!labelLoading){setLabelFiles(fs);processLabels(fs,pltLabelFiles)}}} style={{display:"flex",flexDirection:"column",alignItems:"center",justifyContent:"center",padding:"24px 16px",border:"1px dashed var(--color-border-secondary)",borderRadius:"var(--border-radius-md)",cursor:labelLoading?"default":"pointer",gap:6,background:labelLoading?"var(--color-background-secondary)":"transparent",minHeight:100}}>
                <span style={{fontSize:28}}>{labelLoading?"⏳":"📄"}</span>
                <span style={{fontSize:12,color:"var(--color-text-secondary)",textAlign:"center"}}>{labelLoading?labelStatus:labelFiles.length?`${labelFiles.length}개 파일 (${labelStatus})`:"카톤 라벨 업로드"}</span>
              </div>
              <input ref={labelRef} type="file" accept="application/pdf" multiple style={{display:"none"}} onChange={e=>{const fs=Array.from(e.target.files||[]);if(fs.length){setLabelFiles(fs);processLabels(fs,pltLabelFiles)}}}/>
            </div>

            {/* 파렛트 라벨 드롭존 */}
            <div>
              <div style={{fontSize:11,color:"var(--color-text-tertiary)",marginBottom:6,fontWeight:500}}>파렛트 라벨 PDF <span style={{color:"var(--color-text-info)"}}>→ 파렛트 번호 자동 인식</span></div>
              <div onClick={()=>!labelLoading&&pltLabelRef.current?.click()} onDragOver={e=>e.preventDefault()} onDrop={e=>{e.preventDefault();const fs=Array.from(e.dataTransfer.files).filter(f=>f.type==='application/pdf');if(fs.length&&!labelLoading){setPltLabelFiles(fs);processPalletOnly(fs)}}} style={{display:"flex",flexDirection:"column",alignItems:"center",justifyContent:"center",padding:"24px 16px",border:"1px dashed var(--color-border-info)",borderRadius:"var(--border-radius-md)",cursor:labelLoading?"default":"pointer",gap:6,background:pltLabelFiles.length?"rgba(219,234,254,0.15)":"transparent",minHeight:100}}>
                <span style={{fontSize:28}}>{pltLabelFiles.length?"🏷️":"📋"}</span>
                <span style={{fontSize:12,color:"var(--color-text-secondary)",textAlign:"center"}}>{pltLabelFiles.length?`${pltLabelFiles.length}개 파일 — 파렛트 번호 자동 반영됨`:"파렛트 라벨 업로드 (업로드 즉시 2-5에 반영)"}</span>
              </div>
              <input ref={pltLabelRef} type="file" accept="application/pdf" multiple style={{display:"none"}} onChange={e=>{const fs=Array.from(e.target.files||[]);if(fs.length){setPltLabelFiles(fs);processPalletOnly(fs)}}}/>
            </div>
          </div>
          {/* 카톤 라벨 결과 */}
          {Object.keys(labelGroups).filter(k=>!k.startsWith('__PLT__')).length>0&&!labelLoading&&(
            <div style={{marginBottom:12}}>
              <div style={{fontSize:11,fontWeight:600,color:"var(--color-text-primary)",padding:"6px 10px",background:"var(--color-background-secondary)",borderRadius:"var(--border-radius-md) var(--border-radius-md) 0 0",borderBottom:"0.5px solid var(--color-border-tertiary)",border:"0.5px solid var(--color-border-tertiary)"}}>📦 카톤 라벨</div>
              <div style={{border:"0.5px solid var(--color-border-tertiary)",borderTop:"none",borderRadius:"0 0 var(--border-radius-md) var(--border-radius-md)",overflow:"hidden"}}>
                <table style={{width:"100%",borderCollapse:"collapse",fontSize:12}}>
                  <thead><tr>
                    <th style={TH}>BOX CODE (realSku)</th>
                    <th style={{...TH,textAlign:"right"}}>라벨 수</th>
                    <th style={{...TH,textAlign:"right"}}>1차 가공 수량</th>
                    <th style={{...TH,textAlign:"right"}}>차이</th>
                    <th style={{...TH,textAlign:"right"}}>A4용지 (6개/장)</th>
                    <th style={{...TH,textAlign:"right"}}>파일 크기</th>
                    <th style={TH}></th>
                  </tr></thead>
                  <tbody>{Object.entries(labelGroups).filter(([k])=>!k.startsWith('__PLT__')).sort(([a],[b])=>a.localeCompare(b)).map(([loc,bytes],i)=>{
                    const s2qty = s2rows.filter(r=>r.loc===loc).reduce((s,r)=>s+r.total,0)
                    const labelQty = labelCounts[loc]||0
                    const diff = labelQty - s2qty
                    const hasData = s2qty > 0
                    const isOk = diff === 0
                    const sheets = Math.ceil(labelQty/6)
                    return(<tr key={loc} style={{background:i%2===0?"transparent":"var(--color-background-secondary)"}}>
                      <td style={{...TD,fontWeight:500}}>{locLabel(loc)}</td>
                      <td style={{...TD,textAlign:"right",fontWeight:500}}>{labelQty.toLocaleString()}장</td>
                      <td style={{...TD,textAlign:"right",color:"var(--color-text-secondary)"}}>{hasData?s2qty.toLocaleString()+"개":"—"}</td>
                      <td style={{...TD,textAlign:"right",fontWeight:500,color:!hasData?"var(--color-text-tertiary)":isOk?"var(--color-text-success)":Math.abs(diff)<=3?"#f59e0b":"var(--color-text-danger)"}}>
                        {!hasData?"—":isOk?"✅ 일치":diff>0?`+${diff} 초과`:`${diff} 부족`}
                      </td>
                      <td style={{...TD,textAlign:"right",fontWeight:600,color:"#dc2626"}}>{sheets}장</td>
                      <td style={{...TD,textAlign:"right",color:"var(--color-text-secondary)"}}>{(bytes.length/1024).toFixed(0)}KB</td>
                      <td style={TD}><button onClick={()=>dlLabel(loc,bytes)} style={{fontSize:11,padding:"2px 10px",cursor:"pointer",borderRadius:4,border:"0.5px solid var(--color-border-secondary)",background:"transparent"}}>↓ PDF</button></td>
                    </tr>)
                  })}</tbody>
                  <tfoot><tr style={{background:"var(--color-background-secondary)",borderTop:"1.5px solid var(--color-border-secondary)"}}>
                    <td style={{...TD,fontWeight:500}}>카톤 합계</td>
                    <td style={{...TD,textAlign:"right",fontWeight:500}}>{Object.entries(labelCounts).filter(([k])=>!k.startsWith('__PLT__')).reduce((s,[,c])=>s+c,0).toLocaleString()}장</td>
                    <td style={{...TD,textAlign:"right",fontWeight:500}}>{s2rows.reduce((s,r)=>s+r.total,0).toLocaleString()}개</td>
                    <td style={{...TD,textAlign:"right",fontWeight:500,color:(()=>{const tl=Object.entries(labelCounts).filter(([k])=>!k.startsWith('__PLT__')).reduce((s,[,c])=>s+c,0);const ts=s2rows.reduce((s,r)=>s+r.total,0);return tl===ts?"var(--color-text-success)":"var(--color-text-danger)"})()}}>
                      {(()=>{const tl=Object.entries(labelCounts).filter(([k])=>!k.startsWith('__PLT__')).reduce((s,[,c])=>s+c,0);const ts=s2rows.reduce((s,r)=>s+r.total,0);const d=tl-ts;return tl===ts?"✅ 일치":d>0?`+${d} 초과`:`${d} 부족`})()}
                    </td>
                    <td style={{...TD,textAlign:"right",fontWeight:700,color:"#dc2626"}}>{Math.ceil(Object.entries(labelCounts).filter(([k])=>!k.startsWith('__PLT__')).reduce((s,[,c])=>s+c,0)/6)}장 필요</td>
                    <td style={{...TD,textAlign:"right",fontWeight:500}}>{(Object.entries(labelGroups).filter(([k])=>!k.startsWith('__PLT__')).reduce((s,[,b])=>s+b.length,0)/1024).toFixed(0)}KB</td>
                    <td style={TD}></td>
                  </tr></tfoot>
                </table>
              </div>
            </div>
          )}

          {/* 파렛트 라벨 결과 */}
          {Object.keys(labelGroups).filter(k=>k.startsWith('__PLT__')).length>0&&!labelLoading&&(
            <div style={{marginBottom:12}}>
              <div style={{fontSize:11,fontWeight:600,color:"var(--color-text-info)",padding:"6px 10px",background:"rgba(219,234,254,0.3)",borderRadius:"var(--border-radius-md) var(--border-radius-md) 0 0",border:"0.5px solid var(--color-border-info)",borderBottom:"0.5px solid var(--color-border-tertiary)"}}>🏷️ 파렛트 라벨</div>
              <div style={{border:"0.5px solid var(--color-border-info)",borderTop:"none",borderRadius:"0 0 var(--border-radius-md) var(--border-radius-md)",overflow:"hidden"}}>
                <table style={{width:"100%",borderCollapse:"collapse",fontSize:12}}>
                  <thead><tr>
                    <th style={TH}>FC센터 / FBA ID</th>
                    <th style={{...TH,textAlign:"right"}}>파렛트 수</th>
                    <th style={{...TH,textAlign:"right"}}>A4용지 (6개/장)</th>
                    <th style={{...TH,textAlign:"right"}}>파일 크기</th>
                    <th style={TH}></th>
                  </tr></thead>
                  <tbody>{Object.entries(labelGroups).filter(([k])=>k.startsWith('__PLT__')).map(([loc,bytes],i)=>{
                    const displayKey = loc.replace('__PLT__','')
                    const cnt = labelCounts[loc]||0
                    const sheets = Math.ceil(cnt/6)
                    return(<tr key={loc} style={{background:i%2===0?"rgba(219,234,254,0.08)":"rgba(219,234,254,0.2)"}}>
                      <td style={{...TD,fontWeight:500,color:"var(--color-text-info)"}}>{displayKey}</td>
                      <td style={{...TD,textAlign:"right",fontWeight:600}}>{cnt}파렛트</td>
                      <td style={{...TD,textAlign:"right",fontWeight:700,color:"#dc2626"}}>{sheets}장</td>
                      <td style={{...TD,textAlign:"right",color:"var(--color-text-secondary)"}}>{(bytes.length/1024).toFixed(0)}KB</td>
                      <td style={TD}><button onClick={()=>dlLabel(loc,bytes)} style={{fontSize:11,padding:"2px 10px",cursor:"pointer",borderRadius:4,border:"0.5px solid var(--color-border-info)",background:"transparent"}}>↓ PDF</button></td>
                    </tr>)
                  })}</tbody>
                  <tfoot><tr style={{background:"rgba(219,234,254,0.3)",borderTop:"1.5px solid var(--color-border-info)"}}>
                    <td style={{...TD,fontWeight:500,color:"var(--color-text-info)"}}>파렛트 합계</td>
                    <td style={{...TD,textAlign:"right",fontWeight:500}}>{Object.entries(labelCounts).filter(([k])=>k.startsWith('__PLT__')).reduce((s,[,c])=>s+c,0)}파렛트</td>
                    <td style={{...TD,textAlign:"right",fontWeight:700,color:"#dc2626"}}>{Math.ceil(Object.entries(labelCounts).filter(([k])=>k.startsWith('__PLT__')).reduce((s,[,c])=>s+c,0)/6)}장 필요</td>
                    <td style={{...TD,textAlign:"right",fontWeight:500}}>{(Object.entries(labelGroups).filter(([k])=>k.startsWith('__PLT__')).reduce((s,[,b])=>s+b.length,0)/1024).toFixed(0)}KB</td>
                    <td style={TD}></td>
                  </tr></tfoot>
                </table>
              </div>
            </div>
          )}

          {/* 전체 라벨 리스트 xlsx 저장 */}
          {(Object.keys(labelGroups).filter(k=>!k.startsWith('__PLT__')).length>0||Object.keys(labelGroups).filter(k=>k.startsWith('__PLT__')).length>0)&&!labelLoading&&(
            <div style={{marginBottom:12,display:"flex",justifyContent:"flex-end"}}>
              <button onClick={()=>{
                const hdr=[['구분','BOX CODE / FBA ID','realSku / FC센터','라벨수','A4용지(6개/장)','파일크기(KB)']]
                const cartonRows = Object.entries(labelCounts).filter(([k])=>!k.startsWith('__PLT__')).sort(([a],[b])=>a.localeCompare(b)).map(([loc,cnt])=>{
                  const realSku=locToRealSku(loc)||locToSku(loc)||''
                  return ['카톤라벨', loc, realSku, cnt, Math.ceil(cnt/6), ((labelGroups[loc]?.length||0)/1024).toFixed(0)]
                })
                const pltRows = Object.entries(labelCounts).filter(([k])=>k.startsWith('__PLT__')).map(([loc,cnt])=>{
                  const key=loc.replace('__PLT__','')
                  return ['파렛트라벨', key, key, cnt, Math.ceil(cnt/6), ((labelGroups[loc]?.length||0)/1024).toFixed(0)]
                })
                const totalCarton=Object.entries(labelCounts).filter(([k])=>!k.startsWith('__PLT__')).reduce((s,[,c])=>s+c,0)
                const totalPlt=Object.entries(labelCounts).filter(([k])=>k.startsWith('__PLT__')).reduce((s,[,c])=>s+c,0)
                const totalRow=[['합계','','',totalCarton+totalPlt,Math.ceil(totalCarton/6)+Math.ceil(totalPlt/6),''],
                  ['카톤합계','','',totalCarton,Math.ceil(totalCarton/6),''],
                  ['파렛트합계','','',totalPlt,Math.ceil(totalPlt/6),'']]
                xlsDl([...hdr,...cartonRows,...pltRows,...totalRow],'라벨리스트','라벨리스트_'+new Date().toISOString().slice(0,10)+'.xlsx')
              }} style={{fontSize:11,padding:"5px 14px",cursor:"pointer",border:"0.5px solid var(--color-border-secondary)",borderRadius:"var(--border-radius-md)",background:"var(--color-background-secondary)",color:"var(--color-text-primary)",fontWeight:500}}>
                📊 전체 라벨 리스트 xlsx 저장
              </button>
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
                            {d.note&&<div style={{fontSize:9,color:"var(--color-text-tertiary)",marginTop:2}}>{d.note}</div>}
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

      {/* ══ ②.5 물류 전달 ══ */}
      {mode==='logistics'&&(()=>{
        // 컨테이너별로 묶고, 각 행에 약호별 FC센터 표시
        // 필요 컬럼: 약호, ASIN, 신박스코드, 카톤수량, 파렛트당카톤, 파렛트, FC센터, 파렛트번호(수기)
        const groups = buildS3groups()

        function expLogistics() {
          const hdr = [['컨테이너','약호','신박스코드','카톤수량','파렛트당카톤','파렛트','FC센터','파렛트번호']]
          const body: unknown[][] = []
          for (const g of groups) {
            for (const r of g.rows) {
              const m = master[String(r.sku)] || {} as MasterItem
              const cpp = parseFloat(String(m.cpp)) || 16
              const pallets = Math.ceil(r.quantity / cpp)
              const meta = s2meta[String(r.sku)] || {}
              const key = `${g.no}_${String(r.sku)}`
              body.push([
                g.container || `컨${g.no}`,
                String(r.sku),
                String(r.location || m.loc || ''),
                r.quantity, cpp, pallets,
                meta.fc || '', pltNotes[key] || ''
              ])
            }
          }
          // 합계행
          const totalQty = groups.reduce((s,g)=>s+g.rows.reduce((s2,r)=>s2+r.quantity,0),0)
          const totalPlt = groups.reduce((s,g)=>s+g.rows.reduce((s2,r)=>{const cpp=parseFloat(String((master[String(r.sku)]||{}).cpp))||16;return s2+Math.ceil(r.quantity/cpp)},0),0)
          body.push(['합계','','',totalQty,'',totalPlt,'',''])
          xlsDl([...hdr, ...body], '물류전달', '물류전달_' + new Date().toISOString().slice(0,10) + '.xlsx')
        }

        function printLogistics() {
          const groups2 = buildS3groups()
          type PRow = {ctn:string, ctnNo:number, sku:string, realSku:string, loc:string, qty:number, pallets:number, fc:string, plt:string}
          // 컨테이너별로 그룹핑
          const ctnGroups: {no:number, ctn:string, date:string, dest:string, rows:PRow[]}[] = []
          for (const g of groups2) {
            const rows: PRow[] = g.rows.map(r => {
              const m = master[String(r.sku)] || {} as MasterItem
              const cpp = parseFloat(String(m.cpp)) || 16
              const pallets = Math.ceil(r.quantity / cpp)
              const meta = s2meta[String(r.sku)] || {}
              const key = `${g.no}_${String(r.sku)}`
              return { ctn:`컨${g.no}`, ctnNo:g.no, sku:String(r.sku), realSku:m.sku||String(r.sku), loc:String(r.location||m.loc||''), qty:r.quantity, pallets, fc:meta.fc||'', plt:pltNotes[key]||'' }
            })
            ctnGroups.push({ no:g.no, ctn:g.container||`컨${g.no}`, date:g.shipment_date||'', dest:g.destination||'', rows })
          }

          const totalQty = ctnGroups.reduce((s,g)=>s+g.rows.reduce((s2,r)=>s2+r.qty,0),0)
          const totalPlt = ctnGroups.reduce((s,g)=>s+g.rows.reduce((s2,r)=>s2+r.pallets,0),0)

          const FC_BG: Record<string,string> = { HIY1:'#1e3a8a', TPB5:'#14532d', VJNB:'#78350f', TPB8:'#6b21a8' }
          const FC_ROW: Record<string,string> = { HIY1:'#dbeafe', TPB5:'#dcfce7', VJNB:'#fef9c3', TPB8:'#f3e8ff' }

          // 자연스럽게 흐르는 레이아웃 — break-inside:avoid로 찢기지 않게
          const ctnBlocks = ctnGroups.map((g) => {
            const ctnSumQ = g.rows.reduce((s,r)=>s+r.qty,0)
            const ctnSumP = g.rows.reduce((s,r)=>s+r.pallets,0)
            const dataRows = g.rows.map(r => {
              const fc = r.fc.toUpperCase()
              const fcBg = FC_BG[fc] || '#374151'
              const rowBg = FC_ROW[fc] || '#f9fafb'
              return `<tr style="background:${rowBg}">
                <td class="sku">${r.sku}</td>
                <td class="rsku">${r.realSku}</td>
                <td class="loc">${r.loc}</td>
                <td class="num">${r.qty.toLocaleString()}</td>
                <td class="plt">${r.pallets}</td>
                <td class="fc" style="background:${fcBg}">${r.fc}</td>
                <td class="pltno">${r.plt||'—'}</td>
              </tr>`
            }).join('')
            return `<div class="ctn-block">
              <div class="ctn-title"><span class="ctn-badge">컨${g.no}</span></div>
              <table>
                <colgroup>
                  <col style="width:100px"><col style="width:155px"><col style="width:75px">
                  <col style="width:78px"><col style="width:65px"><col style="width:60px"><col style="width:87px">
                </colgroup>
                <thead>
                  <tr class="gh">
                    <th colspan="3" class="carton-h">카톤 라벨</th>
                    <th colspan="2" class="carton-h">수량</th>
                    <th colspan="2" class="pallet-h">파렛트 라벨</th>
                  </tr>
                  <tr class="ch">
                    <th>약호</th><th>Merchant SKU</th><th>신박스코드</th>
                    <th>카톤</th><th>파렛트</th>
                    <th>FC센터</th><th>파렛트번호</th>
                  </tr>
                </thead>
                <tbody>${dataRows}
                  <tr class="sub-total">
                    <td colspan="3">소계</td>
                    <td>${ctnSumQ.toLocaleString()}</td>
                    <td>${ctnSumP}</td>
                    <td colspan="2"></td>
                  </tr>
                </tbody>
              </table>
            </div>`
          }).join('\n')

          const html = `<!DOCTYPE html>
<html><head><meta charset="utf-8">
<style>
  @page{size:A4 landscape;margin:7mm 9mm}
  *{box-sizing:border-box;-webkit-print-color-adjust:exact;print-color-adjust:exact}
  body{font-family:'Malgun Gothic','Arial',sans-serif;font-size:13px;margin:0;padding:0;color:#111}
  .ctn-block{margin-bottom:10px;break-inside:avoid;page-break-inside:avoid}
  .ctn-title{margin-bottom:0}
  .ctn-badge{font-size:16px;font-weight:800;background:#0f172a;color:#fff;padding:3px 12px;border-radius:4px 4px 0 0;display:inline-block}
  table{width:100%;border-collapse:collapse}
  th,td{border:1.5px solid #94a3b8;padding:6px 7px;vertical-align:middle}
  tr.gh th{font-size:12px;font-weight:700;text-align:center;padding:5px 7px}
  .carton-h{background:#1e293b;color:#fff}
  .pallet-h{background:#1e40af;color:#fff}
  tr.ch th{background:#334155;color:#e2e8f0;font-size:11px;text-align:center;padding:5px 7px}
  td.sku{font-weight:700;font-size:13px;white-space:nowrap}
  td.rsku{color:#475569;font-size:11px;white-space:nowrap}
  td.loc{font-weight:800;color:#1e40af;text-align:center;font-size:14px}
  td.num{text-align:right;font-size:14px}
  td.plt{text-align:right;font-weight:800;font-size:15px}
  td.fc{text-align:center;font-weight:800;font-size:14px;color:#fff;border:2px solid #0f172a}
  td.pltno{text-align:center;font-weight:800;font-size:15px;color:#1e40af}
  tr.sub-total td{background:#1e293b;color:#fff;font-weight:700;font-size:12px;text-align:right}
  tr.sub-total td:first-child{text-align:left;padding-left:10px}
  .total-block{background:#0f172a;color:#fff;padding:7px 12px;border-radius:4px;margin-top:8px;display:flex;gap:24px;font-size:13px;font-weight:700;break-inside:avoid;page-break-inside:avoid}
</style>
</head><body>
${ctnBlocks}
<div class="total-block">
  <span>합계</span>
  <span>카톤 ${totalQty.toLocaleString()}개</span>
  <span>파렛트 ${totalPlt}개</span>
  <span style="margin-left:auto;font-size:11px;opacity:0.6">${new Date().toLocaleDateString('ko-KR')} &nbsp; ${activeSheet||''}</span>
</div>
</body></html>`

          const w = window.open('', '_blank')
          if (w) { w.document.write(html); w.document.close(); setTimeout(()=>w.print(), 500) }
        }

        return (
          <div>
            <div style={{display:"flex",alignItems:"center",gap:8,marginBottom:14,flexWrap:"wrap"}}>
              <span style={{fontSize:12,color:"var(--color-text-secondary)"}}>라벨 배송 정보 정리 · 컨테이너별 분류 · FC센터 표시</span>
              <button onClick={printLogistics} style={{marginLeft:"auto",fontSize:11,padding:"3px 10px",cursor:"pointer",border:"0.5px solid var(--color-border-secondary)",borderRadius:"var(--border-radius-md)",background:"transparent"}}>🖨 인쇄</button>
              <button onClick={expLogistics} style={{fontSize:11,padding:"3px 10px"}}>xlsx 저장</button>
            </div>
            <SheetTabs/>

            {!file ? (
              <div style={{textAlign:"center",padding:"4rem 0",color:"var(--color-text-tertiary)"}}>
                <div style={{fontSize:36,marginBottom:12}}>📋</div>
                <p>파일을 먼저 업로드하고 1차 가공에서 FC CENTER를 입력해주세요</p>
              </div>
            ) : (
              <div style={{display:"flex",flexDirection:"column",gap:12}}>
                {groups.map((g) => {
                  const ci = (g.no - 1) % CB.length
                  const sumQ = g.rows.reduce((s,r)=>s+r.quantity, 0)
                  const sumP = g.rows.reduce((s,r)=>{
                    const cpp = parseFloat(String((master[String(r.sku)]||{}).cpp))||16
                    return s + Math.ceil(r.quantity/cpp)
                  }, 0)
                  return (
                    <div key={g.no} style={{border:"0.5px solid var(--color-border-tertiary)",borderRadius:"var(--border-radius-md)",overflow:"hidden"}}>
                      {/* 컨테이너 헤더 */}
                      <div style={{background:CB[ci],color:CT[ci],padding:"8px 14px",display:"flex",alignItems:"center",gap:12,flexWrap:"wrap"}}>
                        <span style={{fontWeight:600,fontSize:13}}>컨{g.no}</span>
                        {g.container&&<span style={{fontWeight:500,fontFamily:"var(--font-mono)",fontSize:12}}>{g.container}</span>}
                        {g.shipment_date&&<span style={{fontSize:12}}>{g.shipment_date}</span>}
                        {g.destination&&<span style={{fontSize:12}}>{g.destination}</span>}
                        {g.etd&&<span style={{fontSize:11,opacity:0.8}}>ETD {g.etd}</span>}
                        <span style={{marginLeft:"auto",fontWeight:500,fontSize:12}}>{sumQ.toLocaleString()}개 · {sumP}파렛트</span>
                      </div>

                      {/* 테이블 */}
                      <div style={{overflowX:"auto"}}>
                        <table style={{width:"100%",borderCollapse:"collapse",fontSize:12}}>
                          <thead>
                            <tr>
                              {["약호","신박스코드","카톤수량","파렛트당카톤","파렛트","FC센터"].map(h=>(
                                <th key={h} style={TH}>{h}</th>
                              ))}
                              <th style={{...TH,...IBLU}}>파렛트 번호 <span style={{fontSize:9,fontWeight:400,color:"var(--color-text-info)"}}>(파렛트 라벨 업로드시 자동입력)</span></th>
                            </tr>
                          </thead>
                          <tbody>
                            {g.rows.map((r, ri) => {
                              const m = master[String(r.sku)] || {} as MasterItem
                              const cpp = parseFloat(String(m.cpp)) || 16
                              const pallets = Math.ceil(r.quantity / cpp)
                              const meta = s2meta[String(r.sku)] || {}
                              const key = `${g.no}_${String(r.sku)}`
                              return (
                                <tr key={ri} style={{background:ri%2===0?"transparent":"var(--color-background-secondary)"}}>
                                  <td style={{...TD,fontWeight:500,minWidth:130}}>{String(r.sku)}</td>
                                  <td style={{...TD,color:"var(--color-text-info)",fontWeight:500,minWidth:70}}>{String(r.location||m.loc||"")}</td>
                                  <td style={{...TD,textAlign:"right",fontWeight:500,minWidth:60}}>{r.quantity.toLocaleString()}</td>
                                  <td style={{...TD,textAlign:"center",minWidth:60}}>{cpp}</td>
                                  <td style={{...TD,textAlign:"right",fontWeight:500,minWidth:50}}>{pallets}</td>
                                  <td style={{...TD,minWidth:60,fontWeight:500,color:meta.fc?"var(--color-text-primary)":"var(--color-text-tertiary)"}}>{meta.fc||"—"}</td>
                                  <td style={{...TD,...IBLU,minWidth:120}}>
                                    <span style={{fontSize:11,color:pltNotes[key]?"var(--color-text-info)":"var(--color-text-tertiary)",fontFamily:"var(--font-mono)"}}>
                                      {pltNotes[key]||`예) ${pallets}개`}
                                    </span>
                                  </td>
                                </tr>
                              )
                            })}
                          </tbody>
                          <tfoot>
                            <tr style={{background:CB[ci],borderTop:"1px solid "+CT[ci]+"44"}}>
                              <td colSpan={2} style={{...TD,color:CT[ci],fontWeight:500,textAlign:"right",fontSize:11}}>컨{g.no} 소계</td>
                              <td style={{...TD,color:CT[ci],fontWeight:500,textAlign:"right"}}>{sumQ.toLocaleString()}</td>
                              <td style={TD}></td>
                              <td style={{...TD,color:CT[ci],fontWeight:500,textAlign:"right"}}>{sumP}파렛트</td>
                              <td colSpan={2} style={TD}></td>
                            </tr>
                          </tfoot>
                        </table>
                      </div>
                    </div>
                  )
                })}

                {/* 전체 합계 */}
                <div style={{padding:"10px 16px",background:"var(--color-background-secondary)",borderRadius:"var(--border-radius-md)",border:"0.5px solid var(--color-border-tertiary)",display:"flex",gap:24,fontSize:12,flexWrap:"wrap"}}>
                  <span style={{fontWeight:500}}>전체 합계</span>
                  <span>총 {fd.reduce((s,r)=>s+r.quantity,0).toLocaleString()}개</span>
                  <span>총 {s2rows.reduce((s,r)=>s+r.pallets,0)}파렛트</span>
                  <span style={{color:"var(--color-text-tertiary)"}}>{groups.length}개 컨테이너</span>
                  <span style={{color:"var(--color-text-tertiary)"}}>
                    FC: {[...new Set(s2rows.map(r=>(s2meta[r.sku]||{}).fc).filter(Boolean))].join(', ')||"미입력"}
                  </span>
                </div>
              </div>
            )}
          </div>
        )
      })()}

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

      {/* ══ 3차 가공 (역산시트) ══ */}
      {mode==='s3b'&&(()=>{
        // 부킹여부 컬럼으로 BL 구분
        // 같은 booking_ref 값 → 같은 BL
        // BL별로 약호+수량 합산 → 역산시트 행 생성

        // BL 그룹핑
        type BLRow = { no: number; blId: string; sku: string; asin: string; loc: string; qty: number; gw: number; price: number; fba: number }
        const blGroups: Record<string, { rows: BLRow[]; blId: string }> = {}
        let rowNo = 0

        // fd에서 BL 기준으로 집계
        const blAgg: Record<string, Record<string, { qty: number; blId: string; gw: number; pallets: number }>> = {}
        for (const r of fd) {
          const sku = String(r.sku || '')
          if (!sku) continue
          const blId = String(r.booking_ref || r.destination || r.container_no || 'BL1')
          if (!blAgg[blId]) blAgg[blId] = {}
          if (!blAgg[blId][sku]) blAgg[blId][sku] = { qty: 0, blId, gw: 0, pallets: 0 }
          const m = master[sku] || {} as MasterItem
          const cpp = parseFloat(String(m.cpp)) || 16
          blAgg[blId][sku].qty += r.quantity
          // 행별 calcGW 누적 (합산 후 calcGW 금지)
          blAgg[blId][sku].gw += calcGW(r.quantity, parseFloat(String(m.kg)) || 0)
          blAgg[blId][sku].pallets += Math.ceil(r.quantity / cpp)
        }

        // BL별 행 생성
        const allBLRows: BLRow[] = []
        let blNo = 1
        for (const [blId, skuMap] of Object.entries(blAgg)) {
          for (const [sku, { qty, gw, pallets }] of Object.entries(skuMap)) {
            const m = master[sku] || {} as MasterItem
            rowNo++
            allBLRows.push({
              no: rowNo,
              blId,
              sku,
              asin: m.asin || '',
              loc: m.loc || '',
              qty,
              gw: Math.round(gw * 10) / 10,
              price: parseFloat(String(m.price)) || 0,
              fba: parseFloat(String(m.fba)) || 0,
            })
          }
          blNo++
        }

        // BL 목록
        const blIds = [...new Set(fd.map(r => String(r.booking_ref || r.destination || r.container_no || 'BL1')))]

        function expS3B() {
          const hdr = [['NO','상품명(ASIN号)','ASIN No.','약호','신박스코드','UNIT(PCS)','G.W(KG)','판매가격(JPY)','FBA배송료(JPY)','BL']]
          const body = allBLRows.map(r => [r.no, `${r.sku} ${r.asin}`, r.asin, r.sku, r.loc, r.qty, r.gw, r.price ? `¥${r.price.toLocaleString()}` : '', r.fba ? `¥${r.fba.toLocaleString()}` : '', r.blId])
          xlsDl([...hdr, ...body], '역산시트', '역산시트_' + new Date().toISOString().slice(0,10) + '.xlsx')
        }

        // BL별 소계
        const blTotals: Record<string, {qty:number, gw:number}> = {}
        for (const r of allBLRows) {
          if (!blTotals[r.blId]) blTotals[r.blId] = {qty:0, gw:0}
          blTotals[r.blId].qty += r.qty
          blTotals[r.blId].gw += r.gw
        }

        return (
          <div>
            <div style={{display:"flex",alignItems:"center",gap:8,marginBottom:14,flexWrap:"wrap"}}>
              <span style={{fontSize:12,color:"var(--color-text-secondary)"}}>BL 기준 역산시트 · 부킹여부 컬럼으로 BL 구분</span>
              {blIds.length > 0 && blIds.map((bid,i) => (
                <span key={bid} style={{fontSize:11,padding:"2px 8px",borderRadius:10,background:CB[i%CB.length],color:CT[i%CT.length],border:"0.5px solid var(--color-border-tertiary)"}}>BL{i+1}: {bid}</span>
              ))}
              <button onClick={expS3B} style={{marginLeft:"auto",fontSize:11,padding:"3px 10px"}}>xlsx 저장</button>
            </div>
            <SheetTabs/>

            {!file ? (
              <p style={{color:"var(--color-text-tertiary)",fontSize:13,textAlign:"center",padding:"3rem 0"}}>파일을 먼저 업로드해주세요</p>
            ) : allBLRows.length === 0 ? (
              <div style={{textAlign:"center",padding:"3rem 0",color:"var(--color-text-tertiary)"}}>
                <p>부킹여부 컬럼이 없거나 데이터가 없습니다.</p>
                <p style={{fontSize:11,marginTop:8}}>업로드한 엑셀에 "부킹여부" 컬럼이 있어야 BL별로 구분됩니다.</p>
              </div>
            ) : (
              <div style={{overflowX:"auto",border:"0.5px solid var(--color-border-tertiary)",borderRadius:"var(--border-radius-md)"}}>
                <table style={{width:"100%",borderCollapse:"collapse",fontSize:12,minWidth:900}}>
                  <thead>
                    <tr>
                      <th style={{...TH,width:40}}>NO</th>
                      <th style={{...TH,minWidth:200}}>Goods of Description<br/><span style={{fontWeight:400,fontSize:10,opacity:0.7}}>(일문상품명 ASIN号)</span></th>
                      <th style={{...TH,minWidth:100}}>ASIN No.</th>
                      <th style={{...TH,minWidth:130}}>약호</th>
                      <th style={{...TH,minWidth:70}}>신박스코드</th>
                      <th style={{...TH,textAlign:"right",minWidth:70}}>UNIT<br/><span style={{fontWeight:400,fontSize:10}}>(PCS)</span></th>
                      <th style={{...TH,textAlign:"right",minWidth:90}}>G.W<br/><span style={{fontWeight:400,fontSize:10}}>(KG)</span></th>
                      <th style={{...TH,textAlign:"right",minWidth:100}}>★販売単価<br/><span style={{fontWeight:400,fontSize:10}}>판매가격 (JPY)</span></th>
                      <th style={{...TH,textAlign:"right",minWidth:110}}>②配送料 FBA<br/><span style={{fontWeight:400,fontSize:10}}>배송료 (JPY)</span></th>
                      <th style={{...TH,minWidth:80}}>BL</th>
                    </tr>
                  </thead>
                  <tbody>
                    {(() => {
                      let lastBl = ''
                      return allBLRows.map((r, i) => {
                        const isNewBl = r.blId !== lastBl
                        lastBl = r.blId
                        const blIdx = blIds.indexOf(r.blId)
                        const ci = blIdx % CB.length
                        return (
                          <>
                            {isNewBl && (
                              <tr key={`blheader_${r.blId}`} style={{background:CB[ci]}}>
                                <td colSpan={10} style={{...TD,color:CT[ci],fontWeight:700,fontSize:12,padding:"6px 12px"}}>
                                  BL{blIdx+1} — {r.blId}
                                </td>
                              </tr>
                            )}
                            <tr key={i} style={{background:i%2===0?"transparent":"var(--color-background-secondary)"}}>
                              <td style={{...TD,textAlign:"center",color:"var(--color-text-tertiary)"}}>{r.no}</td>
                              <td style={{...TD,fontSize:11}}>
                                <div style={{fontWeight:500}}>{r.sku}</div>
                                <div style={{fontSize:10,color:"var(--color-text-tertiary)",fontFamily:"var(--font-mono)"}}>{r.asin}</div>
                              </td>
                              <td style={{...TD,fontFamily:"var(--font-mono)",fontSize:11,color:"var(--color-text-secondary)"}}>{r.asin}</td>
                              <td style={{...TD,fontWeight:500}}>{r.sku}</td>
                              <td style={{...TD,color:"var(--color-text-info)",fontWeight:500}}>{r.loc}</td>
                              <td style={{...TD,textAlign:"right",fontWeight:500}}>{r.qty.toLocaleString()}</td>
                              <td style={{...TD,textAlign:"right"}}>{r.gw.toFixed(1)}</td>
                              <td style={{...TD,textAlign:"right",color:r.price>0?"var(--color-text-primary)":"var(--color-text-tertiary)"}}>
                                {r.price > 0 ? `¥${r.price.toLocaleString()}` : '—'}
                              </td>
                              <td style={{...TD,textAlign:"right",color:r.fba>0?"var(--color-text-primary)":"var(--color-text-tertiary)"}}>
                                {r.fba > 0 ? `¥${r.fba.toLocaleString()}` : '—'}
                              </td>
                              <td style={{...TD,fontSize:11,color:"var(--color-text-secondary)"}}>{r.blId}</td>
                            </tr>
                          </>
                        )
                      })
                    })()}
                    {/* BL별 소계 */}
                    {blIds.map((blId, bi) => {
                      const tot = blTotals[blId]
                      if (!tot) return null
                      const ci = bi % CB.length
                      return (
                        <tr key={`subtotal_${blId}`} style={{background:CB[ci],borderTop:"1.5px solid var(--color-border-secondary)"}}>
                          <td colSpan={5} style={{...TD,color:CT[ci],fontWeight:600,textAlign:"right"}}>BL{bi+1} 소계</td>
                          <td style={{...TD,color:CT[ci],fontWeight:700,textAlign:"right"}}>{tot.qty.toLocaleString()}</td>
                          <td style={{...TD,color:CT[ci],fontWeight:700,textAlign:"right"}}>{tot.gw.toFixed(1)}</td>
                          <td colSpan={3} style={TD}></td>
                        </tr>
                      )
                    })}
                  </tbody>
                  <tfoot>
                    <tr style={{background:"var(--color-background-secondary)",borderTop:"2px solid var(--color-border-secondary)"}}>
                      <td colSpan={5} style={{...TD,fontWeight:500,textAlign:"right"}}>전체 합계</td>
                      <td style={{...TD,fontWeight:500,textAlign:"right"}}>{allBLRows.reduce((s,r)=>s+r.qty,0).toLocaleString()}</td>
                      <td style={{...TD,fontWeight:500,textAlign:"right"}}>{allBLRows.reduce((s,r)=>s+r.gw,0).toFixed(1)}</td>
                      <td colSpan={3} style={TD}></td>
                    </tr>
                  </tfoot>
                </table>
              </div>
            )}

            {/* ── 비용 계산 (輸送費・通関諸費用) ── */}
            {file && s2rows.length > 0 && (()=>{
              // FC별 파렛트·주소 집계 (s2rows 직접 사용 → NaN 방지)
              const fcPallets: Record<string,number> = {}
              const fcAddr: Record<string,string> = {}
              for (const r of s2rows) {
                const fc = (s2meta[r.sku]||{}).fc?.trim()||''
                if (!fc) continue
                fcPallets[fc] = (fcPallets[fc]||0) + r.pallets
                if (!fcAddr[fc]) fcAddr[fc] = (s2meta[r.sku]||{}).address||''
              }

              // BL별 FC 파렛트 집계 (fd + booking_ref 기반)
              const blFcPallets: Record<string,Record<string,number>> = {}
              for (const r of fd) {
                const sku = String(r.sku||''); if(!sku) continue
                const blId = String(r.booking_ref||r.destination||r.container_no||'BL1')
                const fc = (s2meta[sku]||{}).fc?.trim()||''
                if (!fc) continue
                const cpp = parseFloat(String((master[sku]||{}).cpp))||16
                const plts = Math.ceil(r.quantity/cpp)
                if (!blFcPallets[blId]) blFcPallets[blId] = {}
                blFcPallets[blId][fc] = (blFcPallets[blId][fc]||0) + plts
              }

              // BL별 컨테이너 수
              const blCtnCount: Record<string,number> = {}
              for (const blId of blIds) {
                const nos = new Set(fd.filter(r=>String(r.booking_ref||r.destination||r.container_no||'BL1')===blId).map(r=>r.container_no))
                blCtnCount[blId] = Math.max(nos.size,1)
              }

              // 거리 분류: 관서(오사카 출발) vs 관동(사이타마 출발) 기준이 다름
              // 관서 기준: 도쿄/요코하마=장거리, 나고야=중거리, 고베=단거리
              // 관동 기준: 도쿄/요코하마=단거리, 나고야=중장거리, 고베=장거리
              function getDistanceClass(fc: string, address: string, port: 'kansai'|'kanto'): 'long'|'mid'|'short' {
                const a = (address||'').toUpperCase()+fc.toUpperCase()
                const isTokyoArea = /東京|TOKYO|神奈川|KANAGAWA|YOKOHAMA|横浜|埼玉|SAITAMA|千葉|CHIBA/.test(a) || ['HIY1','TPB5'].includes(fc)
                const isNagoyaArea = /愛知|AICHI|NAGOYA|名古屋|三重|岐阜/.test(a) || ['TPB8'].includes(fc)
                const isOsakaArea = /兵庫|HYOGO|神戸|KOBE|大阪|OSAKA|京都/.test(a) || ['VJNB'].includes(fc)
                if (port === 'kansai') {
                  if (isTokyoArea) return 'long'
                  if (isNagoyaArea) return 'mid'
                  if (isOsakaArea) return 'short'
                } else { // kanto (사이타마 출발)
                  if (isTokyoArea) return 'short'   // 사이타마→요코하마/도쿄 단거리
                  if (isNagoyaArea) return 'mid'    // 사이타마→나고야 중장거리
                  if (isOsakaArea) return 'long'    // 사이타마→고베 장거리
                }
                return 'long'
              }

              function calcTrucks(pallets: number, dc: 'long'|'mid'|'short') {
                const t10 = Math.floor(pallets/14)
                const rem = pallets - t10*14
                let t4=0, t2=0
                if (rem>0) {
                  if (dc==='short' && rem<=2) t2=1
                  else if (rem<=6) t4=1
                  else return {t10:t10+1,t4:0,t2:0}
                }
                return {t10,t4,t2}
              }

              function guessPort(blId: string): 'kansai'|'kanto' {
                for (const fc of Object.keys(blFcPallets[blId]||{})) {
                  if (/東京|TOKYO|神奈川|KANAGAWA|YOKOHAMA|横浜|埼玉|SAITAMA/.test((fcAddr[fc]||'').toUpperCase())) return 'kanto'
                  if (['HIY1','TPB5'].includes(fc)) return 'kanto'
                }
                return 'kansai'
              }

              const DRAY_KANSAI=62000, DRAY_KANTO=62000, TAX=1.1, INOUT=500
              // 배송 단가 — 관서/관동 공통 (출발지만 다름, 단가는 요청 기준 동일 가정)
              // 실제 관동 단가는 CJ 견적서 확인 후 수정
              const RATES: Record<string,{t10:number,t4:number,t2:number}> = {
                HIY1:{t10:55000,t4:35000,t2:25000},
                TPB5:{t10:55000,t4:45000,t2:25000},
                VJNB:{t10:55000,t4:45000,t2:25000},
                TPB8:{t10:55000,t4:45000,t2:25000},
              }

              type CR = {item:string;ko:string;qty:number;unit:string;up:number;sub:number;taxed:boolean;total:number;note?:string}

              const blSections: {blId:string;blNo:number;port:'kansai'|'kanto';rows:CR[];sub:number}[] = []

              blIds.forEach((blId,bi)=>{
                const port=(blPortMode[blId] as 'kansai'|'kanto')||guessPort(blId)
                const ctnCnt=blCtnCount[blId]||1
                const dray=port==='kansai'?DRAY_KANSAI:DRAY_KANTO
                const portLabel=port==='kansai'?'관서 (오사카 남항)':'관동 (도쿄항)'
                const fcPlt=blFcPallets[blId]||{}
                const blPlts=Object.values(fcPlt).reduce((s,v)=>s+(v||0),0)

                const rows: CR[] = [
                  {item:'通関申告料（1申告/2HS）',ko:'통관신고료 (1신고/2HS코드)',qty:1,unit:'件',up:11800,sub:11800,taxed:false,total:11800},
                  {item:'輸入取扱料',ko:'수입취급료',qty:1,unit:'件',up:10000,sub:10000,taxed:true,total:Math.round(10000*TAX)},
                  {item:'評価申告料（1申告/2HS）',ko:'평가신고료 (1신고/2HS코드)',qty:1,unit:'件',up:11800,sub:11800,taxed:false,total:11800},
                  {item:'AN費用',ko:'AN비용',qty:1,unit:'件',up:0,sub:0,taxed:false,total:0,note:'견적'},
                  dray>0
                    ?{item:`40Fドレー+デバン費用 (${portLabel} ${ctnCnt}본)`,ko:`40ft 드레이지+디베닝 (${portLabel})`,qty:ctnCnt,unit:'本',up:dray,sub:ctnCnt*dray,taxed:true,total:Math.round(ctnCnt*dray*TAX)}
                    :{item:`40Fドレー+デバン費用 (${portLabel} ${ctnCnt}본)`,ko:`40ft 드레이지+디베닝 (${portLabel}) — 견적요청`,qty:ctnCnt,unit:'本',up:0,sub:0,taxed:true,total:0,note:'견적'},
                ]

                for (const [fc,plts] of Object.entries(fcPlt)) {
                  if (!plts) continue
                  const dc=getDistanceClass(fc,fcAddr[fc]||'',port)
                  const tr=calcTrucks(plts,dc)
                  const r=RATES[fc]||{t10:55000,t4:45000,t2:25000}
                  if(tr.t10>0){const s=tr.t10*r.t10;rows.push({item:`配送料金 (${fc}) 10T × ${tr.t10}台`,ko:`배송료 (${fc}) 10톤 × ${tr.t10}대`,qty:tr.t10,unit:'本',up:r.t10,sub:s,taxed:true,total:Math.round(s*TAX)})}
                  if(tr.t4>0){const s=tr.t4*r.t4;rows.push({item:`配送料金 (${fc}) 4T × ${tr.t4}台`,ko:`배송료 (${fc}) 4톤 × ${tr.t4}대`,qty:tr.t4,unit:'本',up:r.t4,sub:s,taxed:true,total:Math.round(s*TAX)})}
                  if(tr.t2>0){const s=tr.t2*r.t2;rows.push({item:`配送料金 (${fc}) 2T × ${tr.t2}台`,ko:`배송료 (${fc}) 2톤 × ${tr.t2}대`,qty:tr.t2,unit:'本',up:r.t2,sub:s,taxed:true,total:Math.round(s*TAX)})}
                }

                if(blPlts>0){
                  rows.push({item:`入庫料 (${blPlts}PLT)`,ko:`입고료 (${blPlts}파렛트)`,qty:blPlts,unit:'PLT',up:INOUT,sub:blPlts*INOUT,taxed:true,total:Math.round(blPlts*INOUT*TAX)})
                  rows.push({item:`出荷料 (${blPlts}PLT)`,ko:`출고료 (${blPlts}파렛트)`,qty:blPlts,unit:'PLT',up:INOUT,sub:blPlts*INOUT,taxed:true,total:Math.round(blPlts*INOUT*TAX)})
                }

                const sub=rows.reduce((s,r)=>s+(isNaN(r.total)?0:r.total),0)
                blSections.push({blId,blNo:bi+1,port,rows,sub})
              })

              const grand=blSections.reduce((s,b)=>s+(isNaN(b.sub)?0:b.sub),0)

              return (
                <div style={{marginTop:20}}>
                  <div style={{display:"flex",alignItems:"center",gap:8,marginBottom:10}}>
                    <span style={{fontWeight:700,fontSize:13}}>輸送費・輸入通関諸費用 / 운송비·수입통관비용 (BL별)</span>
                    <span style={{fontSize:11,color:"var(--color-text-tertiary)"}}>※ 입항지 자동 감지 · 수동 변경 가능</span>
                  </div>
                  {blSections.map((bl,si)=>{
                    const ci=si%CB.length
                    return (
                      <div key={bl.blId} style={{border:"0.5px solid var(--color-border-tertiary)",borderRadius:"var(--border-radius-md)",overflow:"hidden",marginBottom:12}}>
                        <div style={{background:"#1e293b",color:"#fff",padding:"8px 14px",display:"flex",alignItems:"center",gap:10,flexWrap:"wrap"}}>
                          <span style={{fontWeight:700,fontSize:13}}>BL{bl.blNo} — {bl.blId}</span>
                          <div style={{display:"flex",alignItems:"center",gap:4,background:"rgba(255,255,255,0.12)",borderRadius:6,padding:"2px 4px"}}>
                            <span style={{fontSize:11,opacity:0.8}}>입항지:</span>
                            {(['kansai','kanto'] as const).map(p=>(
                              <button key={p} onClick={()=>setBlPortMode(prev=>({...prev,[bl.blId]:p}))} style={{fontSize:11,padding:"2px 10px",borderRadius:4,border:"none",cursor:"pointer",fontWeight:bl.port===p?700:400,background:bl.port===p?(p==='kansai'?'#3b82f6':'#10b981'):'rgba(255,255,255,0.2)',color:"#fff"}}>
                                {p==='kansai'?'관서 (오사카)':'관동 (도쿄)'}
                              </button>
                            ))}
                          </div>
                          {Object.entries(blFcPallets[bl.blId]||{}).map(([fc,plt])=>{
                            if(!plt) return null
                            const tr=calcTrucks(plt,getDistanceClass(fc,fcAddr[fc]||'',bl.port))
                            return <span key={fc} style={{fontSize:10,padding:"1px 7px",borderRadius:8,background:"rgba(255,255,255,0.15)"}}>{fc}: {plt}PLT→{tr.t10>0?`10T×${tr.t10}`:''}{tr.t4>0?` 4T×${tr.t4}`:''}{tr.t2>0?` 2T×${tr.t2}`:''}</span>
                          })}
                          <span style={{marginLeft:"auto",fontWeight:700,color:"#fbbf24"}}>¥{bl.sub.toLocaleString()}</span>
                        </div>
                        <table style={{width:"100%",borderCollapse:"collapse",fontSize:12}}>
                          <thead><tr>
                            <th style={{...TH,minWidth:200}}>項目 / 항목</th>
                            <th style={{...TH,textAlign:"right",width:50}}>건</th>
                            <th style={{...TH,width:50}}>단위</th>
                            <th style={{...TH,textAlign:"right",width:90}}>단가</th>
                            <th style={{...TH,textAlign:"right",width:100}}>금액</th>
                            <th style={{...TH,width:75}}>세금</th>
                            <th style={{...TH,textAlign:"right",width:110}}>TOTAL</th>
                          </tr></thead>
                          <tbody>{bl.rows.map((r,i)=>(
                            <tr key={i} style={{background:i%2===0?"transparent":"var(--color-background-secondary)"}}>
                              <td style={TD}>
                                <div style={{fontWeight:500,fontSize:12}}>{r.item}</div>
                                <div style={{fontSize:10,color:"var(--color-text-tertiary)"}}>{r.ko}</div>
                              </td>
                              <td style={{...TD,textAlign:"right"}}>{r.qty}</td>
                              <td style={TD}>{r.unit}</td>
                              <td style={{...TD,textAlign:"right"}}>{r.up>0?`¥${r.up.toLocaleString()}`:'—'}</td>
                              <td style={{...TD,textAlign:"right"}}>{r.sub>0?`¥${r.sub.toLocaleString()}`:'—'}</td>
                              <td style={{...TD,fontSize:10,color:r.taxed?"#ef4444":"var(--color-text-tertiary)"}}>{r.taxed?'課税 10%':'免税'}</td>
                              <td style={{...TD,textAlign:"right",fontWeight:500}}>
                                {r.note==='견적'?<span style={{color:"var(--color-text-tertiary)"}}>견적</span>:r.total>0?`¥${r.total.toLocaleString()}`:'—'}
                              </td>
                            </tr>
                          ))}</tbody>
                          <tfoot><tr style={{background:CB[ci]}}>
                            <td colSpan={5} style={{...TD,color:CT[ci],fontWeight:600,textAlign:"right"}}>BL{bl.blNo} 소계</td>
                            <td style={TD}></td>
                            <td style={{...TD,color:CT[ci],fontWeight:700,fontSize:13,textAlign:"right"}}>¥{bl.sub.toLocaleString()}</td>
                          </tr></tfoot>
                        </table>
                      </div>
                    )
                  })}
                  <div style={{background:"#1e293b",color:"#fff",padding:"10px 16px",borderRadius:"var(--border-radius-md)",display:"flex",justifyContent:"space-between",alignItems:"center"}}>
                    <span style={{fontWeight:600}}>전체 합계 ({blIds.length}BL)</span>
                    <span style={{fontWeight:800,fontSize:16,color:"#fbbf24"}}>¥{grand.toLocaleString()}</span>
                  </div>
                  <div style={{marginTop:6,fontSize:10,color:"var(--color-text-tertiary)"}}>
                    ※ 입출고료 각 ¥500/PLT · 드레이지: 관서 ¥62,000/본, 관동 ¥62,000/본 (미확정) · 트럭: 관서기준 도쿄=장거리/고베=단거리, 관동기준 도쿄=단거리/고베=장거리 · 실제 단가는 견적서 확인 후 수정
                  </div>
                </div>
              )
            })()}
          </div>
        )
      })()}

      {/* ══ 1.5. CJ 출고 ══ */}
      {mode==='cj'&&(()=>{
        const SHEET_ID = '1FanL4UnHfRoUj-d-8C1efLdywqB_qmsx-9eEeYuvNNs'
        const SHEET_URL = `https://docs.google.com/spreadsheets/d/${SHEET_ID}/edit`

        type CJRow = {
          ctnNo: number; etd: string; eta: string; pod: string; blNo: string;
          cntrNo: string; teu: number
          products: string; boxCodes: string
          plt: number; ctn: number; fc: string
        }

        const cjRows: CJRow[] = ctnNums.map(no => {
          const rows = fd.filter(r=>r.container_no===no)
          const f0 = rows[0] || {}
          const meta = ctnMeta[no] || {}

          const dest = String(f0.destination||''  ).toLowerCase()
          let pod = ''
          if (dest.includes('오사카')||dest.includes('osaka')) pod = 'JPOSA'
          else if (dest.includes('도쿄')||dest.includes('tokyo')) pod = 'JPTYO'
          else if (dest.includes('함부르크')||dest.includes('hamburg')) pod = 'DEHAM'
          else pod = String(f0.destination||''  )

          const ftVal = String(f0.ft||meta.container||''  )
          const teu = ftVal.includes('40') ? 2 : ftVal.includes('20') ? 1 : 2

          const skuMap: Record<string,{plt:number;ctn:number;loc:string;fc:string}> = {}
          for (const r of rows) {
            const sku = String(r.sku||''  ); if(!sku) continue
            const m = master[sku] || {} as MasterItem
            const cpp = parseFloat(String(m.cpp))||16
            const loc = String(r.location||m.loc||''  )
            const fc = (s2meta[sku]||{}).fc||''
            if (!skuMap[sku]) skuMap[sku]={plt:0,ctn:0,loc,fc}
            skuMap[sku].plt += Math.ceil(r.quantity/cpp)
            skuMap[sku].ctn += r.quantity
          }

          const skuList = Object.entries(skuMap)
          const fcSet = new Set(skuList.map(([,v])=>v.fc).filter(Boolean))

          return {
            ctnNo: no,
            etd: String(f0.etd||''  ),
            eta: String(f0.eta||''  ),
            pod,
            blNo: String(f0.booking_ref||''  ),
            cntrNo: String(meta.container||''  ),
            teu,
            products: skuList.map(([sku])=>sku).join('\n'),
            boxCodes: skuList.map(([,v])=>v.loc).join('\n'),
            plt: skuList.reduce((s,[,v])=>s+v.plt,0),
            ctn: skuList.reduce((s,[,v])=>s+v.ctn,0),
            fc: [...fcSet].join(', '),
          }
        })

        function tsvCell(val: string) {
          if (val.includes('\n')||val.includes('"')||val.includes('\t')) return '"' + val.replace(/"/g, '\"\\"\"') + '"' 
          return val
        }

        function copyForSheets() {
          const tsv = cjRows.map(r =>
            [r.etd, r.eta, r.pod, r.blNo, r.cntrNo, r.teu,
              tsvCell(r.products), tsvCell(r.boxCodes),
              r.plt, r.ctn, r.fc
            ].join('\t')
          ).join('\n')
          navigator.clipboard.writeText(tsv).then(()=>alert('클립보드에 복사됐습니다!\n구글 시트의 ETD 첫 번째 빈 셀에서 붙여넣기(Ctrl+V) 하세요.'))
        }

        return (
          <div>
            <div style={{display:"flex",alignItems:"center",gap:8,marginBottom:14,flexWrap:"wrap"}}>
              <span style={{fontSize:12,color:"var(--color-text-secondary)"}}>컨테이너별 출고 데이터 → SPEED RACK_CJLJ埼玉センター 시트</span>
              <a href={SHEET_URL} target="_blank" rel="noreferrer" style={{fontSize:11,padding:"3px 10px",borderRadius:"var(--border-radius-md)",border:"0.5px solid var(--color-border-info)",color:"var(--color-text-info)",textDecoration:"none",background:"rgba(219,234,254,0.2)"}}>📊 시트 열기</a>
              <button onClick={copyForSheets} style={{fontSize:11,padding:"3px 10px",borderRadius:"var(--border-radius-md)",border:"0.5px solid var(--color-border-success)",color:"var(--color-text-success)",background:"rgba(209,250,229,0.2)",cursor:"pointer"}}>📋 TSV 복사 (시트 붙여넣기용)</button>
            </div>
            <SheetTabs/>
            {!file ? (
              <p style={{textAlign:"center",padding:"3rem 0",color:"var(--color-text-tertiary)"}}>파일을 먼저 업로드해주세요</p>
            ) : (
              <>
                <div style={{overflowX:"auto",border:"0.5px solid var(--color-border-tertiary)",borderRadius:"var(--border-radius-md)",marginBottom:12}}>
                  <table style={{width:"100%",borderCollapse:"collapse",fontSize:12}}>
                    <thead><tr>
                      {['ETD','ETA','POD','B/L No.','CNTR No.','TEU','Product code','BOX CODE','PLT','CTN','FC'].map(h=>(
                        <th key={h} style={{...TH,textAlign:['PLT','CTN','TEU'].includes(h)?'right':'left'}}>{h}</th>
                      ))}
                    </tr></thead>
                    <tbody>{cjRows.map((r,i)=>{
                      const ci=(r.ctnNo-1)%CB.length
                      return(
                        <tr key={i} style={{background:i%2===0?"transparent":"var(--color-background-secondary)"}}>
                          <td style={{...TD,minWidth:80,borderLeft:`3px solid ${CB[ci]}`}}>
                            <div style={{fontSize:10,color:"var(--color-text-tertiary)",marginBottom:2}}>컨{r.ctnNo}</div>
                            <input defaultValue={r.etd} style={{...INP,width:"100%"}} placeholder="05월 14일"/>
                          </td>
                          <td style={{...TD,minWidth:80}}><input defaultValue={r.eta} style={{...INP,width:"100%"}} placeholder="05월 20일"/></td>
                          <td style={{...TD,minWidth:60}}><input defaultValue={r.pod} style={{...INP,width:"100%"}} placeholder="JPOSA"/></td>
                          <td style={{...TD,minWidth:120}}><input defaultValue={r.blNo} style={{...INP,width:"100%"}} placeholder="B/L No."/></td>
                          <td style={{...TD,minWidth:110}}><input defaultValue={r.cntrNo} style={{...INP,width:"100%"}} placeholder="ONEU6916846"/></td>
                          <td style={{...TD,textAlign:"right"}}>{r.teu}</td>
                          <td style={{...TD,minWidth:130,fontSize:11,whiteSpace:"pre-line",lineHeight:"1.6"}}>{r.products}</td>
                          <td style={{...TD,minWidth:80,fontSize:11,color:"var(--color-text-info)",fontWeight:500,whiteSpace:"pre-line",lineHeight:"1.6"}}>{r.boxCodes}</td>
                          <td style={{...TD,textAlign:"right",fontWeight:500}}>{r.plt}</td>
                          <td style={{...TD,textAlign:"right"}}>{r.ctn.toLocaleString()}</td>
                          <td style={{...TD,color:"var(--color-text-info)",fontWeight:500}}>{r.fc}</td>
                        </tr>
                      )
                    })}</tbody>
                  </table>
                </div>
                <div style={{padding:"8px 12px",background:"rgba(219,234,254,0.2)",borderRadius:"var(--border-radius-md)",fontSize:11,color:"var(--color-text-secondary)",border:"0.5px solid var(--color-border-info)"}}>
                  💡 <strong>사용 방법:</strong> ETD/ETA/CNTR No./B/L No. 직접 입력 → 📋 TSV 복사 → 구글 시트 ETD 첫 빈 셀 클릭 → Ctrl+V
                  &nbsp;·&nbsp; Product code/BOX CODE는 셀 안에 줄바꿈으로 입력됩니다
                </div>
              </>
            )}
          </div>
        )
      })()}

            {/* ══ 매핑 관리 ══ */}
      {mode==='map'&&(
        <div>
          <div style={{display:"flex",alignItems:"center",gap:8,marginBottom:14}}>
            <span style={{fontSize:13,fontWeight:500}}>SKU 마스터</span><span style={{fontSize:12,color:"var(--color-text-tertiary)"}}>{Object.keys(master).length}개</span>
            <button onClick={()=>{if(!newSku.sku)return;setMaster(p=>({...p,[newSku.sku]:{sku:newSku.realSku||newSku.sku,asin:newSku.asin,to:newSku.to,loc:newSku.loc,cpp:+newSku.cpp,fba:+newSku.fba,price:+newSku.price,kg:+newSku.kg,bx:+newSku.bx,by:+newSku.by,bz:+newSku.bz}}));setNewSku({...EMPTY_SKU})}} style={{marginLeft:"auto",fontSize:11,padding:"3px 10px"}}>+ 추가</button>
            <button onClick={()=>{if(confirm("초기화하시겠습니까?"))setMaster(Object.fromEntries(Object.entries(MASTER_INIT).map(([k,v])=>[k,{...v}])))}} style={{fontSize:11,padding:"3px 10px",color:"var(--color-text-danger)",border:"0.5px solid var(--color-border-danger)",borderRadius:"var(--border-radius-md)",background:"transparent",cursor:"pointer"}}>초기화</button>
            <button onClick={()=>xlsDl(Object.entries(master).map(([sku,m])=>({약호:sku,SKU:m.sku,ASIN:m.asin,구박스:m.to,신박스:m.loc,PLT당카톤:m.cpp,FBA비용:m.fba,판매가:m.price,무게kg:m.kg,박스가로:m.bx,박스세로:m.by,박스높이:m.bz})),"마스터","SKU마스터.xlsx")} style={{fontSize:11,padding:"3px 10px"}}>xlsx 저장</button>
          </div>
          {/* JSON 백업/복원 */}
          <div style={{marginBottom:8,padding:"8px 12px",background:"var(--color-background-secondary)",borderRadius:"var(--border-radius-md)",fontSize:11,color:"var(--color-text-secondary)",border:"0.5px solid var(--color-border-tertiary)",display:"flex",alignItems:"center",gap:8,flexWrap:"wrap"}}>
            <span>💾 재배포 후 데이터 초기화 방지 — JSON으로 백업 후 복원하세요</span>
            <button onClick={()=>{
              const data={master,s2meta,ctnMeta,_v:1,_date:new Date().toISOString()}
              const blob=new Blob([JSON.stringify(data,null,2)],{type:"application/json"})
              const a=document.createElement("a");a.href=URL.createObjectURL(blob)
              a.download="shipment_backup_"+new Date().toISOString().slice(0,10)+".json";a.click()
            }} style={{fontSize:11,padding:"3px 10px",cursor:"pointer",border:"0.5px solid var(--color-border-secondary)",borderRadius:"var(--border-radius-md)",background:"var(--color-background-primary)",color:"var(--color-text-primary)",fontWeight:500}}>⬇ JSON 백업</button>
            <label style={{fontSize:11,padding:"3px 10px",cursor:"pointer",border:"0.5px solid var(--color-border-info)",borderRadius:"var(--border-radius-md)",background:"rgba(219,234,254,0.3)",color:"var(--color-text-info)",fontWeight:500}}>
              ⬆ JSON 복원
              <input type="file" accept=".json" style={{display:"none"}} onChange={e=>{
                const f=e.target.files?.[0];if(!f)return
                const rd=new FileReader();rd.onload=ev=>{
                  try{
                    const d=JSON.parse(ev.target?.result as string)
                    if(d.master)setMaster(d.master)
                    if(d.s2meta)setS2meta(d.s2meta)
                    if(d.ctnMeta)setCtnMeta(d.ctnMeta)
                    alert("복원 완료!")
                  }catch{alert("JSON 파일이 올바르지 않습니다")}
                };rd.readAsText(f)
                e.target.value=""
              }}/>
            </label>
          </div>
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

      {/* ══ FBA 자동 등록 ══ */}
      {mode==='fba-api'&&(
        <div>
          <div style={{display:"flex",alignItems:"center",gap:8,marginBottom:14,flexWrap:"wrap"}}>
            <span style={{fontSize:12,color:"var(--color-text-secondary)"}}>SP-API → FBA Inbound v2024-03-20 자동 등록</span>
            {fbaPlanId&&<span style={{fontSize:11,padding:"2px 8px",borderRadius:10,background:"rgba(219,234,254,0.4)",border:"0.5px solid var(--color-border-info)",color:"var(--color-text-info)"}}>Plan: {fbaPlanId}</span>}
          </div>
          <SheetTabs/>

          {/* 환경변수 안내 */}
          {!file&&(
            <div style={{textAlign:"center",padding:"4rem 0",color:"var(--color-text-tertiary)"}}>
              <div style={{fontSize:36,marginBottom:12}}>🚀</div>
              <p style={{fontSize:14,fontWeight:500,color:"var(--color-text-secondary)"}}>SP-API FBA Inbound 자동 등록</p>
              <p style={{fontSize:12,marginTop:8}}>파일을 먼저 업로드하고 1차 가공에서 SKU 데이터를 확인하세요</p>
              <div style={{marginTop:16,padding:"12px 20px",background:"var(--color-background-secondary)",borderRadius:"var(--border-radius-md)",display:"inline-block",textAlign:"left",fontSize:11,color:"var(--color-text-secondary)",border:"0.5px solid var(--color-border-tertiary)"}}>
                <div style={{fontWeight:500,marginBottom:6}}>필요 환경변수 (.env.local)</div>
                {["SP_API_CLIENT_ID","SP_API_CLIENT_SECRET","SP_API_REFRESH_TOKEN","SP_API_AWS_ACCESS_KEY","SP_API_AWS_SECRET_KEY","SP_API_ROLE_ARN","SP_API_SOURCE_NAME","SP_API_SOURCE_ADDRESS1","SP_API_SOURCE_CITY","SP_API_SOURCE_COUNTRY","SP_API_SOURCE_POSTAL"].map(k=>(
                  <div key={k} style={{fontFamily:"var(--font-mono)",padding:"1px 0"}}>{k}=...</div>
                ))}
              </div>
            </div>
          )}

          {file&&(
            <div style={{display:"flex",flexDirection:"column",gap:12}}>
              {/* 데이터 미리보기 */}
              <div style={{padding:"10px 14px",background:"var(--color-background-secondary)",borderRadius:"var(--border-radius-md)",border:"0.5px solid var(--color-border-tertiary)"}}>
                <div style={{fontSize:11,fontWeight:500,marginBottom:6,color:"var(--color-text-secondary)"}}>등록 예정 품목 ({s2rows.length}개 SKU · {totC.toLocaleString()}개)</div>
                <div style={{display:"flex",flexWrap:"wrap",gap:4}}>
                  {s2rows.map(r=>(
                    <span key={r.sku} style={{fontSize:11,padding:"2px 8px",borderRadius:10,background:"var(--color-background-primary)",border:"0.5px solid var(--color-border-tertiary)"}}>
                      {r.sku} <span style={{color:"var(--color-text-info)",fontWeight:600}}>{r.total}개</span>
                      {!master[r.sku]?.sku&&<span style={{color:"var(--color-text-danger)",marginLeft:4}}>⚠️SKU없음</span>}
                    </span>
                  ))}
                </div>
              </div>

              {/* Step 진행 표시 */}
              {(()=>{
                const stepConfig:{key:FbaStepKey,label:string}[]=[
                  {key:'createPlan',label:'① 출하 플랜 생성'},
                  {key:'packingOptions',label:'② 패킹 옵션 확정'},
                  {key:'setPackingInfo',label:'③ 박스 내용물 입력'},
                  {key:'placementOptions',label:'④ FC 배치 (수동 선택)'},
                  {key:'transport',label:'⑤ 운송 설정 + PRO 번호'},
                  {key:'labels',label:'⑥ 라벨 URL 취득'},
                  {key:'bol',label:'⑦ BOL URL 취득'},
                ]
                return(
                  <div style={{border:"0.5px solid var(--color-border-tertiary)",borderRadius:"var(--border-radius-md)",overflow:"hidden"}}>
                    {stepConfig.map(({key,label})=>{
                      const step=fbaSteps[key]
                      const icon={idle:'○',running:'⏳',done:'✅',error:'❌',waiting:'⏸️'}[step.status]
                      const bg={idle:'transparent',running:'rgba(219,234,254,0.3)',done:'rgba(220,252,231,0.3)',error:'rgba(254,202,202,0.3)',waiting:'rgba(254,243,199,0.3)'}[step.status]
                      const col={idle:'var(--color-text-tertiary)',running:'var(--color-text-info)',done:'var(--color-text-success)',error:'var(--color-text-danger)',waiting:'#92400e'}[step.status]
                      return(
                        <div key={key} style={{display:"flex",alignItems:"center",gap:10,padding:"8px 14px",background:bg,borderBottom:"0.5px solid var(--color-border-tertiary)"}}>
                          <span style={{fontSize:16,width:22,textAlign:"center"}}>{icon}</span>
                          <span style={{fontSize:12,fontWeight:500,minWidth:180,color:col}}>{label}</span>
                          {step.msg&&<span style={{fontSize:11,color:col,opacity:0.85}}>{step.msg}</span>}
                        </div>
                      )
                    })}
                  </div>
                )
              })()}

              {/* FC 배치 옵션 선택 (④ waiting 상태일 때) */}
              {fbaSteps.placementOptions.status==='waiting'&&fbaPlacementOptions.length>0&&(
                <div style={{padding:"14px",border:"1px solid #f59e0b",borderRadius:"var(--border-radius-md)",background:"rgba(254,243,199,0.3)"}}>
                  <div style={{fontSize:12,fontWeight:600,marginBottom:10,color:"#92400e"}}>⏸️ FC 배치 옵션 선택 — 선택 후 [계속] 버튼을 클릭하세요</div>
                  <div style={{display:"flex",flexDirection:"column",gap:6,marginBottom:12}}>
                    {fbaPlacementOptions.map((opt,i)=>(
                      <label key={opt.placementOptionId} style={{display:"flex",alignItems:"center",gap:8,padding:"8px 12px",border:`1px solid ${fbaSelectedPlacement===opt.placementOptionId?"#f59e0b":"var(--color-border-tertiary)"}`,borderRadius:"var(--border-radius-md)",background:fbaSelectedPlacement===opt.placementOptionId?"rgba(254,243,199,0.5)":"var(--color-background-primary)",cursor:"pointer"}}>
                        <input type="radio" name="placement" value={opt.placementOptionId}
                          checked={fbaSelectedPlacement===opt.placementOptionId}
                          onChange={()=>setFbaSelectedPlacement(opt.placementOptionId)}/>
                        <span style={{fontSize:12,fontWeight:500}}>옵션 {i+1}</span>
                        <span style={{fontSize:11,fontFamily:"var(--font-mono)",color:"var(--color-text-secondary)"}}>{opt.placementOptionId}</span>
                        <span style={{fontSize:11,padding:"1px 6px",borderRadius:8,background:"var(--color-background-secondary)",border:"0.5px solid var(--color-border-tertiary)"}}>{opt.status}</span>
                        {opt.fees&&<span style={{fontSize:11,color:"var(--color-text-danger)"}}>{JSON.stringify(opt.fees)}</span>}
                      </label>
                    ))}
                  </div>
                  <button
                    onClick={runFbaPhase2}
                    disabled={!fbaSelectedPlacement||fbaRunning}
                    style={{padding:"8px 24px",fontSize:13,fontWeight:600,background:fbaSelectedPlacement?"#f59e0b":"var(--color-background-secondary)",color:fbaSelectedPlacement?"#fff":"var(--color-text-tertiary)",border:"none",borderRadius:"var(--border-radius-md)",cursor:fbaSelectedPlacement?"pointer":"default",opacity:fbaRunning?0.6:1}}>
                    {fbaRunning?"처리 중...":"▶ 계속 (운송 설정 + 라벨 다운로드)"}
                  </button>
                </div>
              )}

              {/* 라벨 + BOL 다운로드 */}
              {(Object.keys(fbaLabelUrls).length>0||Object.keys(fbaBolUrls).length>0)&&(
                <div style={{border:"0.5px solid var(--color-border-success)",borderRadius:"var(--border-radius-md)",overflow:"hidden"}}>
                  <div style={{padding:"8px 14px",background:"rgba(220,252,231,0.4)",fontSize:11,fontWeight:600,color:"var(--color-text-success)",borderBottom:"0.5px solid var(--color-border-success)"}}>📥 다운로드</div>
                  <div style={{padding:"12px 14px",display:"flex",flexWrap:"wrap",gap:8}}>
                    {Object.entries(fbaLabelUrls).map(([key,url])=>(
                      <a key={key} href={url} target="_blank" rel="noreferrer"
                        style={{fontSize:11,padding:"4px 12px",borderRadius:4,border:"0.5px solid var(--color-border-success)",background:"rgba(220,252,231,0.3)",color:"var(--color-text-success)",textDecoration:"none",fontWeight:500}}>
                        {key.startsWith('carton')?'📦 카톤 라벨':'🏷️ 파렛트 라벨'} ({key.split('_').slice(1).join('_')})
                      </a>
                    ))}
                    {Object.entries(fbaBolUrls).map(([shipmentId,url])=>(
                      <a key={shipmentId} href={url} target="_blank" rel="noreferrer"
                        style={{fontSize:11,padding:"4px 12px",borderRadius:4,border:"0.5px solid var(--color-border-info)",background:"rgba(219,234,254,0.3)",color:"var(--color-text-info)",textDecoration:"none",fontWeight:500}}>
                        📄 BOL ({shipmentId})
                      </a>
                    ))}
                  </div>
                </div>
              )}

              {/* 출고 창고 + 운송 모드 선택 + 실행 버튼 */}
              <div style={{display:"flex",gap:8,alignItems:"center",flexWrap:"wrap"}}>
                {/* 창고 선택 */}
                <div style={{display:"flex",border:"0.5px solid var(--color-border-tertiary)",borderRadius:"var(--border-radius-md)",overflow:"hidden"}}>
                  {(Object.entries(FBA_WAREHOUSES) as [typeof fbaWarehouse, {label:string;city:string}][]).map(([key,wh])=>(
                    <button key={key} onClick={()=>setFbaWarehouse(key)}
                      style={{padding:"7px 16px",fontSize:12,fontWeight:fbaWarehouse===key?600:400,border:"none",cursor:"pointer",background:fbaWarehouse===key?"var(--color-text-primary)":"transparent",color:fbaWarehouse===key?"var(--color-background-primary)":"var(--color-text-secondary)",transition:"all 0.15s"}}>
                      🏭 {wh.label}
                    </button>
                  ))}
                </div>
                {/* 운송 모드 선택 */}
                <div style={{display:"flex",border:"0.5px solid var(--color-border-tertiary)",borderRadius:"var(--border-radius-md)",overflow:"hidden"}}>
                  {FBA_SHIPPING_MODES.map(m=>(
                    <button key={m.key} onClick={()=>setFbaShippingMode(m.key)}
                      title={m.desc}
                      style={{padding:"7px 14px",fontSize:12,fontWeight:fbaShippingMode===m.key?600:400,border:"none",cursor:"pointer",background:fbaShippingMode===m.key?"#0891b2":"transparent",color:fbaShippingMode===m.key?"#fff":"var(--color-text-secondary)",transition:"all 0.15s"}}>
                      🚛 {m.label}
                    </button>
                  ))}
                </div>
                {/* 출하 예정일 — readyToShipWindow 시작일 */}
                <label style={{display:"flex",alignItems:"center",gap:5,fontSize:12,color:"var(--color-text-secondary)"}}>
                  📅 출하 예정일
                  <input type="date" value={fbaShipDate} onChange={e=>setFbaShipDate(e.target.value)}
                    style={{fontSize:12,padding:"5px 8px",border:"0.5px solid var(--color-border-tertiary)",borderRadius:"var(--border-radius-md)",background:"var(--color-background-primary)",color:"var(--color-text-primary)",outline:"none"}}/>
                </label>
                <button
                  onClick={runFbaInbound}
                  disabled={fbaRunning||!s2rows.length}
                  style={{padding:"10px 28px",fontSize:13,fontWeight:600,background:s2rows.length?"#2563eb":"var(--color-background-secondary)",color:s2rows.length?"#fff":"var(--color-text-tertiary)",border:"none",borderRadius:"var(--border-radius-md)",cursor:fbaRunning||!s2rows.length?"default":"pointer",opacity:fbaRunning?0.6:1}}>
                  {fbaRunning?"처리 중...":"🚀 FBA Inbound 자동 등록 시작"}
                </button>
                {fbaSteps.createPlan.status!=='idle'&&(
                  <button onClick={()=>{setFbaSteps(initSteps());setFbaPlanId('');setFbaShipmentIds([]);setFbaPlacementOptions([]);setFbaSelectedPlacement('');setFbaTransportOptions({});setFbaLabelUrls({});setFbaBolUrls({})}}
                    style={{fontSize:11,padding:"6px 14px",cursor:"pointer",border:"0.5px solid var(--color-border-tertiary)",borderRadius:"var(--border-radius-md)",background:"transparent",color:"var(--color-text-secondary)"}}>
                    초기화
                  </button>
                )}
                {s2rows.some(r=>!master[r.sku]?.sku)&&(
                  <span style={{fontSize:11,color:"var(--color-text-danger)"}}>⚠️ 마스터에 SKU(Merchant SKU) 없는 항목 있음 → 매핑 관리에서 입력 필요</span>
                )}
              </div>
            </div>
          )}
        </div>
      )}

      {/* ══ 총합 탭 ══ */}
      {mode==='total'&&(
        <div style={{padding:"16px 0",display:"flex",flexDirection:"column",gap:16}}>
          {/* 시트 선택 */}
          <div style={{display:"flex",gap:6,flexWrap:"wrap",alignItems:"center"}}>
            <span style={{fontSize:11,fontWeight:600,color:"var(--color-text-secondary)",whiteSpace:"nowrap"}}>시트 선택</span>
            {sheetNames.length===0&&<span style={{fontSize:12,color:"var(--color-text-tertiary)"}}>엑셀 파일을 먼저 업로드해주세요</span>}
            {sheetNames.map(sn=>{
              const sel=!totalExcludedSheets.has(sn)
              return(<button key={sn} onClick={()=>setTotalExcludedSheets(prev=>{const n=new Set(prev);if(n.has(sn))n.delete(sn);else n.add(sn);return n})}
                style={{padding:"3px 10px",fontSize:11,borderRadius:20,cursor:"pointer",border:`1px solid ${sel?"#06b6d4":"var(--color-border-tertiary)"}`,background:sel?"rgba(6,182,212,0.12)":"transparent",color:sel?"#0891b2":"var(--color-text-tertiary)",fontWeight:sel?600:400}}>{sn}</button>)
            })}
            {sheetNames.length>0&&<>
              <button onClick={()=>setTotalExcludedSheets(new Set())} style={{fontSize:11,padding:"3px 10px",cursor:"pointer",border:"0.5px solid var(--color-border-tertiary)",borderRadius:6,background:"transparent",color:"var(--color-text-secondary)"}}>전체 선택</button>
              <button onClick={()=>setTotalExcludedSheets(new Set(sheetNames))} style={{fontSize:11,padding:"3px 10px",cursor:"pointer",border:"0.5px solid var(--color-border-tertiary)",borderRadius:6,background:"transparent",color:"var(--color-text-secondary)"}}>전체 해제</button>
              <button onClick={()=>setTotalAdjustments({})} style={{fontSize:11,padding:"3px 10px",cursor:"pointer",border:"0.5px solid var(--color-border-tertiary)",borderRadius:6,background:"transparent",color:"var(--color-text-secondary)"}}>조정 초기화</button>
              <button onClick={expTotal} style={{fontSize:11,padding:"3px 12px",cursor:"pointer",border:"0.5px solid #16a34a",borderRadius:6,background:"rgba(22,163,74,0.08)",color:"#16a34a",fontWeight:600,marginLeft:8}}>⬇ 엑셀 다운로드</button>
            </>}
          </div>
          {/* 합산 표 */}
          {sheetNames.length>0&&(()=>{
            const rows=buildTotalRows()
            const grandTotal=rows.reduce((s,r)=>s+r.final,0)
            const totalAdj=rows.reduce((s,r)=>s+r.adj,0)
            const totalRaw=rows.reduce((s,r)=>s+r.total,0)
            const selSheets=sheetNames.filter(sn=>!totalExcludedSheets.has(sn))
            return(<>
              <div style={{fontSize:11,color:"var(--color-text-tertiary)"}}>
                선택된 시트: {selSheets.length?selSheets.join(', '):'없음'}{' · '}SKU {rows.length}종{' · '}합계 {grandTotal.toLocaleString()}개
              </div>
              <div style={{overflowX:"auto"}}>
                <table style={{borderCollapse:"collapse",width:"100%",fontSize:12,tableLayout:"fixed"}}>
                  <colgroup><col style={{width:"190px"}}/><col style={{width:"100px"}}/><col style={{width:"120px"}}/><col style={{width:"110px"}}/></colgroup>
                  <thead>
                    <tr>
                      {(["약호","합계","조정(+/-)","최종합계"] as const).map((h,i)=>(
                        <th key={h} style={{...TH,textAlign:i===0?"left":"right"}}>{h}</th>
                      ))}
                    </tr>
                  </thead>
                  <tbody>
                    {rows.map(r=>(
                      <tr key={r.sku} style={{background:r.adj!==0?"rgba(6,182,212,0.05)":undefined}}>
                        <td style={TD}>{r.sku}</td>
                        <td style={{...TD,textAlign:"right"}}>{r.total.toLocaleString()}</td>
                        <td style={{...TD,padding:"3px 8px",textAlign:"right"}}>
                          <input type="number" value={r.adj===0?"":r.adj} placeholder="0"
                            onChange={e=>{const v=parseInt(e.target.value)||0;setTotalAdjustments(prev=>({...prev,[r.sku]:v}))}}
                            style={{width:"100%",textAlign:"right",fontSize:12,padding:"2px 4px",border:"0.5px solid var(--color-border-tertiary)",borderRadius:4,background:r.adj>0?"rgba(22,163,74,0.08)":r.adj<0?"rgba(220,38,38,0.08)":"var(--color-background-primary)",color:r.adj>0?"#16a34a":r.adj<0?"#dc2626":"var(--color-text-primary)",outline:"none",fontFamily:"inherit"}}/>
                        </td>
                        <td style={{...TD,textAlign:"right",fontWeight:600,color:r.adj>0?"#16a34a":r.adj<0?"#dc2626":"var(--color-text-primary)"}}>{r.final.toLocaleString()}</td>
                      </tr>
                    ))}
                    <tr style={{borderTop:"2px solid var(--color-border-secondary)",background:"var(--color-background-secondary)"}}>
                      <td style={{...TD,fontWeight:700}}>합계</td>
                      <td style={{...TD,textAlign:"right",fontWeight:600}}>{totalRaw.toLocaleString()}</td>
                      <td style={{...TD,textAlign:"right",fontWeight:600,color:totalAdj>0?"#16a34a":totalAdj<0?"#dc2626":"var(--color-text-primary)"}}>{totalAdj>0?"+":""}{totalAdj.toLocaleString()}</td>
                      <td style={{...TD,textAlign:"right",fontWeight:700,fontSize:13}}>{grandTotal.toLocaleString()}</td>
                    </tr>
                  </tbody>
                </table>
              </div>
            </>)
          })()}
        </div>
      )}

      {/* ══ CI/PL 변환기 ══ */}
      {mode==='cipl'&&(
        <div style={{padding:"16px 0",display:"flex",flexDirection:"column",gap:16}}>
          <div style={{display:"flex",alignItems:"center",gap:8}}>
            <span style={{fontSize:13,fontWeight:600,color:"var(--color-text-primary)"}}>CI/PL 변환기</span>
            <span style={{fontSize:11,color:"var(--color-text-tertiary)"}}>한국용 오더짜기 Excel → 일본용 CI/PL 변환</span>
          </div>

          {/* 파일 업로드 */}
          <div style={{display:"flex",alignItems:"center",gap:10,flexWrap:"wrap"}}>
            <div onClick={()=>ciplRef.current?.click()} style={{display:"flex",alignItems:"center",gap:8,padding:"7px 16px",border:"1px dashed var(--color-border-secondary)",borderRadius:"var(--border-radius-md)",cursor:"pointer",fontSize:13}}>
              <span>+</span>
              <span style={{color:ciplFile?"var(--color-text-primary)":"var(--color-text-tertiary)"}}>{ciplFile?ciplFile.name:"한국용 CI/PL Excel 업로드 (.xlsx)"}</span>
            </div>
            <input ref={ciplRef} type="file" accept=".xlsx" style={{display:"none"}} onChange={e=>e.target.files?.[0]&&loadCiplFile(e.target.files[0])}/>
            {ciplWB&&<span style={{fontSize:11,padding:"3px 9px",borderRadius:20,background:"rgba(22,163,74,0.1)",border:"0.5px solid #16a34a",color:"#16a34a"}}>✓ 파싱 완료</span>}
          </div>

          {/* 항구 선택 */}
          <div style={{display:"flex",alignItems:"center",gap:8}}>
            <span style={{fontSize:12,color:"var(--color-text-secondary)",fontWeight:600}}>CJ 수입 담당 지점</span>
            {([['kanto','관동 (SAITAMA / 도쿄)'],['kansai','관서 (OSAKA)']] as const).map(([k,l])=>(
              <button key={k} onClick={()=>setCiplPort(k)} style={{padding:"5px 14px",fontSize:12,borderRadius:6,cursor:"pointer",border:`1px solid ${ciplPort===k?"var(--color-border-info)":"var(--color-border-tertiary)"}`,background:ciplPort===k?"var(--color-background-info)":"transparent",color:ciplPort===k?"var(--color-text-info)":"var(--color-text-secondary)",fontWeight:ciplPort===k?600:400}}>{l}</button>
            ))}
          </div>

          {/* 변환 미리보기 */}
          {ciplWB&&(()=>{
            const inv=ciplWB.Sheets['Invoice ']
            if(!inv) return null
            const invNo=String(inv['H3']?.v||'')
            const dept=inv['A17']?.v
            const deptStr=dept?(typeof dept==='number'?XLSX.SSF.format('yyyy-mm-dd',dept):String(dept)):''
            const from=String(inv['A20']?.v||'')
            const to=String(inv['F20']?.v||'')
            const vessel=String(inv['A23']?.v||'')
            const total=inv['N37']?.v
            // 품목 수 카운트
            let itemCount=0; let r=27
            while(inv[`E${r}`]?.v){ itemCount++; r++ }
            return(
              <div style={{background:"var(--color-background-secondary)",borderRadius:8,padding:"12px 16px",fontSize:12,display:"flex",flexDirection:"column",gap:6}}>
                <div style={{fontWeight:600,marginBottom:4,fontSize:13}}>파싱 결과 미리보기</div>
                <div style={{display:"grid",gridTemplateColumns:"auto 1fr",gap:"3px 16px"}}>
                  <span style={{color:"var(--color-text-tertiary)"}}>Invoice No.</span><span>{invNo}</span>
                  <span style={{color:"var(--color-text-tertiary)"}}>출항일</span><span>{deptStr}</span>
                  <span style={{color:"var(--color-text-tertiary)"}}>From → To</span><span>{from} → {to}</span>
                  <span style={{color:"var(--color-text-tertiary)"}}>선박</span><span>{vessel}</span>
                  <span style={{color:"var(--color-text-tertiary)"}}>품목 수</span><span>{itemCount}종</span>
                  <span style={{color:"var(--color-text-tertiary)"}}>합계</span><span>US$ {total?.toLocaleString?.()??total}</span>
                  <span style={{color:"var(--color-text-tertiary)"}}>수입자 (변환 후)</span><span style={{color:"var(--color-text-info)"}}>{JP_CONSIGNEE_LINES[0]}</span>
                  <span style={{color:"var(--color-text-tertiary)"}}>통지처 (변환 후)</span><span style={{color:"var(--color-text-info)"}}>{(ciplPort==='kanto'?JP_NOTIFY_KANTO:JP_NOTIFY_KANSAI)[0]}</span>
                </div>
              </div>
            )
          })()}

          {/* 다운로드 버튼 */}
          <div>
            <button onClick={downloadCipl} disabled={!ciplWB} style={{padding:"8px 22px",fontSize:13,fontWeight:600,cursor:ciplWB?"pointer":"not-allowed",border:"0.5px solid #16a34a",borderRadius:6,background:ciplWB?"rgba(22,163,74,0.1)":"var(--color-background-secondary)",color:ciplWB?"#16a34a":"var(--color-text-tertiary)"}}>
              ⬇ 일본용 CI/PL Excel 다운로드
            </button>
            <span style={{fontSize:11,color:"var(--color-text-tertiary)",marginLeft:12}}>Invoice + Packing 시트 포함</span>
          </div>

          <div style={{fontSize:11,color:"var(--color-text-tertiary)",borderTop:"0.5px solid var(--color-border-tertiary)",paddingTop:10}}>
            변환 내용: Consignee → Homedant (CJ ON BEHALF OF) · ACP 정보 추가 · 원산지 SOUTH KOREA · {ciplPort==='kanto'?'통지처 SAITAMA':'통지처 OSAKA'} · 제품 데이터 그대로 유지
          </div>
        </div>
      )}

      {/* ══ 납품처 정보 ══ */}
      {mode==='delivery'&&(
        <div style={{padding:"16px 0",display:"flex",flexDirection:"column",gap:20}}>
          <div style={{display:"flex",alignItems:"center",gap:8}}>
            <span style={{fontSize:13,fontWeight:600,color:"var(--color-text-primary)"}}>납품처 정보 자동 작성</span>
            <span style={{fontSize:11,color:"var(--color-text-tertiary)"}}>세관 운송선 일람표 (輸入申告に係る運送先一覧表) — B/L 번호만 입력</span>
          </div>

          {/* 1차가공 FC 현황 */}
          {(()=>{
            const s2rows=buildS2rows()
            const usedFCs=[...new Set(s2rows.map(r=>r.fc).filter(Boolean))]
            if(!usedFCs.length) return(
              <div style={{fontSize:12,color:"var(--color-text-tertiary)"}}>1차가공에서 FC 정보를 입력하면 여기에 표시됩니다</div>
            )
            return(
              <div style={{background:"var(--color-background-secondary)",borderRadius:8,padding:"10px 14px",fontSize:12}}>
                <div style={{fontWeight:600,marginBottom:6}}>현재 1차가공 FC 현황</div>
                <div style={{display:"flex",gap:8,flexWrap:"wrap"}}>
                  {usedFCs.map(fc=>{
                    const addr=s2rows.find(r=>r.fc===fc)?.address||''
                    return(<span key={fc} style={{padding:"3px 10px",borderRadius:20,background:"rgba(6,182,212,0.1)",border:"0.5px solid #0891b2",color:"#0891b2",fontSize:11}}>{fc}{addr?` — ${addr.slice(0,30)}`:''}</span>)
                  })}
                </div>
              </div>
            )
          })()}

          {/* 세관 양식 섹션 */}
          <div style={{display:"flex",flexDirection:"column",gap:12}}>
            {([['saitama','SAITAMA (관동)','埼玉 / 東京 행',dlBL_Saitama,setDlBL_Saitama],['osaka','OSAKA (관서)','大阪 행',dlBL_Osaka,setDlBL_Osaka]] as const).map(([type,label,hint,bl,setBL])=>(
              <div key={type} style={{display:"flex",alignItems:"center",gap:10,padding:"12px 16px",border:"0.5px solid var(--color-border-tertiary)",borderRadius:8,flexWrap:"wrap"}}>
                <span style={{fontSize:12,fontWeight:600,minWidth:130,color:"var(--color-text-primary)"}}>{label}</span>
                <span style={{fontSize:11,color:"var(--color-text-tertiary)",minWidth:80}}>{hint}</span>
                <input
                  type="text"
                  placeholder="B/L 번호 입력 (예: COKR26007992)"
                  value={bl}
                  onChange={e=>setBL(e.target.value)}
                  style={{fontSize:12,padding:"5px 10px",border:"0.5px solid var(--color-border-secondary)",borderRadius:4,background:"var(--color-background-primary)",color:"var(--color-text-primary)",outline:"none",width:240,fontFamily:"inherit"}}
                />
                <button
                  onClick={()=>downloadDelivery(type,bl)}
                  disabled={!bl.trim()}
                  style={{padding:"6px 16px",fontSize:12,fontWeight:600,cursor:bl.trim()?"pointer":"not-allowed",border:"0.5px solid #16a34a",borderRadius:6,background:bl.trim()?"rgba(22,163,74,0.1)":"var(--color-background-secondary)",color:bl.trim()?"#16a34a":"var(--color-text-tertiary)"}}>
                  ⬇ 다운로드
                </button>
              </div>
            ))}
          </div>

          <div style={{fontSize:11,color:"var(--color-text-tertiary)",borderTop:"0.5px solid var(--color-border-tertiary)",paddingTop:10}}>
            세관 정부 양식 원본 형식 그대로 유지 · B/L 번호만 입력됨 · 和文住所用 + 英文住所用 두 시트 모두 적용
          </div>
        </div>
      )}
    </div>
  )
}
