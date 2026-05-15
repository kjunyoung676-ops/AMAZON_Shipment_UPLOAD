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

    // 배경
    coverPage.drawRectangle({ x: 0, y: 0, width: 595, height: 842, color: pdfRgb(0.96, 0.96, 0.96) })
    // 상단 짙은 회색 바
    coverPage.drawRectangle({ x: 0, y: 742, width: 595, height: 100, color: pdfRgb(0.15, 0.15, 0.15) })
    // CARTON LABEL 헤더
    const title = 'CARTON LABEL'
    const tw = font.widthOfTextAtSize(title, 26)
    coverPage.drawText(title, { x: (595 - tw) / 2, y: 790, size: 26, font, color: pdfRgb(0.85, 0.85, 0.85) })

    // BOX CODE (크게)
    const locSize = loc.length <= 7 ? 100 : 72
    const locW = font.widthOfTextAtSize(loc, locSize)
    coverPage.drawText(loc, { x: (595 - locW) / 2, y: 560, size: locSize, font, color: pdfRgb(0.08, 0.08, 0.08) })

    // 구분선
    coverPage.drawLine({ start: {x: 60, y: 540}, end: {x: 535, y: 540}, thickness: 2, color: pdfRgb(0.6, 0.6, 0.65) })

    // realSku (중간 크기)
    const label2 = realSku || sku || ''
    if (label2) {
      const fs2 = label2.length > 25 ? 24 : 32
      const lw2 = font.widthOfTextAtSize(label2, fs2)
      coverPage.drawText(label2, { x: Math.max(40, (595 - lw2) / 2), y: 460, size: fs2, font, color: pdfRgb(0.25, 0.25, 0.3) })
    }

    // 하단 바
    coverPage.drawRectangle({ x: 0, y: 0, width: 595, height: 40, color: pdfRgb(0.15, 0.15, 0.15) })
    const sub2 = 'FBA CARTON SHIPPING LABEL'
    const sw2 = font.widthOfTextAtSize(sub2, 13)
    coverPage.drawText(sub2, { x: (595 - sw2) / 2, y: 13, size: 13, font, color: pdfRgb(0.65, 0.65, 0.65) })
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
      const { PDFDocument, rgb } = await import('pdf-lib')

      type PageEntry = {
        fileBytes: Uint8Array
        pageIdx: number
        labelsToKeep: number[]
        totalLabels: number
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

        // 앞 표지 페이지
        await makeCoverPage(outDoc, loc, false)

        for (const entry of entries) {
          const srcDoc = await PDFDocument.load(entry.fileBytes.slice())
          const [copiedPage] = await outDoc.copyPages(srcDoc, [entry.pageIdx])

          if (entry.labelsToKeep.length > 0) {
            const { width, height } = copiedPage.getSize()
            const cols = 2
            const rows = Math.ceil(entry.totalLabels / cols)
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
    const active=mode===m,avail=m==='1'||m==='map'||m==='label'||m==='logistics'||m==='s3b'||!!file
    return(<button onClick={()=>avail&&setMode(m)} style={{padding:"8px 16px",fontSize:12,fontWeight:active?500:400,cursor:avail?"pointer":"default",border:"none",background:active?"var(--color-text-primary)":"transparent",color:active?"var(--color-background-primary)":avail?"var(--color-text-secondary)":"var(--color-text-tertiary)"}}>{label}</button>)
  }

  return (
    <div style={{padding:"0.75rem 0",fontFamily:"var(--font-sans)"}}>
      {/* 내비 */}
      <div style={{display:"flex",alignItems:"center",gap:8,marginBottom:16}}>
        <div style={{border:"0.5px solid var(--color-border-tertiary)",borderRadius:"var(--border-radius-md)",overflow:"hidden",display:"flex"}}>
          <NavBtn m="1" label="1. 업로드"/><div style={{width:"0.5px",background:"var(--color-border-tertiary)"}}/>
          <NavBtn m="2" label="2. 1차 가공 (준비)"/><div style={{width:"0.5px",background:"var(--color-border-tertiary)"}}/>
          <NavBtn m="label" label="2. 라벨 분류"/><div style={{width:"0.5px",background:"var(--color-border-tertiary)"}}/>
          <NavBtn m="logistics" label="2-5. 물류 전달"/><div style={{width:"0.5px",background:"var(--color-border-tertiary)"}}/>
          <NavBtn m="3" label="3. 2차 가공 (적재)"/><div style={{width:"0.5px",background:"var(--color-border-tertiary)"}}/>
          <NavBtn m="s3b" label="3. 3차 가공 (역산)"/>
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
            {/* 변동값 JSON 백업/복원 */}
            <button onClick={exportMetaJson} style={{fontSize:11,padding:"3px 10px",cursor:"pointer",border:"0.5px solid var(--color-border-secondary)",borderRadius:"var(--border-radius-md)",background:"transparent",color:"var(--color-text-secondary)"}}>⬇ 메타 JSON 백업</button>
            <label style={{fontSize:11,padding:"3px 10px",cursor:"pointer",border:"0.5px solid var(--color-border-info)",borderRadius:"var(--border-radius-md)",background:"rgba(219,234,254,0.25)",color:"var(--color-text-info)"}}>
              ⬆ 메타 JSON 복원
              <input ref={metaJsonRef} type="file" accept=".json" style={{display:"none"}} onChange={e=>{const f=e.target.files?.[0];if(f){importMetaJson(f);e.target.value=""}}}/>
            </label>
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
                    <th style={{...TH,textAlign:"right"}}>파일 크기</th>
                    <th style={TH}></th>
                  </tr></thead>
                  <tbody>{Object.entries(labelGroups).filter(([k])=>k.startsWith('__PLT__')).map(([loc,bytes],i)=>{
                    const displayKey = loc.replace('__PLT__','')
                    const cnt = labelCounts[loc]||0
                    return(<tr key={loc} style={{background:i%2===0?"rgba(219,234,254,0.08)":"rgba(219,234,254,0.2)"}}>
                      <td style={{...TD,fontWeight:500,color:"var(--color-text-info)"}}>{displayKey}</td>
                      <td style={{...TD,textAlign:"right",fontWeight:600}}>{cnt}파렛트</td>
                      <td style={{...TD,textAlign:"right",color:"var(--color-text-secondary)"}}>{(bytes.length/1024).toFixed(0)}KB</td>
                      <td style={TD}><button onClick={()=>dlLabel(loc,bytes)} style={{fontSize:11,padding:"2px 10px",cursor:"pointer",borderRadius:4,border:"0.5px solid var(--color-border-info)",background:"transparent"}}>↓ PDF</button></td>
                    </tr>)
                  })}</tbody>
                </table>
              </div>
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
        const blAgg: Record<string, Record<string, { qty: number; blId: string }>> = {}
        for (const r of fd) {
          const sku = String(r.sku || '')
          if (!sku) continue
          const blId = String(r.booking_ref || r.destination || r.container_no || 'BL1')
          if (!blAgg[blId]) blAgg[blId] = {}
          if (!blAgg[blId][sku]) blAgg[blId][sku] = { qty: 0, blId }
          blAgg[blId][sku].qty += r.quantity
        }

        // BL별 행 생성
        const allBLRows: BLRow[] = []
        let blNo = 1
        for (const [blId, skuMap] of Object.entries(blAgg)) {
          for (const [sku, { qty }] of Object.entries(skuMap)) {
            const m = master[sku] || {} as MasterItem
            rowNo++
            allBLRows.push({
              no: rowNo,
              blId,
              sku,
              asin: m.asin || '',
              loc: m.loc || '',
              qty,
              gw: calcGW(qty, parseFloat(String(m.kg)) || 0),
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
            {file && allBLRows.length > 0 && (()=>{
              const fcPallets: Record<string, number> = {}
              for (const r of s2rows) {
                const fc = (s2meta[r.sku]||{}).fc || ''
                if (!fc) continue
                if (!fcPallets[fc]) fcPallets[fc] = 0
                fcPallets[fc] += r.pallets
              }
              // FC센터 주소 기반 거리 분류
              // 장거리(도쿄/요코하마 400km+): 10T만 사용
              // 중거리(나고야 ~200km): 10T + 4T
              // 단거리(오사카/고베 ~50km 이내): 10T + 4T + 2T
              function getDistanceClass(fc: string, address: string): 'long' | 'mid' | 'short' {
                const addr = (address || '').toUpperCase() + fc.toUpperCase()
                // 도쿄권 (장거리)
                if (/東京|TOKYO|神奈川|KANAGAWA|YOKOHAMA|横浜|埼玉|SAITAMA|千葉|CHIBA/.test(addr)) return 'long'
                // 나고야권 (중거리)
                if (/愛知|AICHI|NAGOYA|名古屋|三重|MIE|岐阜|GIFU/.test(addr)) return 'mid'
                // 오사카/고베 (단거리)
                if (/兵庫|HYOGO|神戸|KOBE|大阪|OSAKA|京都|KYOTO/.test(addr)) return 'short'
                // FC코드로 fallback
                if (['HIY1','TPB5'].includes(fc.toUpperCase())) return 'long'
                if (['TPB8'].includes(fc.toUpperCase())) return 'mid'
                if (['VJNB'].includes(fc.toUpperCase())) return 'short'
                return 'long' // 기본값은 장거리 (안전하게)
              }

              function calcTrucks(pallets: number, distClass: 'long' | 'mid' | 'short') {
                const t10 = Math.floor(pallets / 14)
                const rem = pallets - t10 * 14
                let t4 = 0, t2 = 0
                if (rem > 0) {
                  if (distClass === 'long') {
                    // 장거리: 2T 없음, 나머지는 4T 또는 10T 추가
                    if (rem <= 6) t4 = 1
                    else { return { t10: t10 + 1, t4: 0, t2: 0 } }
                  } else if (distClass === 'mid') {
                    // 중거리: 4T까지 사용
                    if (rem <= 6) t4 = 1
                    else { return { t10: t10 + 1, t4: 0, t2: 0 } }
                  } else {
                    // 단거리: 2T도 사용 가능
                    if (rem <= 2) t2 = 1
                    else if (rem <= 6) t4 = 1
                    else { return { t10: t10 + 1, t4: 0, t2: 0 } }
                  }
                }
                return { t10, t4, t2 }
              }
              // ── 드레이지 단가: 관동/관서 구분 ──
              // 관서: 오사카 남항 → CJ 오사카센터 → 40HQ ¥62,000
              // 관동: 도쿄항 → CJ 사이타마센터 → 별도 견적 (일단 동일 기준 또는 요청)
              const DRAY_KANSAI = 62000  // 관서 40HQ/본
              const DRAY_KANTO  = 62000  // 관동 40HQ/본 (확인 필요 — 동일 가정)
              const TAX = 1.1
              const INOUT = 500

              // BL별 입항지 자동 추정 (FC주소 기반) + 수동 override
              // FC가 관동권(도쿄/요코하마/사이타마)이면 관동, 관서권이면 관서
              function guessPort(blId: string): 'kansai' | 'kanto' {
                const blRows = allBLRows.filter(r => r.blId === blId)
                for (const r of blRows) {
                  const fc = r.sku ? (s2meta[r.sku]||{}).fc || '' : ''
                  const addr = (s2meta[r.sku]||{}).address || ''
                  const combined = (addr + fc).toUpperCase()
                  if (/東京|TOKYO|神奈川|KANAGAWA|YOKOHAMA|横浜|埼玉|SAITAMA|千葉|CHIBA/.test(combined)) return 'kanto'
                  if (/兵庫|HYOGO|神戸|KOBE|大阪|OSAKA|愛知|AICHI/.test(combined)) return 'kansai'
                  if (['HIY1','TPB5'].includes(fc)) return 'kanto'
                }
                return 'kansai' // 기본값
              }

              // 컨테이너 타입별 드레이지: 이번 건은 40HQ 기준
              // BL별로 컨테이너 수 집계
              const blCtnCount: Record<string, number> = {}
              for (const [blId] of Object.entries(blAgg)) {
                // 컨테이너 중 해당 BL 소속인 것 세기
                // fd에서 booking_ref가 blId인 row의 container_no 고유값 카운트
                const blCtnNos = new Set(fd.filter(r=>String(r.booking_ref||r.destination||r.container_no||'BL1')===blId).map(r=>r.container_no))
                blCtnCount[blId] = blCtnNos.size || 1
              }

              // BL별 FC 파렛트 집계
              const blFcPallets: Record<string, Record<string,number>> = {}
              for (const r of allBLRows) {
                if (!blFcPallets[r.blId]) blFcPallets[r.blId] = {}
                const fc = (s2meta[r.sku]||{}).fc || '?'
                if (!blFcPallets[r.blId][fc]) blFcPallets[r.blId][fc] = 0
                blFcPallets[r.blId][fc] += r.pallets
              }

              type CR = {item:string;qty:number;unit:string;up:number;sub:number;taxed:boolean;total:number;note?:string}

              // BL별 비용 계산
              const blCostSections: {blId:string; blNo:number; port:'kansai'|'kanto'; rows:CR[]; subtotal:number}[] = []

              blIds.forEach((blId, bi) => {
                const port = (blPortMode[blId] as 'kansai'|'kanto') || guessPort(blId)
                const ctnCnt = blCtnCount[blId] || 1
                const drayRate = port === 'kansai' ? DRAY_KANSAI : DRAY_KANTO
                const portLabel = port === 'kansai' ? '관서 (오사카 남항)' : '관동 (도쿄항)'
                const fcPlt = blFcPallets[blId] || {}
                const blTotalPlts = Object.values(fcPlt).reduce((s,v)=>s+v,0)

                const rows: CR[] = [
                  {item:'通関申告料（1申告/2HS）통관신고료', qty:1,unit:'件',up:11800,sub:11800,taxed:false,total:11800},
                  {item:'輸入取扱料 수입취급료', qty:1,unit:'件',up:10000,sub:10000,taxed:true,total:Math.round(10000*TAX)},
                  {item:'評価申告料（1申告/2HS）평가신고료', qty:1,unit:'件',up:11800,sub:11800,taxed:false,total:11800},
                  {item:'AN費用 AN비용', qty:1,unit:'件',up:0,sub:0,taxed:false,total:0,note:'견적'},
                  {item:`40Fドレー+デバン費用 ${portLabel} (${ctnCnt}본)`, qty:ctnCnt,unit:'本',up:drayRate,sub:ctnCnt*drayRate,taxed:true,total:Math.round(ctnCnt*drayRate*TAX)},
                ]

                for (const [fc, plts] of Object.entries(fcPlt)) {
                  const addr = (s2meta[Object.keys(s2meta).find(k=>(s2meta[k]?.fc||'')===fc)||'']||{}).address || ''
                  const distClass = getDistanceClass(fc, addr)
                  const tr = calcTrucks(plts, distClass)
                  const RATES: Record<string,{t10:number,t4:number,t2:number}> = {
                    HIY1:{t10:55000,t4:35000,t2:25000}, TPB5:{t10:55000,t4:45000,t2:25000},
                    VJNB:{t10:55000,t4:45000,t2:25000}, TPB8:{t10:55000,t4:45000,t2:25000},
                  }
                  const r = RATES[fc]||{t10:55000,t4:45000,t2:25000}
                  if (tr.t10>0){const s=tr.t10*r.t10;rows.push({item:`配送料金 (${fc}) 10T × ${tr.t10}台`,qty:tr.t10,unit:'本',up:r.t10,sub:s,taxed:true,total:Math.round(s*TAX)})}
                  if (tr.t4>0){const s=tr.t4*r.t4;rows.push({item:`配送料金 (${fc}) 4T × ${tr.t4}台`,qty:tr.t4,unit:'本',up:r.t4,sub:s,taxed:true,total:Math.round(s*TAX)})}
                  if (tr.t2>0){const s=tr.t2*r.t2;rows.push({item:`配送料金 (${fc}) 2T × ${tr.t2}台`,qty:tr.t2,unit:'本',up:r.t2,sub:s,taxed:true,total:Math.round(s*TAX)})}
                }
                rows.push({item:`入庫料 입고료 (${blTotalPlts}PLT)`,qty:blTotalPlts,unit:'PLT',up:INOUT,sub:blTotalPlts*INOUT,taxed:true,total:Math.round(blTotalPlts*INOUT*TAX)})
                rows.push({item:`出荷料 출고료 (${blTotalPlts}PLT)`,qty:blTotalPlts,unit:'PLT',up:INOUT,sub:blTotalPlts*INOUT,taxed:true,total:Math.round(blTotalPlts*INOUT*TAX)})
                const subtotal = rows.reduce((s,r)=>s+r.total,0)
                blCostSections.push({blId, blNo:bi+1, port, rows, subtotal})
              })

              const grandTotal = blCostSections.reduce((s,b)=>s+b.subtotal,0)

              return (
                <div style={{marginTop:20}}>
                  <div style={{display:"flex",alignItems:"center",gap:8,marginBottom:10}}>
                    <span style={{fontWeight:700,fontSize:13,color:"var(--color-text-primary)"}}>輸送費・輸入通関諸費用 (BL별)</span>
                    <span style={{fontSize:11,color:"var(--color-text-tertiary)"}}>※ 입항지 자동 감지 · 수동 변경 가능</span>
                  </div>
                  {blCostSections.map((bl, si) => {
                    const ci = si % CB.length
                    return (
                      <div key={bl.blId} style={{border:"0.5px solid var(--color-border-tertiary)",borderRadius:"var(--border-radius-md)",overflow:"hidden",marginBottom:12}}>
                        {/* BL 헤더 */}
                        <div style={{background:"#1e293b",color:"#fff",padding:"8px 14px",display:"flex",alignItems:"center",gap:10,flexWrap:"wrap"}}>
                          <span style={{fontWeight:700,fontSize:13}}>BL{bl.blNo} — {bl.blId}</span>
                          {/* 입항지 선택 */}
                          <div style={{display:"flex",alignItems:"center",gap:6,marginLeft:8,background:"rgba(255,255,255,0.12)",borderRadius:6,padding:"2px 4px"}}>
                            <span style={{fontSize:11,opacity:0.8}}>입항지:</span>
                            {(['kansai','kanto'] as const).map(p=>(
                              <button key={p} onClick={()=>setBlPortMode(prev=>({...prev,[bl.blId]:p}))} style={{fontSize:11,padding:"2px 10px",borderRadius:4,border:"none",cursor:"pointer",fontWeight:bl.port===p?700:400,background:bl.port===p?(p==='kansai'?'#3b82f6':'#10b981'):'rgba(255,255,255,0.2)',color:"#fff"}}>
                                {p==='kansai'?'관서 (오사카)':'관동 (도쿄)'}
                              </button>
                            ))}
                          </div>
                          {/* FC별 파렛트 요약 */}
                          {Object.entries(blFcPallets[bl.blId]||{}).map(([fc,plt])=>{
                            const addr=(s2meta[Object.keys(s2meta).find(k=>(s2meta[k]?.fc||'')===fc)||'']||{}).address||''
                            const dc=getDistanceClass(fc,addr)
                            const tr=calcTrucks(plt,dc)
                            return<span key={fc} style={{fontSize:10,padding:"1px 6px",borderRadius:8,background:"rgba(255,255,255,0.15)"}}>{fc}: {plt}PLT→{tr.t10>0?`10T×${tr.t10}`:''}{tr.t4>0?` 4T×${tr.t4}`:''}{tr.t2>0?` 2T×${tr.t2}`:''}</span>
                          })}
                          <span style={{marginLeft:"auto",fontWeight:700,color:"#fbbf24"}}>¥{bl.subtotal.toLocaleString()}</span>
                        </div>
                        <table style={{width:"100%",borderCollapse:"collapse",fontSize:12}}>
                          <thead><tr>{['項目 항목','件 건','単位','単価','金額','税','TOTAL金額'].map(h=><th key={h} style={{...TH,textAlign:['件 건','単価','金額','TOTAL金額'].includes(h)?'right':'left'}}>{h}</th>)}</tr></thead>
                          <tbody>{bl.rows.map((r,i)=>(
                            <tr key={i} style={{background:i%2===0?"transparent":"var(--color-background-secondary)"}}>
                              <td style={{...TD,minWidth:300}}>{r.item}</td>
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
                            <td style={{...TD,color:CT[ci],fontWeight:700,fontSize:13,textAlign:"right"}}>¥{bl.subtotal.toLocaleString()}</td>
                          </tr></tfoot>
                        </table>
                      </div>
                    )
                  })}
                  {/* 전체 합계 */}
                  <div style={{background:"#1e293b",color:"#fff",padding:"10px 16px",borderRadius:"var(--border-radius-md)",display:"flex",justifyContent:"space-between",alignItems:"center"}}>
                    <span style={{fontWeight:600}}>전체 합계 ({blIds.length}BL)</span>
                    <span style={{fontWeight:800,fontSize:16,color:"#fbbf24"}}>¥{grandTotal.toLocaleString()}</span>
                  </div>
                  <div style={{marginTop:6,fontSize:10,color:"var(--color-text-tertiary)"}}>
                    ※ 입출고료 각 ¥500/PLT · 트럭: 10T=14PLT, 4T=6PLT(장거리2T불가), 2T=2PLT(오사카권만) · AN費用·실제 배송단가는 견적서 확인 후 수정
                  </div>
                </div>
              )
            })()}
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
    </div>
  )
}
