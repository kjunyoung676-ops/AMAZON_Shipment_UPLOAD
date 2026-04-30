"use client"

import { useState, useRef } from "react"
import * as XLSX from "xlsx"

// ── SKU 정렬 순서 ─────────────────────────────────────────
const SKU_ORDER = [
  "HS124018Z-5W","HS124018Z-5B","HS124518W-5B","HS124518W-5W",
  "HS604012-4B","HS604012-4W","HS604015HP-5W","HS604015HPW-5W",
  "HS604015W-5B","HS604015W-5W","HS804016.5HPW-5W","HS804018HP-5W",
  "HS904016.5W-5W","HS904016.5W-5B","HS904018-5B","HS904018-5W",
]

// ── 마스터 데이터 ─────────────────────────────────────────
const MASTER_INIT: Record<string, MasterItem> = {
  "HS604012-4B":      {sku:"HS604012-4B-1Pack-JP",           asin:"B0CTLLCPF8", to:"H8",  loc:"J-07-B", cpp:24, fba:1020, price:12300, kg:15.50, bx:65,   by:41.5, bz:12  },
  "HS604012-4W":      {sku:"HS604012-4W-1Pack",              asin:"B0CQZJCCCB", to:"H6",  loc:"J-07-W", cpp:24, fba:1020, price:12300, kg:15.50, bx:65,   by:41.5, bz:12  },
  "HS904018-5B":      {sku:"HS904018-5B-1Pack-JP",           asin:"B0CTLDTN5V", to:"H9",  loc:"J-09-B", cpp:16, fba:1100, price:16400, kg:24.38, bx:94.5, by:41.5, bz:12.5},
  "HS904018-5W":      {sku:"US-HSBWB6-090-040-180-5S-1Pack", asin:"B09NLJMGDZ", to:"H2",  loc:"J-09-W", cpp:16, fba:1100, price:16400, kg:22.00, bx:94.5, by:41.5, bz:12.5},
  "HS124018Z-5B":     {sku:"HS124018Z-5B-1Pack-JP",          asin:"B0CTM7TKRC", to:"HG1", loc:"J-10-B", cpp:16, fba:1756, price:20300, kg:31.44, bx:125,  by:41.5, bz:12.5},
  "HS124018Z-5W":     {sku:"HK124018Z-5W-1Pack",             asin:"B0BYJ6C1PG", to:"H3",  loc:"J-10-W", cpp:16, fba:1756, price:20300, kg:31.44, bx:125,  by:41.5, bz:12.5},
  "HS604015W-5B":     {sku:"HS604015W-5B-JP",                asin:"B0CTKFC5LR", to:"H4",  loc:"J-12-B", cpp:24, fba:1020, price:14800, kg:12.90, bx:80,   by:41.5, bz:13  },
  "HS604015W-5W":     {sku:"HS604015W-5W-JAPAN",             asin:"B0CWQ8QX1Y", to:"H11", loc:"J-12-W", cpp:24, fba:1100, price:14800, kg:20.00, bx:80,   by:41.5, bz:13  },
  "HS904016.5W-5B":   {sku:"HS9040165W-5B-JP",              asin:"B0CTKDFQ47", to:"H5",  loc:"J-13-B", cpp:16, fba:1532, price:18800, kg:27.20, bx:42,   by:96,   bz:13  },
  "HS904016.5W-5W":   {sku:"HS904016.5W-5W-JAPAN",           asin:"B0CWQF2XC9", to:"H12", loc:"J-13-W", cpp:16, fba:1532, price:18800, kg:27.20, bx:96,   by:42,   bz:13  },
  "HS124518W-5B":     {sku:"HS124518W-5B-JAPAN",             asin:"B0CWQ9X3QB", to:"H10", loc:"J-14-B", cpp:16, fba:1756, price:22300, kg:35.30, bx:125,  by:47,   bz:13.5},
  "HS124518W-5W":     {sku:"HS124518W-5W-JAPAN",             asin:"B0CWQGKPMG", to:"H13", loc:"J-14-W", cpp:16, fba:1756, price:22300, kg:35.30, bx:125,  by:47,   bz:13.5},
  "HS604015HP-5W":    {sku:"HS604015HP-5W",                  asin:"B0GVSSFYRZ", to:"",    loc:"J-15-W", cpp:24, fba:1100, price:16300, kg:22.00, bx:79.8, by:41.8, bz:13.7},
  "HS804018HP-5W":    {sku:"HS804018HP-5W",                  asin:"B0GVSWH65N", to:"",    loc:"J-16-W", cpp:16, fba:1100, price:19400, kg:27.00, bx:94.5, by:42.5, bz:13.7},
  "HS604015HPW-5W":   {sku:"HS604015HPW-5W",                 asin:"B0GVT1R5HY", to:"",    loc:"J-17-W", cpp:18, fba:1532, price:18800, kg:24.00, bx:79.8, by:41.5, bz:16.8},
  "HS804016.5HPW-5W": {sku:"HS804016.5HPW-5W",               asin:"B0GVSPLDG3", to:"",    loc:"J-18-W", cpp:12, fba:1532, price:21800, kg:29.10, bx:94.5, by:42.0, bz:16.8},
}

// ── 컬럼 별칭 ─────────────────────────────────────────────
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

// ── 컬러 팔레트 ───────────────────────────────────────────
const CB = ["#dbeafe","#d1fae5","#fef3c7","#fce7f3","#ede9fe","#ffedd5","#e0f2fe","#dcfce7"]
const CT = ["#1e40af","#065f46","#92400e","#9d174d","#4c1d95","#9a3412","#075985","#166534"]

// ── 인라인 스타일 상수 ────────────────────────────────────
const TH: React.CSSProperties = {padding:"7px 10px",textAlign:"left",fontWeight:500,fontSize:11,color:"var(--color-text-primary)",borderBottom:"1.5px solid var(--color-border-secondary)",whiteSpace:"nowrap",background:"var(--color-background-secondary)",position:"sticky",top:0}
const TD: React.CSSProperties = {padding:"5px 10px",borderBottom:"0.5px solid var(--color-border-tertiary)",fontSize:12,overflow:"hidden",textOverflow:"ellipsis",whiteSpace:"nowrap",color:"var(--color-text-primary)"}
const INP: React.CSSProperties = {border:"none",background:"transparent",fontSize:12,padding:"1px 0",outline:"none",color:"var(--color-text-primary)",fontFamily:"inherit",width:"100%"}
const IBLU: React.CSSProperties = {background:"rgba(219,234,254,0.35)"}
const EMPTY_SKU = {sku:"",realSku:"",asin:"",to:"",loc:"",cpp:16,fba:0,price:0,kg:0,bx:0,by:0,bz:0}

// ── 타입 ─────────────────────────────────────────────────
interface MasterItem {
  sku: string; asin: string; to: string; loc: string; cpp: number;
  fba: number; price: number; kg: number; bx: number; by: number; bz: number;
}
interface RowData {
  [key: string]: unknown;
  __groupStart?: boolean;
  quantity: number;
  container_no?: number;
}

// ── 계산 ─────────────────────────────────────────────────
const calcGW  = (qty: number, nw: number) => Math.trunc(qty * nw * 1.01 * 10) / 10
const calcCBM = (qty: number, bx: number, by: number, bz: number) => Math.round(bx * by * bz / 1000000 * qty * 100) / 100

// ── 날짜 포맷 ─────────────────────────────────────────────
function fmtDate(v: unknown): unknown {
  if (v instanceof Date && !isNaN(v.getTime()))
    return String(v.getMonth()+1).padStart(2,'0') + "월 " + String(v.getDate()).padStart(2,'0') + "일"
  if (typeof v === 'number' && v > 40000 && v < 60000) {
    const d = new Date(Date.UTC(1899,11,30) + v*86400000)
    if (!isNaN(d.getTime())) return String(d.getUTCMonth()+1).padStart(2,'0') + "월 " + String(d.getUTCDate()).padStart(2,'0') + "일"
  }
  return v
}

// ── 헤더 정규화 ───────────────────────────────────────────
function normHeaders(rows: RowData[]) {
  if (!rows.length) return {data:[]}
  const lm: Record<string,string> = {}
  Object.keys(rows[0]).forEach(k => { lm[k.trim().toLowerCase()] = k })
  const rn: Record<string,string> = {}
  for (const [can, als] of Object.entries(ALIASES))
    for (const a of als)
      if (a.toLowerCase() in lm) { rn[lm[a.toLowerCase()]] = can; break }
  return {
    data: rows.map(r => {
      const nr: RowData = {quantity:0}
      for (const [k,v] of Object.entries(r)) nr[rn[k]||k] = v
      nr.__groupStart = r.__groupStart || false
      nr.quantity = parseFloat(String(nr.quantity)) || 0
      return nr
    })
  }
}

// ── 시트 파싱 ─────────────────────────────────────────────
function parseSheet(ws: XLSX.WorkSheet): RowData[] {
  if (!ws['!ref']) return []
  const merges = ws['!merges'] || []
  const range = XLSX.utils.decode_range(ws['!ref'])
  const hr = range.s.r
  const gsr = new Set<number>()
  for (const m of merges)
    if (m.s.c <= 2 && m.e.r > m.s.r && m.s.r > hr) gsr.add(m.s.r)
  const raw = XLSX.utils.sheet_to_json(ws, {defval:"", cellDates:true}) as Record<string,unknown>[]
  return raw.map((r, i) => {
    const nr: RowData = {quantity:0}
    for (const [k,v] of Object.entries(r)) nr[k] = fmtDate(v)
    nr.__groupStart = gsr.has(i + hr + 1)
    return nr
  })
}

// ── forward-fill ──────────────────────────────────────────
function forwardFill(rows: RowData[]): RowData[] {
  const res = rows.map(r => ({...r}))
  const last: Record<string,unknown> = {}
  let ctnNo = 0
  for (const row of res) {
    if (row.__groupStart) ctnNo++
    if (ctnNo === 0) ctnNo = 1
    for (const col of FF_COLS) {
      const v = row[col]
      if (v !== undefined && v !== "" && v !== null) last[col] = v
      else if (col in last) row[col] = last[col]
    }
    row.container_no = ctnNo
  }
  console.log("[v0] forwardFill 결과: 총", res.length, "행, 컨번호별:", res.reduce((acc,r) => { const k=String(r.container_no); acc[k]=(acc[k]||0)+1; return acc }, {} as Record<string,number>))
  return res
}

// ── xlsx 다운로드 ─────────────────────────────────────────
function xlsDl(data: unknown[], sn: string, fn: string) {
  const wb = XLSX.utils.book_new()
  const ws = Array.isArray(data[0]) ? XLSX.utils.aoa_to_sheet(data as unknown[][]) : XLSX.utils.json_to_sheet(data as Record<string,unknown>[])
  XLSX.utils.book_append_sheet(wb, ws, sn||"data")
  XLSX.writeFile(wb, fn)
}

// ══════════════════════════════════════════════════════════
export default function ShipmentApp() {
  const [sheets, setSh]      = useState<Record<string, RowData[]>>({})
  const [sheetNames, setSn]  = useState<string[]>([])
  const [activeSheet, setAs] = useState<string|null>(null)
  const [file, setFile]      = useState<File|null>(null)
  const [mode, setMode]      = useState('1')
  const [master, setMaster]  = useState<Record<string, MasterItem>>(() => {
    const m = Object.fromEntries(Object.entries(MASTER_INIT).map(([k,v]) => [k,{...v}]))
    console.log("[v0] master init 첫번째 항목:", Object.entries(m)[0])
    return m
  })
  const [s2meta, setS2meta]    = useState<Record<string, Record<string,string>>>({})
  const [ctnMeta, setCtnMeta]  = useState<Record<number, Record<string,string>>>({})
  const [coll, setColl]        = useState<Record<string, boolean>>({})
  const [newSku, setNewSku]    = useState({...EMPTY_SKU})
  const [fbaLoading, setFbaLoading] = useState(false)
  const fileRef = useRef<HTMLInputElement>(null)

  // ── 파생 데이터 ────────────────────────────────────────
  const raw = (activeSheet && sheets[activeSheet]) || []
  const {data: nd} = raw.length ? normHeaders(raw) : {data:[] as RowData[]}
  const fd = nd.length ? forwardFill(nd) : []
  const ctnNums = [...new Set(fd.map(r => r.container_no as number))].sort((a,b) => a-b)

  // ── 파일 로드 ──────────────────────────────────────────
  function loadFile(f: File) {
    setFile(f)
    const rd = new FileReader()
    rd.onload = e => {
      const wb = XLSX.read((e.target as FileReader).result, {type:'array'})
      const ns: Record<string, RowData[]> = {}
      for (const sn of wb.SheetNames) ns[sn] = parseSheet(wb.Sheets[sn])
      setSh(ns); setSn(wb.SheetNames); setAs(wb.SheetNames[0])
      setMode('1'); setS2meta({}); setCtnMeta({})
    }
    rd.readAsArrayBuffer(f)
  }

  // ── 1차 가공 행 ────────────────────────────────────────
  function buildS2rows() {
    const agg: Record<string, {sku:string, total:number}> = {}
    for (const r of fd) {
      const sk = String(r.sku || ""); if (!sk) continue
      if (!agg[sk]) agg[sk] = {sku:sk, total:0}
      agg[sk].total += r.quantity
    }
    return Object.values(agg).sort((a,b) => {
      const ai = SKU_ORDER.indexOf(a.sku), bi = SKU_ORDER.indexOf(b.sku)
      if (ai === -1 && bi === -1) return a.sku.localeCompare(b.sku)
      if (ai === -1) return 1; if (bi === -1) return -1
      return ai - bi
    }).map(({sku, total}) => {
      const m = master[sku] || ({} as MasterItem)
      const cpp = parseFloat(String(m.cpp)) || 16
      const meta = s2meta[sku] || {}
      return {
        sku, total, cpp,
        asin:    m.asin    || "",
        to:      m.to      || "",
        loc:     m.loc     || "",
        pallets: Math.ceil(total / cpp),
        pages:   parseFloat((total/6).toFixed(2)),
        gw:      calcGW(total, parseFloat(String(m.kg))||0),
        cbm:     calcCBM(total, parseFloat(String(m.bx))||0, parseFloat(String(m.by))||0, parseFloat(String(m.bz))||0),
        fc:      meta.fc       || "",
        address: meta.address  || "",
        fbaId:   meta.fbaId    || "",
        amazonId:meta.amazonId || "",
      }
    })
  }

  // ── 2차 가공 그룹 ──────────────────────────────────────
  function buildS3groups() {
    return ctnNums.map(no => {
      const rows = fd.filter(r => r.container_no === no)
      const f0 = rows[0] || {}
      const cm = ctnMeta[no] || {}
      return {
        no,
        container:     cm.container     || "",
        sealNo:        cm.sealNo        || "",
        shipment_date: String(f0.shipment_date || ""),
        shipment_time: String(f0.shipment_time || ""),
        destination:   String(f0.destination   || ""),
        etd: String(f0.etd || ""),
        eta: String(f0.eta || ""),
        rows,
      }
    })
  }

  // ── 헬퍼 ──────────────────────────────────────────────
  const updM       = (sku: string, f: string, v: string|number) => setMaster(p => ({...p, [sku]:{...(p[sku]||{}), [f]:v} as MasterItem}))
  const updMeta    = (sku: string, f: string, v: string) => setS2meta(p => ({...p, [sku]:{...(p[sku]||{}), [f]:v}}))
  const updCtnMeta = (no: number, f: string, v: string)  => setCtnMeta(p => ({...p, [no]:{...(p[no]||{}), [f]:v}}))
  const togCtn     = (k: string) => setColl(c => ({...c, [k]:!c[k]}))
  const isOpen     = (k: string) => !coll[k]

  function tag(l: string, v: unknown) {
    if (!v || v==="" || v===0 || v==="0") return null
    return <span style={{fontSize:12,opacity:0.85}}><span style={{opacity:0.6,marginRight:2}}>{l}</span>{String(v)}</span>
  }

  // ── 익스포트 ───────────���──────────────────────────────
  function expS1() {
    xlsDl(fd.map(r => ({
      시트:activeSheet, 컨:"컨"+r.container_no,
      출항일:r.shipment_date||"", 목적지:r.destination||"",
      BOM품번:r.model_code||"", 품목:r.sku||"",
      색상:r.color||"", 수량:r.quantity,
      TYPE구:r.type_old||"", TYPE신:r.location||"",
      ETD:r.etd||"", ETA:r.eta||""
    })), activeSheet||"S1", "1단계_"+(activeSheet||"data")+".xlsx")
  }

  function expS2() {
    const rows = buildS2rows()
    const hdr = [["약호","ASIN","구박스","신박스","카톤","PLT당카톤","팔레트","G.W(kg)","CBM","FC CENTER","주소","FBA ID","아마존 ID"]]
    const body = rows.map(r => [r.sku,r.asin,r.to,r.loc,r.total,r.cpp,r.pallets,r.gw,r.cbm,r.fc,r.address,r.fbaId,r.amazonId])
    const tC=rows.reduce((s,r)=>s+r.total,0), tP=rows.reduce((s,r)=>s+r.pallets,0)
    const tW=rows.reduce((s,r)=>s+r.gw,0), tB=Math.round(rows.reduce((s,r)=>s+r.cbm,0)*100)/100
    xlsDl([...hdr,...body,["합계","","","",tC,"",tP,tW,tB]], activeSheet||"S2", "1차가공_"+(activeSheet||"data")+".xlsx")
  }

  function expS3() {
    const hdr = [["CONTAINER","SEAL NO.","약호","BOX CODE","ASIN","FBA ID","아마존 ID","FC CENTER","주소","PLT","CT","CTN/PLT","G.W(kg)","CBM"]]
    const body: unknown[][] = []
    for (const g of buildS3groups()) {
      let first = true
      for (const r of g.rows) {
        const m = master[String(r.sku)] || ({} as MasterItem)
        const cpp = parseFloat(String(m.cpp))||16
        const sk = s2meta[String(r.sku)] || {}
        body.push([
          first?g.container:"", first?g.sealNo:"",
          r.sku, r.location||m.loc||"", m.asin||"",
          sk.fbaId||"", sk.amazonId||"", sk.fc||"", first?(sk.address||""):"",
          Math.ceil(r.quantity/cpp), r.quantity, cpp+"CTN/PLT",
          calcGW(r.quantity, parseFloat(String(m.kg))||0),
          calcCBM(r.quantity, parseFloat(String(m.bx))||0, parseFloat(String(m.by))||0, parseFloat(String(m.bz))||0)
        ])
        first = false
      }
    }
    xlsDl([...hdr,...body], activeSheet||"S3", "2차가공_"+(activeSheet||"data")+".xlsx")
  }

 async function expFbaUpload() {
  setFbaLoading(true)
  try {
    const res = await fetch("/UPLOAD_FORMAT.xlsx")  // ← GitHub Pages → 로컬 public/
    if (!res.ok) throw new Error("파일 로드 실패 (" + res.status + ")")
    const buf = await res.arrayBuffer()
    const wb = XLSX.read(buf, {
      type: "array",
      cellStyles: true,
      cellNF: true,
      cellDates: true,
      sheetStubs: true,
    })
    const templateSheetName = wb.SheetNames.find(n => n.toLowerCase().includes("template"))
    if (!templateSheetName) throw new Error("template 시트를 찾을 수 없습니다. 시트 목록: " + wb.SheetNames.join(", "))
    const ws = wb.Sheets[templateSheetName]
    const rows = buildS2rows()

    rows.forEach((r, i) => {
      const row = 9 + i
      const realSku = master[r.sku]?.sku || r.sku
      ws["A" + row] = { ...(ws["A" + row] || {}), t: "s", v: realSku, w: realSku }
      ws["B" + row] = { ...(ws["B" + row] || {}), t: "n", v: r.total, w: String(r.total) }
    })

    // ← 핵심 수정: !ref 범위를 실제 데이터 행 수만큼 확장
    if (rows.length > 0) {
      const lastDataRow = 9 + rows.length - 1  // 0-indexed 아님, xlsx 행번호
      const currentRef = ws["!ref"] || "A1:B8"
      const decoded = XLSX.utils.decode_range(currentRef)
      decoded.e.r = Math.max(decoded.e.r, lastDataRow - 1)  // decode_range는 0-indexed
      ws["!ref"] = XLSX.utils.encode_range(decoded)
    }

    XLSX.writeFile(wb, "UPLOAD_FORMAT_filled.xlsx", {
      cellStyles: true,
      compression: true,
    })
  } catch (e) {
    alert("오류: " + (e as Error).message)
  } finally {
    setFbaLoading(false)
  }
}

  // ── 시트 탭 ────────────────────────────────────────────
  function SheetTabs() {
    if (!sheetNames.length) return null
    return (
      <div style={{display:"flex",gap:2,borderBottom:"0.5px solid var(--color-border-tertiary)",marginBottom:16}}>
        {sheetNames.map(sn => {
          const {data:d} = (sheets[sn]||[]).length ? normHeaders(sheets[sn]) : {data:[] as RowData[]}
          const ff2 = d.length ? forwardFill(d) : []
          const cc = new Set(ff2.map(r => r.container_no)).size
          return (
            <button key={sn} onClick={() => setAs(sn)} style={{
              padding:"6px 14px", fontSize:12, fontWeight:activeSheet===sn?500:400,
              cursor:"pointer", background:"transparent", border:"none",
              color:activeSheet===sn?"var(--color-text-primary)":"var(--color-text-secondary)",
              borderBottom:activeSheet===sn?"2px solid var(--color-text-primary)":"2px solid transparent",
              marginBottom:-1, display:"flex", alignItems:"center", gap:5
            }}>
              {sn}
              {cc > 0 && <span style={{fontSize:10,padding:"1px 5px",borderRadius:10,background:"var(--color-background-secondary)",border:"0.5px solid var(--color-border-tertiary)",color:"var(--color-text-tertiary)"}}>{cc}컨</span>}
            </button>
          )
        })}
      </div>
    )
  }

  const s2rows = buildS2rows()
  const totC   = s2rows.reduce((s,r) => s+r.total,   0)
  const totP   = s2rows.reduce((s,r) => s+r.pallets,  0)
  const totW   = s2rows.reduce((s,r) => s+r.gw,       0)
  const totCBM = Math.round(s2rows.reduce((s,r) => s+r.cbm, 0)*100)/100

  function NavBtn({m, label}: {m:string, label:string}) {
    const active = mode===m, avail = m==='1'||m==='map'||!!file
    return (
      <button onClick={() => avail && setMode(m)} style={{
        padding:"8px 16px", fontSize:12, fontWeight:active?500:400,
        cursor:avail?"pointer":"default", border:"none",
        background:active?"var(--color-text-primary)":"transparent",
        color:active?"var(--color-background-primary)":avail?"var(--color-text-secondary)":"var(--color-text-tertiary)",
      }}>{label}</button>
    )
  }

  // ══════════════════════════════════════════════════════
  return (
    <div style={{padding:"0.75rem 0",fontFamily:"var(--font-sans)"}}>

      {/* 내비 */}
      <div style={{display:"flex",alignItems:"center",gap:8,marginBottom:16}}>
        <div style={{border:"0.5px solid var(--color-border-tertiary)",borderRadius:"var(--border-radius-md)",overflow:"hidden",display:"flex"}}>
          <NavBtn m="1" label="① 업로드" />
          <div style={{width:"0.5px",background:"var(--color-border-tertiary)"}}/>
          <NavBtn m="2" label="② 1차 가공" />
          <div style={{width:"0.5px",background:"var(--color-border-tertiary)"}}/>
          <NavBtn m="3" label="③ 2차 가공" />
        </div>
        <button onClick={() => setMode('map')} style={{marginLeft:"auto",fontSize:11,padding:"6px 14px",background:mode==='map'?"var(--color-background-secondary)":"transparent",border:"0.5px solid var(--color-border-tertiary)",borderRadius:"var(--border-radius-md)",cursor:"pointer",color:mode==='map'?"var(--color-text-primary)":"var(--color-text-secondary)"}}>
          매핑 관리
        </button>
      </div>

      {/* ══ STEP 1 ══ */}
      {mode==='1' && (
        <div>
          <div style={{display:"flex",alignItems:"center",gap:8,marginBottom:14,flexWrap:"wrap"}}>
            <div
              onClick={() => fileRef.current?.click()}
              onDragOver={e => e.preventDefault()}
              onDrop={e => {e.preventDefault(); const f=e.dataTransfer.files[0]; if(f) loadFile(f)}}
              style={{display:"flex",alignItems:"center",gap:8,padding:"6px 14px",border:"1px dashed var(--color-border-secondary)",borderRadius:"var(--border-radius-md)",cursor:"pointer",fontSize:13}}
            >
              <span>+</span>
              <span style={{color:file?"var(--color-text-primary)":"var(--color-text-tertiary)"}}>{file?file.name:"파일 업로드 (xlsx · csv)"}</span>
            </div>
            <input ref={fileRef} type="file" accept=".xlsx,.xls,.csv" style={{display:"none"}} onChange={e => e.target.files?.[0] && loadFile(e.target.files[0])} />
            {file && <>
              <span style={{fontSize:11,padding:"3px 9px",borderRadius:20,background:"var(--color-background-secondary)",border:"0.5px solid var(--color-border-tertiary)",color:"var(--color-text-secondary)"}}>
                {raw.length}행 · {ctnNums.length}컨 · {fd.reduce((s,r)=>s+r.quantity,0).toLocaleString()}개
              </span>
              <button onClick={expS1} style={{fontSize:11,padding:"3px 10px"}}>xlsx 저장</button>
              <button onClick={() => setMode('2')} style={{marginLeft:"auto",fontSize:12,padding:"5px 14px",background:"var(--color-background-info)",border:"0.5px solid var(--color-border-info)",color:"var(--color-text-info)",borderRadius:"var(--border-radius-md)",cursor:"pointer",fontWeight:500}}>
                1차 가공으로 →
              </button>
            </>}
          </div>
          <SheetTabs />
          {!file ? (
            <div style={{textAlign:"center",padding:"5rem 0",color:"var(--color-text-tertiary)"}}>
              <div style={{fontSize:48,marginBottom:12}}>📦</div>
              <p style={{fontSize:15}}>쉽먼트 파일을 업로드해주세요</p>
              <p style={{fontSize:12,marginTop:4}}>시트가 여러 개면 탭으로 구분됩니다</p>
            </div>
          ) : (
            <div style={{display:"flex",flexDirection:"column",gap:10}}>
              {ctnNums.map(no => {
                const rows = fd.filter(r => r.container_no===no)
                const f0=rows[0], ci=(no-1)%CB.length
                const sumQ=rows.reduce((s,r)=>s+r.quantity,0)
                const dQ=parseFloat(String(f0.qty_total))||sumQ
                const key="s1_"+no, open=isOpen(key)
                return (
                  <div key={no} style={{border:"0.5px solid var(--color-border-tertiary)",borderRadius:"var(--border-radius-md)",overflow:"hidden"}}>
                    <div onClick={() => togCtn(key)} style={{background:CB[ci],color:CT[ci],padding:"9px 14px",display:"flex",alignItems:"center",gap:12,flexWrap:"wrap",cursor:"pointer",userSelect:"none"}}>
                      <span style={{fontWeight:500,fontSize:13}}>컨{no}</span>
                      <span style={{fontWeight:500}}>{String(f0.shipment_date||"")}</span>
                      {tag("",f0.shipment_time)}
                      <span style={{fontWeight:500}}>{String(f0.destination||"")}</span>
                      {tag("ETD ",f0.etd)}{tag("ETA ",f0.eta)}{tag("선사 ",f0.carrier)}
                      <span style={{marginLeft:"auto",display:"flex",alignItems:"center",gap:8}}>
                        <span style={{fontWeight:500,fontSize:13}}>
                          {dQ.toLocaleString()}개{f0.ctn_count&&f0.ctn_count!=="0"?" · "+f0.ctn_count+"컨":""}{f0.ft&&f0.ft!=="0"?" · "+f0.ft+"FT":""}
                        </span>
                        <span style={{fontSize:11,opacity:0.6}}>{open?"▲":"▼"}</span>
                      </span>
                    </div>
                    {open && (
                      <div style={{overflowX:"auto"}}>
                        <table style={{width:"100%",borderCollapse:"collapse",fontSize:12,minWidth:600}}>
                          <thead><tr>{["BOM품번","약호(SKU)","색상","수량","TYPE(구)","신박스코드","ETD","ETA"].map(h=><th key={h} style={TH}>{h}</th>)}</tr></thead>
                          <tbody>
                            {rows.map((r,i) => (
                              <tr key={i} style={{background:i%2===0?"transparent":"var(--color-background-secondary)"}}>
                                <td style={{...TD,fontFamily:"var(--font-mono)",fontSize:11,color:"var(--color-text-secondary)",minWidth:130}}>{String(r.model_code||"")}</td>
                                <td style={{...TD,fontWeight:500,minWidth:120}}>{String(r.sku||"")}</td>
                                <td style={{...TD,minWidth:50}}>{String(r.color||"")}</td>
                                <td style={{...TD,textAlign:"right",fontWeight:500,minWidth:50}}>{r.quantity.toLocaleString()}</td>
                                <td style={{...TD,color:"var(--color-text-secondary)"}}>{String(r.type_old||"")}</td>
                                <td style={{...TD,color:"var(--color-text-info)",fontWeight:500}}>{String(r.location||"")}</td>
                                <td style={TD}>{String(r.etd||"")}</td>
                                <td style={TD}>{String(r.eta||"")}</td>
                              </tr>
                            ))}
                          </tbody>
                          <tfoot>
                            <tr style={{background:"var(--color-background-secondary)",borderTop:"1px solid var(--color-border-secondary)"}}>
                              <td colSpan={3} style={{...TD,textAlign:"right",fontSize:11,color:"var(--color-text-secondary)"}}>소계</td>
                              <td style={{...TD,textAlign:"right",fontWeight:500}}>{sumQ.toLocaleString()}</td>
                              <td colSpan={4} style={TD}></td>
                            </tr>
                          </tfoot>
                        </table>
                      </div>
                    )}
                  </div>
                )
              })}
            </div>
          )}
        </div>
      )}

      {/* ══ STEP 2 ══ */}
      {mode==='2' && (
        <div>
          <div style={{display:"flex",alignItems:"center",gap:8,marginBottom:8,flexWrap:"wrap"}}>
            <span style={{fontSize:12,color:"var(--color-text-secondary)"}}>전체 합산 · <span style={{color:"var(--color-text-info)"}}>파란 셀</span> = 직접 입력</span>
            <button onClick={expS2} style={{marginLeft:"auto",fontSize:11,padding:"3px 10px"}}>xlsx 저장</button>
            <button
              onClick={expFbaUpload}
              disabled={fbaLoading}
              style={{fontSize:11,padding:"3px 10px",background:"var(--color-background-success)",border:"0.5px solid var(--color-border-success)",color:"var(--color-text-success)",borderRadius:"var(--border-radius-md)",cursor:fbaLoading?"default":"pointer",opacity:fbaLoading?0.6:1}}
            >
              {fbaLoading ? "처리 중..." : "FBA 업로드 양식"}
            </button>
            <button onClick={() => setMode('3')} style={{fontSize:12,padding:"5px 14px",background:"var(--color-background-success)",border:"0.5px solid var(--color-border-success)",color:"var(--color-text-success)",borderRadius:"var(--border-radius-md)",cursor:"pointer",fontWeight:500}}>
              2차 가공으로 →
            </button>
          </div>
          <SheetTabs />
          <div style={{overflowX:"auto",border:"0.5px solid var(--color-border-tertiary)",borderRadius:"var(--border-radius-md)"}}>
            <table style={{width:"100%",borderCollapse:"collapse",fontSize:12,minWidth:960}}>
              <thead>
                <tr>
                  <th colSpan={9} style={{...TH,borderRight:"2px solid var(--color-border-secondary)"}}>자동 계산</th>
                  <th colSpan={4} style={{...TH,...IBLU,color:"var(--color-text-info)"}}>직접 입력 (변동값)</th>
                </tr>
                <tr>
                  {["약호","ASIN","구박스","신박스","카톤","PLT당카톤","팔레트","G.W(kg)","CBM"].map(h => <th key={h} style={TH}>{h}</th>)}
                  {["FC CENTER","주소","FBA ID","아마존 ID"].map(h => <th key={h} style={{...TH,...IBLU}}>{h}</th>)}
                </tr>
              </thead>
              <tbody>
                {s2rows.map((r,i) => (
                  <tr key={r.sku} style={{background:i%2===0?"transparent":"var(--color-background-secondary)"}}>
                    <td style={{...TD,fontWeight:500,minWidth:130}}>{r.sku}</td>
                    <td style={{...TD,fontFamily:"var(--font-mono)",fontSize:11,minWidth:100}}>{r.asin}</td>
                    <td style={{...TD,minWidth:45}}>{r.to}</td>
                    <td style={{...TD,color:"var(--color-text-info)",fontWeight:500,minWidth:70}}>{r.loc}</td>
                    <td style={{...TD,textAlign:"right",fontWeight:500,minWidth:55}}>{r.total.toLocaleString()}</td>
                    <td style={{...TD,textAlign:"center",minWidth:55}}>{r.cpp}</td>
                    <td style={{...TD,textAlign:"right",fontWeight:500,minWidth:50}}>{r.pallets}</td>
                    <td style={{...TD,textAlign:"right",minWidth:70,borderRight:"1px solid var(--color-border-tertiary)"}}>{r.gw.toFixed(1)}</td>
                    <td style={{...TD,textAlign:"right",minWidth:60,borderRight:"2px solid var(--color-border-secondary)"}}>{r.cbm}</td>
                    {(["fc","address","fbaId","amazonId"] as const).map((f, idx) => {
                      const placeholders = ["FC CENTER","주소","FBA ID","아마존 ID"]
                      return (
                        <td key={f} style={{...TD,minWidth:[80,150,110,90][idx],...IBLU}}>
                          <input value={r[f]} onChange={e => updMeta(r.sku,f,e.target.value)} style={INP} placeholder={placeholders[idx]} />
                        </td>
                      )
                    })}
                  </tr>
                ))}
              </tbody>
              <tfoot>
                <tr style={{background:"var(--color-background-secondary)",borderTop:"2px solid var(--color-border-secondary)"}}>
                  <td colSpan={4} style={{...TD,fontWeight:500,textAlign:"right"}}>합계</td>
                  <td style={{...TD,fontWeight:500,textAlign:"right"}}>{totC.toLocaleString()}</td>
                  <td style={TD}></td>
                  <td style={{...TD,fontWeight:500,textAlign:"right"}}>{totP}</td>
                  <td style={{...TD,fontWeight:500,textAlign:"right"}}>{totW.toFixed(1)}</td>
                  <td style={{...TD,fontWeight:500,textAlign:"right"}}>{totCBM}</td>
                  <td colSpan={4} style={TD}></td>
                </tr>
              </tfoot>
            </table>
          </div>
        </div>
      )}

      {/* ══ STEP 3 ══ */}
      {mode==='3' && (
        <div>
          <div style={{display:"flex",alignItems:"center",gap:8,marginBottom:14}}>
            <span style={{fontSize:12,color:"var(--color-text-secondary)"}}>업로드 컨테이너별 분류 · CONTAINER/SEAL 직접 입력</span>
            <button onClick={expS3} style={{marginLeft:"auto",fontSize:11,padding:"3px 10px"}}>xlsx 저장</button>
          </div>
          <SheetTabs />
          {!file ? (
            <p style={{color:"var(--color-text-tertiary)",fontSize:13,textAlign:"center",padding:"3rem 0"}}>파일을 먼저 업로드해주세요</p>
          ) : (
            <div style={{overflowX:"auto",border:"0.5px solid var(--color-border-tertiary)",borderRadius:"var(--border-radius-md)"}}>
              <table style={{width:"100%",borderCollapse:"collapse",fontSize:12,minWidth:950}}>
                <thead>
                  <tr>
                    <th style={{...TH,minWidth:100}}>CONTAINER</th>
                    <th style={{...TH,minWidth:90,borderRight:"2px solid var(--color-border-secondary)"}}>SEAL NO.</th>
                    {["약호","BOX CODE","ASIN","FBA ID","아마존 ID","FC CENTER","PLT","CT","CTN/PLT","G.W(kg)","CBM"].map(h => (
                      <th key={h} style={TH}>{h}</th>
                    ))}
                  </tr>
                </thead>
                <tbody>
                  {buildS3groups().map(g => {
                    const rows = g.rows
                    const ci   = (g.no-1) % CB.length
                    const sumQ = rows.reduce((s,r) => s+r.quantity, 0)
                    const sumP = rows.reduce((s,r) => { const cpp=parseFloat(String((master[String(r.sku)]||{}).cpp))||16; return s+Math.ceil(r.quantity/cpp) }, 0)
                    const sumW = rows.reduce((s,r) => s+calcGW(r.quantity, parseFloat(String((master[String(r.sku)]||{}).kg))||0), 0)
                    const sumCBM = Math.round(rows.reduce((s,r) => {
                      const m=master[String(r.sku)]||({} as MasterItem)
                      return s+calcCBM(r.quantity, parseFloat(String(m.bx))||0, parseFloat(String(m.by))||0, parseFloat(String(m.bz))||0)
                    }, 0)*100)/100

                    return [
                      ...rows.map((r,i) => {
                        const m   = master[String(r.sku)] || ({} as MasterItem)
                        const cpp = parseFloat(String(m.cpp))||16
                        const sk  = s2meta[String(r.sku)] || {}
                        const gw  = calcGW(r.quantity, parseFloat(String(m.kg))||0)
                        const cbm = calcCBM(r.quantity, parseFloat(String(m.bx))||0, parseFloat(String(m.by))||0, parseFloat(String(m.bz))||0)
                        return (
                          <tr key={g.no+"_"+i} style={{background:i%2===0?"transparent":"var(--color-background-secondary)"}}>
                            {i===0 && (
                              <>
                                <td rowSpan={rows.length+1} style={{...TD,verticalAlign:"middle",textAlign:"center",background:CB[ci],padding:"6px 8px",borderRight:"0.5px solid var(--color-border-tertiary)"}}>
                                  <div style={{fontSize:10,color:CT[ci],opacity:0.7,marginBottom:3}}>컨{g.no}</div>
                                  <input value={ctnMeta[g.no]?.container||""} onChange={e => updCtnMeta(g.no,"container",e.target.value)} placeholder="CONTAINER"
                                    style={{...INP,textAlign:"center",fontWeight:500,fontSize:11,color:CT[ci],borderBottom:"1px solid "+CT[ci]+"66",width:88}} />
                                </td>
                                <td rowSpan={rows.length+1} style={{...TD,verticalAlign:"middle",textAlign:"center",background:CB[ci],padding:"6px 8px",borderRight:"2px solid var(--color-border-secondary)"}}>
                                  <input value={ctnMeta[g.no]?.sealNo||""} onChange={e => updCtnMeta(g.no,"sealNo",e.target.value)} placeholder="SEAL NO."
                                    style={{...INP,textAlign:"center",fontSize:11,color:CT[ci],borderBottom:"1px solid "+CT[ci]+"66",width:78}} />
                                </td>
                              </>
                            )}
                            <td style={{...TD,fontWeight:500,minWidth:130}}>{String(r.sku)}</td>
                            <td style={{...TD,color:"var(--color-text-info)",fontWeight:500,minWidth:70}}>{String(r.location||m.loc||"")}</td>
                            <td style={{...TD,fontFamily:"var(--font-mono)",fontSize:11,minWidth:100}}>{m.asin||""}</td>
                            <td style={{...TD,fontFamily:"var(--font-mono)",fontSize:11,minWidth:110}}>{sk.fbaId||""}</td>
                            <td style={{...TD,fontFamily:"var(--font-mono)",fontSize:11,minWidth:90}}>{sk.amazonId||""}</td>
                            <td style={{...TD,fontWeight:500,minWidth:70}}>{sk.fc||""}</td>
                            <td style={{...TD,textAlign:"right",minWidth:40}}>{Math.ceil(r.quantity/cpp)}</td>
                            <td style={{...TD,textAlign:"right",fontWeight:500,minWidth:40}}>{r.quantity.toLocaleString()}</td>
                            <td style={{...TD,textAlign:"center",color:"var(--color-text-secondary)"}}>{cpp}CTN/PLT</td>
                            <td style={{...TD,textAlign:"right",color:"var(--color-text-secondary)"}}>{gw.toFixed(1)}</td>
                            <td style={{...TD,textAlign:"right",color:"var(--color-text-secondary)"}}>{cbm}</td>
                          </tr>
                        )
                      }),
                      <tr key={"sub_"+g.no} style={{background:CB[ci],borderTop:"1px solid "+CT[ci]+"44"}}>
                        <td colSpan={6} style={{...TD,color:CT[ci],fontWeight:500,textAlign:"right",fontSize:11}}>
                          {(g.container||"컨"+g.no)+" 소계"}
                          {g.shipment_date?" · "+g.shipment_date:""}
                          {g.destination?"  "+g.destination:""}
                          {g.etd?"  ETD "+g.etd:""}
                        </td>
                        <td style={{...TD,color:CT[ci],fontWeight:500,textAlign:"right"}}>{sumP}</td>
                        <td style={{...TD,color:CT[ci],fontWeight:500,textAlign:"right"}}>{sumQ.toLocaleString()}</td>
                        <td style={TD}></td>
                        <td style={{...TD,color:CT[ci],fontWeight:500,textAlign:"right"}}>{sumW.toFixed(1)}</td>
                        <td style={{...TD,color:CT[ci],fontWeight:500,textAlign:"right"}}>{sumCBM}</td>
                      </tr>
                    ]
                  })}
                  {(() => {
                    const gC   = fd.reduce((s,r) => s+r.quantity, 0)
                    const gP   = fd.reduce((s,r) => { const cpp=parseFloat(String((master[String(r.sku)]||{}).cpp))||16; return s+Math.ceil(r.quantity/cpp) }, 0)
                    const gW   = fd.reduce((s,r) => s+calcGW(r.quantity, parseFloat(String((master[String(r.sku)]||{}).kg))||0), 0)
                    const gCBM = Math.round(fd.reduce((s,r) => { const m=master[String(r.sku)]||({} as MasterItem); return s+calcCBM(r.quantity,parseFloat(String(m.bx))||0,parseFloat(String(m.by))||0,parseFloat(String(m.bz))||0) },0)*100)/100
                    return (
                      <tr style={{background:"var(--color-background-secondary)",borderTop:"2px solid var(--color-border-secondary)"}}>
                        <td colSpan={8} style={{...TD,fontWeight:500,textAlign:"right"}}>전체 합계</td>
                        <td style={{...TD,fontWeight:500,textAlign:"right"}}>{gP}</td>
                        <td style={{...TD,fontWeight:500,textAlign:"right"}}>{gC.toLocaleString()}</td>
                        <td style={TD}></td>
                        <td style={{...TD,fontWeight:500,textAlign:"right"}}>{gW.toFixed(1)}</td>
                        <td style={{...TD,fontWeight:500,textAlign:"right"}}>{gCBM}</td>
                      </tr>
                    )
                  })()}
                </tbody>
              </table>
            </div>
          )}
        </div>
      )}

      {/* ══ 매핑 관리 ══ */}
      {mode==='map' && (
        <div>
          <div style={{display:"flex",alignItems:"center",gap:8,marginBottom:14}}>
            <span style={{fontSize:13,fontWeight:500}}>SKU 마스터</span>
            <span style={{fontSize:12,color:"var(--color-text-tertiary)"}}>{Object.keys(master).length}개</span>
            <button onClick={() => {
              if (!newSku.sku) return
              setMaster(p => ({...p, [newSku.sku]:{sku:newSku.realSku||newSku.sku,asin:newSku.asin,to:newSku.to,loc:newSku.loc,cpp:+newSku.cpp,fba:+newSku.fba,price:+newSku.price,kg:+newSku.kg,bx:+newSku.bx,by:+newSku.by,bz:+newSku.bz}}))
              setNewSku({...EMPTY_SKU})
            }} style={{marginLeft:"auto",fontSize:11,padding:"3px 10px"}}>+ 추가</button>
            <button onClick={() => xlsDl(
              Object.entries(master).map(([sku,m]) => ({약호:sku,SKU:m.sku,ASIN:m.asin,구박스:m.to,신박스:m.loc,PLT당카톤:m.cpp,FBA비용:m.fba,판매가:m.price,무게kg:m.kg,박스가로:m.bx,박스세로:m.by,박스높이:m.bz})),
              "마스터","SKU마스터.xlsx"
            )} style={{fontSize:11,padding:"3px 10px"}}>xlsx 저장</button>
          </div>
          <div style={{overflowX:"auto",border:"0.5px solid var(--color-border-tertiary)",borderRadius:"var(--border-radius-md)"}}>
            <table style={{width:"100%",borderCollapse:"collapse",fontSize:12}}>
              <thead>
                <tr>{["약호","SKU","ASIN","구박스","신박스","PLT당카톤","FBA비용(¥)","판매가(¥)","무게(kg)","박스가로","박스세로","박스높이",""].map(h => <th key={h} style={TH}>{h}</th>)}</tr>
              </thead>
              <tbody>
                {Object.entries(master).map(([sku,m],i) => (
                  <tr key={sku} style={{background:i%2===0?"transparent":"var(--color-background-secondary)"}}>
                    <td style={{...TD,fontWeight:500,minWidth:130}}>{sku}</td>
                    <td style={{...TD,...IBLU,minWidth:160}}>
                      <input
                        value={m.sku ?? ""}
                        onChange={e => updM(sku,"sku",e.target.value)}
                        style={{...INP, color: m.sku ? "var(--color-text-primary)" : "var(--color-text-secondary)"}}
                        placeholder="Merchant SKU"
                      />
                    </td>
                    {(["asin","to","loc"] as const).map((f) => (
                      <td key={f} style={{...TD,...IBLU}}>
                        <input value={m[f]||""} onChange={e => updM(sku,f,e.target.value)} style={INP} />
                      </td>
                    ))}
                    {(["cpp","fba","price","kg","bx","by","bz"] as const).map((f) => (
                      <td key={f} style={{...TD,...IBLU}}>
                        <input type="number" value={m[f]||0} onChange={e => updM(sku,f,Number(e.target.value))} style={{...INP,textAlign:"right"}} />
                      </td>
                    ))}
                    <td style={TD}>
                      <button onClick={() => setMaster(p => { const n={...p}; delete n[sku]; return n })} style={{fontSize:11,padding:"1px 6px",cursor:"pointer",borderRadius:4,background:"transparent",color:"var(--color-text-danger)",border:"0.5px solid var(--color-border-danger)"}}>삭제</button>
                    </td>
                  </tr>
                ))}

              </tbody>
            </table>
          </div>
        </div>
      )}

    </div>
  )
}
