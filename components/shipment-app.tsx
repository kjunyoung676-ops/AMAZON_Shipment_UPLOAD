diff --git a/components/shipment-app.tsx b/components/shipment-app.tsx
index 5495ea4709ee6790921b872a83f0839efd86710f..17f4ad1767a9daf49a8ce6255de1b1222c66d507 100644
--- a/components/shipment-app.tsx
+++ b/components/shipment-app.tsx
@@ -52,50 +52,51 @@ const CB = ["#dbeafe","#d1fae5","#fef3c7","#fce7f3","#ede9fe","#ffedd5","#e0f2fe
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
+  note?: string
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
@@ -191,178 +192,242 @@ export default function ShipmentApp() {
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
 
+  function cleanSkuCandidate(s: string): string {
+    return s.replace(/\\n/g, '').replace(/\n/g, '').replace(/\s+/g, '').trim()
+  }
+
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
+      const injectSyntheticLabels = (target: { sku: string; y: number; x: number }[], skus: string[]) => {
+        skus.forEach((sku, idx) => target.push({ sku, y: 10000 - idx * 40, x: idx % 2 === 0 ? 0 : 1000 }))
+      }
+
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
+          const tokenTexts = items.map((it)=>String(it.str || ""))
+          const fullTextSpaced = tokenTexts.join(" ")
+          const fullTextNoSpace = tokenTexts.join("").replace(/\s+/g, "")
 
           // 각 라벨의 SKU 추출
-          // 「単一のSKU」 뒤 3토큰 안에서 SKU처럼 보이는 문자열 찾기
+          // 1) 토큰 기반 추출
+          // 2) 실패 시 전체 텍스트 정규식 fallback
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
-                  sku = s
+                  // SKU가 토큰 분리되는 경우(HK124...-5W- / 1Pack)를 합쳐서 시도
+                  const s2 = ((items[k + 1]?.str as string) || '').trim()
+                  const s3 = ((items[k + 2]?.str as string) || '').trim()
+                  const merged = cleanSkuCandidate([s, s2, s3].join(''))
+                  sku = merged.length >= s.length ? merged : cleanSkuCandidate(s)
                   break
                 }
               }
               if (sku) labelInfos.push({ sku, y, x })
             }
           }
 
+          // fallback: 텍스트 전체에서 「単一のSKU <SKU> 数量」 패턴 추출
           if (labelInfos.length === 0) {
+            const re = /単一のSKU\s*([A-Za-z0-9.\-_/]+?)\s*(?:数量|JAN|FBA|$)/g
+            const fallbackSkus: string[] = []
+            let m: RegExpExecArray | null
+            while ((m = re.exec(fullTextSpaced)) !== null) {
+              const s = (m[1] || "").trim()
+              const cleaned = cleanSkuCandidate(s)
+              if (cleaned.length > 4) fallbackSkus.push(cleaned)
+            }
+            injectSyntheticLabels(labelInfos, fallbackSkus)
+          }
+
+          // fallback2: marker가 전혀 없어도 페이지 텍스트에서 마스터 realSku를 직접 탐색
+          if (labelInfos.length === 0) {
+            const matchedMasterSkus: string[] = []
+            for (const m of Object.values(master)) {
+              if (!m.sku) continue
+              const norm = normalizeSku(m.sku)
+              if (norm.length < 6) continue
+              if (fullTextNoSpace.toUpperCase().includes(norm)) matchedMasterSkus.push(m.sku)
+            }
+            injectSyntheticLabels(labelInfos, [...new Set(matchedMasterSkus)])
+          }
+
+          // fallback3: marker/마스터 직접탐색 모두 실패하면 일반 SKU 패턴 후보를 추출
+          if (labelInfos.length === 0) {
+            const genericCandidates = (fullTextSpaced.match(/[A-Z0-9]+(?:[-.][A-Z0-9]+){2,}(?:-?1PACK(?:-?JP)?|JAPAN|JP)?/gi) || [])
+              .map(s => s.trim())
+              .filter(s => {
+                const u = s.toUpperCase()
+                if (u.startsWith("FBA15")) return false // FBA 라벨 ID 제외
+                if (u.startsWith("B0")) return false    // ASIN 제외
+                return /[A-Z]/.test(u) && /[-.]/.test(u)
+              })
+            injectSyntheticLabels(labelInfos, [...new Set(genericCandidates)])
+          }
+
+          if (labelInfos.length === 0) {
+            debugLog.push({
+              file: f.name,
+              page: i + 1,
+              skusOnPage: [],
+              skusNorm: [],
+              matchedLoc: null,
+              labelCount: 0,
+              note: `SKU marker/candidate not found (textLen=${fullTextNoSpace.length})`,
+            })
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
+            note: matchedLocs.length ? undefined : "SKU extracted but no master mapping match",
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
