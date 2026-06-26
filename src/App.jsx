import { useState, useRef, useCallback } from "react";

const LOCATIONS = ["Grapevine", "Denton", "McKinney"];

// ── API call via Vercel proxy ─────────────────────────────────────────────────
async function callAPI(b64) {
  const prompt = `You are extracting purchase order data from a THML clothing sales order PDF.

RULES:
- Vendor is always "THML"
- Style = the style code after "Style:" (e.g. "JH2424-3") — strip the word "Style:" and the # symbol
- Color = exactly as written in the PDF (e.g. "PK", "CR", "BR", "T", "N", "BL", "GN", "VE", "ST", "BK") — do NOT translate or expand abbreviations
- Sizes: XS, S, M, L, XL, 2XL — use 0 if not mentioned for that size
- Extract ALL style/color rows across ALL shipping months (IMMEDIATE, JUL, AUG, SEP, OCT, etc.)
- Location is determined by the "Ship to" address in the header:
  * "Grapevine" or "421 S Main" or "76051" → location: "Grapevine"
  * "McKinney" or "MC KINNEY" or "206 East Louisiana" or "75069" → location: "McKinney"
  * "Denton" or "119 N Elm" or "76201" → location: "Denton"

Return ONLY valid JSON, no markdown, no explanation, no code fences:
{
  "location": "Grapevine",
  "items": [
    { "style": "JH2424-3", "color": "PK", "xs": 1, "s": 2, "m": 2, "l": 1, "xl": 0, "xxl": 0 }
  ]
}`;

  const controller = new AbortController();
  const timeout = setTimeout(() => controller.abort(), 55000);

  let resp;
  try {
    resp = await fetch("/api/extract", {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      signal: controller.signal,
      body: JSON.stringify({
        model: "claude-sonnet-4-6",
        max_tokens: 8000,
        messages: [{
          role: "user",
          content: [
            { type: "document", source: { type: "base64", media_type: "application/pdf", data: b64 } },
            { type: "text", text: prompt }
          ]
        }]
      })
    });
  } catch (err) {
    clearTimeout(timeout);
    throw new Error(`Network error: ${err.message}`);
  }
  clearTimeout(timeout);

  let data;
  try { data = await resp.json(); }
  catch (e) { throw new Error(`Server error (status ${resp.status})`); }
  if (data.error) throw new Error(data.error.message || JSON.stringify(data.error));
  if (!data.content) throw new Error(`Unexpected response: ${JSON.stringify(data).slice(0, 200)}`);

  const raw = (data.content || []).map(c => c.text || "").join("").trim();
  const clean = raw.replace(/```json|```/g, "").trim();
  const start = clean.indexOf("{");
  const end = clean.lastIndexOf("}");
  if (start === -1) throw new Error(`No JSON found. Got: ${raw.slice(0, 300)}`);
  return JSON.parse(clean.substring(start, end + 1));
}

// ── Chunk large PDFs ─────────────────────────────────────────────────────────
async function splitPDFToChunks(file, chunkSize = 3) {
  const { PDFDocument } = await import("https://cdn.jsdelivr.net/npm/pdf-lib@1.17.1/+esm");
  const arrayBuffer = await file.arrayBuffer();
  const srcDoc = await PDFDocument.load(arrayBuffer);
  const totalPages = srcDoc.getPageCount();
  const chunks = [];
  for (let start = 0; start < totalPages; start += chunkSize) {
    const end = Math.min(start + chunkSize, totalPages);
    const chunkDoc = await PDFDocument.create();
    const pages = await chunkDoc.copyPages(srcDoc, Array.from({ length: end - start }, (_, i) => start + i));
    pages.forEach(p => chunkDoc.addPage(p));
    const bytes = await chunkDoc.save();
    const b64 = btoa(new Uint8Array(bytes).reduce((d, byte) => d + String.fromCharCode(byte), ""));
    chunks.push({ b64, pages: `${start + 1}-${end}` });
  }
  return { chunks, totalPages };
}

async function toBase64(file) {
  return new Promise((res, rej) => {
    const r = new FileReader();
    r.onload = () => res(r.result.split(",")[1]);
    r.onerror = () => rej(new Error("File read failed"));
    r.readAsDataURL(file);
  });
}

// ── XLSX builder (PO Entry layout) ───────────────────────────────────────────
async function buildAndDownloadXLSX(allData) {
  // allData: { Grapevine: [{style,color,xs,s,m,l,xl,xxl},...], Denton:[...], McKinney:[...] }
  const XLSX = await import("https://cdn.jsdelivr.net/npm/xlsx@0.18.5/+esm");

  // Collect unique style+color keys in location order
  const seen = new Set();
  const keys = [];
  LOCATIONS.forEach(loc => {
    (allData[loc] || []).forEach(item => {
      const k = `${item.style}||${item.color}`;
      if (!seen.has(k)) { seen.add(k); keys.push({ style: item.style, color: item.color }); }
    });
  });

  function getSizes(loc, style, color) {
    const found = (allData[loc] || []).find(i => i.style === style && i.color === color);
    if (!found) return [0, 0, 0, 0, 0, 0];
    return [found.xs || 0, found.s || 0, found.m || 0, found.l || 0, found.xl || 0, found.xxl || 0];
  }

  const wsData = [];

  // Row 1: Title
  wsData.push(["PALM TREE  |  PO ENTRY – SIZING BY LOCATION", ...Array(29).fill("")]);

  // Row 2: Meta
  wsData.push(["PO / Invoice #:", "", "", "Season:", "", "", "Date Entered:", "", "", "", "Entered by:", ...Array(19).fill("")]);

  // Row 3: Group headers
  wsData.push([
    "Style #", "Color / Description",
    "Grapevine", "", "", "", "", "", "Loc Total",
    "Denton", "", "", "", "", "", "Loc Total",
    "Mc Kinney", "", "", "", "", "", "Loc Total",
    "XS", "S", "M", "L", "XL", "2XL",
    "GRAND TOTAL"
  ]);

  // Row 4: Size headers
  wsData.push([
    "Style #", "Color",
    "XS", "S", "M", "L", "XL", "2XL", "Loc Total",
    "XS", "S", "M", "L", "XL", "2XL", "Loc Total",
    "XS", "S", "M", "L", "XL", "2XL", "Loc Total",
    "XS", "S", "M", "L", "XL", "2XL",
    "All Locs"
  ]);

  // Data rows
  keys.forEach(({ style, color }) => {
    const gv = getSizes("Grapevine", style, color);
    const dn = getSizes("Denton", style, color);
    const mk = getSizes("McKinney", style, color);
    const gvTot = gv.reduce((a, b) => a + b, 0);
    const dnTot = dn.reduce((a, b) => a + b, 0);
    const mkTot = mk.reduce((a, b) => a + b, 0);
    const sizeTots = [0, 1, 2, 3, 4, 5].map(i => gv[i] + dn[i] + mk[i]);
    const grand = gvTot + dnTot + mkTot;

    wsData.push([
      style, color,
      gv[0] || "", gv[1] || "", gv[2] || "", gv[3] || "", gv[4] || "", gv[5] || "", gvTot || "",
      dn[0] || "", dn[1] || "", dn[2] || "", dn[3] || "", dn[4] || "", dn[5] || "", dnTot || "",
      mk[0] || "", mk[1] || "", mk[2] || "", mk[3] || "", mk[4] || "", mk[5] || "", mkTot || "",
      sizeTots[0] || "", sizeTots[1] || "", sizeTots[2] || "", sizeTots[3] || "", sizeTots[4] || "", sizeTots[5] || "",
      grand || ""
    ]);
  });

  // Totals row
  const gvSizeTots = [0,1,2,3,4,5].map(i => keys.reduce((a, { style, color }) => a + getSizes("Grapevine", style, color)[i], 0));
  const dnSizeTots = [0,1,2,3,4,5].map(i => keys.reduce((a, { style, color }) => a + getSizes("Denton", style, color)[i], 0));
  const mkSizeTots = [0,1,2,3,4,5].map(i => keys.reduce((a, { style, color }) => a + getSizes("McKinney", style, color)[i], 0));
  const allSizeTots = [0,1,2,3,4,5].map(i => gvSizeTots[i] + dnSizeTots[i] + mkSizeTots[i]);
  const allGvTot = gvSizeTots.reduce((a, b) => a + b, 0);
  const allDnTot = dnSizeTots.reduce((a, b) => a + b, 0);
  const allMkTot = mkSizeTots.reduce((a, b) => a + b, 0);

  wsData.push([
    "TOTALS", "",
    ...gvSizeTots, allGvTot,
    ...dnSizeTots, allDnTot,
    ...mkSizeTots, allMkTot,
    ...allSizeTots,
    allGvTot + allDnTot + allMkTot
  ]);

  const ws = XLSX.utils.aoa_to_sheet(wsData);

  // Merges
  ws["!merges"] = [
    { s: { r: 0, c: 0 }, e: { r: 0, c: 29 } },  // Title
    { s: { r: 2, c: 2 }, e: { r: 2, c: 8 } },    // Grapevine
    { s: { r: 2, c: 9 }, e: { r: 2, c: 15 } },   // Denton
    { s: { r: 2, c: 16 }, e: { r: 2, c: 22 } },  // McKinney
    { s: { r: 2, c: 23 }, e: { r: 2, c: 28 } },  // Size totals
  ];

  // Column widths
  ws["!cols"] = [
    { wch: 14 }, { wch: 9 },
    { wch: 5 }, { wch: 5 }, { wch: 5 }, { wch: 5 }, { wch: 5 }, { wch: 5 }, { wch: 8 },
    { wch: 5 }, { wch: 5 }, { wch: 5 }, { wch: 5 }, { wch: 5 }, { wch: 5 }, { wch: 8 },
    { wch: 5 }, { wch: 5 }, { wch: 5 }, { wch: 5 }, { wch: 5 }, { wch: 5 }, { wch: 8 },
    { wch: 5 }, { wch: 5 }, { wch: 5 }, { wch: 5 }, { wch: 5 }, { wch: 5 },
    { wch: 10 }
  ];

  const wb = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(wb, ws, "PO Entry");

  const buf = XLSX.write(wb, { type: "array", bookType: "xlsx" });
  const blob = new Blob([new Uint8Array(buf)], { type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet" });
  const url = URL.createObjectURL(blob);
  const a = document.createElement("a");
  a.href = url;
  a.download = "PalmTree_PO_Entry.xlsx";
  document.body.appendChild(a);
  a.click();
  document.body.removeChild(a);
  URL.revokeObjectURL(url);
}

// ── Main App ──────────────────────────────────────────────────────────────────
export default function App() {
  const [files, setFiles] = useState([null, null, null]);
  const [assigned, setAssigned] = useState([null, null, null]);
  const [status, setStatus] = useState("");
  const [processing, setProcessing] = useState(false);
  const [summary, setSummary] = useState(null); // [{loc, count}]
  const [dragging, setDragging] = useState(false);
  const inputRefs = [useRef(), useRef(), useRef()];

  const usedLocations = assigned.filter(Boolean);
  const availableFor = idx => LOCATIONS.filter(l => !usedLocations.includes(l) || assigned[idx] === l);

  function handleFile(idx, file) {
    if (!file || file.type !== "application/pdf") return;
    const newFiles = [...files];
    newFiles[idx] = file;
    setFiles(newFiles);
  }

  function handleDrop(e) {
    e.preventDefault();
    setDragging(false);
    const dropped = Array.from(e.dataTransfer.files).filter(f => f.type === "application/pdf");
    if (!dropped.length) return;
    const newFiles = [...files];
    dropped.forEach(f => {
      const emptySlot = newFiles.findIndex(s => s === null);
      if (emptySlot !== -1) newFiles[emptySlot] = f;
    });
    setFiles(newFiles);
  }

  function removeFile(idx) {
    const newFiles = [...files];
    const newAssigned = [...assigned];
    newFiles[idx] = null;
    newAssigned[idx] = null;
    setFiles(newFiles);
    setAssigned(newAssigned);
  }

  function setLocation(idx, loc) {
    const newAssigned = [...assigned];
    newAssigned[idx] = loc || null;
    setAssigned(newAssigned);
  }

  const canProcess = files.every(f => f !== null)
    && assigned.every(a => a !== null)
    && new Set(assigned).size === 3;

  async function processAll() {
    setProcessing(true);
    setSummary(null);
    setStatus("");
    const allData = {};

    for (let i = 0; i < 3; i++) {
      const loc = assigned[i];
      const file = files[i];
      try {
        // Check page count
        const { PDFDocument } = await import("https://cdn.jsdelivr.net/npm/pdf-lib@1.17.1/+esm");
        const srcDoc = await PDFDocument.load(await file.arrayBuffer());
        const totalPages = srcDoc.getPageCount();

        let result;
        if (totalPages <= 4) {
          setStatus(`Extracting ${loc} (${i + 1}/3)…`);
          const b64 = await toBase64(file);
          result = await callAPI(b64);
          allData[loc] = result.items || [];
        } else {
          setStatus(`${loc}: splitting ${totalPages} pages into chunks…`);
          const { chunks } = await splitPDFToChunks(file, 3);
          const mergedItems = {};
          for (let c = 0; c < chunks.length; c++) {
            setStatus(`${loc} (${i + 1}/3): chunk ${c + 1}/${chunks.length} (pages ${chunks[c].pages})…`);
            try {
              const chunkResult = await callAPI(chunks[c].b64);
              (chunkResult.items || []).forEach(item => {
                const key = `${item.style}||${item.color}`;
                if (!mergedItems[key]) {
                  mergedItems[key] = { ...item };
                } else {
                  // Sum sizes if same style/color appears in multiple chunks
                  mergedItems[key].xs  = (mergedItems[key].xs  || 0) + (item.xs  || 0);
                  mergedItems[key].s   = (mergedItems[key].s   || 0) + (item.s   || 0);
                  mergedItems[key].m   = (mergedItems[key].m   || 0) + (item.m   || 0);
                  mergedItems[key].l   = (mergedItems[key].l   || 0) + (item.l   || 0);
                  mergedItems[key].xl  = (mergedItems[key].xl  || 0) + (item.xl  || 0);
                  mergedItems[key].xxl = (mergedItems[key].xxl || 0) + (item.xxl || 0);
                }
              });
            } catch (chunkErr) {
              setStatus(`⚠️ ${loc} chunk ${chunks[c].pages} failed: ${chunkErr.message}`);
              await new Promise(r => setTimeout(r, 1000));
            }
          }
          allData[loc] = Object.values(mergedItems);
        }
      } catch (err) {
        setStatus(`❌ Error on ${loc}: ${err.message}`);
        setProcessing(false);
        return;
      }
    }

    setStatus("Building XLSX…");
    try {
      await buildAndDownloadXLSX(allData);
      setSummary(LOCATIONS.map(loc => ({ loc, count: (allData[loc] || []).length })));
      setStatus("✅ Done! PalmTree_PO_Entry.xlsx downloaded.");
    } catch (err) {
      setStatus(`❌ XLSX error: ${err.message}`);
    }
    setProcessing(false);
  }

  return (
    <div style={s.root}>
      <style>{`@keyframes spin { to { transform: rotate(360deg); } }`}</style>

      {/* Header */}
      <div style={s.header}>
        <div style={s.headerLeft}>
          <div style={s.logoBlock}><span style={s.logoEmoji}>🌴</span></div>
          <div>
            <div style={s.logoTitle}>The Palm Tree</div>
            <div style={s.logoSub}>PO Extractor · Grapevine · Denton · McKinney</div>
          </div>
        </div>
        <div style={s.headerBadge}>THML · PO Entry XLSX</div>
      </div>

      <div style={s.main}>
        <div style={s.hero}>
          <h1 style={s.title}>Extract purchase orders<br /><em style={s.titleEm}>across all locations</em></h1>
          <p style={s.desc}>Upload one THML PDF per location, assign each to its store, then extract. Output is a single PO Entry XLSX matching the Palm Tree format.</p>
          <div style={s.locationPills}>
            {LOCATIONS.map(loc => (
              <div key={loc} style={s.locationPill}><span style={s.locationDot} />{loc}</div>
            ))}
          </div>
        </div>

        {/* Drop zone */}
        <div
          style={{ ...s.dropZone, ...(dragging ? s.dropZoneActive : {}) }}
          onDragOver={e => { e.preventDefault(); setDragging(true); }}
          onDragLeave={() => setDragging(false)}
          onDrop={handleDrop}
        >
          <div style={s.dropIcon}>📄</div>
          <p style={s.dropTitle}>Drag & drop up to 3 PDFs here</p>
          <p style={s.dropSub}>or use the upload buttons below — one PDF per location</p>
        </div>

        {/* File slots */}
        <div style={s.slots}>
          {[0, 1, 2].map(idx => (
            <div key={idx} style={s.slot}>
              <div style={s.slotHeader}>
                <span style={s.slotNum}>PDF {idx + 1}</span>
                {files[idx] && (
                  <button style={s.removeBtn} onClick={() => removeFile(idx)}>✕</button>
                )}
              </div>

              {!files[idx] ? (
                <div style={s.uploadArea} onClick={() => inputRefs[idx].current.click()}>
                  <div style={s.uploadPlus}>+</div>
                  <div style={s.uploadText}>Click to upload PDF</div>
                  <input
                    ref={inputRefs[idx]}
                    type="file"
                    accept=".pdf"
                    style={{ display: "none" }}
                    onChange={e => handleFile(idx, e.target.files[0])}
                  />
                </div>
              ) : (
                <div style={s.fileInfo}>
                  <div style={s.fileEmoji}>📄</div>
                  <div style={s.fileName}>{files[idx].name}</div>
                  <div style={s.fileSize}>{(files[idx].size / 1024).toFixed(1)} KB</div>
                </div>
              )}

              <select
                style={{ ...s.select, ...(assigned[idx] ? s.selectActive : {}) }}
                value={assigned[idx] || ""}
                onChange={e => setLocation(idx, e.target.value)}
                disabled={!files[idx]}
              >
                <option value="">— Assign location —</option>
                {availableFor(idx).map(loc => (
                  <option key={loc} value={loc}>{loc}</option>
                ))}
              </select>
            </div>
          ))}
        </div>

        {/* Hint */}
        {files.some(f => f) && !canProcess && (
          <div style={s.hint}>
            {files.filter(f => f).length < 3
              ? `Upload ${3 - files.filter(f => f).length} more PDF(s)`
              : "Assign all 3 PDFs to different locations"}
          </div>
        )}

        {/* Extract button */}
        <button
          style={{ ...s.extractBtn, ...(!canProcess || processing ? s.extractBtnDisabled : {}) }}
          onClick={processAll}
          disabled={!canProcess || processing}
        >
          {processing
            ? <span style={s.spinRow}><span style={s.spinner} /> Extracting…</span>
            : "⚡ Extract & Download PO Entry XLSX"}
        </button>

        {/* Status */}
        {status && (
          <p style={{
            ...s.statusText,
            ...(status.startsWith("✅") ? s.statusSuccess : status.startsWith("❌") ? s.statusError : {})
          }}>
            {status}
          </p>
        )}

        {/* Summary */}
        {summary && (
          <div style={s.summaryCard}>
            <div style={s.summaryTitle}>Extraction Summary</div>
            <div style={s.summaryGrid}>
              {summary.map(({ loc, count }) => (
                <div key={loc} style={s.summaryItem}>
                  <div style={s.summaryLoc}>{loc}</div>
                  <div style={s.summaryCount}>{count}</div>
                  <div style={s.summaryLabel}>style/color rows</div>
                </div>
              ))}
            </div>
          </div>
        )}
      </div>
    </div>
  );
}

// ── Styles ─────────────────────────────────────────────────────────────────────
const s = {
  root: { fontFamily: "'DM Mono', monospace", background: "#fdf8f0", minHeight: "100vh" },
  header: { background: "#2c1a0e", padding: "1.2rem 2rem", display: "flex", alignItems: "center", justifyContent: "space-between", position: "sticky", top: 0, zIndex: 100 },
  headerLeft: { display: "flex", alignItems: "center", gap: 14 },
  logoBlock: { width: 36, height: 36, background: "#8b5e3c", borderRadius: 8, display: "flex", alignItems: "center", justifyContent: "center" },
  logoEmoji: { fontSize: 20 },
  logoTitle: { fontFamily: "serif", fontSize: "1.05rem", color: "#fdf8f0", fontWeight: 600 },
  logoSub: { fontSize: "0.62rem", color: "rgba(253,248,240,0.4)", letterSpacing: "0.08em", textTransform: "uppercase" },
  headerBadge: { fontSize: "0.6rem", letterSpacing: "0.1em", textTransform: "uppercase", border: "1px solid rgba(255,255,255,0.12)", padding: "4px 12px", borderRadius: 20, color: "rgba(253,248,240,0.5)" },
  main: { maxWidth: 900, margin: "0 auto", padding: "3rem 1.5rem 5rem" },
  hero: { marginBottom: "2.5rem" },
  title: { fontFamily: "serif", fontSize: "2.2rem", fontWeight: 400, color: "#2c1a0e", lineHeight: 1.25, marginBottom: "0.75rem", letterSpacing: "-0.02em" },
  titleEm: { fontStyle: "italic", color: "#8b5e3c" },
  desc: { fontSize: "0.78rem", color: "#8a7a6a", letterSpacing: "0.02em", lineHeight: 1.6, marginBottom: "1rem" },
  locationPills: { display: "flex", gap: 8, flexWrap: "wrap" },
  locationPill: { display: "flex", alignItems: "center", gap: 6, background: "#fff", border: "1px solid rgba(139,94,60,0.2)", borderRadius: 20, padding: "4px 12px", fontSize: "0.72rem", color: "#8b5e3c", fontWeight: 500 },
  locationDot: { width: 6, height: 6, borderRadius: "50%", background: "#8b5e3c" },
  dropZone: { border: "1.5px dashed rgba(44,26,14,0.2)", borderRadius: 20, padding: "2.5rem 2rem", textAlign: "center", background: "#f5ede0", transition: "all 0.2s", marginBottom: "1.5rem" },
  dropZoneActive: { background: "#ecdfc8", borderColor: "#8b5e3c" },
  dropIcon: { fontSize: 32, marginBottom: 10 },
  dropTitle: { fontSize: "0.88rem", fontWeight: 500, color: "#2c1a0e", marginBottom: 4 },
  dropSub: { fontSize: "0.7rem", color: "#8a7a6a" },
  slots: { display: "grid", gridTemplateColumns: "repeat(3,1fr)", gap: 16, marginBottom: 16 },
  slot: { background: "#fff", borderRadius: 12, padding: 16, boxShadow: "0 1px 4px rgba(0,0,0,.07)", display: "flex", flexDirection: "column", gap: 12 },
  slotHeader: { display: "flex", justifyContent: "space-between", alignItems: "center" },
  slotNum: { fontWeight: 700, fontSize: 13, color: "#2c1a0e" },
  removeBtn: { background: "none", border: "none", color: "#c62828", cursor: "pointer", fontSize: 13, fontWeight: 700 },
  uploadArea: { border: "2px dashed #d7c5b0", borderRadius: 8, padding: "20px 10px", textAlign: "center", cursor: "pointer" },
  uploadPlus: { fontSize: 24, color: "#8b5e3c" },
  uploadText: { fontSize: 12, color: "#8a7a6a", marginTop: 4 },
  fileInfo: { background: "#fdf5ec", borderRadius: 8, padding: "10px 12px", display: "flex", flexDirection: "column", gap: 4 },
  fileEmoji: { fontSize: 18 },
  fileName: { fontSize: 11, fontWeight: 600, color: "#2c1a0e", wordBreak: "break-all" },
  fileSize: { fontSize: 10, color: "#8a7a6a" },
  select: { width: "100%", padding: "7px 10px", borderRadius: 7, border: "1px solid #ddd", fontSize: 12, color: "#555", background: "#fafafa", cursor: "pointer", fontFamily: "'DM Mono', monospace" },
  selectActive: { borderColor: "#8b5e3c", color: "#2c1a0e", fontWeight: 600, background: "#fdf5ec" },
  hint: { textAlign: "center", color: "#8b5e3c", fontSize: 12, marginBottom: 12, background: "#fdf5ec", padding: "8px 16px", borderRadius: 8 },
  extractBtn: { display: "block", width: "100%", padding: "14px", background: "#2c1a0e", color: "#fdf8f0", border: "none", borderRadius: 12, fontSize: "0.78rem", fontWeight: 500, letterSpacing: "0.08em", textTransform: "uppercase", cursor: "pointer", fontFamily: "'DM Mono', monospace", marginBottom: 14 },
  extractBtnDisabled: { background: "#d8cfc4", color: "#8a7a6a", cursor: "not-allowed" },
  spinRow: { display: "flex", alignItems: "center", gap: 10, justifyContent: "center" },
  spinner: { display: "inline-block", width: 14, height: 14, border: "2px solid rgba(253,248,240,0.2)", borderTopColor: "#fdf8f0", borderRadius: "50%", animation: "spin 0.7s linear infinite" },
  statusText: { textAlign: "center", fontSize: "0.72rem", color: "#8a7a6a", letterSpacing: "0.04em", background: "#f5ede0", padding: "10px 16px", borderRadius: 8, marginBottom: 12 },
  statusSuccess: { background: "#f0fdf4", color: "#166534" },
  statusError: { background: "#fef2f2", color: "#991b1b" },
  summaryCard: { background: "#fff", borderRadius: 16, padding: 20, boxShadow: "0 1px 4px rgba(0,0,0,.07)", marginTop: 8 },
  summaryTitle: { fontWeight: 700, fontSize: 14, color: "#2c1a0e", marginBottom: 14 },
  summaryGrid: { display: "grid", gridTemplateColumns: "repeat(3,1fr)", gap: 12 },
  summaryItem: { background: "#fdf5ec", borderRadius: 10, padding: "14px", textAlign: "center" },
  summaryLoc: { fontWeight: 700, color: "#2c1a0e", fontSize: 13, marginBottom: 6 },
  summaryCount: { fontSize: 30, fontWeight: 800, color: "#8b5e3c", lineHeight: 1 },
  summaryLabel: { fontSize: 10, color: "#8a7a6a", marginTop: 4 },
};
