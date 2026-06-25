import { useState, useRef, useCallback } from "react";

const LOCATIONS = ["Grapevine", "Denton", "McKinney"];

const LOCATION_KEYWORDS = {
  Grapevine: ["grapevine", "421 s main", "76051"],
  Denton: ["denton"],
  McKinney: ["mckinney", "mc kinney"],
};

function detectLocation(shipTo) {
  if (!shipTo) return "Unknown";
  const lower = shipTo.toLowerCase();
  for (const [loc, keywords] of Object.entries(LOCATION_KEYWORDS)) {
    if (keywords.some(k => lower.includes(k))) return loc;
  }
  return "Unknown";
}

const SYSTEM_PROMPT = `You are an expert PO/invoice extraction AI for a multi-location retail store. The document may contain MULTIPLE ORDERS from MULTIPLE BRANDS/COMPANIES. Read ALL pages before responding.

Return ONLY a single valid JSON array — one object per company/brand per location found. No markdown, no code fences, no explanation.

Each object in the array:
{"company":"","po_number":"","po_date":"","currency":"","ship_to":"","billing_address":"","items":[{"sku":"","description":"","color":"","size":"","quantity":0,"unit_price":0,"msrp":0,"line_total":0}],"subtotal":0,"freight":0,"tax":null,"total_amount":0,"notes":""}

RULES:
1. READ ALL PAGES - collect every line item from every brand.
2. MULTIPLE BRANDS: if the PDF contains orders from multiple brands/companies, create a SEPARATE object for each brand.
3. company: the brand/supplier name. PRIORITY ORDER:
      1st — The large centered LOGO or brand name at the top center of the page.
      2nd — Brand Information section if no clear logo.
      3rd — The document header/title.
      Use the logo/brand name in Title Case (e.g. "Mother", "Simkhai", "Hunter Bell").
4. po_number: use PO# first; fallback to Order#, Invoice#, Reference.
5. ship_to: copy the FULL Ship To address block exactly as it appears. This is critical for location detection.
6. SIZES:
   a. INDIVIDUAL SIZES: Sizes can be letters (XS, S, M, L, XL, XXL) OR numbers (0, 2, 4, 6, 8, 10, 12, 14, 16, 23, 24, 25, 26, 27, 28, etc.). Create ONE row per size.
   b. PRE-PACK (e.g. "1 Prepack of 6 containing 1XS/2S/2M/1L"): EXPAND into individual size rows.
      Example: 1 pack "1XS/2S/2M/1L" -> {size:"XS",qty:1}, {size:"S",qty:2}, {size:"M",qty:2}, {size:"L",qty:1}
      unit_price = pack price / total units in pack.
7. SKU RULES:
   a. If explicit SKU/product code exists -> use it as-is.
   b. If no SKU, use Style # only. NEVER append color. Strip the # symbol.
8. description: full product name/description.
9. color: the color name. If repeated (e.g. "Saddle Saddle"), use it ONCE only.
10. unit_price: wholesale price per piece.
11. msrp: the retail price per item. PRIORITY ORDER:
    1st — MSRP field if present (printed or handwritten).
    2nd — Sugg. Retail or Suggested Retail field if present (printed or handwritten).
    3rd — If neither exists, leave as null.
12. line_total: unit_price x quantity for each row.
13. subtotal, freight, total_amount: from the order totals.
14. dates: YYYY-MM-DD format.
15. numbers: plain numbers only, strip commas and currency symbols.
16. missing fields: null.
17. currency: infer from symbols ($=USD, EUR, GBP).
18. notes: payment terms, sales rep, ship via, delivery window, order type.
CRITICAL: Return a JSON ARRAY even if only one company. Include ALL items from ALL pages. Do not truncate.`;

function fmtMoney(v) {
  if (v == null) return "—";
  const n = parseFloat(v);
  if (isNaN(n)) return "—";
  return n.toLocaleString("en-US", { minimumFractionDigits: 2, maximumFractionDigits: 2 });
}

function toTitleCase(s) {
  if (!s) return "";
  return s.toLowerCase().replace(/\b\w/g, c => c.toUpperCase());
}

function csvCell(v) {
  const s = (v == null ? "" : String(v)).replace(/"/g, '""');
  return s.includes(",") || s.includes('"') || s.includes("\n") ? `"${s}"` : s;
}

function toBase64(file) {
  return new Promise((res, rej) => {
    const r = new FileReader();
    r.onload = () => res(r.result.split(",")[1]);
    r.onerror = () => rej(new Error("Read failed"));
    r.readAsDataURL(file);
  });
}

function downloadFile(content, filename, mime) {
  const bytes = new TextEncoder().encode(content);
  const b64 = btoa(String.fromCharCode.apply(null, bytes));
  const a = document.createElement("a");
  a.href = `data:${mime};base64,${b64}`;
  a.download = filename;
  document.body.appendChild(a);
  a.click();
  document.body.removeChild(a);
}

// Group companies by brand name, merging location quantities
function groupByBrand(companies) {
  const map = {};
  companies.forEach(co => {
    const location = co.assignedLocation || detectLocation(co.ship_to);
    const key = (co.company || "Unknown").toLowerCase().trim();
    if (!map[key]) {
      map[key] = {
        company: co.company,
        po_number: co.po_number,
        po_date: co.po_date,
        currency: co.currency,
        notes: co.notes,
        subtotal: co.subtotal,
        total_amount: co.total_amount,
        locations: {},
        items: {},
      };
    }
    map[key].locations[location] = co;
    // Merge items by sku+size key
    (co.items || []).forEach(item => {
      const itemKey = `${item.sku}|${item.description}|${item.color}|${item.size}`;
      if (!map[key].items[itemKey]) {
        map[key].items[itemKey] = {
          sku: item.sku,
          description: item.description,
          color: item.color,
          size: item.size,
          unit_price: item.unit_price,
          msrp: item.msrp,
          locationQty: {},
        };
      }
      map[key].items[itemKey].locationQty[location] = (map[key].items[itemKey].locationQty[location] || 0) + (item.quantity || 0);
    });
  });
  return Object.values(map);
}

async function buildStockyXLSX(brandGroups) {
  const XLSX = await import("https://cdn.jsdelivr.net/npm/xlsx@0.18.5/+esm");
  const wb = XLSX.utils.book_new();

  brandGroups.forEach(brand => {
    const sheetName = (brand.company || "Unknown").slice(0, 31).replace(/[:\\/?*[\]]/g, "-");
    const headers = ["Variant SKU", "Product Name", "Supplier", ...LOCATIONS, "Total Qty", "Cost Price", "msrp", "color", "size"];
    const rows = [headers];

    Object.values(brand.items).forEach(item => {
      const locQtys = LOCATIONS.map(loc => item.locationQty[loc] || 0);
      const totalQty = locQtys.reduce((a, b) => a + b, 0);
      rows.push([
        item.sku || "",
        toTitleCase(item.description || ""),
        brand.company || "",
        ...locQtys,
        totalQty,
        item.unit_price != null ? Number(item.unit_price).toFixed(2) : "",
        item.msrp != null ? Number(item.msrp).toFixed(2) : "",
        toTitleCase(item.color || ""),
        item.size || "",
      ]);
    });

    const ws = XLSX.utils.aoa_to_sheet(rows);
    // Style header row
    ws["!cols"] = headers.map((h, i) => ({ wch: i < 2 ? 30 : 12 }));
    XLSX.utils.book_append_sheet(wb, ws, sheetName);
  });

  const buf = XLSX.write(wb, { type: "array", bookType: "xlsx" });
  return new Uint8Array(buf);
}

async function buildShopifyXLSX(brandGroups) {
  const XLSX = await import("https://cdn.jsdelivr.net/npm/xlsx@0.18.5/+esm");
  const COLS = [
    "Handle","Title","Body (HTML)","Vendor","Product Category","Type","Tags",
    "Image Src","Published","Option1 Value","Option2 Value",
    "Variant SKU","Variant Price","Cost per item","Variant Grams",
    "Option2 Name","Option1 Name",
    "Variant Inventory Tracker","Variant Inventory Qty","Variant Inventory Policy",
    "Variant Fulfillment Service","Status","Variant Compare At Price",
    "Variant Requires Shipping","Variant Taxable","Variant Barcode",
  ];
  const wb = XLSX.utils.book_new();

  brandGroups.forEach(brand => {
    const sheetName = (brand.company || "Unknown").slice(0, 31).replace(/[:\\/?*[\]]/g, "-");
    const rows = [COLS];

    Object.values(brand.items).forEach((item, idx) => {
      const sku = item.sku || "";
      const desc = item.description || "";
      const color = item.color || "";
      const size = item.size || "";
      const vendor = brand.company || "";
      const price = item.unit_price != null ? Number(item.unit_price).toFixed(2) : "";
      const title = toTitleCase([vendor, desc].filter(Boolean).join(" ").trim());
      const handle = title;
      const totalQty = LOCATIONS.reduce((a, loc) => a + (item.locationQty[loc] || 0), 0);
      const row = {};
      COLS.forEach(c => row[c] = "");
      Object.assign(row, {
        Handle: handle,
        Title: title,
        Vendor: vendor,
        "Option1 Value": size,
        "Option2 Value": color,
        "Option1 Name": "Size",
        "Option2 Name": "Color",
        "Variant SKU": sku,
        "Variant Price": price,
        "Cost per item": price,
        "Variant Inventory Tracker": "Shopify",
        "Variant Inventory Qty": totalQty,
        "Variant Inventory Policy": "deny",
        "Variant Fulfillment Service": "manual",
        Status: "draft",
      });
      rows.push(COLS.map(c => row[c]));
    });

    const ws = XLSX.utils.aoa_to_sheet(rows);
    XLSX.utils.book_append_sheet(wb, ws, sheetName);
  });

  const buf = XLSX.write(wb, { type: "array", bookType: "xlsx" });
  return new Uint8Array(buf);
}

async function downloadXLSX(brandGroups, baseName, type) {
  const bytes = type === "stocky" ? await buildStockyXLSX(brandGroups) : await buildShopifyXLSX(brandGroups);
  const blob = new Blob([bytes], { type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet" });
  const url = URL.createObjectURL(blob);
  const a = document.createElement("a");
  a.href = url;
  a.download = `${type}_${baseName}.xlsx`;
  document.body.appendChild(a);
  a.click();
  document.body.removeChild(a);
  URL.revokeObjectURL(url);
}

function BrandCard({ brand }) {
  const items = Object.values(brand.items);
  const totalUnits = items.reduce((s, item) => s + LOCATIONS.reduce((a, loc) => a + (item.locationQty[loc] || 0), 0), 0);

  return (
    <div style={styles.brandCard}>
      <div style={styles.brandHeader}>
        <div>
          <div style={styles.brandName}>{brand.company || "Unknown Brand"}</div>
          <div style={styles.brandMeta}>
            PO <b style={{ color: "rgba(255,255,255,0.9)" }}>{brand.po_number || "—"}</b>
            &nbsp;·&nbsp; {brand.po_date || "—"}
            &nbsp;·&nbsp; {brand.currency || "USD"}
          </div>
        </div>
        <div style={styles.totalBadge}>{totalUnits} units</div>
      </div>

      <div style={{ overflowX: "auto" }}>
        <table style={styles.table}>
          <thead>
            <tr>
              {["SKU", "Description", "Color", "Size", ...LOCATIONS, "Total", "Cost", "MSRP"].map(h => (
                <th key={h} style={{ ...styles.th, textAlign: ["SKU","Description","Color","Size"].includes(h) ? "left" : "center" }}>{h}</th>
              ))}
            </tr>
          </thead>
          <tbody>
            {items.map((item, i) => {
              const total = LOCATIONS.reduce((a, loc) => a + (item.locationQty[loc] || 0), 0);
              return (
                <tr key={i} style={{ borderBottom: "1px solid rgba(0,0,0,0.06)", background: i % 2 === 0 ? "#fff" : "#fafaf8" }}>
                  <td style={{ ...styles.td, fontFamily: "monospace", fontSize: "0.72rem", color: "#8b5e3c", fontWeight: 600 }}>{item.sku || "—"}</td>
                  <td style={styles.td}>{toTitleCase(item.description || "")}</td>
                  <td style={styles.td}>{toTitleCase(item.color || "—")}</td>
                  <td style={{ ...styles.td, textAlign: "center" }}>{item.size || "—"}</td>
                  {LOCATIONS.map(loc => (
                    <td key={loc} style={{ ...styles.td, textAlign: "center", fontWeight: item.locationQty[loc] ? 600 : 400, color: item.locationQty[loc] ? "#1c2b1e" : "#ccc" }}>
                      {item.locationQty[loc] || 0}
                    </td>
                  ))}
                  <td style={{ ...styles.td, textAlign: "center", fontWeight: 700 }}>{total}</td>
                  <td style={{ ...styles.td, textAlign: "right" }}>{fmtMoney(item.unit_price)}</td>
                  <td style={{ ...styles.td, textAlign: "right" }}>{fmtMoney(item.msrp)}</td>
                </tr>
              );
            })}
          </tbody>
        </table>
      </div>
    </div>
  );
}

function ResultCard({ result }) {
  const { filename, brandGroups } = result;
  const [activeTab, setActiveTab] = useState(0);
  const baseName = filename.replace(/\.pdf$/i, "").replace(/[^a-z0-9]/gi, "_");

  return (
    <div style={styles.resultCard}>
      <div style={styles.resultHeader}>
        <div>
          <div style={styles.filename}>{filename}</div>
          <div style={styles.resultMeta}>
            {brandGroups.length} brand{brandGroups.length !== 1 ? "s" : ""} found
          </div>
        </div>
        <div style={styles.downloadRow}>
          <button style={{ ...styles.dlBtn, ...styles.shopifyBtn }} onClick={() => downloadXLSX(brandGroups, baseName, "shopify")}>↓ Shopify XLSX</button>
          <button style={{ ...styles.dlBtn, ...styles.stockyBtn }} onClick={() => downloadXLSX(brandGroups, baseName, "stocky")}>↓ Stocky XLSX</button>
          <button style={{ ...styles.dlBtn, ...styles.jsonBtn }} onClick={() => downloadFile(JSON.stringify(brandGroups, null, 2), `po_${baseName}.json`, "application/json")}>↓ JSON</button>
        </div>
      </div>

      {brandGroups.length > 1 && (
        <div style={styles.tabBar}>
          {brandGroups.map((b, i) => (
            <button key={i} style={{ ...styles.tab, ...(activeTab === i ? styles.tabActive : {}) }} onClick={() => setActiveTab(i)}>
              {b.company || `Brand ${i + 1}`}
              <span style={styles.tabCount}>{Object.keys(b.items).length}</span>
            </button>
          ))}
        </div>
      )}

      <BrandCard brand={brandGroups[activeTab]} />
    </div>
  );
}

export default function App() {
  const [files, setFiles] = useState([]); // [{file, location}]
  const [dragging, setDragging] = useState(false);
  const [status, setStatus] = useState("");
  const [processing, setProcessing] = useState(false);
  const [results, setResults] = useState([]);
  const inputRef = useRef();

  const addFiles = useCallback((newFiles) => {
    const pdfs = Array.from(newFiles).filter(f => f.type === "application/pdf");
    setFiles(prev => {
      const existing = new Set(prev.map(f => f.file.name));
      const newEntries = pdfs.filter(f => !existing.has(f.name)).map(f => ({ file: f, location: LOCATIONS[0] }));
      return [...prev, ...newEntries];
    });
  }, []);

  const removeFile = (i) => setFiles(prev => prev.filter((_, idx) => idx !== i));
  const setFileLocation = (i, location) => setFiles(prev => prev.map((f, idx) => idx === i ? { ...f, location } : f));

  const handleDrop = (e) => {
    e.preventDefault();
    setDragging(false);
    addFiles(e.dataTransfer.files);
  };

  const callAPI = async (b64) => {
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
          system: SYSTEM_PROMPT,
          messages: [{ role: "user", content: [
            { type: "document", source: { type: "base64", media_type: "application/pdf", data: b64 } },
            { type: "text", text: "Extract all PO data from every brand in this document. Capture the full Ship To address for each order. Return a JSON array." }
          ]}]
        })
      });
    } catch (fetchErr) {
      clearTimeout(timeout);
      const isTimeout = fetchErr.name === "AbortError";
      throw new Error(isTimeout ? `Request timed out after 55s — try a smaller PDF or check your API key` : `Network error: ${fetchErr.message}`);
    }
    clearTimeout(timeout);
    let data;
    try { data = await resp.json(); }
    catch (e) { throw new Error(`Server returned non-JSON (status ${resp.status}) — check Vercel function logs`); }
    if (data.error) throw new Error(`API error: ${data.error.message || JSON.stringify(data.error)}`);
    if (!data.content) throw new Error(`Unexpected API response: ${JSON.stringify(data).slice(0, 300)}`);
    const raw = (data.content || []).map(c => c.text || "").join("").trim();
    if (!raw) throw new Error("API returned empty response — check ANTHROPIC_API_KEY in Vercel settings");
    const clean = raw.replace(/```json|```/g, "").trim();
    const arrStr = clean.substring(clean.indexOf("["), clean.lastIndexOf("]") + 1);
    if (!arrStr) throw new Error(`No JSON array in response. Raw: ${raw.slice(0, 400)}`);
    return JSON.parse(arrStr);
  };

  const splitPDFToChunks = async (file, chunkSize = 3) => {
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
      const b64 = btoa(new Uint8Array(bytes).reduce((data, byte) => data + String.fromCharCode(byte), ""));
      chunks.push({ b64, pages: `${start + 1}-${end}` });
    }
    return chunks;
  };

  const mergeCompanies = (allCompanies) => {
    const map = {};
    allCompanies.forEach(co => {
      const key = `${(co.company || "").toLowerCase()}|${detectLocation(co.ship_to)}`;
      if (!map[key]) {
        map[key] = { ...co, items: [...(co.items || [])] };
      } else {
        map[key].items.push(...(co.items || []));
      }
    });
    return Object.values(map);
  };

  const runExtraction = async () => {
    if (!files.length) return;
    setProcessing(true);

    // Process all files, tag each company with its assigned location
    const allTaggedCompanies = [];
    const errors = [];

    for (let i = 0; i < files.length; i++) {
      const { file, location } = files[i];
      try {
        setStatus(`Processing ${file.name} → ${location} (${i + 1} of ${files.length})…`);
        const b64 = await toBase64(file);
        const { PDFDocument } = await import("https://cdn.jsdelivr.net/npm/pdf-lib@1.17.1/+esm");
        const srcDoc = await PDFDocument.load(await file.arrayBuffer());
        const totalPages = srcDoc.getPageCount();

        // Always process in chunks to avoid timeouts
        setStatus(`${file.name} (${location}): splitting into chunks…`);
        const chunks = await splitPDFToChunks(file, 3);
        const chunkResults = [];
        let lastChunkError = "";
        for (let c = 0; c < chunks.length; c++) {
          setStatus(`${file.name} (${location}): pages ${chunks[c].pages} (${c + 1}/${chunks.length})…`);
          try {
            const cos = await callAPI(chunks[c].b64);
            chunkResults.push(...cos);
          } catch (chunkErr) {
            lastChunkError = chunkErr.message;
            setStatus(`Chunk ${chunks[c].pages} error: ${chunkErr.message.slice(0, 100)}`);
            await new Promise(r => setTimeout(r, 2000)); // show error for 2s
          }
        }
        if (chunkResults.length === 0) throw new Error(lastChunkError || `All ${chunks.length} chunks failed`);
        const companies = chunkResults;

        // Tag each company with the assigned location
        companies.forEach(co => {
          allTaggedCompanies.push({ ...co, assignedLocation: location });
        });
      } catch (err) {
        errors.push({ filename: file.name, location, msg: err.message });
      }
    }

    // Group all companies across all files by brand, with location quantities
    const brandGroups = groupByBrand(allTaggedCompanies);
    const newResults = [];
    if (brandGroups.length > 0) {
      newResults.push({ type: "success", brandGroups, filename: `${files.length} file(s)` });
    }
    errors.forEach(e => newResults.push({ type: "error", filename: `${e.filename} (${e.location})`, msg: e.msg }));

    setResults(newResults);
    setStatus(`Done — ${files.length} file${files.length > 1 ? "s" : ""} processed.`);
    setProcessing(false);
  };

  return (
    <div style={styles.root}>
      <style>{`@keyframes spin { to { transform: rotate(360deg); } }`}</style>

      <div style={styles.header}>
        <div style={styles.headerLeft}>
          <div style={styles.logoBlock}>
            <span style={styles.logoEmoji}>🌴</span>
          </div>
          <div>
            <div style={styles.logoTitle}>The Palm Tree</div>
            <div style={styles.logoSub}>PO Extractor · Grapevine · Denton · McKinney</div>
          </div>
        </div>
        <div style={styles.headerBadge}>Multi-Location · Shopify · Stocky</div>
      </div>

      <div style={styles.main}>
        <div style={styles.hero}>
          <h1 style={styles.title}>Extract purchase orders<br /><em style={styles.titleEm}>across all locations</em></h1>
          <p style={styles.desc}>Upload PDFs — quantities are automatically split by location based on the Ship To address.</p>

          <div style={styles.locationPills}>
            {LOCATIONS.map(loc => (
              <div key={loc} style={styles.locationPill}>
                <span style={styles.locationDot} />
                {loc}
              </div>
            ))}
          </div>
        </div>

        <div
          style={{ ...styles.dropZone, ...(dragging ? styles.dropZoneActive : {}) }}
          onClick={() => inputRef.current.click()}
          onDragOver={(e) => { e.preventDefault(); setDragging(true); }}
          onDragLeave={() => setDragging(false)}
          onDrop={handleDrop}
        >
          <div style={styles.dropIcon}>📄</div>
          <p style={styles.dropTitle}>Drop PDFs here or click to upload</p>
          <p style={styles.dropSub}>Multiple files supported · Auto-detects location from Ship To address</p>
        </div>
        <input ref={inputRef} type="file" multiple accept="application/pdf" style={{ display: "none" }}
          onChange={(e) => { addFiles(e.target.files); e.target.value = ""; }} />

        {files.length > 0 && (
          <div style={styles.fileList}>
            {files.map((entry, i) => (
              <div key={entry.file.name} style={styles.fileChip}>
                <div style={styles.fileIcon}>PDF</div>
                <span style={styles.fileName}>{entry.file.name}</span>
                <span style={styles.fileSize}>{(entry.file.size / 1024).toFixed(1)} KB</span>
                <select
                  value={entry.location}
                  onChange={e => setFileLocation(i, e.target.value)}
                  style={styles.locationSelect}
                >
                  {LOCATIONS.map(loc => (
                    <option key={loc} value={loc}>{loc}</option>
                  ))}
                </select>
                <button onClick={() => removeFile(i)} style={styles.removeBtn}>×</button>
              </div>
            ))}
          </div>
        )}

        <button
          disabled={!files.length || processing}
          onClick={runExtraction}
          style={{ ...styles.extractBtn, ...((!files.length || processing) ? styles.extractBtnDisabled : {}) }}
        >
          {processing ? (
            <span style={{ display: "flex", alignItems: "center", gap: 10, justifyContent: "center" }}>
              <span style={styles.spinner} /> Extracting…
            </span>
          ) : "Extract PO Data"}
        </button>

        {status && <p style={styles.statusText}>{status}</p>}

        {results.length > 0 && (
          <div style={{ marginTop: "2.5rem" }}>
            <div style={styles.divider}>
              <div style={styles.dividerLine} />
              <span style={styles.dividerLabel}>Extracted Results</span>
              <div style={styles.dividerLine} />
            </div>
            <div style={{ display: "flex", flexDirection: "column", gap: "2rem" }}>
              {results.map((r, i) =>
                r.type === "success"
                  ? <ResultCard key={i} result={r} />
                  : <div key={i} style={styles.errorCard}><strong>{r.filename}</strong><br />{r.msg}</div>
              )}
            </div>
          </div>
        )}
      </div>
    </div>
  );
}

const styles = {
  root: { fontFamily: "'DM Mono', monospace", background: "#fdf8f0", minHeight: "100vh" },
  header: { background: "#2c1a0e", padding: "1.2rem 2rem", display: "flex", alignItems: "center", justifyContent: "space-between", position: "sticky", top: 0, zIndex: 100 },
  headerLeft: { display: "flex", alignItems: "center", gap: 14 },
  logoBlock: { width: 36, height: 36, background: "#8b5e3c", borderRadius: 8, display: "flex", alignItems: "center", justifyContent: "center" },
  logoEmoji: { fontSize: 20 },
  logoTitle: { fontFamily: "serif", fontSize: "1.05rem", color: "#fdf8f0", fontWeight: 600 },
  logoSub: { fontSize: "0.62rem", color: "rgba(253,248,240,0.4)", letterSpacing: "0.08em", textTransform: "uppercase" },
  headerBadge: { fontSize: "0.6rem", letterSpacing: "0.1em", textTransform: "uppercase", border: "1px solid rgba(255,255,255,0.12)", padding: "4px 12px", borderRadius: 20, color: "rgba(253,248,240,0.5)" },
  main: { maxWidth: 1000, margin: "0 auto", padding: "3rem 1.5rem 5rem" },
  hero: { marginBottom: "2.5rem" },
  title: { fontFamily: "serif", fontSize: "2.2rem", fontWeight: 400, color: "#2c1a0e", lineHeight: 1.25, marginBottom: "0.75rem", letterSpacing: "-0.02em" },
  titleEm: { fontStyle: "italic", color: "#8b5e3c" },
  desc: { fontSize: "0.78rem", color: "#8a7a6a", letterSpacing: "0.02em", lineHeight: 1.6, marginBottom: "1rem" },
  locationPills: { display: "flex", gap: 8, flexWrap: "wrap" },
  locationPill: { display: "flex", alignItems: "center", gap: 6, background: "#fff", border: "1px solid rgba(139,94,60,0.2)", borderRadius: 20, padding: "4px 12px", fontSize: "0.72rem", color: "#8b5e3c", fontWeight: 500 },
  locationDot: { width: 6, height: 6, borderRadius: "50%", background: "#8b5e3c" },
  dropZone: { border: "1.5px dashed rgba(44,26,14,0.2)", borderRadius: 20, padding: "3rem 2rem", textAlign: "center", cursor: "pointer", background: "#f5ede0", transition: "all 0.2s" },
  dropZoneActive: { background: "#ecdfc8", borderColor: "#8b5e3c" },
  dropIcon: { fontSize: 32, marginBottom: 12 },
  dropTitle: { fontSize: "0.88rem", fontWeight: 500, color: "#2c1a0e", marginBottom: 4 },
  dropSub: { fontSize: "0.7rem", color: "#8a7a6a" },
  fileList: { marginTop: 12, display: "flex", flexDirection: "column", gap: 6 },
  fileChip: { display: "flex", alignItems: "center", gap: 10, background: "#fff", border: "1px solid rgba(44,26,14,0.1)", borderRadius: 10, padding: "8px 14px" },
  fileIcon: { fontSize: "0.55rem", fontWeight: 600, background: "#8b5e3c", color: "#fff", padding: "2px 5px", borderRadius: 3 },
  fileName: { flex: 1, fontSize: "0.75rem", color: "#2c1a0e", overflow: "hidden", textOverflow: "ellipsis", whiteSpace: "nowrap" },
  fileSize: { fontSize: "0.68rem", color: "#8a7a6a" },
  removeBtn: { background: "none", border: "none", cursor: "pointer", color: "#8a7a6a", fontSize: "1.1rem", lineHeight: 1, padding: "0 2px" },
  locationSelect: { fontFamily: "'DM Mono', monospace", fontSize: "0.68rem", background: "#fdf8f0", border: "1px solid rgba(139,94,60,0.3)", borderRadius: 6, padding: "3px 8px", color: "#8b5e3c", cursor: "pointer" },
  extractBtn: { marginTop: "1.25rem", width: "100%", padding: "14px 20px", fontFamily: "'DM Mono', monospace", fontSize: "0.78rem", fontWeight: 500, letterSpacing: "0.08em", textTransform: "uppercase", background: "#2c1a0e", color: "#fdf8f0", border: "none", borderRadius: 12, cursor: "pointer" },
  extractBtnDisabled: { background: "#d8cfc4", color: "#8a7a6a", cursor: "not-allowed" },
  spinner: { display: "inline-block", width: 14, height: 14, border: "2px solid rgba(253,248,240,0.2)", borderTopColor: "#fdf8f0", borderRadius: "50%", animation: "spin 0.7s linear infinite" },
  statusText: { textAlign: "center", fontSize: "0.7rem", color: "#8a7a6a", marginTop: 10, letterSpacing: "0.04em" },
  divider: { display: "flex", alignItems: "center", marginBottom: "1.5rem" },
  dividerLine: { flex: 1, height: 1, background: "rgba(44,26,14,0.1)" },
  dividerLabel: { fontSize: "0.6rem", letterSpacing: "0.14em", textTransform: "uppercase", color: "#8a7a6a", padding: "0 14px", whiteSpace: "nowrap" },
  resultCard: { background: "#fff", border: "1px solid rgba(44,26,14,0.1)", borderRadius: 20, overflow: "hidden" },
  resultHeader: { background: "#2c1a0e", padding: "16px 20px", display: "flex", alignItems: "flex-start", justifyContent: "space-between", gap: 12, flexWrap: "wrap" },
  filename: { fontFamily: "serif", fontSize: "1rem", color: "#fdf8f0", fontWeight: 600, marginBottom: 3 },
  resultMeta: { fontSize: "0.68rem", color: "rgba(253,248,240,0.45)" },
  downloadRow: { display: "flex", gap: 6, flexWrap: "wrap" },
  dlBtn: { border: "1px solid rgba(255,255,255,0.15)", borderRadius: 6, padding: "5px 10px", fontFamily: "'DM Mono', monospace", fontSize: "0.65rem", letterSpacing: "0.05em", textTransform: "uppercase", cursor: "pointer" },
  shopifyBtn: { background: "rgba(149,191,71,0.15)", borderColor: "rgba(149,191,71,0.4)", color: "rgba(149,191,71,0.9)" },
  stockyBtn: { background: "rgba(59,130,246,0.12)", borderColor: "rgba(59,130,246,0.35)", color: "rgba(147,197,253,0.9)" },
  jsonBtn: { background: "rgba(255,255,255,0.06)", borderColor: "rgba(255,255,255,0.15)", color: "rgba(253,248,240,0.6)" },
  tabBar: { display: "flex", overflowX: "auto", borderBottom: "1px solid rgba(44,26,14,0.1)", background: "#fdf8f0" },
  tab: { padding: "10px 18px", fontSize: "0.72rem", fontFamily: "'DM Mono', monospace", letterSpacing: "0.04em", background: "none", border: "none", borderBottom: "2px solid transparent", cursor: "pointer", color: "#8a7a6a", whiteSpace: "nowrap", display: "flex", alignItems: "center", gap: 7 },
  tabActive: { color: "#2c1a0e", borderBottomColor: "#8b5e3c", background: "#fff" },
  tabCount: { background: "#8b5e3c", color: "#fff", borderRadius: 10, padding: "1px 6px", fontSize: "0.6rem" },
  brandCard: { background: "#fff" },
  brandHeader: { padding: "14px 20px", background: "#5c3418", display: "flex", alignItems: "center", justifyContent: "space-between", flexWrap: "wrap", gap: 10 },
  brandName: { fontFamily: "serif", fontSize: "1.1rem", color: "#fdf8f0", fontWeight: 600 },
  brandMeta: { fontSize: "0.68rem", color: "rgba(253,248,240,0.5)", marginTop: 2 },
  totalBadge: { background: "rgba(253,248,240,0.15)", border: "1px solid rgba(253,248,240,0.3)", color: "#fdf8f0", borderRadius: 8, padding: "4px 12px", fontSize: "0.78rem", fontWeight: 600 },
  table: { width: "100%", borderCollapse: "collapse", fontSize: "0.75rem" },
  th: { padding: "8px 12px", fontSize: "0.58rem", fontWeight: 500, letterSpacing: "0.1em", textTransform: "uppercase", color: "#8a7a6a", borderBottom: "1px solid rgba(44,26,14,0.1)", background: "#fdf8f0" },
  td: { padding: "9px 12px", color: "#2c1a0e", verticalAlign: "top" },
  errorCard: { background: "#fff8f8", border: "1px solid #f5c2c2", borderRadius: 16, padding: "16px 20px", fontSize: "0.78rem", color: "#8b2020" },
};
