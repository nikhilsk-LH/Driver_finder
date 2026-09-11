import React, { useState, useMemo, useEffect, useRef, useCallback } from "react";
import * as XLSX from "xlsx";
import Papa from "papaparse";

/* ------------------------------------------------------------------ */
/*  Styling — Netflix palette + glassmorphism in one injected sheet.   */
/* ------------------------------------------------------------------ */
const CSS = `
:root{
  --bg0:#0a0a0a; --bg1:#141414; --red:#e50914; --red-2:#7a0007;
  --txt:#f5f5f5; --muted:#9a9a9a; --line:rgba(255,255,255,.12);
  --glass:rgba(255,255,255,.055); --glass-2:rgba(255,255,255,.09);
}
*{box-sizing:border-box}
.dn-root{
  position:relative; min-height:100vh; width:100%;
  background:
    radial-gradient(1100px 600px at 12% -10%, rgba(229,9,20,.16), transparent 60%),
    radial-gradient(900px 700px at 100% 0%, rgba(229,9,20,.10), transparent 55%),
    linear-gradient(180deg,#0b0b0b 0%, #141414 40%, #0a0a0a 100%);
  color:var(--txt);
  font-family:"Helvetica Neue",Helvetica,Arial,-apple-system,BlinkMacSystemFont,"Segoe UI",Roboto,sans-serif;
  overflow-x:hidden;
}
.dn-wrap{position:relative; z-index:2; max-width:1100px; margin:0 auto; padding:26px 22px 80px}
.blob{position:fixed; border-radius:50%; filter:blur(90px); opacity:.5; z-index:0; pointer-events:none; animation:float 18s ease-in-out infinite}
.blob.a{width:420px;height:420px;background:rgba(229,9,20,.35); top:-140px; left:-120px}
.blob.b{width:360px;height:360px;background:rgba(120,0,7,.4); bottom:-160px; right:-120px; animation-delay:-6s}

.topbar{display:flex; align-items:center; justify-content:space-between; gap:16px; margin-bottom:22px; flex-wrap:wrap}
.brand{display:flex; align-items:baseline; gap:10px}
.brand h1{margin:0; font-size:26px; font-weight:900; letter-spacing:-1.4px}
.brand h1 span{color:var(--red)}
.brand small{color:var(--muted); font-weight:600; letter-spacing:.2px}
.src-chip{display:flex; align-items:center; gap:10px; font-size:13px; color:var(--muted)}
.src-chip b{color:var(--txt); font-weight:700}
.linklike{background:none;border:none;color:var(--red);cursor:pointer;font-weight:700;font-size:13px;padding:0}
.linklike:hover{text-decoration:underline}

.glass{
  background:var(--glass); border:1px solid var(--line); border-radius:18px;
  -webkit-backdrop-filter:blur(18px) saturate(120%); backdrop-filter:blur(18px) saturate(120%);
  box-shadow:0 18px 40px rgba(0,0,0,.45);
}

.hero{padding:40px 8px 26px; text-align:center}
.hero h2{font-size:clamp(30px,5vw,52px); line-height:1.02; margin:0 0 12px; font-weight:900; letter-spacing:-1.8px}
.hero h2 em{font-style:normal;color:var(--red)}
.hero p{max-width:560px;margin:0 auto;color:var(--muted);font-size:16px;line-height:1.5}
.dropzone{margin:30px auto 0; max-width:640px; padding:44px 26px; text-align:center; cursor:pointer;
  border:1.5px dashed rgba(255,255,255,.22); transition:.25s ease; position:relative; overflow:hidden}
.dropzone:hover,.dropzone.over{border-color:var(--red); background:var(--glass-2); transform:translateY(-2px)}
.dropzone.over::after{content:"";position:absolute;inset:0;background:linear-gradient(120deg,transparent,rgba(229,9,20,.14),transparent);animation:shimmer 1.1s linear infinite}
.dz-ic{width:58px;height:58px;margin:0 auto 14px;border-radius:16px;display:grid;place-items:center;
  background:linear-gradient(135deg,var(--red),var(--red-2)); box-shadow:0 10px 30px rgba(229,9,20,.45); font-size:26px}
.dz-t{font-size:19px;font-weight:800}
.dz-s{color:var(--muted);font-size:13.5px;margin-top:6px}
.fmt{display:flex;gap:8px;justify-content:center;flex-wrap:wrap;margin-top:16px}
.fmt span{font-size:11px;font-weight:700;color:var(--muted);border:1px solid var(--line);border-radius:999px;padding:4px 11px;letter-spacing:.4px}

.btn{display:inline-flex;align-items:center;gap:8px;border:none;border-radius:10px;cursor:pointer;
  font-weight:800;font-size:14px;padding:12px 20px;transition:.18s ease;color:#fff}
.btn.red{background:var(--red);box-shadow:0 8px 24px rgba(229,9,20,.4)}
.btn.red:hover{transform:translateY(-2px);background:#f6121d}
.btn.ghost{background:var(--glass-2);border:1px solid var(--line)}
.btn.ghost:hover{background:rgba(255,255,255,.14)}
.cta-row{display:flex;gap:12px;justify-content:center;flex-wrap:wrap;margin-top:20px}
.gs-row{display:flex;gap:10px;max-width:640px;margin:16px auto 0;flex-wrap:wrap}
.gs-row input{flex:1;min-width:220px;background:rgba(0,0,0,.35);border:1px solid var(--line);color:var(--txt);
  border-radius:10px;padding:11px 14px;font-size:13.5px;outline:none}
.gs-row input:focus{border-color:var(--red)}

.eyebrow{color:var(--muted);font-size:12.5px;font-weight:800;letter-spacing:.6px;margin:0 0 10px}
.sec{margin-top:26px}

/* facility dropdown */
.selectwrap{position:relative; max-width:440px}
.fac-select{appearance:none;-webkit-appearance:none;width:100%;background:transparent;color:var(--txt);
  border:none;border-radius:16px;padding:16px 46px 16px 18px;font-size:17px;font-weight:800;cursor:pointer;outline:none;letter-spacing:-.3px}
.fac-select option{background:#141414;color:#fff;font-weight:700}
.selectwrap::after{content:"";position:absolute;right:20px;top:50%;width:9px;height:9px;border-right:2px solid var(--red);
  border-bottom:2px solid var(--red);transform:translateY(-70%) rotate(45deg);pointer-events:none}

/* driver header */
.dhead{display:flex;align-items:center;justify-content:space-between;gap:14px;flex-wrap:wrap;margin-bottom:14px}
.dhead h3{margin:0;font-size:22px;font-weight:900;letter-spacing:-.8px}
.dhead h3 .tag{color:var(--red)}
.dhead h3 .c{color:var(--muted);font-weight:700;font-size:16px}

.grid{display:grid;grid-template-columns:repeat(auto-fill,minmax(240px,1fr));gap:14px}
.dcard{padding:18px;position:relative;overflow:hidden;transition:transform .2s ease,border-color .2s ease,box-shadow .2s ease;
  animation:fadeUp .5s cubic-bezier(.2,.8,.2,1) both}
.dcard:hover{transform:translateY(-6px) scale(1.02);border-color:rgba(229,9,20,.6);
  box-shadow:0 22px 46px rgba(0,0,0,.55),0 0 0 1px rgba(229,9,20,.35),0 14px 40px rgba(229,9,20,.28)}
.nameplate{position:relative;overflow:hidden;border-radius:14px;padding:15px 16px;
  background:linear-gradient(135deg,var(--red),var(--red-2));
  box-shadow:0 8px 22px rgba(229,9,20,.4);color:#fff;font-weight:900;font-size:19px;letter-spacing:-.4px;
  line-height:1.18;min-height:58px;display:flex;align-items:center;word-break:break-word;
  transition:filter .25s ease,box-shadow .25s ease,letter-spacing .25s ease}
.nameplate::before{content:"";position:absolute;top:0;left:-130%;width:55%;height:100%;pointer-events:none;
  background:linear-gradient(120deg,transparent,rgba(255,255,255,.42),transparent);transform:skewX(-20deg);
  transition:left .65s cubic-bezier(.2,.8,.2,1)}
.dcard:hover .nameplate{filter:brightness(1.09) saturate(1.05);box-shadow:0 12px 30px rgba(229,9,20,.6);letter-spacing:-.1px}
.dcard:hover .nameplate::before{left:150%}
.dcard .sub{margin-top:12px;display:flex;gap:6px;align-items:center;flex-wrap:wrap}
.pref{margin-top:11px;font-size:12px;color:var(--muted);font-weight:600;display:flex;align-items:center;gap:6px;flex-wrap:wrap;line-height:1.3}
.pref .star{color:var(--red);font-size:13px}
.pref b{color:#fff;font-weight:800}
.pref .blk{color:var(--red);font-weight:800;background:rgba(229,9,20,.14);border:1px solid rgba(229,9,20,.4);
  border-radius:999px;padding:1px 8px;font-size:11px}
.shift-badge{display:inline-block;padding:2px 9px;border-radius:999px;font-size:11px;font-weight:800;letter-spacing:.4px;
  border:1px solid var(--line);background:rgba(255,255,255,.06);color:var(--muted)}
.multi{display:inline-block;margin-left:6px;padding:2px 8px;border-radius:999px;font-size:10.5px;font-weight:800;
  color:#ffd9db;background:rgba(229,9,20,.16);border:1px solid rgba(229,9,20,.5)}
.meta{margin-top:14px;display:flex;flex-direction:column;gap:7px}
.meta .m{display:flex;justify-content:space-between;gap:10px;font-size:12.5px}
.meta .m .k{color:var(--muted);font-weight:600}
.meta .m .v{font-weight:700;text-align:right;overflow:hidden;text-overflow:ellipsis;white-space:nowrap;max-width:60%}

.empty{padding:46px 20px;text-align:center;color:var(--muted)}
.empty b{color:var(--txt);display:block;font-size:16px;margin-bottom:6px}

.err{background:rgba(229,9,20,.14);border:1px solid rgba(229,9,20,.5);color:#ffd9db;border-radius:12px;
  padding:12px 16px;font-size:13.5px;margin-top:16px}
.spin{width:16px;height:16px;border-radius:50%;border:2px solid rgba(255,255,255,.3);border-top-color:#fff;
  animation:spin .7s linear infinite;display:inline-block}
.foot{margin-top:40px;text-align:center;color:var(--muted);font-size:12px}

@keyframes float{0%,100%{transform:translate(0,0)}50%{transform:translate(30px,26px)}}
@keyframes shimmer{0%{transform:translateX(-100%)}100%{transform:translateX(100%)}}
@keyframes fadeUp{from{opacity:0;transform:translateY(16px)}to{opacity:1;transform:translateY(0)}}
@keyframes spin{to{transform:rotate(360deg)}}
@media (prefers-reduced-motion:reduce){*{animation:none!important;transition:none!important}}
`;

/* ------------------------------------------------------------------ */
/*  Demo network — first/last columns + a few multi-facility drivers.  */
/* ------------------------------------------------------------------ */
const FIRST = ["Aarav","Mohammed","James","Daniel","Oliver","Krish","Samuel","Tomasz","Ade","Chen","Ravi","Luca","Ethan","Noah","Ibrahim","George","Marcus","Dmitri","Sanjay","Kofi","Andrei","Hassan","Leo","Ryan","Priya","Amara","Sofia","Grace","Fatima","Hannah","Zara","Nadia","Elena","Bilal","Omar","Jakub","Kwame","Diego","Yusuf","Arjun"];
const LAST = ["Sharma","Khan","Smith","Nowak","Patel","Adeyemi","Wang","Silva","O'Brien","Kumar","Rossi","Brown","Okafor","Ivanov","Ali","Mensah","Dubois","Nguyen","Kowalski","Reyes","Haddad","Marsh","Bello","Petrov","Costa","Hughes","Sane","Diallo","Fernandez","Osei"];
const FACS = ["LUX","GSD","WRS","OLD","HDC","DRP"];
const SHIFTS = ["AM","PM","NGT"];

function buildDemo() {
  const rows = [];
  let n = 0;
  const counts = [["LUX", 11], ["GSD", 9], ["WRS", 8], ["OLD", 7], ["HDC", 9], ["DRP", 6]];
  counts.forEach(([fac, c]) => {
    for (let i = 0; i < c; i++) {
      let facility = fac;
      if (n % 9 === 4) { const alt = FACS[(n + 2) % FACS.length]; if (alt !== fac) facility = fac + ", " + alt; } // multi-facility driver
      rows.push({
        "First Name": FIRST[(n * 7) % FIRST.length],
        "Last Name": LAST[(n * 13) % LAST.length],
        "Facility": facility,
        "Shift": SHIFTS[(n + i) % 3],
        "Driver ID": "LH-" + (2400 + n),
        "Phone": "+44 7" + String(300 + ((n * 37) % 699)) + " " + String(100000 + ((n * 5309) % 899999)),
      });
      n++;
    }
  });
  return rows;
}

/* ------------------------------------------------------------------ */
/*  Column auto-detection                                              */
/* ------------------------------------------------------------------ */
function detectColumns(headers, rows) {
  const facPat = /facil|site|location|hub|depot|station|branch|centre|center|warehouse|store/i;
  const firstPat = /first ?name|fore ?name|given/i;
  const lastPat = /last ?name|surname|family ?name/i;
  const drvPat = /driver|full ?name|employee|rider|courier|agent|staff|partner|name/i;
  const shiftPat = /shift|slot|block|rota/i;

  let facility = headers.find((h) => facPat.test(h));
  const firstCol = headers.find((h) => firstPat.test(h));
  const lastCol = headers.find((h) => lastPat.test(h) && h !== firstCol);

  let driver, lastName = null;
  if (firstCol) { driver = firstCol; lastName = lastCol || null; }
  else { driver = headers.find((h) => drvPat.test(h)) || headers[0]; lastName = lastCol || null; }

  const shift = headers.find((h) => shiftPat.test(h) && h !== driver && h !== facility && h !== lastName) || null;

  if (!facility) {
    let best = null, bestScore = Infinity;
    headers.forEach((h) => {
      if (h === driver || h === lastName) return;
      const uniq = new Set(rows.map((r) => String(r[h] ?? "").trim()).filter(Boolean));
      if (uniq.size > 1 && uniq.size < rows.length && uniq.size < bestScore) { bestScore = uniq.size; best = h; }
    });
    facility = best || headers.find((h) => h !== driver && h !== lastName) || headers[0];
  }
  return { facility, driver, lastName, shift };
}

function normalizeRows(raw) {
  const cleaned = raw
    .map((r) => {
      const o = {};
      Object.keys(r).forEach((k) => { o[String(k).trim()] = typeof r[k] === "string" ? r[k].trim() : r[k]; });
      return o;
    })
    .filter((r) => Object.values(r).some((v) => v !== "" && v != null));
  const headerSet = [];
  cleaned.forEach((r) => Object.keys(r).forEach((k) => { if (k && !headerSet.includes(k)) headerSet.push(k); }));
  return { rows: cleaned, headers: headerSet };
}

const initials = (name) =>
  String(name).split(/\s+/).filter(Boolean).slice(0, 2).map((w) => w[0]).join("").toUpperCase() || "?";

/* ------------------------------------------------------------------ */
export default function DriverNetwork() {
  const [rows, setRows] = useState([]);
  const [headers, setHeaders] = useState([]);
  const [cols, setCols] = useState({ facility: null, driver: null, lastName: null, shift: null });
  const [selected, setSelected] = useState("");
  const [fileName, setFileName] = useState("");
  const [over, setOver] = useState(false);
  const [error, setError] = useState("");
  const [loading, setLoading] = useState(false);
  const [gsUrl, setGsUrl] = useState("");
  const inputRef = useRef(null);

  const finish = useCallback((raw, name) => {
    const { rows: r, headers: h } = normalizeRows(raw);
    if (!r.length) { setError("That file has no readable rows. Check it has a header row and try again."); setLoading(false); return; }
    setRows(r); setHeaders(h); setCols(detectColumns(h, r));
    setFileName(name); setSelected(""); setError(""); setLoading(false);
  }, []);

  const fail = useCallback((msg) => { setError(typeof msg === "string" ? msg : (msg?.message || "Could not read that file.")); setLoading(false); }, []);

  const ingest = useCallback((file) => {
    if (!file) return;
    setLoading(true); setError("");
    const name = file.name.toLowerCase();
    if (/\.(csv|tsv|txt)$/.test(name)) {
      Papa.parse(file, { header: true, skipEmptyLines: "greedy", complete: (res) => finish(res.data, file.name), error: (e) => fail(e) });
    } else if (/\.(xlsx|xls|xlsm)$/.test(name)) {
      const r = new FileReader();
      r.onload = (e) => {
        try {
          const wb = XLSX.read(e.target.result, { type: "array" });
          const ws = wb.Sheets[wb.SheetNames[0]];
          finish(XLSX.utils.sheet_to_json(ws, { defval: "" }), file.name);
        } catch (err) { fail(err); }
      };
      r.onerror = () => fail(r.error); r.readAsArrayBuffer(file);
    } else { fail("Unsupported format. Upload a .csv, .xlsx or .xls file."); }
  }, [finish, fail]);

  const loadGoogleSheet = useCallback(() => {
    const m = gsUrl.match(/\/d\/([a-zA-Z0-9-_]+)/);
    if (!m) { setError("That doesn't look like a Google Sheets link. Paste the full URL, or download the sheet as CSV and upload it."); return; }
    setLoading(true); setError("");
    fetch(`https://docs.google.com/spreadsheets/d/${m[1]}/gviz/tq?tqx=out:csv`)
      .then((res) => { if (!res.ok) throw new Error("blocked"); return res.text(); })
      .then((text) => Papa.parse(text, { header: true, skipEmptyLines: "greedy", complete: (r) => finish(r.data, "Google Sheet") }))
      .catch(() => fail("Couldn't reach that sheet — it may be private. Set link sharing to 'Anyone with the link', or download it as CSV/Excel and upload."));
  }, [gsUrl, finish, fail]);

  const onDrop = (e) => { e.preventDefault(); setOver(false); ingest(e.dataTransfer.files?.[0]); };
  const clearAll = () => { setRows([]); setHeaders([]); setCols({ facility: null, driver: null, lastName: null, shift: null }); setFileName(""); setSelected(""); setError(""); };

  /* facilities: split each cell so multi-facility drivers count for each site */
  const facilityValues = useCallback((r) => {
    const raw = String(r[cols.facility] ?? "").trim();
    if (!raw) return [];
    return raw.split(/[,;/|\n]+/).map((s) => s.trim()).filter(Boolean);
  }, [cols.facility]);

  const facilities = useMemo(() => {
    if (!rows.length || !cols.facility) return [];
    const map = new Map();
    rows.forEach((r) => facilityValues(r).forEach((f) => map.set(f, (map.get(f) || 0) + 1)));
    return [...map.entries()].map(([name, count]) => ({ name, count })).sort((a, b) => a.name.localeCompare(b.name));
  }, [rows, cols.facility, facilityValues]);

  const assignedCount = useMemo(() => rows.filter((r) => facilityValues(r).length > 0).length, [rows, facilityValues]);

  const fullName = useCallback((r) => {
    const f = String(r[cols.driver] ?? "").trim();
    const l = cols.lastName ? String(r[cols.lastName] ?? "").trim() : "";
    return (f + " " + l).trim() || "Unknown";
  }, [cols.driver, cols.lastName]);

  /* per-driver block counts across facilities → preferred facility */
  const driverStats = useMemo(() => {
    const map = new Map();
    rows.forEach((r) => {
      const facs = facilityValues(r);
      if (!facs.length) return;
      const name = fullName(r);
      if (!map.has(name)) map.set(name, new Map());
      const fm = map.get(name);
      facs.forEach((f) => fm.set(f, (fm.get(f) || 0) + 1));
    });
    return map;
  }, [rows, fullName, facilityValues]);

  const preferred = useCallback((name) => {
    const fm = driverStats.get(name);
    if (!fm || fm.size === 0) return null;
    let best = null, bestC = -1;
    fm.forEach((c, f) => { if (c > bestC || (c === bestC && best && f.localeCompare(best) < 0)) { bestC = c; best = f; } });
    return { facility: best, count: bestC };
  }, [driverStats]);

  const inView = useMemo(() => {
    if (!selected) return [];
    return selected === "ALL" ? rows.filter((r) => facilityValues(r).length > 0) : rows.filter((r) => facilityValues(r).includes(selected));
  }, [rows, selected, facilityValues]);

  const detailKeys = useMemo(
    () => headers.filter((h) => ![cols.driver, cols.lastName, cols.facility, cols.shift].includes(h)).slice(0, 4),
    [headers, cols]
  );

  const hasData = rows.length > 0;

  return (
    <div className="dn-root">
      <style>{CSS}</style>
      <div className="blob a" /><div className="blob b" />
      <input ref={inputRef} type="file" accept=".csv,.tsv,.txt,.xlsx,.xls,.xlsm" style={{ display: "none" }}
        onChange={(e) => ingest(e.target.files?.[0])} />

      <div className="dn-wrap">
        <div className="topbar">
          <div className="brand">
            <h1><span>DRIVER</span>NET</h1>
          </div>
          {hasData && (
            <div className="src-chip">
              {loading ? <span className="spin" /> : null}
              <span>Source: <b>{fileName}</b> · {rows.length} rows</span>
              <button className="linklike" onClick={() => inputRef.current?.click()}>Change</button>
              <button className="linklike" onClick={clearAll}>Clear</button>
            </div>
          )}
        </div>

        {!hasData && (
          <>
            <div className="hero">
              <h2>Pick a <em>facility</em>. See its drivers.</h2>
              <p>Drop in your roster and DriverNet groups it by facility — select one and everyone who runs there appears instantly.</p>
            </div>

            <div className={"dropzone glass" + (over ? " over" : "")}
              onClick={() => inputRef.current?.click()}
              onDragOver={(e) => { e.preventDefault(); setOver(true); }}
              onDragLeave={() => setOver(false)} onDrop={onDrop}>
              <div className="dz-ic">{loading ? <span className="spin" /> : "↑"}</div>
              <div className="dz-t">{loading ? "Reading your file…" : "Drop your roster here"}</div>
              <div className="dz-s">or click to browse — CSV, Excel or a Google Sheet</div>
              <div className="fmt"><span>.CSV</span><span>.XLSX</span><span>.XLS</span><span>GOOGLE SHEET</span></div>
            </div>

            <div className="gs-row">
              <input placeholder="…or paste a Google Sheet link (shared: anyone with link)"
                value={gsUrl} onChange={(e) => setGsUrl(e.target.value)}
                onKeyDown={(e) => e.key === "Enter" && loadGoogleSheet()} />
              <button className="btn ghost" onClick={loadGoogleSheet}>Load sheet</button>
            </div>

            <div className="cta-row">
              <button className="btn red" onClick={() => finish(buildDemo(), "Sample network")}>▶ Load sample network</button>
            </div>

            {error && <div className="err">{error}</div>}
          </>
        )}

        {hasData && (
          <>
            {/* facility dropdown */}
            <div className="sec" style={{ marginTop: 6 }}>
              <p className="eyebrow">SELECT A FACILITY</p>
              <div className="selectwrap glass">
                <select className="fac-select" value={selected} onChange={(e) => setSelected(e.target.value)}>
                  <option value="">Select a facility…</option>
                  <option value="ALL">All facilities</option>
                  {facilities.map((f) => <option key={f.name} value={f.name}>{f.name}</option>)}
                </select>
              </div>
            </div>

            {/* drivers */}
            {!selected ? (
              <div className="sec"><div className="empty glass">
                <b>Select a facility to begin</b>
                Choose a facility above to see everyone who runs there.
              </div></div>
            ) : (
              <div className="sec">
                <div className="dhead">
                  <h3>
                    {selected === "ALL" ? "All drivers" : <>Drivers at <span className="tag">{selected}</span></>}{" "}
                    <span className="c">({inView.length})</span>
                  </h3>
                </div>

                {inView.length === 0 ? (
                  <div className="empty glass">
                    <b>No drivers here</b>
                    This facility has no drivers in the roster.
                  </div>
                ) : (
                  <div className="grid" key={selected}>
                    {inView.map((r, i) => {
                      const name = fullName(r);
                      const pref = preferred(name);
                      const shiftVal = cols.shift ? r[cols.shift] : null;
                      const sites = facilityValues(r);
                      const multi = selected !== "ALL" && sites.length > 1;
                      return (
                        <div key={i} className="dcard glass" style={{ animationDelay: Math.min(i * 0.035, 0.5) + "s" }}>
                          <div className="nameplate">{name}</div>
                          {pref && (
                            <div className="pref"><span className="star">★</span>Preferred facility: <b>{pref.facility}</b></div>
                          )}
                          {(shiftVal || multi) && (
                            <div className="sub">
                              {shiftVal ? <span className="shift-badge">{shiftVal}</span> : null}
                              {multi ? <span className="multi" title={"Also runs at " + sites.filter((s) => s !== selected).join(", ")}>+{sites.length - 1} more site{sites.length - 1 > 1 ? "s" : ""}</span> : null}
                            </div>
                          )}
                          {detailKeys.length > 0 && (
                            <div className="meta">
                              {detailKeys.map((k) => (
                                <div className="m" key={k}>
                                  <span className="k">{k}</span>
                                  <span className="v" title={String(r[k] ?? "")}>{String(r[k] ?? "—") || "—"}</span>
                                </div>
                              ))}
                            </div>
                          )}
                        </div>
                      );
                    })}
                  </div>
                )}
              </div>
            )}

            {error && <div className="err">{error}</div>}
          </>
        )}

        <div className="foot">DriverNet · reads your file in the browser — nothing is uploaded to a server.</div>
      </div>
    </div>
  );
}
