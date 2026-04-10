require("dotenv").config();
const express = require("express");
const fetch = require("node-fetch");
const XLSX = require("xlsx");
const path = require("path");

const app = express();
app.use(express.json());
app.use(express.static(path.join(__dirname, "public")));

// ─── Config from env ─────────────────────────────────────────────────────────

const DBX_APP_KEY     = process.env.DROPBOX_APP_KEY;
const DBX_APP_SECRET  = process.env.DROPBOX_APP_SECRET;
const DBX_REFRESH_TOK = process.env.DROPBOX_REFRESH_TOKEN;
const SYNC_MINUTES    = parseInt(process.env.SYNC_MINUTES || "60", 10);

const PATHS = {
  FY1:   process.env.PATH_FY1   || "/Production_Style_Ledger/FY1_Production_Style_Ledger.xlsx",
  NB:    process.env.PATH_NB    || "/Production_Style_Ledger/NB_Production_Style_Ledger.xlsx",
  PC:    process.env.PATH_PC    || "/Production_Style_Ledger/PC_Production_Style_Ledger.xlsx",
  DAVID: process.env.PATH_DAVID || "/Versa Share Files/David - Dropbox/Style Ledger/Style Ledger Template for David.xlsx",
};

const FACTORY_TABS = {
  FY1: ["Style_Ledger", "Style Ledger '26", "Style Ledger '25"],
  NB:  ["Style_Ledger", "Style Ledger '26", "Style Ledger '25"],
  PC:  ["Style_Ledger"],
};

// ─── Dropbox OAuth2 — auto-refreshing access token ──────────────────────────

let accessToken = null;
let tokenExpires = 0;

async function getAccessToken() {
  if (accessToken && Date.now() < tokenExpires - 60000) return accessToken;

  log("Refreshing Dropbox access token...");
  const res = await fetch("https://api.dropboxapi.com/oauth2/token", {
    method: "POST",
    headers: { "Content-Type": "application/x-www-form-urlencoded" },
    body: new URLSearchParams({
      grant_type: "refresh_token",
      refresh_token: DBX_REFRESH_TOK,
      client_id: DBX_APP_KEY,
      client_secret: DBX_APP_SECRET,
    }),
  });

  if (!res.ok) {
    const txt = await res.text();
    throw new Error(`Token refresh failed (${res.status}): ${txt.slice(0, 200)}`);
  }

  const data = await res.json();
  accessToken = data.access_token;
  tokenExpires = Date.now() + (data.expires_in * 1000);
  log("Access token refreshed ✓", "ok");
  return accessToken;
}

// ─── State ───────────────────────────────────────────────────────────────────

let state = {
  lastSync: null,
  nextSync: null,
  syncing: false,
  error: null,
  logs: [],
  analysis: null,
  factoryCounts: {},
  davidCount: 0,
};

function log(msg, type = "info") {
  const entry = { msg, type, ts: new Date().toISOString() };
  state.logs.unshift(entry);
  if (state.logs.length > 100) state.logs.length = 100;
  const prefix = type === "ok" ? "✓" : type === "error" ? "✗" : type === "warn" ? "⚠" : "·";
  console.log(`[${prefix}] ${msg}`);
}

// ─── Dropbox Download ────────────────────────────────────────────────────────

async function dropboxDownload(filePath) {
  const token = await getAccessToken();
  const res = await fetch("https://content.dropboxapi.com/2/files/download", {
    method: "POST",
    headers: {
      Authorization: `Bearer ${token}`,
      "Dropbox-API-Arg": JSON.stringify({ path: filePath }),
    },
  });
  if (!res.ok) {
    const txt = await res.text();
    throw new Error(`Dropbox ${res.status} for ${filePath}: ${txt.slice(0, 200)}`);
  }
  return Buffer.from(await res.arrayBuffer());
}

// ─── Parsing ─────────────────────────────────────────────────────────────────

function findTab(wb, key) {
  const pref = FACTORY_TABS[key] || [];
  for (const t of pref) { if (wb.SheetNames.includes(t)) return t; }
  return wb.SheetNames.find(s => s.toLowerCase().includes("style") && s.toLowerCase().includes("ledger"));
}

function sheetToRows(wb, tabName, factory) {
  const sheet = wb.Sheets[tabName];
  if (!sheet) return [];
  const raw = XLSX.utils.sheet_to_json(sheet, { header: 1 });
  if (raw.length < 2) return [];
  const hdr = raw[0].map(h => (h || "").toString().trim());
  return raw.slice(1).filter(r => r && r[0]).map(r => {
    const obj = {};
    hdr.forEach((h, j) => { obj[h] = r[j]; });
    if (factory) { obj._factory = factory; obj._tab = tabName; }
    return obj;
  });
}

function parseFactory(buf, key) {
  const wb = XLSX.read(buf);
  const tab = findTab(wb, key);
  return tab ? sheetToRows(wb, tab, key) : [];
}

function parseDavid(buf) {
  const wb = XLSX.read(buf);
  const tab = wb.SheetNames.find(s => s.toLowerCase().includes("style") && s.toLowerCase().includes("ledger"));
  return tab ? sheetToRows(wb, tab) : [];
}

// ─── Date / Number Helpers ───────────────────────────────────────────────────

function toDate(v) {
  if (!v) return null;
  if (v instanceof Date) return v;
  if (typeof v === "number") return new Date((v - 25569) * 86400000);
  const d = new Date(String(v).slice(0, 10) + "T00:00:00");
  return isNaN(d.getTime()) ? null : d;
}

function fmtDate(v) {
  const d = toDate(v);
  return d ? d.toISOString().slice(0, 10) : "";
}

function num(v) {
  return v == null || v === "" || isNaN(Number(v)) ? null : Number(v);
}

// ─── Brand Normalization ─────────────────────────────────────────────────────

const BRAND_ALIASES = {
  "DKNY": "DK", "DK": "DK",
};

function normBrand(b) {
  const v = (b || "").toString().trim().toUpperCase();
  return BRAND_ALIASES[v] || v;
}

// ─── Analysis Engine ─────────────────────────────────────────────────────────

function runAnalysis(factoryRows, davidRows) {
  const davidMap = {};
  davidRows.forEach(r => {
    const k = (r["STYLE"] || "").toString().trim();
    if (k) davidMap[k] = r;
  });

  const now = new Date();
  const cutoff = new Date(now.getTime() - 60 * 86400000);

  const factoryMap = {};
  factoryRows.forEach(r => {
    const k = (r["STYLE"] || "").toString().trim();
    if (k) { if (!factoryMap[k]) factoryMap[k] = []; factoryMap[k].push(r); }
  });

  const newStyles = [], updates = [], matched = new Set(), ignoredOld = [];

  Object.entries(factoryMap).forEach(([style, fRows]) => {
    const f = fRows[0], d = davidMap[style];
    if (!d) {
      const etd = toDate(f["ETD"] || f["EX-FACTORY.DATE"]);
      if (etd && etd > cutoff) newStyles.push(cleanRow(f, style));
      else ignoredOld.push(style);
      return;
    }
    matched.add(style);
    const changes = [];

    // Only track changes where BOTH sides have real values (not empty→filled)
    const fU = num(f["Ship Units"]) ?? num(f["PO Units"]);
    const dU = num(d["Ship Units"]);
    if (fU != null && dU != null && fU !== dU) changes.push({ field: "Units", from: dU.toLocaleString(), to: fU.toLocaleString(), delta: fU - dU });

    const fE = fmtDate(f["ETD"] || f["EX-FACTORY.DATE"]);
    const dE = fmtDate(d["ETD"]);
    if (fE && dE && fE !== dE) changes.push({ field: "ETD", from: dE, to: fE });

    const fP = (f["Production#"] || "").toString().trim(), dP = (d["Production#"] || "").toString().trim();
    if (fP && dP && fP !== dP) changes.push({ field: "Prod#", from: dP, to: fP });

    if (changes.length) {
      updates.push({
        style,
        factory: f._factory,
        changes,
        row: cleanRow(f, style),
      });
    }
  });

  const missing = davidRows.filter(r => {
    const s = (r["STYLE"] || "").toString().trim();
    return s && !factoryMap[s];
  }).map(r => ({
    prod: (r["Production#"] || "").toString(),
    po: (r["PO NAME"] || "").toString(),
    style: (r["STYLE"] || "").toString(),
    units: num(r["Ship Units"]),
    brand: normBrand(r["Brand"]),
    etd: fmtDate(r["ETD"]),
  }));

  return {
    newStyles,
    updates,
    missing,
    matched: matched.size,
    ignoredOld: ignoredOld.length,
    totalFactory: Object.keys(factoryMap).length,
    totalDavid: Object.keys(davidMap).length,
    cutoffDate: fmtDate(cutoff),
  };
}

function cleanRow(f, style) {
  return {
    prod: (f["Production#"] || "").toString(),
    po: (f["PO NAME"] || "").toString(),
    style,
    units: num(f["Ship Units"]) ?? num(f["PO Units"]),
    brand: normBrand(f["Brand"]),
    etd: fmtDate(f["ETD"] || f["EX-FACTORY.DATE"]),
    factory: f._factory,
}

// ─── Sync Orchestrator ───────────────────────────────────────────────────────

async function doSync(source) {
  if (!DBX_APP_KEY || !DBX_APP_SECRET || !DBX_REFRESH_TOK) {
    log("Missing DROPBOX_APP_KEY, DROPBOX_APP_SECRET, or DROPBOX_REFRESH_TOKEN", "error");
    state.error = "Missing Dropbox credentials";
    return;
  }
  state.syncing = true;
  state.error = null;
  log(`[${source}] Sync started`);
  const t0 = Date.now();

  try {
    const allFactory = [];

    for (const key of ["FY1", "NB", "PC"]) {
      const p = PATHS[key];
      if (!p) { log(`Skip ${key} — no path`, "warn"); continue; }
      log(`Downloading ${key}...`);
      const buf = await dropboxDownload(p);
      const rows = parseFactory(buf, key);
      allFactory.push(...rows);
      state.factoryCounts[key] = rows.length;
      log(`${key}: ${rows.length} styles`, "ok");
    }

    log("Downloading David's template...");
    const dBuf = await dropboxDownload(PATHS.DAVID);
    const davidRows = parseDavid(dBuf);
    state.davidCount = davidRows.length;
    log(`Template: ${davidRows.length} styles`, "ok");

    state.analysis = runAnalysis(allFactory, davidRows);
    state.lastSync = new Date().toISOString();
    const next = new Date(Date.now() + SYNC_MINUTES * 60000);
    state.nextSync = next.toISOString();
    log(`Done in ${((Date.now() - t0) / 1000).toFixed(1)}s — ${state.analysis.newStyles.length} new, ${state.analysis.updates.length} updates`, "ok");
  } catch (e) {
    state.error = e.message;
    log(`Failed: ${e.message}`, "error");
  }

  state.syncing = false;
}

// ─── API Routes ──────────────────────────────────────────────────────────────

app.get("/api/status", (req, res) => {
  res.json(state);
});

app.post("/api/sync", async (req, res) => {
  if (state.syncing) return res.json({ ok: false, msg: "Already syncing" });
  doSync("manual-api");
  res.json({ ok: true, msg: "Sync started" });
});

app.get("/api/export", (req, res) => {
  if (!state.analysis) return res.status(400).json({ error: "No data yet" });
  const a = state.analysis;
  const hdr = ["Status", "Factory", "Production#", "PO NAME", "STYLE", "Ship Units", "Brand", "ETD", "Changes"];
  const rows = [hdr];
  a.newStyles.forEach(r => rows.push(["NEW", r.factory, r.prod, r.po, r.style, r.units ?? "", r.brand, r.etd, ""]));
  a.updates.forEach(u => rows.push(["UPDATED", u.factory, u.row.prod, u.row.po, u.style, u.row.units ?? "", u.row.brand, u.row.etd, u.changes.map(c => `${c.field}: ${c.from} → ${c.to}`).join("; ")]));
  a.missing.forEach(r => rows.push(["MISSING", "—", r.prod, r.po, r.style, r.units ?? "", r.brand, r.etd, ""]));

  const ws = XLSX.utils.aoa_to_sheet(rows);
  ws["!cols"] = [{ wch: 9 }, { wch: 5 }, { wch: 13 }, { wch: 44 }, { wch: 16 }, { wch: 10 }, { wch: 7 }, { wch: 11 }, { wch: 55 }];
  const wb = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(wb, ws, "Sync Report");
  const buf = XLSX.write(wb, { type: "buffer", bookType: "xlsx" });

  res.setHeader("Content-Type", "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet");
  res.setHeader("Content-Disposition", `attachment; filename=Ledger_Sync_${new Date().toISOString().slice(0, 10)}.xlsx`);
  res.send(buf);
});

// Serve frontend for all other routes
app.get("/api/debug/list-folder", async (req, res) => {
  try {
    const token = await getAccessToken();
    const folder = req.query.path || "/Production_Style_Ledger";
    const r = await fetch("https://api.dropboxapi.com/2/files/list_folder", {
      method: "POST",
      headers: { Authorization: `Bearer ${token}`, "Content-Type": "application/json" },
      body: JSON.stringify({ path: folder }),
    });
    const data = await r.json();
    const files = (data.entries || []).map(e => ({ name: e.name, path: e.path_display, type: e[".tag"] }));
    res.json({ folder, files });
  } catch(e) { res.status(500).json({ error: e.message }); }
});

app.get("*", (req, res) => {
  res.sendFile(path.join(__dirname, "public", "index.html"));
});

// ─── Start ───────────────────────────────────────────────────────────────────

const PORT = process.env.PORT || 3000;
app.listen(PORT, () => {
  console.log(`Production Sync running on port ${PORT}`);
  console.log(`Sync interval: every ${SYNC_MINUTES} minutes`);
  console.log(`FY1: ${PATHS.FY1}`);
  console.log(`NB:  ${PATHS.NB}`);
  console.log(`PC:  ${PATHS.PC}`);
  console.log(`David: ${PATHS.DAVID}`);

  // Initial sync
  doSync("startup");

  // Auto-sync interval
  setInterval(() => doSync("auto"), SYNC_MINUTES * 60000);
});
