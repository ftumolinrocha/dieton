
import express from "express";
import session from "express-session";
import multer from "multer";
import XLSX from "xlsx";
import path from "path";
import fs from "fs/promises";
import fssync from "fs";
import crypto from "crypto";
import bwipjs from "bwip-js";
import { fileURLToPath } from "url";

const __filename = fileURLToPath(import.meta.url);
const __dirname = path.dirname(__filename);

const PORT = process.env.PORT ? Number(process.env.PORT) : 4000;
// BUILD: sempre derive do package.json para não ter divergência (npm start vs badge na UI)
let BUILD = "0.0.0";
try {
  const pkgPath = path.join(__dirname, "package.json");
  const pkg = JSON.parse(fssync.readFileSync(pkgPath, "utf8"));
  if (pkg && pkg.version) {
    const v = String(pkg.version);
    // Mostra como "3.5" quando o package.json está "3.5.0" (padrão semver), para seguir seu padrão.
    const m = v.match(/^\s*(\d+)\.(\d+)\.0\s*$/);
    BUILD = m ? `${m[1]}.${m[2]}` : v;
  }
} catch {
  // fallback mantém 0.0.0
}
const CORE_PATH = path.join(__dirname, "data", "marmitaria.json");
const BD_DIR = path.join(__dirname, "bd");
// MRP agora fica dentro de /bd para facilitar backup/restauração do "BD"
const BD_MRP_PATH = path.join(BD_DIR, "mrp.json");
const BD_RAW_PATH = path.join(BD_DIR, "estoque_mp.json");
const BD_FG_PATH = path.join(BD_DIR, "estoque_pf.json");
const BD_UNITS_PATH = path.join(BD_DIR, "units.json");
// Pedidos de Venda / Pontos (freezers)
const BD_SALES_POINTS_PATH = path.join(BD_DIR, "sales_points.json");
const BD_SALES_ORDERS_PATH = path.join(BD_DIR, "sales_orders.json");
const BD_SALES_POINT_MOVES_PATH = path.join(BD_DIR, "sales_point_moves.json");
// Ordens de Compra arquivadas (separado do mrp.json para manter a lista principal limpa)
const BD_PO_ARCH_PATH = path.join(BD_DIR, "purchase_orders_archived.json");
// Ordens de Produção arquivadas (separado do mrp.json para manter a lista principal limpa)
const BD_OP_ARCH_PATH = path.join(BD_DIR, "production_orders_archived.json");
const BD_PHOTOS_DIR = path.join(BD_DIR, "photos");

// Simple in-process lock to serialize MRP file read/write (avoid partial reads while writing)
let __mrpLock = Promise.resolve();
function withMrpLock(fn){
  __mrpLock = __mrpLock.then(fn, fn);
  return __mrpLock;
}

// -------------------- qty helpers (OC/OP) --------------------
function unitKey(u) {
  return String(u || "").trim().toLowerCase();
}

// Normaliza unidades para uso interno (principalmente para quantização/decimais)
function fmtUnit(u){
  return unitKey(u);
}

function qtyDecimalsForUnit(u) {
  // UI trabalha com 3 casas para decimais (kg/l etc) e inteiro para "un".
  return unitKey(u) === "un" ? 0 : 3;
}

function quantizeQty(qty, unit) {
  const q = Number(qty || 0);
  if (!Number.isFinite(q)) return 0;
  const d = qtyDecimalsForUnit(unit);
  if (d === 0) return Math.round(q);
  const m = 10 ** d;
  // Arredonda de forma estável para evitar diferenças invisíveis no UI (ex.: 0.155127 vs 0.155)
  return Math.round(q * m) / m;
}

function calcFinalQty(ord, adj) {
  return (Number(ord || 0) || 0) + (Number(adj || 0) || 0);
}


const app = express();
app.use(express.json({ limit: "2mb" }));

// -------------------- robust JSON IO helpers --------------------
function sleep(ms) {
  return new Promise((r) => setTimeout(r, ms));
}

async function readJsonStrict(p) {
  const raw = await fs.readFile(p, "utf8");
  if (!raw || !raw.trim()) {
    const e = new Error("empty");
    e.code = "EMPTY";
    throw e;
  }
  return JSON.parse(raw);
}

// Per-file async mutex to prevent concurrent writes to the same file.
// Without this, two overlapping writes may fight over the same "*.tmp" path,
// causing ENOENT when one rename removes the temp file before the other copies it.
const __fileLocks = new Map();
function withFileLock(filePath, fn) {
  const prev = __fileLocks.get(filePath) || Promise.resolve();
  const next = prev
    .catch(() => {})
    .then(fn)
    .finally(() => {
      // Only clear if this promise is still the latest for this path
      if (__fileLocks.get(filePath) === next) __fileLocks.delete(filePath);
    });
  __fileLocks.set(filePath, next);
  return next;
}

async function safeReadJsonFile(filePath, {
  retries = 8,
  delayMs = 120,
  allowMissing = true,
  tryBak = true,
  ensureBdDir = false,
} = {}) {
  if (ensureBdDir) await ensureBD();
  const bakPath = `${filePath}.bak`;
  let lastErr = null;

  for (let i = 0; i <= retries; i++) {
    try {
      return await readJsonStrict(filePath);
    } catch (e) {
      lastErr = e;
      if (e?.code === "ENOENT" && allowMissing) return null;

      const transient =
        e?.code === "EMPTY" ||
        e instanceof SyntaxError ||
        /Unexpected end of JSON input/i.test(String(e?.message || ""));

      if (transient && i < retries) {
        await sleep(delayMs);
        continue;
      }
      break;
    }
  }

  if (tryBak) {
    try {
      return await readJsonStrict(bakPath);
    } catch (e) {
      lastErr = lastErr || e;
    }
  }

  const err = new Error(
    `Falha ao ler JSON em ${filePath}. O arquivo pode estar vazio/truncado (muito comum ao extrair/copy "por cima"). ` +
      `Feche o servidor, re-extraia o bd e tente novamente. Se existir ${bakPath}, ele pode ser usado para recuperar.`
  );
  err.cause = lastErr;
  throw err;
}

async function writeJsonAtomic(filePath, obj, { ensureBdDir = false } = {}) {
  return withFileLock(filePath, async () => {
    if (ensureBdDir) await ensureBD();

    // Ensure parent directory exists (fresh extraction may not include /data or /bd).
    await fs.mkdir(path.dirname(filePath), { recursive: true });

    // Unique temp name avoids collisions even if something calls this twice quickly.
    const tmpPath = `${filePath}.${process.pid}.${Date.now()}.${Math.random().toString(16).slice(2)}.tmp`;
    const bakPath = `${filePath}.bak`;
    const payload = JSON.stringify(obj, null, 2);

    // 1) escreve temp
    await fs.writeFile(tmpPath, payload, "utf8");

    // 2) backup (best-effort)
    // Use copy (not rename) so we never leave a window where the main file disappears.
    try {
      await fs.copyFile(filePath, bakPath);
    } catch (e) {
      // ignore (missing file, permission, etc.)
    }

    // 3) publica temp -> oficial
    try {
      // On Windows, rename is atomic but can fail in some edge cases; fallback keeps us safe.
      await fs.rename(tmpPath, filePath);
    } catch (e) {
      await fs.copyFile(tmpPath, filePath);
      await fs.unlink(tmpPath).catch(() => {});
    }
  });
}

// Health check (used by the front-end to detect if the local server is up)
app.get("/api/health", (req, res) => {
  res.json({ ok: true, build: BUILD, at: new Date().toISOString() });
});

// Build info (front-end badge sync)
app.get("/api/build", (req, res) => {
  res.json({ ok: true, build: BUILD, at: new Date().toISOString() });
});

app.use(
  session({
    secret: process.env.SESSION_SECRET || "marmitaria-local-dev-secret",
    resave: false,
    saveUninitialized: false,
    cookie: { httpOnly: true, sameSite: "lax" },
  })
);

app.use(express.static(path.join(__dirname, "public")));

// Fotos do PF (marmitas): servidas a partir de /bd/photos
app.use(
  "/photos",
  express.static(BD_PHOTOS_DIR, {
    etag: false,
    cacheControl: false,
    setHeaders(res) {
      res.setHeader("Cache-Control", "no-store");
    },
  })
);

// -------------------- DB helpers (JSON file) --------------------
async function readCore() {
  return safeReadJsonFile(CORE_PATH, { allowMissing: true, tryBak: true });
}

async function writeCore(db) {
  db.meta = db.meta || {};
  db.meta.updatedAt = new Date().toISOString();
  await writeJsonAtomic(CORE_PATH, db);
}

// -------------------- MRP (local JSON em /bd) --------------------
async function readMrp() {
  await ensureBD();
  const bakPath = `${BD_MRP_PATH}.bak`;

  const tryParse = async (p) => {
    const raw = await fs.readFile(p, "utf8");
    if (!raw || !raw.trim()) throw new Error("empty");
    return JSON.parse(raw);
  };

  // IMPORTANT: nunca trate JSON parse error como "sem arquivo".
  // Se falhar, tentamos o .bak; se ainda falhar, propagamos erro.
  return withMrpLock(async () => {
    try {
      return await tryParse(BD_MRP_PATH);
    } catch (e) {
      // Se o arquivo principal falhar, tenta o backup.
      try {
        return await tryParse(bakPath);
      } catch (_) {
        if (e.code == "ENOENT") return null;
        // se for vazio/truncado/corrompido e não houver backup válido,
        // deixe o erro aparecer (melhor do que apagar dados).
        throw e;
      }
    }
  });
}

async function writeMrp(mrp) {
  await ensureBD();
  mrp.meta = mrp.meta || {};
  mrp.meta.updatedAt = new Date().toISOString();

  // Optional granular wipe permissions (used for intentional deletions of OP/OC lists).
  // This is NOT persisted (removed before writing).
  const allowWipe = mrp?.meta?.allowWipe || {};

  // Safety: refuse accidental wipe (ex.: bug/erro de JS chamando PUT com arrays vazios)
  // unless explicitly requested via mrp.meta.forceReset = true.
  const force = !!mrp?.meta?.forceReset;
  if (!force) {
    try {
      const cur = await safeReadJsonFile(BD_MRP_PATH, { allowMissing: true, tryBak: true, ensureBdDir: true });
      const curRecipes = Array.isArray(cur?.recipes) ? cur.recipes.length : 0;
      const curPO = Array.isArray(cur?.productionOrders) ? cur.productionOrders.length : 0;
      const curOC = Array.isArray(cur?.purchaseOrders) ? cur.purchaseOrders.length : 0;

      const nextRecipes = Array.isArray(mrp?.recipes) ? mrp.recipes.length : 0;
      const nextPO = Array.isArray(mrp?.productionOrders) ? mrp.productionOrders.length : 0;
      const nextOC = Array.isArray(mrp?.purchaseOrders) ? mrp.purchaseOrders.length : 0;

      const wouldWipeRecipes = (curRecipes > 0 && nextRecipes === 0);
      const wouldWipePO = (curPO > 0 && nextPO === 0);
      const wouldWipeOC = (curOC > 0 && nextOC === 0);

      const force = !!mrp?.meta?.forceReset;
      const blocked =
        !force && (
          wouldWipeRecipes ||
          (wouldWipePO && !allowWipe.productionOrders) ||
          (wouldWipeOC && !allowWipe.purchaseOrders)
        );

      if (blocked) {
        // dump attempted payload for debugging
        const dump = path.join(BD_DIR, `mrp.rejected.${Date.now()}.json`);
        await fs.writeFile(dump, JSON.stringify(mrp, null, 2), "utf8").catch(() => {});
        const e = new Error(
          "Bloqueado: tentativa de sobrescrever o MRP com dados vazios. " +
          "Isso normalmente indica algum bug/erro no app. O arquivo atual foi preservado."
        );
        e.code = "MRP_WIPE_BLOCKED";
        throw e;
      }
    } catch (e) {
      // If the guard itself failed due to read error, bubble up (better than writing).
      throw e;
    }
  }

  // Never persist wipe permissions.
  if (mrp?.meta?.allowWipe) {
    try { delete mrp.meta.allowWipe; } catch (_) {}
  }

  // Use the common atomic writer (unique tmp + per-file lock)
  await writeJsonAtomic(BD_MRP_PATH, mrp, { ensureBdDir: true });
}

async function ensureMrpFile(legacyCoreMrp) {
  // Prefer BD file; if missing, migrate legacy MRP that used to live in /data/marmitaria.json
  let mrp = null;
  try {
    mrp = await readMrp();
  } catch (e) {
    // IMPORTANT: nunca sobrescrever o MRP por "default" se não conseguimos ler.
    // Isso evita perda de dados em caso de arquivo inválido.
    throw e;
  }

  let dirty = false;

  if (!mrp) {
    const legacy = legacyCoreMrp || null;
    const hasLegacyData = !!(
      legacy &&
      ((Array.isArray(legacy.recipes) && legacy.recipes.length) ||
        (Array.isArray(legacy.productionOrders) && legacy.productionOrders.length) ||
        (Array.isArray(legacy.purchaseOrders) && legacy.purchaseOrders.length))
    );

    mrp = {
      meta: { createdAt: new Date().toISOString(), migratedFromCore: hasLegacyData ? true : false },
      recipes: [],
      productionOrders: [],
      purchaseOrders: [],
    };

    if (hasLegacyData) {
      mrp.recipes = Array.isArray(legacy.recipes) ? legacy.recipes : [];
      mrp.productionOrders = Array.isArray(legacy.productionOrders) ? legacy.productionOrders : [];
      mrp.purchaseOrders = Array.isArray(legacy.purchaseOrders) ? legacy.purchaseOrders : [];
    }

    dirty = true;
  }

  // Guarantee shape (sem apagar dados existentes)
  if (!mrp.meta || typeof mrp.meta !== "object") {
    mrp.meta = {};
    dirty = true;
  }
  if (!Array.isArray(mrp.recipes)) {
    mrp.recipes = [];
    dirty = true;
  }
  if (!Array.isArray(mrp.productionOrders)) {
    mrp.productionOrders = [];
    dirty = true;
  }
  if (!Array.isArray(mrp.purchaseOrders)) {
    mrp.purchaseOrders = [];
    dirty = true;
  }

  // Numbering (OP/OC): keep UUIDs internally, but expose sequential numbers to the user.
  // Persist counters in mrp.meta to remain stable across restarts.
  const ensureNumbering = (list, metaKeyNext) => {
    const arr = Array.isArray(list) ? list : [];
    let maxN = 0;
    for (const o of arr) {
      const n = Number(o?.number);
      if (Number.isFinite(n) && n > maxN) maxN = n;
    }

    let changed = false;

    // Assign numbers to legacy orders that don't have it yet (stable-ish: old->new by createdAt)
    const missing = arr
      .filter((o) => !(Number.isFinite(Number(o?.number)) && Number(o.number) > 0))
      .sort((a, b) => String(a.createdAt || "").localeCompare(String(b.createdAt || "")));
    for (const o of missing) {
      maxN += 1;
      o.number = maxN;
      changed = true;
    }

    const prevNext = Number(mrp.meta[metaKeyNext]) || 0;
    const computedNext = maxN + 1;
    if (!prevNext) {
      mrp.meta[metaKeyNext] = computedNext;
      changed = true;
    } else if (prevNext <= maxN) {
      mrp.meta[metaKeyNext] = maxN + 1;
      changed = true;
    }

    return changed;
  };

  if (ensureNumbering(mrp.productionOrders, "nextOpNumber")) dirty = true;
  if (ensureNumbering(mrp.purchaseOrders, "nextOcNumber")) dirty = true;

  // Numbering (Lotes): usados para etiquetas de código de barras ao encerrar OP.
  // Mantém contador em mrp.meta.nextLotNumber (não depende de node_modules/BD externo).
  {
    let maxLot = 0;
    for (const o of (mrp.productionOrders || [])) {
      const n = Number(o?.lotNumber);
      if (Number.isFinite(n) && n > maxLot) maxLot = n;
    }
    const prevNext = Number(mrp.meta.nextLotNumber) || 0;
    const computedNext = maxLot + 1 || 1;
    if (!prevNext) {
      mrp.meta.nextLotNumber = computedNext;
      dirty = true;
    } else if (prevNext <= maxLot) {
      mrp.meta.nextLotNumber = maxLot + 1;
      dirty = true;
    }
  }


  if (dirty) await writeMrp(mrp);
  return mrp;
}

function pad6(n) {
  const x = Number(n);
  if (!Number.isFinite(x) || x <= 0) return "";
  return String(Math.trunc(x)).padStart(6, "0");
}

function padLot(n) {
  const x = Number(n);
  if (!Number.isFinite(x) || x <= 0) return "";
  return String(Math.trunc(x)).padStart(5, "0");
}

function formatLotCode(n) {
  const p = padLot(n);
  if (!p) return "";
  return `LOTE${p}`;
}

async function nextLotNumber(mrp) {
  mrp.meta = mrp.meta || {};
  // base counter
  let next = Number(mrp.meta.nextLotNumber) || 1;

  const used = new Set();
  const collect = (arr) => {
    for (const o of (arr || [])) {
      const n = Number(o?.lotNumber);
      if (Number.isFinite(n) && n > 0) used.add(Math.trunc(n));
    }
  };
  collect(mrp.productionOrders);

  // Considera arquivadas para evitar duplicidade
  try {
    const arch = await ensureProductionOrdersArchivedFile();
    collect(arch.productionOrders);
  } catch {}

  // Avança até achar um livre
  while (used.has(next)) next += 1;

  mrp.meta.nextLotNumber = next + 1;
  return next;
}

async function nextMrpNumber(mrp, kind) {
  mrp.meta = mrp.meta || {};
  const key = kind === "oc" ? "nextOcNumber" : "nextOpNumber";

  const used = new Set();
  const list = kind === "oc" ? (mrp.purchaseOrders || []) : (mrp.productionOrders || []);
  for (const o of list) {
    const n = Number(o?.number);
    if (Number.isFinite(n) && n > 0) used.add(Math.trunc(n));
  }

  // Considera também arquivadas, para evitar duplicidade de número.
  try {
    if (kind === "oc") {
      const arch = await ensurePurchaseOrdersArchivedFile();
      for (const o of (arch.purchaseOrders || [])) {
        const n = Number(o?.number);
        if (Number.isFinite(n) && n > 0) used.add(Math.trunc(n));
      }
    } else {
      const arch = await ensureProductionOrdersArchivedFile();
      for (const o of (arch.productionOrders || [])) {
        const n = Number(o?.number);
        if (Number.isFinite(n) && n > 0) used.add(Math.trunc(n));
      }
    }
  } catch (_) { /* noop */ }

  let cur = 1;
  while (used.has(cur)) cur += 1;

  mrp.meta[key] = cur + 1;
  return cur;
}

async function ensureBD() {
  await fs.mkdir(BD_DIR, { recursive: true });
  await fs.mkdir(BD_PHOTOS_DIR, { recursive: true });
}

// -------------------- Units (local JSON) --------------------
const DEFAULT_UNITS = [
  { v: "kg", l: "kg (kilograma)" },
  { v: "un", l: "un (unidade)" },
  { v: "ml", l: "ml (mililitro)" },
  { v: "l",  l: "l (litro)" },
];

async function readUnits() {
  return safeReadJsonFile(BD_UNITS_PATH, { allowMissing: true, tryBak: true, ensureBdDir: true });
}

async function writeUnits(unitsDb) {
  await ensureBD();
  unitsDb.meta = unitsDb.meta || {};
  unitsDb.meta.updatedAt = new Date().toISOString();
  await writeJsonAtomic(BD_UNITS_PATH, unitsDb, { ensureBdDir: true });
}

async function ensureUnitsFile() {
  let u = await readUnits();
  if (u && Array.isArray(u.units)) return u;
  u = { meta: { createdAt: new Date().toISOString() }, units: DEFAULT_UNITS };
  await writeUnits(u);
  return u;
}

// -------------------- Pedidos de Venda / Pontos (local JSON em /bd) --------------------
function pad3(n){
  const v = String(Math.trunc(Number(n) || 0));
  return v.padStart(3, "0");
}

async function readSalesPoints() {
  return safeReadJsonFile(BD_SALES_POINTS_PATH, { allowMissing: true, tryBak: true, ensureBdDir: true });
}

async function writeSalesPoints(db){
  await ensureBD();
  db = db || {};
  db.meta = db.meta || {};
  db.meta.updatedAt = new Date().toISOString();
  await writeJsonAtomic(BD_SALES_POINTS_PATH, db, { ensureBdDir: true });
}

async function ensureSalesPointsFile(){
  let db = await readSalesPoints();
  if (!db || !Array.isArray(db.points)) {
    db = { meta: { createdAt: new Date().toISOString(), nextPointCode: 1 }, points: [] };
    await writeSalesPoints(db);
    return db;
  }
  db.meta = db.meta || {};
  if (!db.meta.createdAt) db.meta.createdAt = new Date().toISOString();

  const out = sanitizeSalesPointsDb(db);
  if (out.changed) {
    await writeSalesPoints(out.db);
  }
  return out.db;
}

function normalizeSalesPointCode(raw){
  // IMPORTANTE: esta função deve ser SÍNCRONA. (Bug v4.52: async sem await gerava [object Object])
  const s = String(raw ?? '').trim();
  if (!s) return '';
  // Aceita: P001, p1, 001, 1
  let m = s.match(/^P\s*(\d{1,6})$/i);
  if (!m) m = s.match(/^(\d{1,6})$/);
  if (!m) return '';
  const num = Number(m[1]);
  if (!Number.isFinite(num) || num <= 0) return '';
  if (num <= 999) return `P${pad3(num)}`;
  return `P${Math.trunc(num)}`;
}

function parseSalesPointCodeNumber(code){
  const m = String(code ?? '').trim().match(/^P\s*(\d{1,6})$/i);
  if (!m) return null;
  const num = Number(m[1]);
  if (!Number.isFinite(num) || num <= 0) return null;
  return Math.trunc(num);
}

function buildSalesPointCode(num){
  const n = Math.trunc(Number(num) || 0);
  if (!Number.isFinite(n) || n <= 0) return '';
  if (n <= 999) return `P${pad3(n)}`;
  return `P${n}`;
}

function sanitizeSalesPointsDb(db){
  db = db || {};
  db.meta = db.meta || {};
  db.points = Array.isArray(db.points) ? db.points : [];

  let changed = false;
  const usedNums = new Set();
  const usedCodes = new Set();

  // 1) normaliza códigos existentes e detecta duplicados/invalidos
  for (const p of db.points) {
    const before = String(p.code ?? '');
    const norm = normalizeSalesPointCode(before);
    if (norm && !usedCodes.has(norm)) {
      if (norm !== before) {
        p.code = norm;
        changed = true;
      }
      usedCodes.add(norm);
      const n = parseSalesPointCodeNumber(norm);
      if (n != null) usedNums.add(n);
      p.__needsCode = false;
    } else {
      p.__needsCode = true;
    }
  }

  // 2) reatribui códigos inválidos/duplicados sem duplicar
  let next = 1;
  const takeNext = () => {
    while (usedNums.has(next)) next += 1;
    const code = buildSalesPointCode(next);
    usedNums.add(next);
    usedCodes.add(code);
    next += 1;
    return code;
  };

  for (const p of db.points) {
    if (!p.__needsCode) continue;
    const newCode = takeNext();
    if (String(p.code || '') !== newCode) {
      p.code = newCode;
      changed = true;
    }
    p.__needsCode = false;
  }

  // 3) ajusta próximo contador para evitar colisões
  const maxUsed = usedNums.size ? Math.max(...Array.from(usedNums)) : 0;
  const desiredNext = Math.max(Number(db.meta.nextPointCode) || 1, maxUsed + 1);
  if ((Number(db.meta.nextPointCode) || 1) !== desiredNext) {
    db.meta.nextPointCode = desiredNext;
    changed = true;
  }

  // cleanup
  for (const p of db.points) {
    if (p && Object.prototype.hasOwnProperty.call(p, '__needsCode')) delete p.__needsCode;
  }

  return { db, changed };
}

function nextSalesPointCode(db){
  db.meta = db.meta || {};
  const used = new Set();
  for (const p of (db.points || [])) {
    const n = parseSalesPointCodeNumber(p.code);
    if (n != null) used.add(n);
  }
  let cur = Number(db.meta.nextPointCode) || 1;
  if (cur < 1) cur = 1;
  while (used.has(cur)) cur += 1;
  db.meta.nextPointCode = cur + 1;
  return buildSalesPointCode(cur);
}

async function readSalesOrders(){
  return safeReadJsonFile(BD_SALES_ORDERS_PATH, { allowMissing: true, tryBak: true, ensureBdDir: true });
}

async function writeSalesOrders(db){
  await ensureBD();
  db = db || {};
  db.meta = db.meta || {};
  db.meta.updatedAt = new Date().toISOString();
  await writeJsonAtomic(BD_SALES_ORDERS_PATH, db, { ensureBdDir: true });
}

function getSalesOrderSeries(o){
  const s = String(o?.series || '').trim().toUpperCase();
  if (s === 'PV' || s === 'PVR') return s;
  const t = String(o?.type || '').trim().toUpperCase();
  return t === 'QUICK' ? 'PVR' : 'PV';
}

function formatSalesOrderCode(series, number){
  const ser = String(series || '').trim().toUpperCase();
  const n = pad6(number);
  if (!n) return '';
  const s = (ser === 'PVR') ? 'PVR' : 'PV';
  return `${s}${n}`;
}

function parseSalesOrderNumberInput(raw){
  const s = String(raw ?? '').trim();
  if (!s) return null;
  // Aceita: 1, 001, PV001, PV000001, PVR001...
  const m = s.match(/^(?:PV|PVR)?\s*(\d{1,6})$/i);
  if (!m) return null;
  const num = Number(m[1]);
  if (!Number.isFinite(num) || num <= 0) return null;
  return Math.trunc(num);
}

function syncSalesOrdersMeta(db){
  db = db || {};
  db.meta = db.meta || {};
  db.orders = Array.isArray(db.orders) ? db.orders : [];

  let changed = false;

  // Normaliza series + archived (compat)
  for (const o of db.orders) {
    const ser = getSalesOrderSeries(o);
    if (o.series !== ser) {
      o.series = ser;
      changed = true;
    }
    if (typeof o.archived !== 'boolean') {
      o.archived = !!o.archived;
      changed = true;
    }
  }

  let maxPv = 0;
  let maxPvr = 0;
  for (const o of db.orders) {
    const ser = getSalesOrderSeries(o);
    const n = Number(o.number);
    if (!Number.isFinite(n) || n <= 0) continue;
    const nn = Math.trunc(n);
    if (ser === 'PVR') maxPvr = Math.max(maxPvr, nn);
    else maxPv = Math.max(maxPv, nn);
  }

  const prevPv = Number(db.meta.nextPvNumber) || 0;
  const prevPvr = Number(db.meta.nextPvrNumber) || 0;

  const nextPv = Math.max(prevPv, maxPv + 1, 1);
  const nextPvr = Math.max(prevPvr, maxPvr + 1, 1);

  if (Number(db.meta.nextPvNumber) !== nextPv) {
    db.meta.nextPvNumber = nextPv;
    changed = true;
  }
  if (Number(db.meta.nextPvrNumber) !== nextPvr) {
    db.meta.nextPvrNumber = nextPvr;
    changed = true;
  }

  return changed;
}

async function ensureSalesOrdersFile(){
  let db = await readSalesOrders();
  if (!db || !Array.isArray(db.orders)) {
    db = { meta: { createdAt: new Date().toISOString(), nextPvNumber: 1, nextPvrNumber: 1 }, orders: [] };
    await writeSalesOrders(db);
    return db;
  }

  db.meta = db.meta || {};
  if (!db.meta.createdAt) db.meta.createdAt = new Date().toISOString();

  const changed = syncSalesOrdersMeta(db);
  if (changed) await writeSalesOrders(db);
  return db;
}

function nextSalesOrderNumber(db, series){
  db.meta = db.meta || {};
  db.orders = Array.isArray(db.orders) ? db.orders : [];
  const ser = (String(series || 'PV').trim().toUpperCase() === 'PVR') ? 'PVR' : 'PV';
  const metaKey = (ser === 'PVR') ? 'nextPvrNumber' : 'nextPvNumber';

  const used = new Set();
  for (const o of db.orders) {
    if (getSalesOrderSeries(o) !== ser) continue;
    const n = Number(o.number);
    if (Number.isFinite(n) && n > 0) used.add(Math.trunc(n));
  }

  let cur = Number(db.meta[metaKey]) || 1;
  if (cur < 1) cur = 1;
  while (used.has(cur)) cur += 1;
  db.meta[metaKey] = cur + 1;
  return cur;
}

async function readSalesPointMoves(){
  return safeReadJsonFile(BD_SALES_POINT_MOVES_PATH, { allowMissing: true, tryBak: true, ensureBdDir: true });
}

async function writeSalesPointMoves(db){
  await ensureBD();
  db = db || {};
  db.meta = db.meta || {};
  db.meta.updatedAt = new Date().toISOString();
  await writeJsonAtomic(BD_SALES_POINT_MOVES_PATH, db, { ensureBdDir: true });
}

async function ensureSalesPointMovesFile(){
  let db = await readSalesPointMoves();
  if (db && Array.isArray(db.moves)) return db;
  db = { meta: { createdAt: new Date().toISOString() }, moves: [] };
  await writeSalesPointMoves(db);
  return db;
}

function computePointInventory(pointMoves, pointId){
  const pid = String(pointId);
  const map = new Map(); // itemId -> qty
  for (const mv of (pointMoves.moves || [])) {
    if (String(mv.pointId) !== pid) continue;
    const itemId = String(mv.itemId);
    const unit = mv.unit;
    const cur = Number(map.get(itemId) || 0);
    const delta = quantizeQty(mv.delta, unit);
    map.set(itemId, quantizeQty(cur + delta, unit));
  }
  return map;
}

// -------------------- Purchase Orders Archived (local JSON em /bd) --------------------
async function readPurchaseOrdersArchived() {
  return safeReadJsonFile(BD_PO_ARCH_PATH, { allowMissing: true, tryBak: true, ensureBdDir: true });
}

async function writePurchaseOrdersArchived(dbArch) {
  await ensureBD();
  dbArch = dbArch || {};
  dbArch.meta = dbArch.meta || {};
  dbArch.meta.updatedAt = new Date().toISOString();
  await writeJsonAtomic(BD_PO_ARCH_PATH, dbArch, { ensureBdDir: true });
}

async function ensurePurchaseOrdersArchivedFile() {
  let a = await readPurchaseOrdersArchived();
  if (a && Array.isArray(a.purchaseOrders)) return a;
  a = { meta: { createdAt: new Date().toISOString() }, purchaseOrders: [] };
  await writePurchaseOrdersArchived(a);
  return a;
}


// -------------------- Production Orders Archived (local JSON em /bd) --------------------
async function readProductionOrdersArchived() {
  return safeReadJsonFile(BD_OP_ARCH_PATH, { allowMissing: true, tryBak: true, ensureBdDir: true });
}

async function writeProductionOrdersArchived(dbArch) {
  await ensureBD();
  dbArch = dbArch || {};
  dbArch.meta = dbArch.meta || {};
  dbArch.meta.updatedAt = new Date().toISOString();
  await writeJsonAtomic(BD_OP_ARCH_PATH, dbArch, { ensureBdDir: true });
}

async function ensureProductionOrdersArchivedFile() {
  let a = await readProductionOrdersArchived();
  if (a && Array.isArray(a.productionOrders)) return a;
  a = { meta: { createdAt: new Date().toISOString() }, productionOrders: [] };
  await writeProductionOrdersArchived(a);
  return a;
}

function normalizeUnitCode(v) {
  return String(v || "").trim();
}

function normalizeUnitLabel(l, v) {
  const s = String(l || "").trim();
  if (s) return s;
  return v;
}

function stockPath(type) {
  return type === "fg" ? BD_FG_PATH : BD_RAW_PATH;
}

async function readStock(type) {
  const p = stockPath(type);
  return safeReadJsonFile(p, { allowMissing: true, tryBak: true, ensureBdDir: true });
}

async function writeStock(type, stockDb) {
  await ensureBD();
  const p = stockPath(type);
  stockDb.meta = stockDb.meta || {};
  stockDb.meta.updatedAt = new Date().toISOString();
  stockDb.meta.type = type;
  await writeJsonAtomic(p, stockDb, { ensureBdDir: true });
}

async function ensureStockFile(type) {
  let s = await readStock(type);
  if (s) {
    // backfill item codes (MPxxx / PFxxx)
    let changed = false;
    for (const it of (s.items || [])) {
      if (!it.code) { it.code = nextItemCode(s.items, type); changed = true; }
      if (type === "raw") {
        if (it.lossPercent === undefined) { it.lossPercent = 0; changed = true; }
      }
      if (type === "fg") {
        if (it.salePrice === undefined) { it.salePrice = 0; changed = true; }
      }
    }
    if (changed) await writeStock(type, s);
    return s;
  }
  s = { meta: { createdAt: new Date().toISOString(), type }, items: [], movements: [] };
  await writeStock(type, s);
  return s;
}


function hashPassword(password, salt) {
  // Local-only: PBKDF2 (built-in crypto), OK for MVP
  // Salt is stored as hex string; use raw bytes for stable cross-language hashing.
  const s = String(salt || "");
  const isHex = /^[0-9a-f]+$/i.test(s) && s.length % 2 === 0;
  const saltBuf = isHex ? Buffer.from(s, "hex") : Buffer.from(s, "utf8");
  const derived = crypto.pbkdf2Sync(String(password), saltBuf, 120000, 32, "sha256");
  return derived.toString("hex");
}

function requireAuth(req, res, next) {
  if (!req.session?.userId) return res.status(401).json({ error: "not_authenticated" });
  next();
}


// Gera código de barras (Code128) para impressão de etiquetas
app.get("/api/barcode", requireAuth, async (req, res) => {
  const text = String(req.query.text || "").trim();
  if (!text) return res.status(400).json({ error: "missing_text" });

  try {
    const png = await bwipjs.toBuffer({
      bcid: "code128",
      text,
      scale: 3,
      height: 10,
      includetext: false,
      backgroundcolor: "FFFFFF",
    });
    res.setHeader("Content-Type", "image/png");
    res.setHeader("Cache-Control", "no-store");
    res.end(png);
  } catch (e) {
    res.status(400).json({ error: "barcode_error", message: String(e?.message || e) });
  }
});

// Some sensitive actions (import/export) require user+password confirmation every time.
// The UI calls /api/reauth immediately before the protected action; the next protected
// request consumes this single-use authorization.
function requireReauthOnce(req, res, next) {
  if (!req.session?.userId) return res.status(401).json({ error: "not_authenticated" });

  const ra = req.session.reauth;
  const now = Date.now();
  if (!ra || !ra.exp || now > ra.exp || !Number.isFinite(ra.remaining) || ra.remaining <= 0) {
    return res.status(401).json({ error: "reauth_required" });
  }

  ra.remaining = Number(ra.remaining) - 1;
  if (ra.remaining <= 0) {
    try { delete req.session.reauth; } catch (_) {}
  } else {
    req.session.reauth = ra;
  }
  next();
}


// Permissions helper: admin overrides everything. Other permissions are boolean flags on user.permissions.
function userHasPermission(user, perm) {
  if (!user || typeof user !== "object") return false;
  const p = (user.permissions && typeof user.permissions === "object" && !Array.isArray(user.permissions)) ? user.permissions : {};
  if (p.admin === true) return true;
  if (p[perm] === true) return true;
  // Legacy: older builds used a single 'mrp' flag for Receitas/OP/OC/Custos
  if ((perm === 'recipes' || perm === 'op' || perm === 'oc' || perm === 'costs') && typeof p[perm] !== 'boolean' && p.mrp === true) return true;
  return false;
}

function requirePerm(perm) {
  return async (req, res, next) => {
    if (!req.session?.userId) return res.status(401).json({ error: "not_authenticated" });
    const db = await ensureDB();
    const user = (db.users || []).find((u) => u && u.id === req.session.userId);
    if (!user) return res.status(401).json({ error: "not_authenticated" });
    if (!userHasPermission(user, perm)) return res.status(403).json({ error: "forbidden" });
    next();
  };
}

// Require any of a set of permissions (admin always passes).
function requireAnyPerm(perms) {
  const list = Array.isArray(perms) ? perms : [perms];
  return async (req, res, next) => {
    if (!req.session?.userId) return res.status(401).json({ error: "not_authenticated" });
    const db = await ensureDB();
    const user = (db.users || []).find((u) => u && u.id === req.session.userId);
    if (!user) return res.status(401).json({ error: "not_authenticated" });
    const ok = list.some((k) => userHasPermission(user, k));
    if (!ok) return res.status(403).json({ error: "forbidden" });
    next();
  };
}



function newId(prefix = "") {
  return prefix + crypto.randomUUID();
}


async function migrateDB(db) {
  // Backward-compatible defaults
  db.users = Array.isArray(db.users) ? db.users : [];
  db.inventory = db.inventory || { items: [], movements: [] };
  db.inventory.items = Array.isArray(db.inventory.items) ? db.inventory.items : [];
  db.inventory.movements = Array.isArray(db.inventory.movements) ? db.inventory.movements : [];
  db.mrp = db.mrp || { recipes: [], productionOrders: [] };
  db.mrp.recipes = Array.isArray(db.mrp.recipes) ? db.mrp.recipes : [];
  db.mrp.productionOrders = Array.isArray(db.mrp.productionOrders) ? db.mrp.productionOrders : [];

  // Default: items without type are raw materials/insumos
  for (const it of stockDb.items) {
    if (!it.type) it.type = "raw"; // raw | fg
  }

  // Ensure existing recipes have an output item
  const itemsByName = new Map(db.inventory.items.map((i) => [String(i.name||"").toLowerCase(), i]));
  for (const r of db.mrp.recipes) {
    if (!r.outputItemId) {
      // Try match by name
      const match = itemsByName.get(String(r.name||"").toLowerCase());
      if (match && match.type === "fg") {
        r.outputItemId = match.id;
      }
    }
  }

  return db;
}




function normalizeItemCodeServer(input, type){
  const prefix = type === "fg" ? "PF" : "MP";
  let c = String(input ?? "").trim().toUpperCase();
  if (!c) return "";
  if (/^\d+$/.test(c)){
    const n = parseInt(c, 10);
    if (!Number.isFinite(n) || n <= 0) return "";
    return prefix + String(n).padStart(3, "0");
  }
  if (c.startsWith(prefix)){
    const tail = c.slice(prefix.length).trim();
    if (/^\d+$/.test(tail)){
      const n = parseInt(tail, 10);
      if (!Number.isFinite(n) || n <= 0) return "";
      return prefix + String(n).padStart(3, "0");
    }
  }
  return c;
}

function nextItemCode(items, type) {
  const prefix = type === "fg" ? "PF" : "MP";
  const nums = (items || []).map((i) => {
    const c = normalizeItemCodeServer(i.code, type);
    if (!c.startsWith(prefix)) return NaN;
    const n = parseInt(c.slice(prefix.length), 10);
    return Number.isFinite(n) ? n : NaN;
  }).filter((n) => Number.isFinite(n) && n > 0);

  const set = new Set(nums);
  const max = nums.length ? Math.max(...nums) : 0;
  let cand = 1;
  while (cand <= max + 1) {
    if (!set.has(cand)) break;
    cand += 1;
  }
  return prefix + String(cand).padStart(3, "0");
}

async function ensureDB() {
  // Core DB (users + mrp)
  let db = await readCore();

  if (!db) {
    // create new core
    const salt = crypto.randomBytes(16).toString("hex");
    const defaultUser = {
      id: "Felipe",
      name: "Felipe",
      salt,
      passwordHash: hashPassword("Mestre", salt),
      permissions: { admin: true, inventory: true, mrp: true, sales: true, canReset: true, canImportExport: true },
      createdAt: new Date().toISOString(),
    };

    db = {
      meta: { createdAt: new Date().toISOString(), app: "dieton-mrp" },
      users: [defaultUser],
      mrp: {
        recipes: [],
        productionOrders: [],
      },
    };

    await writeCore(db);
  } else {
    // normalize core
    db.users = Array.isArray(db.users) ? db.users : [];

    // If core has no users, seed default login
    if (db.users.length === 0) {
      const salt = crypto.randomBytes(16).toString("hex");
      db.users.push({
        id: "Felipe",
        name: "Felipe",
        salt,
        passwordHash: hashPassword("Mestre", salt),
        createdAt: new Date().toISOString(),
      });
      await writeCore(db);
    }
  }

  
  // Permissions (backward compatible): users created before v4.55 may not have permissions.
  // Rule: first user defaults to admin/full access; subsequent users default to module access only.
  try {
    let changedPerms = false;
    const defaultFirst = { admin: true, inventory: true, sales: true, recipes: true, op: true, oc: true, costs: true, canReset: true, canImportExport: true };
    const defaultOther = { admin: false, inventory: true, sales: true, recipes: true, op: true, oc: true, costs: true, canReset: false, canImportExport: false };

    db.users = Array.isArray(db.users) ? db.users : [];
    for (let i = 0; i < db.users.length; i += 1) {
      const u = db.users[i];
      if (!u || typeof u !== "object") continue;
      if (!u.permissions || typeof u.permissions !== "object" || Array.isArray(u.permissions)) {
        u.permissions = {};
        changedPerms = true;
      }
      const base = i === 0 ? defaultFirst : defaultOther;
      // Legacy: older builds used a single 'mrp' flag for Receitas/OP/OC/Custos. If present and new flags are missing, mirror it.
      if (typeof u.permissions.mrp === 'boolean') {
        const legacy = u.permissions.mrp;
        for (const k of ['recipes','op','oc','costs']) {
          if (typeof u.permissions[k] !== 'boolean') {
            u.permissions[k] = legacy;
            changedPerms = true;
          }
        }
      }
      for (const k of Object.keys(base)) {
        if (typeof u.permissions[k] !== "boolean") {
          u.permissions[k] = base[k];
          changedPerms = true;
        }
      }
      if (!u.name) {
        u.name = u.id;
        changedPerms = true;
      }
    }
    if (changedPerms) await writeCore(db);
  } catch (_) {}

// MRP (agora em /bd/mrp.json; migra automaticamente do legado no core)
  db.mrp = await ensureMrpFile(db.mrp);

  // Normaliza OPs legadas: remove status HOLD/READY (não existe mais "Em espera")
  try {
    let changed = false;
    for (const op of (db.mrp.productionOrders || [])) {
      const s = String(op?.status || "").toUpperCase();
      if (s === "HOLD" || s === "READY") {
        op.status = "ISSUED";
        changed = true;
      }
    }
    if (changed) await writeMrp(db.mrp);
  } catch (_) {}

  // Ensure stock files exist
  const rawStock = await ensureStockFile("raw");
  const fgStock = await ensureStockFile("fg");
  // Ensure units file exists
  await ensureUnitsFile();
  // Ensure archived purchase orders file exists
  await ensurePurchaseOrdersArchivedFile();
  // Ensure archived production orders file exists
  await ensureProductionOrdersArchivedFile();

  // One-time migration: if old core had inventory, split to stock files (best effort)
  // (kept backward compatible; new versions ignore db.inventory)
  if (db.inventory && (Array.isArray(db.inventory.items) || Array.isArray(db.inventory.movements))) {
    const items = Array.isArray(db.inventory.items) ? db.inventory.items : [];
    const movs = Array.isArray(db.inventory.movements) ? db.inventory.movements : [];
    const itemsById = new Map(items.map(i => [i.id, i]));
    for (const it of items) {
      const t = it.type === "fg" ? "fg" : "raw";
      const target = t === "fg" ? fgStock : rawStock;
      if (!target.items.find(x => x.id === it.id)) {
        target.items.push({ ...it, type: t });
      }
    }
    for (const mv of movs) {
      const it = itemsById.get(mv.itemId);
      const t = (it?.type === "fg") ? "fg" : "raw";
      const target = t === "fg" ? fgStock : rawStock;
      target.movements.push(mv);
    }
    await writeStock("raw", rawStock);
    await writeStock("fg", fgStock);
    // keep db.inventory but do not use; no destructive change to avoid surprises
  }

  return db;
}

function computeInventory(stockDb) {
  const stock = new Map();
  for (const it of (stockDb.items || [])) stock.set(String(it.id), 0);

  for (const mv of (stockDb.movements || [])) {
    const key = String(mv.itemId);
    const prev = stock.get(key) ?? 0;
    let next = prev;
    if (mv.type === "in") next = prev + mv.qty;
    else if (mv.type === "out") next = prev - mv.qty;
    else if (mv.type === "adjust") next = mv.qty; // absolute set
    stock.set(key, next);
  }

  return stock;
}

// -------------------- Auth --------------------
app.get("/api/me", async (req, res) => {
  if (!req.session?.userId) return res.json({ authenticated: false });
  const db = await ensureDB();
  const user = db.users.find((u) => u.id === req.session.userId);
  if (!user) return res.json({ authenticated: false });
  res.json({ authenticated: true, user: { id: user.id, name: user.name, permissions: user.permissions || {} } });
});

app.post("/api/login", async (req, res) => {
  const { userId, password } = req.body || {};
  if (!userId || !password) return res.status(400).json({ error: "missing_credentials" });

  const db = await ensureDB();
  const user = db.users.find((u) => u.id === userId);
  if (!user) return res.status(401).json({ error: "invalid_credentials" });

  const candidate = hashPassword(password, user.salt);
  if (candidate !== user.passwordHash) return res.status(401).json({ error: "invalid_credentials" });

  req.session.userId = user.id;
  // Ensure session is persisted before replying (avoids races on immediate follow-up requests)
  if (typeof req.session.save === "function") {
    return req.session.save(() => {
      res.json({ ok: true, user: { id: user.id, name: user.name, permissions: user.permissions || {} } });
    });
  }
  res.json({ ok: true, user: { id: user.id, name: user.name, permissions: user.permissions || {} } });
});

// Re-auth (single-use) for sensitive actions like XLSX import/export.
// Requires an active session and the same userId; on success, enables exactly one
// subsequent protected request within the next 60 seconds.
app.post("/api/reauth", requireAuth, async (req, res) => {
  const { userId, password } = req.body || {};
  if (!userId || !password) return res.status(400).json({ error: "missing_credentials" });

  const sessionUser = String(req.session.userId || "");
  if (String(userId) !== sessionUser) return res.status(401).json({ error: "invalid_credentials" });

  const db = await ensureDB();
  const user = db.users.find((u) => u.id === sessionUser);
  if (!user) return res.status(401).json({ error: "invalid_credentials" });

  const candidate = hashPassword(password, user.salt);
  if (candidate !== user.passwordHash) return res.status(401).json({ error: "invalid_credentials" });

  req.session.reauth = { exp: Date.now() + 60 * 1000, remaining: 1 };
  // Persist session before replying (critical for immediate download via iframe/window)
  if (typeof req.session.save === "function") {
    return req.session.save(() => res.json({ ok: true }));
  }
  res.json({ ok: true });
});


app.post("/api/logout", (req, res) => {
  req.session.destroy(() => res.json({ ok: true }));
});

// -------------------- Users (Admin) --------------------
// NOTE: For safety, user management requires ADMIN permission + reauth (single-use).
const ALLOWED_USER_PERMS = ["admin", "inventory", "sales", "recipes", "op", "oc", "costs", "canReset", "canImportExport", "mrp"];

function sanitizeUserPermissions(input, fallback = {}) {
  const out = {};
  const src = (input && typeof input === "object" && !Array.isArray(input)) ? input : {};
  for (const k of ALLOWED_USER_PERMS) {
    if (typeof src[k] === "boolean") out[k] = src[k];
    else if (typeof fallback[k] === "boolean") out[k] = fallback[k];
  }
  // If admin, force full access flags on (keeps UX simple and avoids accidental lockouts).
  if (out.admin === true) {
    out.inventory = true;
    out.sales = true;
    out.recipes = true;
    out.op = true;
    out.oc = true;
    out.costs = true;
    out.canReset = true;
    out.canImportExport = true;
    out.mrp = true; // legacy flag
  }
  return out;
}

function isValidUserId(id) {
  const s = String(id || "").trim();
  if (s.length < 2 || s.length > 32) return false;
  // allow letters, numbers, underscore, dash, dot
  return /^[A-Za-z0-9_.-]+$/.test(s);
}

function userPublic(u) {
  return { id: u.id, name: u.name, permissions: u.permissions || {}, createdAt: u.createdAt || null, updatedAt: u.updatedAt || null };
}

function countAdmins(users) {
  return (users || []).filter((u) => u && u.permissions && u.permissions.admin === true).length;
}

app.get("/api/users", requireAuth, requirePerm("admin"), async (req, res) => {
  const db = await ensureDB();
  const users = (db.users || []).map(userPublic);
  res.json({ ok: true, users });
});

app.post("/api/users", requireAuth, requirePerm("admin"), requireReauthOnce, async (req, res) => {
  const { id, name, password, permissions } = req.body || {};
  const userId = String(id || "").trim();
  const userName = String(name || userId || "").trim();

  if (!isValidUserId(userId)) return res.status(400).json({ error: "invalid_user_id" });
  if (!userName) return res.status(400).json({ error: "missing_name" });
  if (!password || String(password).length < 4) return res.status(400).json({ error: "weak_password" });

  const db = await ensureDB();
  if ((db.users || []).some((u) => u && u.id === userId)) return res.status(409).json({ error: "user_exists" });

  const salt = crypto.randomBytes(16).toString("hex");
  const newUser = {
    id: userId,
    name: userName,
    salt,
    passwordHash: hashPassword(String(password), salt),
    permissions: sanitizeUserPermissions(permissions, { admin: false, inventory: true, mrp: true, sales: true, canReset: false, canImportExport: false }),
    createdAt: new Date().toISOString(),
    updatedAt: new Date().toISOString(),
  };

  db.users.push(newUser);
  await writeCore(db);
  res.json({ ok: true, user: userPublic(newUser) });
});

app.put("/api/users/:id", requireAuth, requirePerm("admin"), requireReauthOnce, async (req, res) => {
  const targetId = String(req.params.id || "").trim();
  const { name, password, permissions } = req.body || {};

  if (!isValidUserId(targetId)) return res.status(400).json({ error: "invalid_user_id" });

  const db = await ensureDB();
  const idx = (db.users || []).findIndex((u) => u && u.id === targetId);
  if (idx < 0) return res.status(404).json({ error: "user_not_found" });

  const u = db.users[idx];
  const newName = (name !== undefined) ? String(name || "").trim() : u.name;
  if (!newName) return res.status(400).json({ error: "missing_name" });

  // Update permissions (if provided)
  if (permissions !== undefined) {
    const nextPerms = sanitizeUserPermissions(permissions, u.permissions || {});
    // Prevent deleting last admin (including self demotion)
    const wasAdmin = !!(u.permissions && u.permissions.admin === true);
    const willAdmin = !!(nextPerms && nextPerms.admin === true);
    if (wasAdmin && !willAdmin) {
      const admins = countAdmins(db.users);
      if (admins <= 1) return res.status(409).json({ error: "cannot_remove_last_admin" });
    }
    u.permissions = nextPerms;
  }

  // Update password (optional)
  if (password !== undefined && String(password || "").length > 0) {
    if (String(password).length < 4) return res.status(400).json({ error: "weak_password" });
    const salt = crypto.randomBytes(16).toString("hex");
    u.salt = salt;
    u.passwordHash = hashPassword(String(password), salt);
  }

  u.name = newName;
  u.updatedAt = new Date().toISOString();

  await writeCore(db);
  res.json({ ok: true, user: userPublic(u) });
});

app.delete("/api/users/:id", requireAuth, requirePerm("admin"), requireReauthOnce, async (req, res) => {
  const targetId = String(req.params.id || "").trim();
  if (!isValidUserId(targetId)) return res.status(400).json({ error: "invalid_user_id" });

  // Cannot delete yourself
  if (String(req.session.userId || "") === targetId) return res.status(409).json({ error: "cannot_delete_self" });

  const db = await ensureDB();
  const idx = (db.users || []).findIndex((u) => u && u.id === targetId);
  if (idx < 0) return res.status(404).json({ error: "user_not_found" });

  // Prevent deleting last admin
  const target = db.users[idx];
  const isAdmin = !!(target.permissions && target.permissions.admin === true);
  if (isAdmin) {
    const admins = countAdmins(db.users);
    if (admins <= 1) return res.status(409).json({ error: "cannot_delete_last_admin" });
  }

  db.users.splice(idx, 1);
  await writeCore(db);
  res.json({ ok: true });
});

// -------------------- Units --------------------

app.get("/api/units", requireAuth, async (req, res) => {
  await ensureDB();
  const u = await ensureUnitsFile();
  res.json({ units: Array.isArray(u.units) ? u.units : [] });
});

app.post("/api/units", requireAuth, async (req, res) => {
  const { v, l } = req.body || {};
  const code = normalizeUnitCode(v);
  if (!code) return res.status(400).json({ error: "missing_unit_code" });

  await ensureDB();
  const unitsDb = await ensureUnitsFile();
  unitsDb.units = Array.isArray(unitsDb.units) ? unitsDb.units : [];
  if (unitsDb.units.find((u) => u.v === code)) return res.status(409).json({ error: "unit_exists" });

  const label = normalizeUnitLabel(l, code);
  unitsDb.units.push({ v: code, l: label });
  await writeUnits(unitsDb);
  res.json({ ok: true, unit: { v: code, l: label } });
});

app.put("/api/units/:code", requireAuth, async (req, res) => {
  const oldCode = normalizeUnitCode(req.params.code);
  const { v, l } = req.body || {};
  const newCode = normalizeUnitCode(v || oldCode);
  if (!oldCode) return res.status(400).json({ error: "missing_unit_code" });
  if (!newCode) return res.status(400).json({ error: "missing_unit_code" });

  await ensureDB();
  const unitsDb = await ensureUnitsFile();
  unitsDb.units = Array.isArray(unitsDb.units) ? unitsDb.units : [];
  const idx = unitsDb.units.findIndex((u) => u.v === oldCode);
  if (idx < 0) return res.status(404).json({ error: "unit_not_found" });

  if (newCode !== oldCode && unitsDb.units.find((u) => u.v === newCode)) {
    return res.status(409).json({ error: "unit_exists" });
  }

  const label = normalizeUnitLabel(l, newCode);
  unitsDb.units[idx] = { v: newCode, l: label };

  // If unit code changed, migrate existing items (raw + fg)
  if (newCode !== oldCode) {
    const raw = await ensureStockFile("raw");
    const fg = await ensureStockFile("fg");
    let changed = false;
    for (const it of raw.items || []) {
      if (String(it.unit || "") === oldCode) { it.unit = newCode; changed = true; }
    }
    for (const it of fg.items || []) {
      if (String(it.unit || "") === oldCode) { it.unit = newCode; changed = true; }
    }
    if (changed) {
      await writeStock("raw", raw);
      await writeStock("fg", fg);
    }
  }

  await writeUnits(unitsDb);
  res.json({ ok: true, unit: { v: newCode, l: label } });
});

app.delete("/api/units/:code", requireAuth, async (req, res) => {
  const code = normalizeUnitCode(req.params.code);
  if (!code) return res.status(400).json({ error: "missing_unit_code" });

  await ensureDB();
  const unitsDb = await ensureUnitsFile();
  unitsDb.units = Array.isArray(unitsDb.units) ? unitsDb.units : [];

  const raw = await ensureStockFile("raw");
  const fg = await ensureStockFile("fg");
  const inUse = [...(raw.items || []), ...(fg.items || [])].some((it) => String(it.unit || "") === code);
  if (inUse) return res.status(409).json({ error: "unit_in_use" });

  const before = unitsDb.units.length;
  unitsDb.units = unitsDb.units.filter((u) => u.v !== code);
  if (unitsDb.units.length === before) return res.status(404).json({ error: "unit_not_found" });
  await writeUnits(unitsDb);
  res.json({ ok: true });
});

// -------------------- Inventory --------------------
app.get("/api/inventory/items", requireAuth, requirePerm("inventory"), async (req, res) => {
  const type = req.query.type === "fg" ? "fg" : "raw";
  await ensureDB();
  const stockDb = await ensureStockFile(type);
  const stock = computeInventory(stockDb);
  const items = stockDb.items.map((it) => ({ ...it, type, currentStock: stock.get(String(it.id)) ?? 0 }));
  res.json({ items });
});

app.post("/api/inventory/items", requireAuth, requirePerm("inventory"), async (req, res) => {
  const type = req.query.type === "fg" ? "fg" : "raw";
  const { code: codeIn, name, unit, sku = "", minStock = 0, cost = 0, lossPercent = 0, cookFactor = 1, salePrice = 0 } = req.body || {};
  if (!name || !unit) return res.status(400).json({ error: "missing_fields" });

  await ensureDB();
  const stockDb = await ensureStockFile(type);
  // Código sequencial por tipo (MP001, MP002... / PF001, PF002...)
  const desiredCode = normalizeItemCodeServer(codeIn, type);
  const prefix = type === "fg" ? "PF" : "MP";
  let nextCode = desiredCode || nextItemCode(stockDb.items, type);
  if (!nextCode || !String(nextCode).startsWith(prefix)) return res.status(400).json({ error: "invalid_code" });
  const codeKey = String(nextCode).trim().toUpperCase();
  const exists = (stockDb.items || []).some((i) => String(i.code || "").trim().toUpperCase() === codeKey);
  if (exists) return res.status(400).json({ error: "code_exists", code: codeKey });
  const item = {
    id: newId("itm_"),
    code: String(nextCode).trim().toUpperCase(),
    name: String(name).trim(),
    unit: String(unit).trim(),
    sku: String(sku || "").trim(),
    lossPercent: type === "raw" ? (Number(lossPercent) || 0) : undefined,
    cookFactor: type === "raw" ? (Number(cookFactor) || 1) : undefined,
    salePrice: type === "fg" ? (Number(salePrice) || 0) : undefined,
    type,
    minStock: Number(minStock) || 0,
    cost: Number(cost) || 0,
    createdAt: new Date().toISOString(),
  };
  stockDb.items.push(item);
  await writeStock(type, stockDb);
  res.json({ item });
});

app.put("/api/inventory/items/:id", requireAuth, requirePerm("inventory"), async (req, res) => {
  const type = req.query.type === "fg" ? "fg" : "raw";
  await ensureDB();
  const stockDb = await ensureStockFile(type);
  const item = stockDb.items.find((i) => i.id === req.params.id);
  if (!item) return res.status(404).json({ error: "not_found" });

  const { name, unit, minStock, cost, sku, lossPercent, cookFactor, salePrice } = req.body || {};
  if (name !== undefined) item.name = String(name).trim();
  if (unit !== undefined) item.unit = String(unit).trim();
  // sku era usado em versões antigas (mantido só para compatibilidade)
  if (sku !== undefined) item.sku = String(sku || "").trim();
  if (minStock !== undefined) item.minStock = Number(minStock) || 0;
  if (cost !== undefined) item.cost = Number(cost) || 0;

  if (type === "raw") {
    if (lossPercent !== undefined) item.lossPercent = Number(lossPercent) || 0;
    if (cookFactor !== undefined) item.cookFactor = Number(cookFactor) || 1;
  } else {
    if (salePrice !== undefined) item.salePrice = Number(salePrice) || 0;
  }

  await writeStock(type, stockDb);
  res.json({ item });
});

app.delete("/api/inventory/items/:id", requireAuth, requirePerm("inventory"), async (req, res) => {
  const type = req.query.type === "fg" ? "fg" : "raw";
  const force = String(req.query.force || "") === "1";

  const db = await ensureDB();
  const stockDb = await ensureStockFile(type);
  const id = String(req.params.id);

  const idx = stockDb.items.findIndex((i) => i.id === id);
  if (idx === -1) return res.status(404).json({ error: "not_found" });

  // Descobrir vínculos no MRP (BOM/OP/OC) e movimentações no estoque
  const links = [];
  const hasMov = (stockDb.movements || []).some((m) => m.itemId === id);

  try {
    const mrp = db.mrp || {};
    const recipes = Array.isArray(mrp.recipes) ? mrp.recipes : [];
    const ops = Array.isArray(mrp.productionOrders) ? mrp.productionOrders : [];
    const ocs = Array.isArray(mrp.purchaseOrders) ? mrp.purchaseOrders : [];

    // BOM: item usado como ingrediente
    for (const r of recipes) {
      const bom = Array.isArray(r.bom) ? r.bom : [];
      if (bom.some((l) => String(l.itemId) === id)) {
        links.push({ kind: "recipe_bom", label: `Receita/BOM: ${r.name || r.id}` });
      }
      // PF: item é outputItemId (somente em fg)
      if (type === "fg" && r.outputItemId && String(r.outputItemId) === id) {
        links.push({ kind: "recipe_output", label: `Receita (PF): ${r.name || r.id}` });
      }
    }

    // OP
    for (const op of ops) {
      const planC = Array.isArray(op?.planned?.consumed) ? op.planned.consumed : [];
      const cons = Array.isArray(op?.consumed) ? op.consumed : [];
      const shorts = Array.isArray(op?.shortages) ? op.shortages : [];
      const planP = op?.planned?.produced ? [op.planned.produced] : [];
      const prod = op?.produced ? [op.produced] : [];
      const hit = [...planC, ...cons, ...shorts, ...planP, ...prod].some((x) => String(x?.itemId) === id);
      if (hit) links.push({ kind: "production_order", label: `OP ${pad6(op.number || 0)} (${op.status || ""})` });
    }

    // OC
    for (const oc of ocs) {
      const items = Array.isArray(oc?.items) ? oc.items : [];
      const recs = Array.isArray(oc?.receipts) ? oc.receipts : [];
      const hit = items.some((x) => String(x?.itemId) === id) || recs.some((x) => String(x?.itemId) === id);
      if (hit) links.push({ kind: "purchase_order", label: `OC ${pad6(oc.number || 0)} (${oc.status || ""})` });
    }
  } catch (_e) {
    // Se algo falhar no scan, não travar o delete
  }

  if (!force && (hasMov || links.length)) {
    return res.status(409).json({
      error: "has_links",
      hasMovements: !!hasMov,
      links,
      message: "Item possui movimentações e/ou vínculos no MRP.",
    });
  }

  // Force delete: remove item, remove movimentações do item e remove referências no MRP.
  const removed = stockDb.items.splice(idx, 1)[0];
  stockDb.movements = (stockDb.movements || []).filter((m) => m.itemId !== id);
  await writeStock(type, stockDb);

  // Atualizar MRP para evitar referências quebradas
  let mrpDirty = false;
  const mrp = db.mrp || {};
  if (Array.isArray(mrp.recipes)) {
    for (const r of mrp.recipes) {
      if (Array.isArray(r.bom)) {
        const before = r.bom.length;
        r.bom = r.bom.filter((l) => String(l.itemId) !== id);
        if (r.bom.length !== before) mrpDirty = true;
      }
      if (type === "fg" && r.outputItemId && String(r.outputItemId) === id) {
        delete r.outputItemId;
        mrpDirty = true;
      }
    }
  }
  if (Array.isArray(mrp.productionOrders)) {
    for (const op of mrp.productionOrders) {
      if (Array.isArray(op?.planned?.consumed)) {
        const b = op.planned.consumed.length;
        op.planned.consumed = op.planned.consumed.filter((x) => String(x.itemId) !== id);
        if (op.planned.consumed.length !== b) mrpDirty = true;
      }
      if (op?.planned?.produced && String(op.planned.produced.itemId) === id) { delete op.planned.produced; mrpDirty = true; }
      if (Array.isArray(op.consumed)) { const b = op.consumed.length; op.consumed = op.consumed.filter((x) => String(x.itemId) !== id); if (op.consumed.length !== b) mrpDirty = true; }
      if (op?.produced && String(op.produced.itemId) === id) { delete op.produced; mrpDirty = true; }
      if (Array.isArray(op.shortages)) { const b = op.shortages.length; op.shortages = op.shortages.filter((x) => String(x.itemId) !== id); if (op.shortages.length !== b) mrpDirty = true; }
    }
  }
  if (Array.isArray(mrp.purchaseOrders)) {
    for (const oc of mrp.purchaseOrders) {
      if (Array.isArray(oc.items)) {
        const b = oc.items.length;
        oc.items = oc.items.filter((x) => String(x.itemId) !== id);
        if (oc.items.length !== b) mrpDirty = true;
      }
      if (Array.isArray(oc.receipts)) {
        const b = oc.receipts.length;
        oc.receipts = oc.receipts.filter((x) => String(x.itemId) !== id);
        if (oc.receipts.length !== b) mrpDirty = true;
      }
    }
  }
  if (mrpDirty) await writeMrp(mrp);

  res.json({ ok: true, removed, forced: !!force });
});

app.get("/api/inventory/movements", requireAuth, requirePerm("inventory"), async (req, res) => {
  const type = req.query.type === "fg" ? "fg" : "raw";
  await ensureDB();
  const stockDb = await ensureStockFile(type);
  const limit = Math.min(Number(req.query.limit || 200), 1000);
  const movements = [...stockDb.movements].sort((a, b) => (a.at < b.at ? 1 : -1)).slice(0, limit);
  res.json({ movements });
});

app.post("/api/inventory/movements", requireAuth, requirePerm("inventory"), async (req, res) => {
  const stockType = req.query.type === "fg" ? "fg" : "raw";
  const { type, itemId, qty, reason = "" } = req.body || {};
  if (!["in", "out", "adjust"].includes(type)) return res.status(400).json({ error: "invalid_type" });
  if (!itemId) return res.status(400).json({ error: "missing_itemId" });

  const q = Number(qty);
  if (!Number.isFinite(q) || q < 0) return res.status(400).json({ error: "invalid_qty" });

  const db = await ensureDB();
  const stockDb = await ensureStockFile(stockType);
  const item = stockDb.items.find((i) => i.id === itemId);
  if (!item) return res.status(404).json({ error: "item_not_found" });

  const user = db.users.find((u) => u.id === req.session.userId);
  const by = user ? { id: user.id, name: user.name } : { id: String(req.session.userId || ""), name: "" };

  // capture stock before/after for audit (esp. ajuste de inventário)
  const beforeQty = computeInventory(stockDb).get(itemId) ?? 0;
  let afterQty = beforeQty;
  if (type === "in") afterQty = beforeQty + q;
  else if (type === "out") afterQty = beforeQty - q;
  else if (type === "adjust") afterQty = q;

  const mv = {
    id: newId("mv_"),
    type,
    itemId,
    qty: q,
    reason: String(reason || ""),
    at: new Date().toISOString(),
    by,
    beforeQty,
    afterQty,
    delta: afterQty - beforeQty,
  };

  if (type === "out") {
    const stock = computeInventory(stockDb).get(itemId) ?? 0;
    if (stock - q < 0) return res.status(400).json({ error: "insufficient_stock", stock });
  }

  stockDb.movements.push(mv);
  await writeStock(stockType, stockDb);
  res.json({ movement: mv });
});

// Histórico de inventário (somente AJUSTES) — MP + PF
app.get("/api/inventory/inventory-history", requireAuth, requirePerm("inventory"), async (req, res) => {
  const limit = Math.min(Number(req.query.limit || 500), 2000);
  const db = await ensureDB(); // ensures auth DB is loaded (for user names)
  const rawStock = await ensureStockFile("raw");
  const fgStock = await ensureStockFile("fg");

  function buildFor(stockDb, stockType){
    const itemsById = new Map((stockDb.items || []).map((i) => [i.id, i]));
    const movs = [...(stockDb.movements || [])].sort((a,b) => (a.at < b.at ? -1 : 1));

    // replay to infer before/after when missing (backward compat)
    const running = new Map();
    const out = [];
    for (const mv of movs){
      const prev = running.get(mv.itemId) ?? 0;
      let next = prev;
      const q = Number(mv.qty || 0) || 0;
      if (mv.type === "in") next = prev + q;
      else if (mv.type === "out") next = prev - q;
      else if (mv.type === "adjust") next = q;
      running.set(mv.itemId, next);

      if (mv.type !== "adjust") continue;
      if (mv.hiddenFromHistory) continue;
      const item = itemsById.get(mv.itemId);
      const by = mv.by || { id: String(mv.userId || ""), name: String(mv.userName || "") };
      const beforeQty = Number(mv.beforeQty ?? prev) || 0;
      const afterQty = Number(mv.afterQty ?? next) || 0;
      const delta = Number(mv.delta ?? (afterQty - beforeQty)) || 0;
      out.push({
        id: mv.id,
        at: mv.at,
        stockType,
        itemId: mv.itemId,
        code: item?.code || "",
        name: item?.name || "",
        unit: item?.unit || "",
        beforeQty,
        afterQty,
        delta,
        by,
        reason: mv.reason || "",
      });
    }
    return out;
  }

  const history = [...buildFor(rawStock, "raw"), ...buildFor(fgStock, "fg")]
    .sort((a,b) => (a.at < b.at ? 1 : -1))
    .slice(0, limit);

  res.json({ history });
});


// Limpar historico de inventario (oculta ajustes no historico, sem afetar estoque atual)
app.post("/api/inventory/inventory-history/clear", requireAuth, requirePerm("canImportExport"), requireReauthOnce, async (req, res) => {
  const rawStock = await ensureStockFile("raw");
  const fgStock = await ensureStockFile("fg");

  function hideAdjust(stockDb){
    let n = 0;
    for (const mv of (stockDb.movements || [])) {
      if (mv.type === "adjust" && !mv.hiddenFromHistory){
        mv.hiddenFromHistory = true;
        n++;
      }
    }
    return n;
  }

  const rawHidden = hideAdjust(rawStock);
  const fgHidden = hideAdjust(fgStock);

  await writeStock("raw", rawStock);
  await writeStock("fg", fgStock);

  res.json({ ok: true, hidden: rawHidden + fgHidden, rawHidden, fgHidden });
});








// -------------------- Inventory XLSX (import/export) --------------------
const uploadXlsx = multer({
  storage: multer.memoryStorage(),
  limits: { fileSize: 25 * 1024 * 1024 },
});

function wbToBuffer(wb) {
  return XLSX.write(wb, { bookType: "xlsx", type: "buffer" });
}

function sendXlsx(res, wb, filename) {
  const buf = wbToBuffer(wb);
  res.setHeader(
    "Content-Type",
    "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
  );
  res.setHeader("Content-Disposition", `attachment; filename="${filename}"`);
  res.send(buf);
}

function stockItemsToSheetRows(stockDb, type) {
  const stock = computeInventory(stockDb);
  return (stockDb.items || [])
    .slice()
    .sort((a, b) => String(a.code || "").localeCompare(String(b.code || ""), "pt-BR"))
    .map((it) => ({
      COD: it.code || "",
      DESCRICAO: it.name || "",
      UN: it.unit || "",
      CUSTO: Number(it.cost || 0) || 0,
      V_VENDA: type === "fg" ? (Number(it.salePrice || 0) || 0) : (Number(it.salePrice || 0) || 0),
      PERDA_PCT: type === "raw" ? (Number(it.lossPercent || 0) || 0) : (Number(it.lossPercent || 0) || 0),
      FC: Number(it.cookFactor || 1) || 1,
      ESTOQUE_MIN: Number(it.minStock || 0) || 0,
      ESTOQUE_ATUAL: Number(stock.get(String(it.id)) ?? 0) || 0,
    }));
}

app.get("/api/inventory/export.xlsx", requireAuth, requirePerm("canImportExport"), requireReauthOnce, async (req, res) => {
  const type = String(req.query.type || "raw"); // raw | fg | all
  await ensureDB();

  const wb = XLSX.utils.book_new();

  if (type === "all") {
    const rawDb = await ensureStockFile("raw");
    const fgDb = await ensureStockFile("fg");
    const rawRows = stockItemsToSheetRows(rawDb, "raw");
    const fgRows = stockItemsToSheetRows(fgDb, "fg");
    XLSX.utils.book_append_sheet(wb, XLSX.utils.json_to_sheet(rawRows), "MP");
    XLSX.utils.book_append_sheet(wb, XLSX.utils.json_to_sheet(fgRows), "PF");
    return sendXlsx(res, wb, "cadastro_geral.xlsx");
  }

  const stockType = type === "fg" ? "fg" : "raw";
  const stockDb = await ensureStockFile(stockType);
  const rows = stockItemsToSheetRows(stockDb, stockType);
  XLSX.utils.book_append_sheet(wb, XLSX.utils.json_to_sheet(rows), stockType === "fg" ? "PF" : "MP");
  return sendXlsx(res, wb, stockType === "fg" ? "cadastro_pf.xlsx" : "cadastro_mp.xlsx");
});

app.post("/api/inventory/import.xlsx", requireAuth, requirePerm("canImportExport"), requireReauthOnce, uploadXlsx.single("file"), async (req, res) => {
  const type = String(req.query.type || "raw"); // raw | fg
  const stockType = type === "fg" ? "fg" : "raw";

  if (!req.file?.buffer) return res.status(400).json({ error: "missing_file" });

  await ensureDB();
  const stockDb = await ensureStockFile(stockType);

  const wb = XLSX.read(req.file.buffer, { type: "buffer" });
  const sheetName = wb.SheetNames[0];
  const ws = wb.Sheets[sheetName];
  const rows = XLSX.utils.sheet_to_json(ws, { defval: "" });

  const byCode = new Map((stockDb.items || []).map((i) => [String(normalizeItemCodeServer(i.code, stockType)).trim().toUpperCase(), i]));
  let created = 0;
  let updated = 0;
  let adjusted = 0;

  const now = new Date().toISOString();

  for (const r of rows) {
    const codeRaw = String(r.COD || r.Codigo || r.CÓDIGO || r.CODE || "").trim();
    const code = normalizeItemCodeServer(codeRaw, stockType);
    const name = String(r.DESCRICAO || r.Descricao || r.DESCRIÇÃO || r.NOME || r.Name || "").trim();
    const unit = String(r.UN || r.Unidade || r.UNIDADE || r.Unit || "").trim();
    if (!code || !name || !unit) continue;

    const key = String(code || '').trim().toUpperCase();
    let it = byCode.get(key);
    const cost = Number(String(r.CUSTO ?? r.COST ?? 0).toString().replace(",", ".")) || 0;
    const salePrice = Number(String(r.V_VENDA ?? r.VVENDA ?? r.VENDA ?? r.SALE ?? 0).toString().replace(",", ".")) || 0;
    const lossPercent = Number(String(r.PERDA_PCT ?? r.PERDA ?? 0).toString().replace(",", ".")) || 0;
    const cookFactor = Number(String(r.FC ?? 1).toString().replace(",", ".")) || 1;
    const minStock = Number(String(r.ESTOQUE_MIN ?? r.MIN ?? 0).toString().replace(",", ".")) || 0;
    const desiredStock = (r.ESTOQUE_ATUAL !== undefined && r.ESTOQUE_ATUAL !== "")
      ? (Number(String(r.ESTOQUE_ATUAL).toString().replace(",", ".")) || 0)
      : null;

    if (!it) {
      it = {
        id: newId("itm_"),
        code: String(code).trim().toUpperCase(),
        name,
        unit,
        sku: "",
        type: stockType,
        minStock,
        cost,
        createdAt: now,
      };
      if (stockType === "raw") {
        it.lossPercent = lossPercent;
        it.cookFactor = cookFactor;
      } else {
        it.salePrice = salePrice;
      }
      stockDb.items.push(it);
      byCode.set(key, it);
      created += 1;
    } else {
      it.name = name;
      it.unit = unit;
      it.cost = cost;
      it.minStock = minStock;
      if (stockType === "raw") {
        it.lossPercent = lossPercent;
        it.cookFactor = cookFactor;
      } else {
        it.salePrice = salePrice;
      }
      updated += 1;
    }

    if (desiredStock !== null) {
      // cria um ajuste de inventário (absoluto)
      const beforeQty = computeInventory(stockDb).get(it.id) ?? 0;
      const afterQty = desiredStock;
      const mv = {
        id: newId("mv_"),
        type: "adjust",
        itemId: it.id,
        qty: afterQty,
        reason: "Import XLSX",
        at: new Date().toISOString(),
        by: { id: String(req.session.userId || ""), name: String(req.session.userId || "") },
        beforeQty,
        afterQty,
        delta: afterQty - beforeQty,
      };
      stockDb.movements.push(mv);
      adjusted += 1;
    }
  }

  await writeStock(stockType, stockDb);
  res.json({ ok: true, created, updated, adjusted });
});

// Importar Cadastro Geral (MP + PF) em um único XLSX (abas MP e PF)
app.post("/api/inventory/import-all.xlsx", requireAuth, requirePerm("canImportExport"), requireReauthOnce, uploadXlsx.single("file"), async (req, res) => {
  if (!req.file?.buffer) return res.status(400).json({ error: "missing_file" });

  await ensureDB();

  const wb = XLSX.read(req.file.buffer, { type: "buffer" });
  const nameMap = new Map(wb.SheetNames.map((n) => [String(n||"").trim().toLowerCase(), n]));

  const pick = (want) => {
    const key = want.toLowerCase();
    if (nameMap.has(key)) return nameMap.get(key);
    // tolerâncias: "materia prima", "produto final", etc
    for (const n of wb.SheetNames){
      const lc = String(n||"").toLowerCase();
      if (want === "MP" && (lc.startsWith("mp") || lc.includes("mat") || lc.includes("raw"))) return n;
      if (want === "PF" && (lc.startsWith("pf") || lc.includes("final") || lc.includes("fg"))) return n;
    }
    return null;
  };

  // fallback: 1ª aba = MP, 2ª aba = PF
  const mpName = pick("MP") || wb.SheetNames[0] || null;
  const pfName = pick("PF") || wb.SheetNames[1] || null;

  const readRows = (sheetName) => {
    if (!sheetName || !wb.Sheets[sheetName]) return [];
    const ws = wb.Sheets[sheetName];
    return XLSX.utils.sheet_to_json(ws, { defval: "" });
  };

  const mpRows = readRows(mpName);
  const pfRows = readRows(pfName);

  const applyRows = async (stockType, rows) => {
    if (!rows || !rows.length) return { created: 0, updated: 0, adjusted: 0 };
    const stockDb = await ensureStockFile(stockType);
    const byCode = new Map((stockDb.items || []).map((i) => [String(normalizeItemCodeServer(i.code, stockType)).trim().toUpperCase(), i]));
    let created = 0, updated = 0, adjusted = 0;
    const now = new Date().toISOString();

    for (const r of rows) {
      const codeRaw = String(r.COD || r.Codigo || r.CÓDIGO || r.CODE || "").trim();
    const code = normalizeItemCodeServer(codeRaw, stockType);
      const name = String(r.DESCRICAO || r.Descricao || r.DESCRIÇÃO || r.NOME || r.Name || "").trim();
      const unit = String(r.UN || r.Unidade || r.UNIDADE || r.Unit || "").trim();
      if (!code || !name || !unit) continue;

      const key = String(code || '').trim().toUpperCase();
      let it = byCode.get(key);

      const cost = Number(String(r.CUSTO ?? r.COST ?? 0).toString().replace(",", ".")) || 0;
      const salePrice = Number(String(r.V_VENDA ?? r.VVENDA ?? r.VENDA ?? r.SALE ?? 0).toString().replace(",", ".")) || 0;
      const lossPercent = Number(String(r.PERDA_PCT ?? r.PERDA ?? 0).toString().replace(",", ".")) || 0;
      const cookFactor = Number(String(r.FC ?? 1).toString().replace(",", ".")) || 1;
      const minStock = Number(String(r.ESTOQUE_MIN ?? r.MIN ?? 0).toString().replace(",", ".")) || 0;
      const desiredStock = (r.ESTOQUE_ATUAL !== undefined && r.ESTOQUE_ATUAL !== "")
        ? (Number(String(r.ESTOQUE_ATUAL).toString().replace(",", ".")) || 0)
        : null;

      if (!it) {
        it = {
          id: newId("itm_"),
          code,
          name,
          unit,
          sku: "",
          type: stockType,
          minStock,
          cost,
          createdAt: now,
        };
        if (stockType === "raw") {
          it.lossPercent = lossPercent;
          it.cookFactor = cookFactor;
        } else {
          it.salePrice = salePrice;
        }
        stockDb.items.push(it);
        byCode.set(key, it);
        created += 1;
      } else {
        it.name = name;
        it.unit = unit;
        it.cost = cost;
        it.minStock = minStock;
        if (stockType === "raw") {
          it.lossPercent = lossPercent;
          it.cookFactor = cookFactor;
        } else {
          it.salePrice = salePrice;
        }
        updated += 1;
      }

      if (desiredStock !== null) {
        const beforeQty = computeInventory(stockDb).get(it.id) ?? 0;
        const afterQty = desiredStock;
        const mv = {
          id: newId("mv_"),
          type: "adjust",
          itemId: it.id,
          qty: afterQty,
          reason: "Import XLSX",
          at: new Date().toISOString(),
          by: { id: String(req.session.userId || ""), name: String(req.session.userId || "") },
          beforeQty,
          afterQty,
          delta: afterQty - beforeQty,
        };
        stockDb.movements.push(mv);
        adjusted += 1;
      }
    }

    await writeStock(stockType, stockDb);
    return { created, updated, adjusted };
  };

  const mpResult = await applyRows("raw", mpRows);
  const pfResult = await applyRows("fg", pfRows);

  res.json({ ok: true, mp: mpResult, pf: pfResult });
});

app.get("/api/inventory/inventory-history.xlsx", requireAuth, requirePerm("canImportExport"), requireReauthOnce, async (req, res) => {
  const limit = Math.min(Number(req.query.limit || 2000), 5000);
  await ensureDB();
  const rawStock = await ensureStockFile("raw");
  const fgStock = await ensureStockFile("fg");

  const itemsByIdRaw = new Map((rawStock.items || []).map((i) => [i.id, i]));
  const itemsByIdFg = new Map((fgStock.items || []).map((i) => [i.id, i]));

  const rows = [];
  const addRows = (stockDb, stockType, itemsById) => {
    const movs = [...(stockDb.movements || [])]
      .filter((m) => m.type === "adjust")
      .sort((a, b) => (a.at < b.at ? 1 : -1));
    for (const mv of movs) {
      const item = itemsById.get(mv.itemId);
      rows.push({
        DATA_HORA: mv.at,
        TIPO: stockType === "fg" ? "PF" : "MP",
        COD: item?.code || "",
        DESCRICAO: item?.name || "",
        UN: item?.unit || "",
        ANTES: Number(mv.beforeQty ?? 0) || 0,
        DEPOIS: Number(mv.afterQty ?? 0) || 0,
        DIF: Number(mv.delta ?? 0) || 0,
        POR: mv.by?.name || mv.by?.id || "",
        OBS: mv.reason || "",
      });
      if (rows.length >= limit) break;
    }
  };

  addRows(rawStock, "raw", itemsByIdRaw);
  addRows(fgStock, "fg", itemsByIdFg);

  const wb = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(wb, XLSX.utils.json_to_sheet(rows), "HISTORICO");
  return sendXlsx(res, wb, "historico_inventario.xlsx");
});


function ensureRecipeOutputItem(fgStock, recipe) {
  // Prefer linking output to an existing PF item when productId is provided
  if (recipe.productId) {
    const pf = fgStock.items.find((i) => i.id === recipe.productId);
    if (pf) {
      pf.type = "fg";
      // Keep PF name; but align unit with recipe yield when set
      if (recipe.yieldUnit) pf.unit = recipe.yieldUnit;
      recipe.outputItemId = pf.id;
      // Keep recipe name in sync with PF for consistency
      recipe.name = pf.name || recipe.name;
      return pf.id;
    }
  }

  // Backward compatibility: use stored outputItemId if it still exists
  if (recipe.outputItemId) {
    const exists = fgStock.items.find((i) => i.id === recipe.outputItemId);
    if (exists) return recipe.outputItemId;
    // if missing, fall through to recreate
  }

  // Fallback: try match by name
  const existing = fgStock.items.find((i) => (i.name || "").toLowerCase() === (recipe.name || "").toLowerCase());
  if (existing) {
    existing.type = "fg";
    existing.unit = recipe.yieldUnit || existing.unit || "un";
    recipe.outputItemId = existing.id;
    return existing.id;
  }

  // Last resort: create a PF item for this recipe
  const item = {
    id: newId("itm_"),
    name: recipe.name,
    unit: recipe.yieldUnit || "un",
    sku: "",
    type: "fg",
    minStock: 0,
    cost: 0,
    createdAt: new Date().toISOString(),
  };
  fgStock.items.push(item);
  recipe.outputItemId = item.id;
  return item.id;
}



// -------------------- MRP / Recipes --------------------
app.get("/api/mrp/recipes", requireAuth, requireAnyPerm(["recipes","op","oc","costs"]), async (req, res) => {
  const db = await ensureDB();
  res.json({ recipes: db.mrp.recipes });
});

app.post("/api/mrp/recipes", requireAuth, requirePerm("recipes"), async (req, res) => {
  const { name, productId = "", yieldQty = 1, yieldUnit = "un", bom = [], notes = "", method = "" } = req.body || {};
  if (!name) return res.status(400).json({ error: "missing_name" });
  if (!Array.isArray(bom) || bom.length === 0) return res.status(400).json({ error: "missing_bom" });

  const db = await ensureDB();

  const recipe = {
    id: newId("rcp_"),
    name: String(name).trim(),
    productId: String(productId || "") || undefined,
    yieldQty: Number(yieldQty) || 1,
    yieldUnit: String(yieldUnit || "un").trim(),
    notes: String(notes || ""),
    method: String(method || ""),
    bom: bom.map((l) => ({
      itemId: String(l.itemId),
      qty: Number(l.qty) || 0,
      pos: (Number(l.pos) > 0 ? Number(l.pos) : undefined),
      fc: (Number(l.fc) > 0 ? Number(l.fc) : undefined),
    })).filter((l) => l.itemId && l.qty > 0),
    createdAt: new Date().toISOString(),
  };

  if (recipe.bom.length === 0) return res.status(400).json({ error: "invalid_bom" });

  const fgStock = await ensureStockFile("fg");
  ensureRecipeOutputItem(fgStock, recipe);

  db.mrp.recipes.push(recipe);

  await writeMrp(db.mrp);
  await writeStock("fg", fgStock);

  res.json({ recipe });
});


app.put("/api/mrp/recipes/:id", requireAuth, requirePerm("recipes"), async (req, res) => {
  const db = await ensureDB();
  const recipe = db.mrp.recipes.find((r) => r.id === req.params.id);
  if (!recipe) return res.status(404).json({ error: "not_found" });

  const { name, productId, yieldQty, yieldUnit, bom, notes, method } = req.body || {};
  if (name !== undefined) recipe.name = String(name).trim();
  if (productId !== undefined) recipe.productId = String(productId || "") || undefined;
  if (yieldQty !== undefined) recipe.yieldQty = Number(yieldQty) || 1;
  if (yieldUnit !== undefined) recipe.yieldUnit = String(yieldUnit || "un").trim();
  if (notes !== undefined) recipe.notes = String(notes || "");
  if (method !== undefined) recipe.method = String(method || "");

  if (bom !== undefined) {
    if (!Array.isArray(bom) || bom.length === 0) return res.status(400).json({ error: "invalid_bom" });
    recipe.bom = bom.map((l) => ({
      itemId: String(l.itemId),
      qty: Number(l.qty) || 0,
      pos: (Number(l.pos) > 0 ? Number(l.pos) : undefined),
      fc: (Number(l.fc) > 0 ? Number(l.fc) : undefined),
    })).filter((l) => l.itemId && l.qty > 0);
    if (recipe.bom.length === 0) return res.status(400).json({ error: "invalid_bom" });
  }

  const fgStock = await ensureStockFile("fg");
  // Ensure output item exists, then sync it
  ensureRecipeOutputItem(fgStock, recipe);
  if (recipe.outputItemId) {
    const outIt = fgStock.items.find((i) => i.id === recipe.outputItemId);
    if (outIt) {
      outIt.type = "fg";
      // If recipe is linked to an existing PF (productId), don't overwrite its name
      if (!recipe.productId || recipe.productId !== outIt.id) {
        outIt.name = recipe.name;
      } else {
        // keep recipe in sync with PF name
        recipe.name = outIt.name || recipe.name;
      }
      outIt.unit = recipe.yieldUnit || outIt.unit;
    }
  }

  await writeMrp(db.mrp);
  await writeStock("fg", fgStock);

  res.json({ recipe });
});

// -------------------- BOM XLSX (por Produto Final) --------------------
function toFriendlyQty(baseQty, baseUnit) {
  const u = String(baseUnit || "").toLowerCase();
  const q = Number(baseQty || 0) || 0;
  if (u === "kg") return { v: q * 1000, u: "g" };
  if (u === "l") return { v: q * 1000, u: "ml" };
  return { v: q, u: u || "" };
}

function fromFriendlyQty(friendlyQty, baseUnit) {
  const u = String(baseUnit || "").toLowerCase();
  const q = Number(friendlyQty || 0) || 0;
  if (u === "kg") return q / 1000; // g -> kg
  if (u === "l") return q / 1000;  // ml -> l
  return q;
}

function parseNumBR(x, def = 0) {
  // XLSX pode entregar número (ex.: 2.6) mesmo que o usuário veja "2,6" no Excel.
  // Se a gente remover '.' cegamente, "2.6" vira "26" (bug). Portanto, detectamos separadores.
  if (x === null || x === undefined) return def;
  if (typeof x === "number") return Number.isFinite(x) ? x : def;

  let s = String(x).trim();
  if (!s) return def;
  s = s.replace(/\s/g, "");
  // Remove qualquer coisa que não seja dígito/sinal/separador
  s = s.replace(/[^\d,\.\-]/g, "");

  const hasComma = s.includes(",");
  const hasDot = s.includes(".");

  if (hasComma && hasDot) {
    // Decide qual é o separador decimal pelo último que aparece.
    const lastComma = s.lastIndexOf(",");
    const lastDot = s.lastIndexOf(".");
    if (lastComma > lastDot) {
      // 1.234,56  -> milhares '.' e decimal ','
      s = s.replace(/\./g, "").replace(",", ".");
    } else {
      // 1,234.56  -> milhares ',' e decimal '.'
      s = s.replace(/,/g, "");
    }
  } else if (hasComma && !hasDot) {
    // 2,6 -> 2.6
    s = s.replace(",", ".");
  } else if (!hasComma && hasDot) {
    // Se for padrão de milhar (1.234.567), remove.
    // Se for decimal (2.6), mantém.
    if (/^\d{1,3}(\.\d{3})+$/.test(s)) {
      s = s.replace(/\./g, "");
    }
  }

  const n = Number(s);
  return Number.isFinite(n) ? n : def;
}

app.get("/api/mrp/recipes/:id/bom.xlsx", requireAuth, requireAnyPerm(["recipes","op","oc","costs"]), requireReauthOnce, async (req, res) => {
  const db = await ensureDB();
  const recipe = db.mrp.recipes.find((r) => r.id === req.params.id);
  if (!recipe) return res.status(404).json({ error: "not_found" });

  const rawDb = await ensureStockFile("raw");
  const rawById = new Map((rawDb.items || []).map((i) => [i.id, i]));

  const rows = (recipe.bom || [])
    .slice()
    .sort((a, b) => Number(a.pos || 0) - Number(b.pos || 0))
    .map((l) => {
      const it = rawById.get(l.itemId);
      const baseUnit = it?.unit || "";
      const fq = toFriendlyQty(Number(l.qty || 0), baseUnit);
      const fc = Number(l.fc || it?.cookFactor || 1) || 1;
      const cooked = { v: fq.v * fc, u: fq.u };
      return {
        POS: Number(l.pos || 0) || "",
        COD: it?.code || "",
        DESCRICAO: it?.name || "",
        UN_COMPRA: baseUnit,
        UN_CONSUMO: fq.u,
        FC: fc,
        QTE_CRUA: fq.v,
        QTE_COZIDA: cooked.v,
      };
    });

  const wb = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(wb, XLSX.utils.json_to_sheet(rows), "BOM");
  const safeCode = String((recipe.name || "BOM").replace(/[^a-z0-9\-_]+/gi, "_").slice(0, 40));
  return sendXlsx(res, wb, `bom_${safeCode}.xlsx`);
});

function parseBomXlsxBuffer(buf, rawDb, { createMissing = false } = {}) {
  const wb = XLSX.read(buf, { type: "buffer" });
  const ws = wb.Sheets[wb.SheetNames[0]];
  const rows = XLSX.utils.sheet_to_json(ws, { defval: "" });

  // Mapa por código (MP###)
  const byCode = new Map((rawDb.items || []).map((i) => [String(i.code || "").trim().toUpperCase(), i]));

  const out = [];
  const missing = [];
  const created = [];

  const normUnit = (u) => {
    const s = String(u || "").trim();
    if (!s) return "un";
    return s.toLowerCase();
  };

  for (const r of rows) {
    const codeRaw = String(r.COD || r.Codigo || r.CÓDIGO || r.CODE || "").trim();
    const code = normalizeItemCodeServer(codeRaw);
    if (!code) continue;

    let it = byCode.get(code.toUpperCase());

    // Se o item não existe no cadastro de MP e o modo createMissing está ligado,
    // criamos automaticamente com os dados básicos do XLSX.
    if (!it && createMissing) {
      const name = String(r.DESCRICAO || r.DESCRIÇÃO || r.Descricao || r.Descrição || "").trim();
      const unitBuy = normUnit(r.UN_COMPRA || r.UN || r.UNIDADE || r.UNID || "");
      // FC: se vier, já salva como cookFactor do item
      const fc = parseNumBR(r.FC ?? "", NaN);

      // Evita criar sem nome (ainda assim permite, mas fica com o código como nome)
      const safeName = name || code;

      it = {
        id: newId("raw_"),
        code,
        name: safeName,
        unit: unitBuy,
        stock: 0,
        minStock: 0,
        lossPercent: 0,
      };
      if (Number.isFinite(fc) && fc > 0) it.cookFactor = fc;

      rawDb.items = rawDb.items || [];
      rawDb.items.push(it);
      byCode.set(code.toUpperCase(), it);
      created.push({ code, name: safeName, unit: unitBuy });
    }

    if (!it) {
      missing.push(code);
      continue;
    }

    const pos = Math.max(1, Math.trunc(parseNumBR(r.POS, 0) || 1));
    const qFriendly = parseNumBR(r.QTE_CRUA ?? r.QTE ?? r.QTD ?? r.QUANTIDADE ?? 0, 0);
    const baseQty = fromFriendlyQty(qFriendly, it.unit);
    const fc = parseNumBR(r.FC ?? "", NaN);

    out.push({
      itemId: it.id,
      qty: Math.max(0, Number(baseQty) || 0),
      pos,
      fc: Number.isFinite(fc) && fc > 0 ? fc : undefined,
    });
  }

  // resolve conflitos de POS de forma determinística
  out.sort((a, b) => a.pos - b.pos || String(a.itemId).localeCompare(String(b.itemId)));
  const used = new Set();
  for (const l of out) {
    let p = l.pos;
    while (used.has(p)) p += 1;
    l.pos = p;
    used.add(p);
  }

  const bom = out.filter((l) => l.itemId && l.qty > 0);
  const missingUniq = Array.from(new Set(missing));
  return { bom, missing: missingUniq, created };
}


app.post("/api/mrp/recipes/:id/bom.xlsx", requireAuth, requirePerm("recipes"), requireReauthOnce, uploadXlsx.single("file"), async (req, res) => {
  if (!req.file?.buffer) return res.status(400).json({ error: "missing_file" });
  try {
    const db = await ensureDB();
    const recipe = db.mrp.recipes.find((r) => r.id === req.params.id);
    if (!recipe) return res.status(404).json({ error: "not_found" });

    const rawDb = await ensureStockFile("raw");
    const { bom, missing, created } = parseBomXlsxBuffer(req.file.buffer, rawDb, { createMissing: true });
    if (!bom.length) return res.status(400).json({ error: "invalid_bom", missing });
    // Se criou itens novos no cadastro de MP, persiste antes de gravar o MRP
    if (created && created.length) {
      await writeStock("raw", rawDb);
    }
    recipe.bom = bom;
    await writeMrp(db.mrp);
    return res.json({ ok: true, bomCount: bom.length, createdCount: (created||[]).length, missing });
  } catch (e) {
    console.error("BOM import failed:", e);
    return res.status(400).json({ error: "import_failed" });
  }
});



// Importar BOM por Produto Final (cria receita se ainda não existir)
app.post("/api/mrp/pf/:productId/bom.xlsx", requireAuth, requirePerm("recipes"), requireReauthOnce, uploadXlsx.single("file"), async (req, res) => {
  if (!req.file?.buffer) return res.status(400).json({ error: "missing_file" });
  const productId = String(req.params.productId || "");
  if (!productId) return res.status(400).json({ error: "missing_product" });

  try {
    const db = await ensureDB();

    const fgStock = await ensureStockFile("fg");
    const pf = (fgStock.items || []).find((i) => String(i.id) === String(productId));

    const rawDb = await ensureStockFile("raw");
    const { bom, missing, created } = parseBomXlsxBuffer(req.file.buffer, rawDb, { createMissing: true });
    if (!bom.length) return res.status(400).json({ error: "invalid_bom", missing });

    // Se criou MPs novos, persiste
    if (created && created.length) {
      await writeStock("raw", rawDb);
    }

    let recipe = (db.mrp.recipes || []).find((r) => String(r.productId || "") === String(productId));
    if (!recipe) {
      recipe = {
        id: newId("rcp_"),
        name: String(pf?.name || (pf?.code ? `Receita ${pf.code}` : "Receita")).trim(),
        productId: String(productId),
        yieldQty: 1,
        yieldUnit: String(pf?.unit || "un").trim(),
        notes: "",
        method: "",
        bom,
        createdAt: new Date().toISOString(),
      };
      ensureRecipeOutputItem(fgStock, recipe);
      db.mrp.recipes.push(recipe);
    } else {
      recipe.bom = bom;
      ensureRecipeOutputItem(fgStock, recipe);
    }

    await writeMrp(db.mrp);
    await writeStock("fg", fgStock);

    return res.json({ ok: true, recipeId: recipe.id, bomCount: bom.length, createdCount: (created || []).length, missing });
  } catch (e) {
    console.error("BOM import (by PF) failed:", e);
    return res.status(400).json({ error: "import_failed" });
  }
});

app.post("/api/mrp/bom/parse.xlsx", requireAuth, requirePerm("recipes"), requireReauthOnce, uploadXlsx.single("file"), async (req, res) => {
  if (!req.file?.buffer) return res.status(400).json({ error: "missing_file" });
  try {
    await ensureDB();
    const rawDb = await ensureStockFile("raw");
    const { bom, missing, created } = parseBomXlsxBuffer(req.file.buffer, rawDb, { createMissing: false });
    return res.json({ ok: true, bom, count: bom.length, missing, createdPreview: created });
  } catch (e) {
    console.error("BOM parse failed:", e);
    return res.status(400).json({ error: "parse_failed" });
  }
});

// Excluir BOM (estrutura) de uma receita/PF.
// Mantém o registro da receita (id) para não quebrar referências, mas zera a estrutura.
// Também remove a foto associada (best-effort).
app.delete("/api/mrp/recipes/:id/bom", requireAuth, requirePerm("recipes"), requireReauthOnce, async (req, res) => {
  const db = await ensureDB();
  const recipe = db.mrp.recipes.find((r) => r.id === req.params.id);
  if (!recipe) return res.status(404).json({ error: "not_found" });

  // remove foto (se existir)
  if (recipe.photoFile) {
    await ensureBD();
    const prev = path.join(BD_PHOTOS_DIR, String(recipe.photoFile));
    if (fssync.existsSync(prev)) {
      await fs.unlink(prev).catch(() => {});
    }
  }

  recipe.bom = [];
  delete recipe.photoFile;
  // campos opcionais (não garantimos nomes)
  delete recipe.observations;
  delete recipe.obs;
  delete recipe.notes;
  delete recipe.method;

  await writeMrp(db.mrp);
  return res.json({ ok: true });
});

// -------------------- Foto do Produto Final (salva em /bd/photos) --------------------
const uploadPhoto = multer({
  storage: multer.memoryStorage(),
  limits: { fileSize: 8 * 1024 * 1024 },
  fileFilter: (req, file, cb) => {
    const ok = ["image/jpeg", "image/png", "image/webp"].includes(String(file.mimetype || ""));
    cb(ok ? null : new Error("invalid_image_type"), ok);
  },
});

function extFromMime(m) {
  if (m === "image/jpeg") return "jpg";
  if (m === "image/png") return "png";
  if (m === "image/webp") return "webp";
  return "bin";
}

app.post("/api/mrp/recipes/:id/photo", requireAuth, requirePerm("recipes"), uploadPhoto.single("file"), async (req, res) => {
  if (!req.file?.buffer) return res.status(400).json({ error: "missing_file" });
  const db = await ensureDB();
  const recipe = db.mrp.recipes.find((r) => r.id === req.params.id);
  if (!recipe) return res.status(404).json({ error: "not_found" });

  await ensureBD();
  const ext = extFromMime(req.file.mimetype);
  const base = String((recipe.productId || recipe.id || "pf")).replace(/[^a-z0-9\-_]+/gi, "_");
  const filename = `${base}_${Date.now()}.${ext}`;
  const outPath = path.join(BD_PHOTOS_DIR, filename);

  await fs.writeFile(outPath, req.file.buffer);

  // remove foto anterior (best-effort)
  if (recipe.photoFile) {
    const prev = path.join(BD_PHOTOS_DIR, String(recipe.photoFile));
    if (fssync.existsSync(prev)) {
      await fs.unlink(prev).catch(() => {});
    }
  }

  recipe.photoFile = filename;
  await writeMrp(db.mrp);
  res.json({ ok: true, photoFile: filename });
});

app.delete("/api/mrp/recipes/:id/photo", requireAuth, requirePerm("recipes"), async (req, res) => {
  const db = await ensureDB();
  const recipe = db.mrp.recipes.find((r) => r.id === req.params.id);
  if (!recipe) return res.status(404).json({ error: "not_found" });

  await ensureBD();
  if (recipe.photoFile) {
    const p = path.join(BD_PHOTOS_DIR, String(recipe.photoFile));
    if (fssync.existsSync(p)) {
      await fs.unlink(p).catch(() => {});
    }
  }
  delete recipe.photoFile;
  await writeMrp(db.mrp);
  res.json({ ok: true });
});


// Calculate requirements for a recipe and qtyToProduce (returns availability)
app.post("/api/mrp/requirements", requireAuth, requirePerm("recipes"), async (req, res) => {
  const { recipeId, qtyToProduce } = req.body || {};
  const qtp = Number(qtyToProduce);
  if (!recipeId || !Number.isFinite(qtp) || qtp <= 0) return res.status(400).json({ error: "invalid_payload" });

  const db = await ensureDB();
  const recipe = db.mrp.recipes.find((r) => r.id === recipeId);
  if (!recipe) return res.status(404).json({ error: "recipe_not_found" });
  const fgStock = await ensureStockFile("fg");
  ensureRecipeOutputItem(fgStock, recipe);
  await writeStock("fg", fgStock);

  const factor = qtp / (Number(recipe.yieldQty) || 1);
  const rawStock = await ensureStockFile("raw");
  const stock = computeInventory(rawStock);
  const itemsById = new Map((rawStock.items || []).map((i) => [String(i.id), i]));

  const requirements = recipe.bom.map((line) => {
    const item = itemsById.get(String(line.itemId));
    const unit = item?.unit || "";
    const lossPercent = Number(item?.lossPercent) || 0;

    const requiredNet = Number(line.qty) * factor;
    const requiredGrossRaw = requiredNet * (1 + (lossPercent / 100));

    // Quantiza para evitar falsos "faltantes" por imprecisão de float (ex.: 0.517999999 vs 0.518)
    const required = quantizeQty(requiredGrossRaw, unit);
    const available = quantizeQty((stock.get(String(line.itemId)) ?? 0), unit);
    const shortage = quantizeQty(Math.max(0, required - available), unit);

    return {
      itemId: String(line.itemId),
      itemCode: item?.code || "",
      itemName: item?.name || "—",
      unit,
      lossPercent,
      requiredNet: Number(requiredNet.toFixed(6)),
      required, // valor usado para checagem (inclui perda)
      available,
      shortage,
      ok: available >= required,
    };
  });

  res.json({ recipe, qtyToProduce: qtp, factor, requirements });
});

app.get("/api/mrp/production-orders", requireAuth, requirePerm("op"), async (req, res) => {
  const db = await ensureDB();
  const orders = [...db.mrp.productionOrders].sort((a, b) => (a.createdAt < b.createdAt ? 1 : -1));
  res.json({ orders });
});


// Archived Production Orders (separado)
app.get("/api/mrp/production-orders/archived", requireAuth, requirePerm("op"), async (req, res) => {
  const arch = await ensureProductionOrdersArchivedFile();
  const orders = [...(arch.productionOrders || [])].sort((a, b) => (
    String(a.archivedAt || a.createdAt || "") < String(b.archivedAt || b.createdAt || "") ? 1 : -1
  ));
  res.json({ orders });
});


// Atualizar OP (status/observação)
app.put("/api/mrp/production-orders/:id", requireAuth, requirePerm("op"), async (req, res) => {
  const db = await ensureDB();
  const op = (db.mrp.productionOrders || []).find((o) => o.id === req.params.id);
  if (!op) return res.status(404).json({ error: "not_found" });

  const { status, note } = req.body || {};
  if (status !== undefined) {
    const normalize = (v) => {
      let s = String(v || "ISSUED").toUpperCase();
      if (s === "READY") s = "ISSUED";
      if (s === "HOLD") s = "ISSUED";
      if (s === "EXECUTED") s = "CLOSED";
      if (!s) s = "ISSUED";
      return s;
    };

    const cur = normalize(op.status || "ISSUED");
    const next = normalize(status);

    const allowed = new Set(["ISSUED", "IN_PRODUCTION", "CLOSED", "CANCELLED"]);
    if (!allowed.has(next)) return res.status(400).json({ error: "invalid_status" });

    // Não permitir reabrir status final
    if ((cur === "CLOSED" || cur === "CANCELLED") && next !== cur) {
      const msg = cur === "CLOSED" ? "OP encerrada não pode ser reaberta." : "OP cancelada não pode ser reaberta.";
      return res.status(400).json({ error: "final_status", message: msg });
    }

    // Nenhuma mudança
    if (next === cur) {
      op.status = next;
    } else {
      // Regras do usuário:
      // - Baixa/consumo ao entrar em IN_PRODUCTION
      // - Entrada do PF ao entrar em CLOSED
      // - Cancelamento somente antes de iniciar produção
      let rawStock = null;
      let fgStock = null;
      let touchedStock = false;

      const recipe = db.mrp.recipes.find((r) => r.id === op.recipeId);
      if (!recipe) return res.status(400).json({ error: "recipe_not_found" });

      const qtp = Number(op.qtyToProduce);
      const factor = qtp / (Number(recipe.yieldQty) || 1);

      const loadStocks = async () => {
        if (!rawStock) rawStock = await ensureStockFile("raw");
        if (!fgStock) fgStock = await ensureStockFile("fg");
      };

      const computeRequirements = async () => {
        await loadStocks();
        const stock = computeInventory(rawStock);
        const itemsById = new Map((rawStock.items || []).map((i) => [i.id, i]));
        const reqs = (recipe.bom || []).map((line) => {
          const item = itemsById.get(String(line.itemId));
          const unit = item?.unit || "";
          const lossPercent = Number(item?.lossPercent) || 0;

          const requiredNet = Number(line.qty) * factor;
          const requiredGrossRaw = requiredNet * (1 + (lossPercent / 100));

          const required = quantizeQty(requiredGrossRaw, unit);
          const available = quantizeQty((stock.get(String(line.itemId)) ?? 0), unit);
          const shortage = quantizeQty(Math.max(0, required - available), unit);

          return {
            itemId: String(line.itemId),
            unit,
            required,
            available,
            shortage,
            ok: available >= required,
          };
        });
        return reqs;
      };

      if (next === "IN_PRODUCTION") {
        if (cur !== "ISSUED") {
          return res.status(400).json({ error: "invalid_transition", message: "Só é possível iniciar produção a partir de Emitida." });
        }

        const reqs = await computeRequirements();
        const shortages = reqs.filter((r) => !r.ok && r.shortage > 0);
        if (shortages.length) {
          // mantém EMITIDA se ainda estiver faltando (não permite iniciar produção)
          return res.status(400).json({ error: "insufficient_stock", shortages });
        }

        // Baixa no estoque (movimentos OUT)
        await loadStocks();
        rawStock.movements = rawStock.movements || [];
        op.consumed = [];
        for (const r of reqs) {
          const mv = {
            id: newId("mv_"),
            type: "out",
            itemId: r.itemId,
            qty: r.required,
            reason: `Consumo OP ${pad6(op.number)} (${op.recipeName || recipe.name})`,
            at: new Date().toISOString(),
          };
          rawStock.movements.push(mv);
          op.consumed.push({ itemId: r.itemId, qty: r.required, movementId: mv.id });
        }
        op.status = "IN_PRODUCTION";
        op.startedAt = new Date().toISOString();
        op.shortages = [];
        op.factor = Number(factor.toFixed(6));
        op.planned = op.planned || {};
        op.planned.consumed = reqs.map((r) => ({ itemId: r.itemId, qty: r.required }));
        touchedStock = true;
      } else if (next === "CLOSED") {
        if (cur !== "IN_PRODUCTION") {
          return res.status(400).json({ error: "invalid_transition", message: "Para encerrar, primeiro coloque a OP em Em produção." });
        }

        await loadStocks();
        fgStock.movements = fgStock.movements || [];
        const outItemId = ensureRecipeOutputItem(fgStock, recipe);
        const mvIn = {
          id: newId("mv_"),
          type: "in",
          itemId: outItemId,
          qty: Number(qtp),
          reason: `Produção OP ${pad6(op.number)} (${op.recipeName || recipe.name})`,
          at: new Date().toISOString(),
        };
        fgStock.movements.push(mvIn);
        op.produced = { itemId: outItemId, qty: Number(qtp), movementId: mvIn.id };
        // Lote / Código de barras (para etiquetas)
        if (!op.lotNumber) {
          const lotN = await nextLotNumber(db.mrp);
          op.lotNumber = lotN;
          op.lotCode = formatLotCode(lotN);
          const rc = String(op.recipeCode || recipe.code || "");
          op.recipeCode = rc;
          op.barcodeValue = rc ? `DIETON-${rc}-${op.lotCode}` : `DIETON-${op.lotCode}`;
          op.lotCreatedAt = new Date().toISOString();
        }

        op.status = "CLOSED";
        op.closedAt = new Date().toISOString();
        touchedStock = true;
      } else if (next === "CANCELLED") {
        // Permite cancelar antes de encerrar. Se já estava em produção, estorna a baixa do estoque.
        if (!(cur === "ISSUED" || cur === "IN_PRODUCTION")) {
          return res.status(400).json({ error: "invalid_transition", message: "Só é possível cancelar antes de encerrar a OP." });
        }

        if (cur === "IN_PRODUCTION") {
          // Estorna consumo: remove movimentos 'out' já lançados e limpa o registro de consumo
          const rawStock = await ensureStockFile("raw");
          const mvIds = new Set((op.consumed || []).map(x => String(x.movementId || "")).filter(Boolean));
          if (mvIds.size) {
            rawStock.movements = (rawStock.movements || []).filter(mv => !mvIds.has(String(mv.id)));
          }
          op.consumed = [];
          try { delete op.startedAt; } catch (_) {}
          await writeStock("raw", rawStock);
        }

        op.status = "CANCELLED";
        op.cancelledAt = new Date().toISOString();
      } else if (next === "ISSUED") {
        // Emitida é o estado padrão. Não existe mais "Em espera".
        // Não permite voltar para Emitida a partir de Em produção ou estados finais.
        if (cur !== "ISSUED") {
          return res.status(400).json({ error: "invalid_transition", message: "Transição inválida para Emitida." });
        }
        op.status = "ISSUED";
      } else {
        return res.status(400).json({ error: "invalid_transition" });
      }

      op.updatedAt = new Date().toISOString();
      await writeMrp(db.mrp);
      if (touchedStock) {
        await writeStock("raw", rawStock);
        await writeStock("fg", fgStock);
      }
      return res.json({ order: op });
    }
  }
  if (note !== undefined) op.note = String(note || "");
  op.updatedAt = new Date().toISOString();

  await writeMrp(db.mrp);
  res.json({ order: op });
});

async function deleteProductionOrderById(db, opId) {
  const idx = (db.mrp.productionOrders || []).findIndex((o) => o.id === opId);
  if (idx < 0) return { ok: false, reason: "not_found" };
  const op = db.mrp.productionOrders[idx];

  const rawStock = await ensureStockFile("raw");
  const fgStock = await ensureStockFile("fg");

  // Remove movements if EXECUTED
  const consumedMvIds = new Set((op.consumed || []).map((c) => c.movementId).filter(Boolean));
  const producedMvId = op.produced?.movementId;
  if (producedMvId) consumedMvIds.add(producedMvId);

  if (consumedMvIds.size) {
    rawStock.movements = (rawStock.movements || []).filter((mv) => !consumedMvIds.has(mv.id));
    fgStock.movements = (fgStock.movements || []).filter((mv) => !consumedMvIds.has(mv.id));
  }

  // Unlink linked OC (do not delete it automatically)
  if (op.linkedPurchaseOrderId) {
    const po = (db.mrp.purchaseOrders || []).find((o) => o.id === op.linkedPurchaseOrderId);
    if (po && po.linkedProductionOrderId === op.id) {
      po.linkedProductionOrderId = null;
      po.updatedAt = new Date().toISOString();
    }
  }

  // Delete OP
  db.mrp.productionOrders.splice(idx, 1);

  // If we deleted the last OP, allow wiping the OP array for this write.
  if ((db.mrp.productionOrders || []).length === 0) {
    db.mrp.meta = db.mrp.meta || {};
    db.mrp.meta.allowWipe = { ...(db.mrp.meta.allowWipe || {}), productionOrders: true };
  }
  await writeMrp(db.mrp);
  await writeStock("raw", rawStock);
  await writeStock("fg", fgStock);
  return { ok: true, op };
}

// Excluir OPs em lote — exige reauth
// (definir ANTES de /:id para não colidir com o parâmetro)
app.delete("/api/mrp/production-orders/batch", requireAuth, requirePerm("op"), requireReauthOnce, async (req, res) => {
  const { ids = [] } = req.body || {};
  if (!Array.isArray(ids) || ids.length === 0) return res.status(400).json({ error: "missing_ids" });
  const db = await ensureDB();
  const deleted = [];
  for (const id of ids) {
    const out = await deleteProductionOrderById(db, String(id));
    if (out.ok) deleted.push(String(id));
  }
  res.json({ ok: true, deleted });
});

// Excluir OP (permanente) — exige reauth
app.delete("/api/mrp/production-orders/:id", requireAuth, requirePerm("op"), requireReauthOnce, async (req, res) => {
  const db = await ensureDB();
  const out = await deleteProductionOrderById(db, String(req.params.id));
  if (!out.ok) return res.status(404).json({ error: "not_found" });
  res.json({ ok: true });
});


// RESET OP/OC + Arquivadas (permanente) — exige reauth
// Escopo padrão: OP/OC (compat). Escopo TOTAL: { scope: "all" }.
// Remove também os movimentos de estoque criados por recebimentos/consumos/produção dessas ordens,
// e (no total) os movimentos de PV/PVR + pontos + histórico.
app.post("/api/mrp/reset-orders", requireAuth, requirePerm("canReset"), requireReauthOnce, async (req, res) => {
  const { scope } = req.body || {};
  const isAll = String(scope || '').trim().toLowerCase() === 'all';

  const db = await ensureDB();

  const rawStock = await ensureStockFile("raw");
  const fgStock = await ensureStockFile("fg");

  // Carrega arquivadas para remover movimentos também
  const archPO = await ensurePurchaseOrdersArchivedFile();
  const archOP = await ensureProductionOrdersArchivedFile();

  const mvRaw = new Set();
  const mvFg = new Set();

  const collectOC = (oc) => {
    for (const rcv of (oc?.receipts || [])) {
      for (const ln of (rcv?.lines || [])) {
        if (ln?.movementId) mvRaw.add(String(ln.movementId));
      }
    }
  };
  const collectOP = (op) => {
    for (const c of (op?.consumed || [])) {
      if (c?.movementId) mvRaw.add(String(c.movementId));
    }
    if (op?.produced?.movementId) mvFg.add(String(op.produced.movementId));
  };

  for (const oc of (db.mrp.purchaseOrders || [])) collectOC(oc);
  for (const oc of (archPO.purchaseOrders || [])) collectOC(oc);

  for (const op of (db.mrp.productionOrders || [])) collectOP(op);
  for (const op of (archOP.productionOrders || [])) collectOP(op);

  // No reset TOTAL, também remove movimentos criados por PV/PVR
  if (isAll) {
    // Marcações novas (refType) + fallback por prefixo de reason (compat com builds antigos)
    for (const mv of (fgStock.movements || [])) {
      const rt = String(mv.refType || '').toUpperCase();
      const reason = String(mv.reason || '');
      if (rt === 'PV' || rt === 'PVR') mvFg.add(String(mv.id));
      else if (/^Despacho\s+PV\b/i.test(reason)) mvFg.add(String(mv.id));
      else if (/^Venda\s+r[aá]pida/i.test(reason)) mvFg.add(String(mv.id));
    }
  }

  if (mvRaw.size) rawStock.movements = (rawStock.movements || []).filter(mv => !mvRaw.has(String(mv.id)));
  if (mvFg.size) fgStock.movements = (fgStock.movements || []).filter(mv => !mvFg.has(String(mv.id)));

  // Wipe orders (main)
  db.mrp.productionOrders = [];
  db.mrp.purchaseOrders = [];
  db.mrp.meta = db.mrp.meta || {};
  db.mrp.meta.nextOpNumber = 1;
  db.mrp.meta.nextOcNumber = 1;
  // Reset também o contador de lotes (etiquetas / rastreabilidade)
  db.mrp.meta.nextLotNumber = 1;

  // Permite wipe das listas nesta gravação
  db.mrp.meta.allowWipe = { ...(db.mrp.meta.allowWipe || {}), productionOrders: true, purchaseOrders: true };

  // Wipe archived
  archPO.purchaseOrders = [];
  archPO.meta = archPO.meta || {};
  archPO.meta.resetAt = new Date().toISOString();

  archOP.productionOrders = [];
  archOP.meta = archOP.meta || {};
  archOP.meta.resetAt = new Date().toISOString();

  await writeMrp(db.mrp);
  await writePurchaseOrdersArchived(archPO);
  await writeProductionOrdersArchived(archOP);
  await writeStock("raw", rawStock);
  await writeStock("fg", fgStock);

  // Reset TOTAL: zera PV/PVR + pontos + movimentos do ponto (sem mexer em MP/PF/BOM)
  if (isAll) {
    const pointsDb = await ensureSalesPointsFile();
    pointsDb.points = [];
    pointsDb.meta = pointsDb.meta || {};
    pointsDb.meta.nextPointCode = 1;
    pointsDb.meta.resetAt = new Date().toISOString();

    const ordersDb = await ensureSalesOrdersFile();
    ordersDb.orders = [];
    ordersDb.meta = ordersDb.meta || {};
    ordersDb.meta.nextPvNumber = 1;
    ordersDb.meta.nextPvrNumber = 1;
    ordersDb.meta.resetAt = new Date().toISOString();

    const movesDb = await ensureSalesPointMovesFile();
    movesDb.moves = [];
    movesDb.meta = movesDb.meta || {};
    movesDb.meta.resetAt = new Date().toISOString();

    await writeSalesPoints(pointsDb);
    await writeSalesOrders(ordersDb);
    await writeSalesPointMoves(movesDb);
  }

  res.json({ ok: true, scope: isAll ? 'all' : 'orders' });
});


app.post("/api/mrp/production-orders", requireAuth, requirePerm("op"), async (req, res) => {
  const {
    recipeId,
    qtyToProduce,
    note = "",
    allowInsufficient = true,
    createPurchaseOrder = true,
    purchaseItems = null, // optional: [{itemId, qty}]
  } = req.body || {};

  const qtp = Number(qtyToProduce);
  if (!recipeId || !Number.isFinite(qtp) || qtp <= 0) return res.status(400).json({ error: "invalid_payload" });

  const db = await ensureDB();
  const recipe = db.mrp.recipes.find((r) => r.id === recipeId);
  if (!recipe) return res.status(404).json({ error: "recipe_not_found" });

  const factor = qtp / (Number(recipe.yieldQty) || 1);

  const rawStock = await ensureStockFile("raw");
  const fgStock = await ensureStockFile("fg");
  const stock = computeInventory(rawStock);
  const itemsById = new Map((rawStock.items || []).map((i) => [String(i.id), i]));

  // Build requirements + shortages (considera % perda do item)
  const requirements = recipe.bom.map((line) => {
    const item = itemsById.get(String(line.itemId));
    const unit = item?.unit || "";
    const lossPercent = Number(item?.lossPercent) || 0;

    const requiredNet = Number(line.qty) * factor;
    const requiredGrossRaw = requiredNet * (1 + (lossPercent / 100));

    const required = quantizeQty(requiredGrossRaw, unit);
    const available = quantizeQty((stock.get(String(line.itemId)) ?? 0), unit);
    const shortage = quantizeQty(Math.max(0, required - available), unit);

    return {
      itemId: String(line.itemId),
      unit,
      requiredNet: Number(requiredNet.toFixed(6)),
      required,
      lossPercent,
      available,
      shortage,
      ok: available >= required,
    };
  });

  const shortages = requirements.filter((r) => !r.ok && r.shortage > 0);

  const opNumber = await nextMrpNumber(db.mrp, "op");
  const order = {
    id: newId("po_"),
    number: opNumber,
    recipeId,
    recipeName: recipe.name,
    recipeCode: String(recipe.code || ""),
    qtyToProduce: qtp,
    note: String(note || ""),
    createdAt: new Date().toISOString(),
    // Status da OP:
    // - ISSUED: emitida (padrão quando gerada)
    // - IN_PRODUCTION: em produção (ao dar baixa no estoque)
    // - CLOSED: encerrada (ao dar entrada do PF no estoque)
    // - CANCELLED: cancelada (manual; se já estava em produção, estorna a baixa)
    // (Compat: READY ≈ ISSUED, EXECUTED ≈ CLOSED)
    status: "ISSUED",
    factor: Number(factor.toFixed(6)),
    planned: {
      consumed: requirements.map((r) => ({ itemId: r.itemId, qty: r.required })), // qty "gross" (com perda)
      produced: null, // preenchido abaixo
    },
    consumed: [],
    produced: null,
    shortages: shortages.map((s) => ({ itemId: s.itemId, shortage: s.shortage, required: s.required, available: s.available })),
    linkedPurchaseOrderId: null,
  };

  // Ensure output item exists
  const outItemId = ensureRecipeOutputItem(fgStock, recipe);
  order.planned.produced = { itemId: outItemId, qty: Number(qtp) };

  // If shortages exist: optionally create Purchase Order (OC) and DO NOT move stock now.
  if (shortages.length) {
    if (!allowInsufficient) {
      // keep backward compatible behavior
      const first = shortages[0];
      return res.status(400).json({
        error: "insufficient_stock",
        itemId: first.itemId,
        required: first.required,
        available: first.available,
      });
    }

    if (createPurchaseOrder) {
      const poItemsFromUI = Array.isArray(purchaseItems) ? purchaseItems : null;
      const items = (poItemsFromUI || shortages.map(s => ({ itemId: s.itemId, qty: s.shortage })))
        .map(x => ({ itemId: String(x.itemId), qtyOrdered: Number(x.qty) || 0 }))
        .filter(x => x.itemId && x.qtyOrdered > 0);

      if (items.length) {
        const ocNumber = await nextMrpNumber(db.mrp, "oc");
        const purchaseOrder = {
          id: newId("oc_"),
          number: ocNumber,
          status: "OPEN", // OPEN | PARTIAL | RECEIVED | CANCELLED
          linkedProductionOrderId: order.id,
          createdAt: new Date().toISOString(),
          updatedAt: new Date().toISOString(),
          note: `Gerado automaticamente a partir da OP ${pad6(opNumber)}.`,
          items: items.map(it => {
            const u = unitKey(itemsById.get(String(it.itemId))?.unit);
            return {
              itemId: String(it.itemId),
              // Quantiza para o que o usuário vê/edita na UI (evita "parece igual" mas fica parcial)
              qtyOrdered: quantizeQty(it.qtyOrdered, u),
              qtyAdjusted: 0,
              qtyReceived: 0,
            };
          }),
          receipts: [],
        };
        db.mrp.purchaseOrders.push(purchaseOrder);
        order.linkedPurchaseOrderId = purchaseOrder.id;
      }
    }

    db.mrp.productionOrders.push(order);
    await writeMrp(db.mrp);
    await writeStock("raw", rawStock);
    await writeStock("fg", fgStock);
    return res.json({ order });
  }

  // No shortages: do NOT move stock now. OP fica como ISSUED e será executada manualmente.
  order.status = "ISSUED";
  order.shortages = [];
  db.mrp.productionOrders.push(order);

  await writeMrp(db.mrp);
  // ensure output PF exists in stock file
  await writeStock("raw", rawStock);
  await writeStock("fg", fgStock);

  res.json({ order });
});


// Executar (baixar/produzir) uma OP que estava em HOLD/READY
app.post("/api/mrp/production-orders/:id/execute", requireAuth, requirePerm("op"), async (req, res) => {
  const db = await ensureDB();
  const order = db.mrp.productionOrders.find((o) => o.id === req.params.id);
  if (!order) return res.status(404).json({ error: "not_found" });

  // idempotente (compat: EXECUTED em BD antiga)
  if (order.status === "CLOSED" || order.status === "EXECUTED") {
    // normalize no retorno
    if (order.status === "EXECUTED") order.status = "CLOSED";
    return res.json({ order });
  }

  const recipe = db.mrp.recipes.find((r) => r.id === order.recipeId);
  if (!recipe) return res.status(400).json({ error: "recipe_not_found" });

  const rawStock = await ensureStockFile("raw");
  const fgStock = await ensureStockFile("fg");
  const stock = computeInventory(rawStock);
  const itemsById = new Map((rawStock.items || []).map((i) => [String(i.id), i]));

  // Recompute requirements to validate
  const qtp = Number(order.qtyToProduce);
  const factor = qtp / (Number(recipe.yieldQty) || 1);

  const requirements = recipe.bom.map((line) => {
    const item = itemsById.get(String(line.itemId));
    const unit = item?.unit || "";
    const lossPercent = Number(item?.lossPercent) || 0;

    const requiredNet = Number(line.qty) * factor;
    const requiredGrossRaw = requiredNet * (1 + (lossPercent / 100));

    const required = quantizeQty(requiredGrossRaw, unit);
    const available = quantizeQty((stock.get(String(line.itemId)) ?? 0), unit);
    const shortage = quantizeQty(Math.max(0, required - available), unit);

    return { itemId: String(line.itemId), unit, required, available, shortage, ok: available >= required };
  });

  const shortages = requirements.filter(r => !r.ok && r.shortage > 0);
  if (shortages.length) {
    return res.status(400).json({ error: "insufficient_stock", shortages });
  }

  // Post movements (idempotente)
  rawStock.movements = rawStock.movements || [];
  fgStock.movements = fgStock.movements || [];

  const hasConsumed = Array.isArray(order.consumed) && order.consumed.length > 0;
  if (!hasConsumed) {
    order.consumed = [];
    for (const r of requirements) {
      const mv = {
        id: newId("mv_"),
        type: "out",
        itemId: r.itemId,
        qty: r.required,
        reason: `Consumo OP ${pad6(order.number)} (${order.recipeName || recipe.name})`,
        at: new Date().toISOString(),
      };
      rawStock.movements.push(mv);
      order.consumed.push({ itemId: r.itemId, qty: r.required, movementId: mv.id });
    }
    order.startedAt = order.startedAt || new Date().toISOString();
  }

  const hasProduced = !!order.produced?.movementId;
  if (!hasProduced) {
    const outItemId = ensureRecipeOutputItem(fgStock, recipe);
    const mvIn = {
      id: newId("mv_"),
      type: "in",
      itemId: outItemId,
      qty: Number(qtp),
      reason: `Produção OP ${pad6(order.number)} (${order.recipeName || recipe.name})`,
      at: new Date().toISOString(),
    };
    fgStock.movements.push(mvIn);
    order.produced = { itemId: outItemId, qty: Number(qtp), movementId: mvIn.id };
  }

  // Lote / Código de barras (para etiquetas)
  if (!order.lotNumber) {
    const lotN = await nextLotNumber(db.mrp);
    order.lotNumber = lotN;
    order.lotCode = formatLotCode(lotN);
    const rc = String(order.recipeCode || recipe.code || "");
    order.recipeCode = rc;
    order.barcodeValue = rc ? `DIETON-${rc}-${order.lotCode}` : `DIETON-${order.lotCode}`;
    order.lotCreatedAt = new Date().toISOString();
  }

  order.status = "CLOSED";
  order.closedAt = new Date().toISOString();
  order.executedAt = order.executedAt || order.closedAt;
  order.shortages = [];
  await writeMrp(db.mrp);
  await writeStock("raw", rawStock);
  await writeStock("fg", fgStock);

  res.json({ order });
});



// Arquivar OP (somente CLOSED/EXECUTED ou CANCELLED) — move da lista principal para /bd/production_orders_archived.json
app.post("/api/mrp/production-orders/:id/archive", requireAuth, requirePerm("op"), async (req, res) => {
  const db = await ensureDB();
  const opId = String(req.params.id);
  const idx = (db.mrp.productionOrders || []).findIndex((o) => o.id === opId);
  if (idx < 0) return res.status(404).json({ error: "not_found" });

  const op = db.mrp.productionOrders[idx];
  const st = String(op.status || "").toUpperCase();
  if (!(st === "CLOSED" || st === "EXECUTED" || st === "CANCELLED")) {
    return res.status(400).json({ error: "not_allowed", message: "Só é possível arquivar OP Encerrada ou Cancelada." });
  }
  // normalize legacy
  if (op.status === "EXECUTED") op.status = "CLOSED";

  // Unlink linked OC (avoid broken references)
  if (op.linkedPurchaseOrderId) {
    const po = (db.mrp.purchaseOrders || []).find((o) => o.id === op.linkedPurchaseOrderId);
    if (po && po.linkedProductionOrderId === op.id) {
      po.linkedProductionOrderId = null;
      po.updatedAt = new Date().toISOString();
    }
  }

  db.mrp.productionOrders.splice(idx, 1);
  if ((db.mrp.productionOrders || []).length === 0) {
    db.mrp.meta = db.mrp.meta || {};
    db.mrp.meta.allowWipe = { ...(db.mrp.meta.allowWipe || {}), productionOrders: true };
  }

  const arch = await ensureProductionOrdersArchivedFile();
  arch.productionOrders = Array.isArray(arch.productionOrders) ? arch.productionOrders : [];
  // de-dup por segurança
  arch.productionOrders = arch.productionOrders.filter((o) => o.id !== op.id);
  op.archivedAt = new Date().toISOString();
  op.updatedAt = new Date().toISOString();
  arch.productionOrders.unshift(op);

  await writeMrp(db.mrp);
  await writeProductionOrdersArchived(arch);
  res.json({ ok: true });
});

// Restaurar OP do arquivo para a lista principal
app.post("/api/mrp/production-orders/archived/:id/restore", requireAuth, requirePerm("op"), async (req, res) => {
  const opId = String(req.params.id);
  const arch = await ensureProductionOrdersArchivedFile();
  arch.productionOrders = Array.isArray(arch.productionOrders) ? arch.productionOrders : [];
  const idx = arch.productionOrders.findIndex((o) => o.id === opId);
  if (idx < 0) return res.status(404).json({ error: "not_found" });

  const op = arch.productionOrders[idx];
  arch.productionOrders.splice(idx, 1);
  try { delete op.archivedAt; } catch {}
  op.updatedAt = new Date().toISOString();

  const db = await ensureDB();
  db.mrp.productionOrders = Array.isArray(db.mrp.productionOrders) ? db.mrp.productionOrders : [];
  if (!db.mrp.productionOrders.find((o) => o.id === op.id)) {
    db.mrp.productionOrders.push(op);
  }

  // best-effort relink OC if it exists and is not linked
  if (op.linkedPurchaseOrderId) {
    const po = (db.mrp.purchaseOrders || []).find((o) => o.id === op.linkedPurchaseOrderId);
    if (po && !po.linkedProductionOrderId) {
      po.linkedProductionOrderId = op.id;
      po.updatedAt = new Date().toISOString();
    }
  }

  await writeMrp(db.mrp);
  await writeProductionOrdersArchived(arch);
  res.json({ ok: true });
});

async function deleteArchivedProductionOrderById(opId) {
  const arch = await ensureProductionOrdersArchivedFile();
  arch.productionOrders = Array.isArray(arch.productionOrders) ? arch.productionOrders : [];
  const idx = arch.productionOrders.findIndex((o) => o.id === opId);
  if (idx < 0) return { ok: false, reason: "not_found" };
  const op = arch.productionOrders[idx];

  const rawStock = await ensureStockFile("raw");
  const fgStock = await ensureStockFile("fg");

  // Remove movements if EXECUTED
  const mvIds = new Set((op.consumed || []).map((c) => c.movementId).filter(Boolean));
  const producedMvId = op.produced?.movementId;
  if (producedMvId) mvIds.add(producedMvId);
  if (mvIds.size) {
    rawStock.movements = (rawStock.movements || []).filter((mv) => !mvIds.has(mv.id));
    fgStock.movements = (fgStock.movements || []).filter((mv) => !mvIds.has(mv.id));
  }

  // Unlink linked OC (do not delete it)
  if (op.linkedPurchaseOrderId) {
    const db = await ensureDB();
    const po = (db.mrp.purchaseOrders || []).find((o) => o.id === op.linkedPurchaseOrderId);
    if (po && po.linkedProductionOrderId === op.id) {
      po.linkedProductionOrderId = null;
      po.updatedAt = new Date().toISOString();
      await writeMrp(db.mrp);
    }
  }

  arch.productionOrders.splice(idx, 1);
  await writeProductionOrdersArchived(arch);
  await writeStock("raw", rawStock);
  await writeStock("fg", fgStock);
  return { ok: true };
}

// Excluir OP arquivada (permanente) — exige reauth
app.delete("/api/mrp/production-orders/archived/:id", requireAuth, requirePerm("op"), requireReauthOnce, async (req, res) => {
  const out = await deleteArchivedProductionOrderById(String(req.params.id));
  if (!out.ok) return res.status(404).json({ error: "not_found" });
  res.json({ ok: true });
});


// ---------- Purchase Orders (OC) ----------
app.get("/api/mrp/purchase-orders", requireAuth, requirePerm("oc"), async (req, res) => {
  const db = await ensureDB();
  const orders = [...(db.mrp.purchaseOrders || [])].sort((a, b) => (a.createdAt < b.createdAt ? 1 : -1));
  res.json({ orders });
});

// Archived Purchase Orders (separado)
app.get("/api/mrp/purchase-orders/archived", requireAuth, requirePerm("oc"), async (req, res) => {
  const arch = await ensurePurchaseOrdersArchivedFile();
  const orders = [...(arch.purchaseOrders || [])].sort((a, b) => (
    String(a.archivedAt || a.createdAt || "") < String(b.archivedAt || b.createdAt || "") ? 1 : -1
  ));
  res.json({ orders });
});

app.post("/api/mrp/purchase-orders", requireAuth, requirePerm("oc"), async (req, res) => {
  const { note = "", linkedProductionOrderId = null, items = [] } = req.body || {};
  if (!Array.isArray(items) || items.length === 0) return res.status(400).json({ error: "missing_items" });

  const db = await ensureDB();
  const rawStock = await ensureStockFile("raw");
  const unitById = new Map((rawStock.items || []).map(i => [String(i.id), unitKey(i.unit)]));
  const ocNumber = await nextMrpNumber(db.mrp, "oc");
  const po = {
    id: newId("oc_"),
    number: ocNumber,
    status: "OPEN",
    linkedProductionOrderId: linkedProductionOrderId ? String(linkedProductionOrderId) : null,
    createdAt: new Date().toISOString(),
    updatedAt: new Date().toISOString(),
    note: String(note || ""),
    items: items.map(it => {
      const itemId = String(it.itemId);
      const u = unitById.get(itemId) || "";
      const ord = quantizeQty(Number(it.qtyOrdered || it.qty || 0), u);
      const adj = quantizeQty(Number(it.qtyAdjusted || 0), u);
      const rec = quantizeQty(Number(it.qtyReceived || 0), u);
      return { itemId, qtyOrdered: ord, qtyAdjusted: adj, qtyReceived: rec };
    }).filter(it => {
      const u = unitById.get(String(it.itemId)) || "";
      const finalQty = quantizeQty(calcFinalQty(it.qtyOrdered, it.qtyAdjusted), u);
      return it.itemId && finalQty > 0;
    }),
    receipts: [],
  };
  // Validação: não permitir que o pedido final fique menor que o já recebido (geralmente 0 no create)
  for (const it of po.items) {
    const u = unitById.get(String(it.itemId)) || "";
    const finalQty = quantizeQty(calcFinalQty(it.qtyOrdered, it.qtyAdjusted), u);
    const rec = quantizeQty(it.qtyReceived, u);
    if (finalQty < rec) {
      return res.status(400).json({ error: "adjusted_below_received", itemId: it.itemId });
    }
  }
  if (po.items.length === 0) return res.status(400).json({ error: "invalid_items" });

  db.mrp.purchaseOrders.push(po);
  await writeMrp(db.mrp);
  res.json({ order: po });
});

app.put("/api/mrp/purchase-orders/:id", requireAuth, requirePerm("oc"), async (req, res) => {
  const db = await ensureDB();
  const po = (db.mrp.purchaseOrders || []).find(o => o.id === req.params.id);
  if (!po) return res.status(404).json({ error: "not_found" });

  const { note, status, items, linkedProductionOrderId } = req.body || {};
  if (note !== undefined) po.note = String(note || "");
  if (status !== undefined) po.status = String(status || po.status);
  if (linkedProductionOrderId !== undefined) po.linkedProductionOrderId = linkedProductionOrderId ? String(linkedProductionOrderId) : null;

  if (items !== undefined) {
    if (!Array.isArray(items) || items.length === 0) return res.status(400).json({ error: "invalid_items" });
    const rawStock = await ensureStockFile("raw");
    const unitById = new Map((rawStock.items || []).map(i => [String(i.id), unitKey(i.unit)]));
    const newItems = items.map(it => {
      const itemId = String(it.itemId);
      const u = unitById.get(itemId) || "";
      const ord = quantizeQty(Number(it.qtyOrdered || it.qty || 0), u);
      const adj = quantizeQty(Number(it.qtyAdjusted || 0), u);
      const rec = quantizeQty(Number(it.qtyReceived || 0), u);
      return { itemId, qtyOrdered: ord, qtyAdjusted: adj, qtyReceived: rec };
    }).filter(it => {
      const u = unitById.get(String(it.itemId)) || "";
      const finalQty = quantizeQty(calcFinalQty(it.qtyOrdered, it.qtyAdjusted), u);
      return it.itemId && finalQty > 0;
    });
    // Validação: não permitir que o pedido final fique menor que o já recebido
    for (const it of newItems) {
      const u = unitById.get(String(it.itemId)) || "";
      const finalQty = quantizeQty(calcFinalQty(it.qtyOrdered, it.qtyAdjusted), u);
      const rec = quantizeQty(it.qtyReceived, u);
      if (finalQty < rec) {
        return res.status(400).json({ error: "adjusted_below_received", itemId: it.itemId });
      }
    }
    if (newItems.length === 0) return res.status(400).json({ error: "invalid_items" });
    po.items = newItems;
  }

  po.updatedAt = new Date().toISOString();
  await writeMrp(db.mrp);
  res.json({ order: po });
});

// Arquivar OC (move da lista principal para /bd/purchase_orders_archived.json)
app.post("/api/mrp/purchase-orders/:id/archive", requireAuth, requirePerm("oc"), async (req, res) => {
  const db = await ensureDB();
  const poId = String(req.params.id);
  const idx = (db.mrp.purchaseOrders || []).findIndex((o) => o.id === poId);
  if (idx < 0) return res.status(404).json({ error: "not_found" });

  const po = db.mrp.purchaseOrders[idx];

  // Unlink linked OP (avoid broken references)
  if (po.linkedProductionOrderId) {
    const op = (db.mrp.productionOrders || []).find((o) => o.id === po.linkedProductionOrderId);
    if (op && op.linkedPurchaseOrderId === po.id) {
      op.linkedPurchaseOrderId = null;
    }
  }

  db.mrp.purchaseOrders.splice(idx, 1);
  if ((db.mrp.purchaseOrders || []).length === 0) {
    db.mrp.meta = db.mrp.meta || {};
    db.mrp.meta.allowWipe = { ...(db.mrp.meta.allowWipe || {}), purchaseOrders: true };
  }

  const arch = await ensurePurchaseOrdersArchivedFile();
  arch.purchaseOrders = Array.isArray(arch.purchaseOrders) ? arch.purchaseOrders : [];
  // de-dup por segurança
  arch.purchaseOrders = arch.purchaseOrders.filter((o) => o.id !== po.id);
  po.archivedAt = new Date().toISOString();
  po.updatedAt = new Date().toISOString();
  arch.purchaseOrders.unshift(po);

  await writeMrp(db.mrp);
  await writePurchaseOrdersArchived(arch);
  res.json({ ok: true });
});

// Restaurar OC do arquivo para a lista principal
app.post("/api/mrp/purchase-orders/archived/:id/restore", requireAuth, requirePerm("oc"), async (req, res) => {
  const poId = String(req.params.id);
  const arch = await ensurePurchaseOrdersArchivedFile();
  arch.purchaseOrders = Array.isArray(arch.purchaseOrders) ? arch.purchaseOrders : [];
  const idx = arch.purchaseOrders.findIndex((o) => o.id === poId);
  if (idx < 0) return res.status(404).json({ error: "not_found" });

  const po = arch.purchaseOrders[idx];
  arch.purchaseOrders.splice(idx, 1);
  try { delete po.archivedAt; } catch {}
  po.updatedAt = new Date().toISOString();

  const db = await ensureDB();
  db.mrp.purchaseOrders = Array.isArray(db.mrp.purchaseOrders) ? db.mrp.purchaseOrders : [];
  if (!db.mrp.purchaseOrders.find((o) => o.id === po.id)) {
    db.mrp.purchaseOrders.push(po);
  }

  await writeMrp(db.mrp);
  await writePurchaseOrdersArchived(arch);
  res.json({ ok: true });
});

async function deleteArchivedPurchaseOrderById(poId) {
  const arch = await ensurePurchaseOrdersArchivedFile();
  arch.purchaseOrders = Array.isArray(arch.purchaseOrders) ? arch.purchaseOrders : [];
  const idx = arch.purchaseOrders.findIndex((o) => o.id === poId);
  if (idx < 0) return { ok: false, reason: "not_found" };
  const po = arch.purchaseOrders[idx];

  const rawStock = await ensureStockFile("raw");

  // Remove stock movements created by receipts (same semantics as delete from main list)
  const mvIds = new Set();
  for (const rcv of (po.receipts || [])) {
    for (const ln of (rcv.lines || [])) {
      if (ln && ln.movementId) mvIds.add(String(ln.movementId));
    }
  }
  if (mvIds.size) {
    rawStock.movements = (rawStock.movements || []).filter((mv) => !mvIds.has(mv.id));
  }

  arch.purchaseOrders.splice(idx, 1);
  await writePurchaseOrdersArchived(arch);
  await writeStock("raw", rawStock);
  return { ok: true };
}

// Excluir OC arquivada (permanente) — exige reauth
app.delete("/api/mrp/purchase-orders/archived/:id", requireAuth, requirePerm("oc"), requireReauthOnce, async (req, res) => {
  const out = await deleteArchivedPurchaseOrderById(String(req.params.id));
  if (!out.ok) return res.status(404).json({ error: "not_found" });
  res.json({ ok: true });
});

async function deletePurchaseOrderById(db, poId) {
  const idx = (db.mrp.purchaseOrders || []).findIndex((o) => o.id === poId);
  if (idx < 0) return { ok: false, reason: "not_found" };
  const po = db.mrp.purchaseOrders[idx];

  const rawStock = await ensureStockFile("raw");

  // Remove stock movements created by receipts
  const mvIds = new Set();
  for (const rcv of (po.receipts || [])) {
    for (const ln of (rcv.lines || [])) {
      if (ln && ln.movementId) mvIds.add(String(ln.movementId));
    }
  }
  if (mvIds.size) {
    rawStock.movements = (rawStock.movements || []).filter((mv) => !mvIds.has(mv.id));
  }

  // Unlink linked OP (do not delete it)
  if (po.linkedProductionOrderId) {
    const op = (db.mrp.productionOrders || []).find((o) => o.id === po.linkedProductionOrderId);
    if (op && op.linkedPurchaseOrderId === po.id) {
      op.linkedPurchaseOrderId = null;
      // Mantém status do OP, mas evita referência quebrada.
    }
  }

  db.mrp.purchaseOrders.splice(idx, 1);

  // If we deleted the last OC, allow wiping the OC array for this write.
  if ((db.mrp.purchaseOrders || []).length === 0) {
    db.mrp.meta = db.mrp.meta || {};
    db.mrp.meta.allowWipe = { ...(db.mrp.meta.allowWipe || {}), purchaseOrders: true };
  }

  await writeMrp(db.mrp);
  await writeStock("raw", rawStock);
  return { ok: true, po };
}

// Excluir OCs em lote — exige reauth
// (definir ANTES de /:id para não colidir com o parâmetro)
app.delete("/api/mrp/purchase-orders/batch", requireAuth, requirePerm("oc"), requireReauthOnce, async (req, res) => {
  const { ids = [] } = req.body || {};
  if (!Array.isArray(ids) || ids.length === 0) return res.status(400).json({ error: "missing_ids" });
  const db = await ensureDB();
  const deleted = [];
  for (const id of ids) {
    const out = await deletePurchaseOrderById(db, String(id));
    if (out.ok) deleted.push(String(id));
  }
  res.json({ ok: true, deleted });
});

// Excluir OC (permanente) — exige reauth
app.delete("/api/mrp/purchase-orders/:id", requireAuth, requirePerm("oc"), requireReauthOnce, async (req, res) => {
  const db = await ensureDB();
  const out = await deletePurchaseOrderById(db, String(req.params.id));
  if (!out.ok) return res.status(404).json({ error: "not_found" });
  res.json({ ok: true });
});

// Receber itens da OC (gera entrada no estoque de MP)
app.post("/api/mrp/purchase-orders/:id/receive", requireAuth, requirePerm("oc"), async (req, res) => {
  const { items = [], note = "", finalize = false } = req.body || {};
  if (!Array.isArray(items) || items.length === 0) return res.status(400).json({ error: "missing_items" });

  const db = await ensureDB();
  const po = (db.mrp.purchaseOrders || []).find(o => o.id === req.params.id);
  if (!po) return res.status(404).json({ error: "not_found" });

  const rawStock = await ensureStockFile("raw");
  const unitById = new Map((rawStock.items || []).map(it => [String(it.id), unitKey(it.unit)]));

  const receipt = {
    id: newId("rcv_"),
    at: new Date().toISOString(),
    note: String(note || ""),
    lines: [],
  };

  for (const line of items) {
    const itemId = String(line.itemId || "");
    const unit = unitById.get(itemId) || "";
    const qtyRaw = Number(line.qty || line.qtyReceived || 0);
    const qty = quantizeQty(qtyRaw, unit);
    if (!itemId || !Number.isFinite(qtyRaw) || qty <= 0) continue;

    const poLine = (po.items || []).find(x => x.itemId === itemId);
    if (!poLine) continue;

    const mv = {
      id: newId("mv_"),
      type: "in",
      itemId,
      qty: Number(qty.toFixed(6)),
      reason: `Entrada OC ${pad6(po.number)}${receipt.note ? " - " + receipt.note : ""}`,
      at: receipt.at,
    };
    rawStock.movements.push(mv);

    poLine.qtyReceived = Number((Number(poLine.qtyReceived || 0) + qty).toFixed(6));
    receipt.lines.push({ itemId, qty: Number(qty.toFixed(6)), movementId: mv.id });
  }

  if (receipt.lines.length === 0) return res.status(400).json({ error: "invalid_items" });

  po.receipts = Array.isArray(po.receipts) ? po.receipts : [];
  po.receipts.push(receipt);

  // Se o usuário optou por 'finalizar' (ex.: gerou nova OC faltante),
  // ajusta o pedido final para bater com o recebido e remove itens não recebidos.
  if (finalize) {
    po.items = (po.items || []).map((x) => {
      const unit = unitById.get(String(x.itemId)) || "";
      const ord = Number(x.qtyOrdered || 0);
      const rec = Number(x.qtyReceived || 0);
      const adj = quantizeQty((rec - ord), unit);
      return { ...x, qtyAdjusted: Number(adj.toFixed(6)) };
    }).filter((x) => {
      const unit = unitById.get(String(x.itemId)) || "";
      const rec = Math.max(0, quantizeQty(Number(x.qtyReceived || 0), unit));
      return rec > 0;
    });
  }

  // Update status (comparação quantizada para evitar falsos "PARCIAL" por diferença de casas decimais)
  const allReceived = (po.items || []).every(x => {
    const unit = unitById.get(String(x.itemId)) || "";
    const ord = Number(x.qtyOrdered || 0);
    const adj = Number(x.qtyAdjusted || 0);
    const finalQty = Math.max(0, quantizeQty(calcFinalQty(ord, adj), unit));
    const rec = Math.max(0, quantizeQty(Number(x.qtyReceived || 0), unit));
    return rec >= finalQty;
  });
  const anyReceived = (po.items || []).some(x => (Number(x.qtyReceived || 0) > 0));
  po.status = allReceived ? "RECEIVED" : (anyReceived ? "PARTIAL" : (po.status || "OPEN"));
  po.updatedAt = new Date().toISOString();

  // Se houver OP vinculada, atualizar faltas (e liberar de HOLD quando estoque já cobre)
  if (po.linkedProductionOrderId) {
    const op = (db.mrp.productionOrders || []).find(o => o.id === po.linkedProductionOrderId);
    const cur = String(op?.status || "").toUpperCase();
    if (op && !(cur === "CLOSED" || cur === "EXECUTED" || cur === "CANCELLED")) {
      const recipe = db.mrp.recipes.find(r => r.id === op.recipeId);
      if (recipe) {
        const stockNow = computeInventory(rawStock);
        const itemsById = new Map((rawStock.items || []).map((i) => [String(i.id), i]));
        const qtp = Number(op.qtyToProduce);
        const factor = qtp / (Number(recipe.yieldQty) || 1);
        const reqs = recipe.bom.map((line) => {
          const item = itemsById.get(String(line.itemId));
          const unit = item?.unit || "";
          const lossPercent = Number(item?.lossPercent) || 0;

          const requiredNet = Number(line.qty) * factor;
          const requiredGrossRaw = requiredNet * (1 + (lossPercent / 100));

          const required = quantizeQty(requiredGrossRaw, unit);
          const available = quantizeQty((stockNow.get(String(line.itemId)) ?? 0), unit);
          const shortage = quantizeQty(Math.max(0, required - available), unit);

          return { itemId: String(line.itemId), unit, required, available, shortage, ok: available >= required };
        });
        const shorts = reqs.filter(r => !r.ok && r.shortage > 0).map(r => ({ itemId: r.itemId, shortage: r.shortage, required: r.required, available: r.available }));
        op.shortages = shorts;
        // Não existe mais "Em espera". A OP permanece EMITIDA até o usuário iniciar a produção.
        // (Se vier de BD antiga com HOLD/READY, normaliza para ISSUED aqui.)
        const cur2 = String(op.status || "").toUpperCase();
        if (cur2 !== "IN_PRODUCTION") op.status = "ISSUED";
      }
    }
  }

  await writeMrp(db.mrp);
  await writeStock("raw", rawStock);
  res.json({ order: po });
});

// -------------------- Pedidos de Venda / Pontos --------------------

app.get('/api/sales/points', requireAuth, requirePerm("sales"), async (req, res) => {
  await ensureBD();
  const db = await ensureSalesPointsFile();
  const points = [...(db.points || [])].sort((a,b) => String(a.code||"").localeCompare(String(b.code||""), 'pt-BR'));
  res.json({ points });
});

app.post('/api/sales/points', requireAuth, requirePerm("sales"), async (req, res) => {
  const { name, address = '', note = '', code: reqCode = '' } = req.body || {};
  const nm = String(name || '').trim();
  if (!nm) return res.status(400).json({ error: 'missing_name' });

const db = await ensureSalesPointsFile();
let code = '';
const userCode = normalizeSalesPointCode(reqCode);
if (userCode) {
  const exists = (db.points || []).some(x => String(x.code || '').toUpperCase() === String(userCode).toUpperCase());
  if (exists) return res.status(409).json({ error: 'code_exists' });
  code = userCode;
} else {
  code = await nextSalesPointCode(db);
}
  const p = {
    id: newId('pt_'),
    code,
    name: nm,
    address: String(address || '').trim(),
    note: String(note || '').trim(),
    createdAt: new Date().toISOString(),
  };
  db.points.push(p);
  await writeSalesPoints(db);
  res.json({ point: p });
});

app.put('/api/sales/points/:id', requireAuth, requirePerm("sales"), async (req, res) => {
  const id = String(req.params.id || '');
  const { name, address = '', note = '' } = req.body || {};
  const nm = String(name || '').trim();
  if (!nm) return res.status(400).json({ error: 'missing_name' });
  const db = await ensureSalesPointsFile();
  const p = (db.points || []).find(x => String(x.id) === id);
  if (!p) return res.status(404).json({ error: 'not_found' });
  p.name = nm;
  p.address = String(address || '').trim();
  p.note = String(note || '').trim();
  p.updatedAt = new Date().toISOString();
  await writeSalesPoints(db);
  res.json({ point: p });
});

app.delete('/api/sales/points/:id', requireAuth, requirePerm("sales"), requireReauthOnce, async (req, res) => {
  const id = String(req.params.id || '');
  const pointsDb = await ensureSalesPointsFile();
  const idx = (pointsDb.points || []).findIndex(x => String(x.id) === id);
  if (idx < 0) return res.status(404).json({ error: 'not_found' });
  const p = pointsDb.points[idx];

  // bloqueia se houver pedidos ativos do ponto
  const ordersDb = await ensureSalesOrdersFile();
  const hasOrders = (ordersDb.orders || []).some(o => String(o.pointId) === id && String(o.status||'').toUpperCase() !== 'CANCELLED');
  if (hasOrders) return res.status(400).json({ error: 'point_has_orders' });

  pointsDb.points.splice(idx, 1);
  await writeSalesPoints(pointsDb);

  // remove movimentos do ponto (limpeza)
  const movesDb = await ensureSalesPointMovesFile();
  movesDb.moves = (movesDb.moves || []).filter(mv => String(mv.pointId) !== id);
  await writeSalesPointMoves(movesDb);

  res.json({ ok: true, deleted: { id: p.id, code: p.code } });
});

app.get('/api/sales/points/:id/stock', requireAuth, requirePerm("sales"), async (req, res) => {
  const id = String(req.params.id || '');
  const pointsDb = await ensureSalesPointsFile();
  const p = (pointsDb.points || []).find(x => String(x.id) === id);
  if (!p) return res.status(404).json({ error: 'not_found' });

  const movesDb = await ensureSalesPointMovesFile();
  const inv = computePointInventory(movesDb, id);

  const fgStock = await ensureStockFile('fg');
  const itemsById = new Map((fgStock.items || []).map(it => [String(it.id), it]));

  const items = [...inv.entries()]
    .map(([itemId, qty]) => {
      const it = itemsById.get(String(itemId));
      const unit = fmtUnit(it?.unit || '');
      return {
        itemId: String(itemId),
        code: it?.code || '',
        name: it?.name || '',
        unit,
        qty: quantizeQty(qty, unit),
      };
    })
    .filter(x => x.qty !== 0)
    .sort((a,b) => String(a.code||'').localeCompare(String(b.code||''), 'pt-BR'));

  res.json({ point: { id: p.id, code: p.code, name: p.name }, items });
});


app.get('/api/sales/orders', requireAuth, requirePerm("sales"), async (req, res) => {
  await ensureBD();
  const db = await ensureSalesOrdersFile();
  // Normaliza campo 'archived' para builds antigos
  const orders = [...(db.orders || [])]
    .map(o => ({ ...o, archived: !!o.archived }))
    .sort((a,b) => String(b.createdAt||'').localeCompare(String(a.createdAt||'')));
  res.json({ orders });
});

app.post('/api/sales/orders/point', requireAuth, requirePerm("sales"), async (req, res) => {
  const { pointId, items, number } = req.body || {};
  const pid = String(pointId || '');
  if (!pid) return res.status(400).json({ error: 'missing_pointId' });
  if (!Array.isArray(items) || !items.length) return res.status(400).json({ error: 'missing_items' });

  const pointsDb = await ensureSalesPointsFile();
  const pt = (pointsDb.points || []).find(x => String(x.id) === pid);
  if (!pt) return res.status(400).json({ error: 'point_not_found' });

  const fgStock = await ensureStockFile('fg');
  const itemsById = new Map((fgStock.items || []).map(it => [String(it.id), it]));

  const lines = [];
  for (const it of items) {
    const itemId = String(it.itemId || '');
    const qty = Number(it.qty);
    if (!itemId) continue;
    if (!Number.isFinite(qty) || qty <= 0) continue;
    const invIt = itemsById.get(String(itemId));
    if (!invIt) return res.status(400).json({ error: 'item_not_found', itemId });
    const unit = fmtUnit(invIt.unit);
    const q = quantizeQty(qty, unit);
    lines.push({ itemId, code: invIt.code || '', name: invIt.name || '', unit, qty: q });
  }
  if (!lines.length) return res.status(400).json({ error: 'invalid_items' });

  const ordersDb = await ensureSalesOrdersFile();

  // Número manual (opcional) — aceita: 1, 001, PV001, PV000001
  let manual = null;
  if (number !== undefined && number !== null && String(number).trim() !== '') {
    manual = parseSalesOrderNumberInput(number);
    if (!manual) return res.status(400).json({ error: 'invalid_number' });

    const exists = (ordersDb.orders || []).some(o => getSalesOrderSeries(o) === 'PV' && Number(o.number) === manual);
    if (exists) return res.status(409).json({ error: 'number_exists' });

    // Se manual for maior que o contador, avança para não colidir no futuro
    const nextPv = Number(ordersDb.meta?.nextPvNumber) || 1;
    if (manual >= nextPv) ordersDb.meta.nextPvNumber = manual + 1;
  }

  const num = manual || nextSalesOrderNumber(ordersDb, 'PV');
  const order = {
    id: newId('so_'),
    series: 'PV',
    number: num,
    type: 'POINT',
    pointId: pid,
    status: 'OPEN',
    archived: false,
    items: lines,
    linkedOps: [],
    createdAt: new Date().toISOString(),
  };
  ordersDb.orders.push(order);
  await writeSalesOrders(ordersDb);
  res.json({ order });
});

app.post('/api/sales/orders/quick', requireAuth, requirePerm("sales"), async (req, res) => {
  const { channel = 'Delivery', items } = req.body || {};
  if (!Array.isArray(items) || !items.length) return res.status(400).json({ error: 'missing_items' });

  const fgStock = await ensureStockFile('fg');
  const itemsById = new Map((fgStock.items || []).map(it => [String(it.id), it]));
  const stockMap = computeInventory(fgStock);

  const user = (await ensureDB()).users.find(u => u.id === req.session.userId);
  const by = user ? { id: user.id, name: user.name } : { id: String(req.session.userId || ''), name: '' };

  // 1) valida estoque e prepara linhas
  const shortages = [];
  const lines = [];
  for (const it of items) {
    const itemId = String(it.itemId || '');
    const qty = Number(it.qty);
    if (!itemId) continue;
    if (!Number.isFinite(qty) || qty <= 0) continue;
    const invIt = itemsById.get(String(itemId));
    if (!invIt) return res.status(400).json({ error: 'item_not_found', itemId });
    const unit = fmtUnit(invIt.unit);
    const q = quantizeQty(qty, unit);
    const available = quantizeQty(Number(stockMap.get(itemId) || 0), unit);
    if (available - q < 0) {
      shortages.push({ itemId, code: invIt.code || '', name: invIt.name || '', unit, available, shortage: quantizeQty(q - available, unit) });
      continue;
    }
    lines.push({ itemId, code: invIt.code || '', name: invIt.name || '', unit, qty: q });
  }

  if (shortages.length) return res.status(400).json({ error: 'insufficient_fg_stock', shortages });
  if (!lines.length) return res.status(400).json({ error: 'invalid_items' });

  // 2) gera PVR e aplica movimentos
  const ordersDb = await ensureSalesOrdersFile();
  const number = nextSalesOrderNumber(ordersDb, 'PVR');
  const orderId = newId('so_');
  const code = formatSalesOrderCode('PVR', number);
  const now = new Date().toISOString();

  const moves = [];
  for (const ln of lines) {
    const itemId = String(ln.itemId);
    const unit = fmtUnit(ln.unit);
    const q = quantizeQty(ln.qty, unit);
    const beforeQty = quantizeQty(Number(stockMap.get(itemId) || 0), unit);
    const afterQty = quantizeQty(beforeQty - q, unit);
    stockMap.set(itemId, afterQty);

    moves.push({
      id: newId('mv_'),
      type: 'out',
      itemId,
      qty: q,
      reason: `Venda rápida ${code} (${String(channel || 'Delivery')})`,
      refType: 'PVR',
      refId: orderId,
      refNumber: number,
      refCode: code,
      at: now,
      by,
      beforeQty,
      afterQty,
      delta: quantizeQty(afterQty - beforeQty, unit),
    });
  }

  for (const mv of moves) fgStock.movements.push(mv);
  await writeStock('fg', fgStock);

  const order = {
    id: orderId,
    series: 'PVR',
    number,
    type: 'QUICK',
    channel: String(channel || 'Delivery'),
    status: 'DONE',
    archived: false,
    items: lines,
    movementIds: moves.map(m => m.id),
    createdAt: now,
  };
  ordersDb.orders.push(order);
  await writeSalesOrders(ordersDb);
  res.json({ order });
});



// Arquivar pedidos/vendas (mantém histórico; não é exclusão)
app.post('/api/sales/orders/archive-batch', requireAuth, requirePerm("sales"), async (req, res) => {
  const { ids } = req.body || {};
  if (!Array.isArray(ids) || !ids.length) return res.status(400).json({ error: 'missing_ids' });
  const want = new Set(ids.map(x => String(x)));

  const ordersDb = await ensureSalesOrdersFile();
  const now = new Date().toISOString();
  const archived = [];

  for (const o of (ordersDb.orders || [])) {
    if (!want.has(String(o.id))) continue;
    if (!o.archived) {
      o.archived = true;
      o.archivedAt = now;
    }
    o.updatedAt = now;
    archived.push({ id: o.id, number: o.number, type: o.type });
  }

  await writeSalesOrders(ordersDb);
  res.json({ ok: true, archived });
});

app.post('/api/sales/orders/:id/archive', requireAuth, requirePerm("sales"), async (req, res) => {
  const id = String(req.params.id || '');
  const ordersDb = await ensureSalesOrdersFile();
  const o = (ordersDb.orders || []).find(x => String(x.id) === id);
  if (!o) return res.status(404).json({ error: 'not_found' });
  const now = new Date().toISOString();
  o.archived = true;
  o.archivedAt = now;
  o.updatedAt = now;
  await writeSalesOrders(ordersDb);
  res.json({ order: o });
});
app.get('/api/sales/orders/:id/plan', requireAuth, requirePerm("sales"), async (req, res) => {
  const id = String(req.params.id || '');
  const ordersDb = await ensureSalesOrdersFile();
  const order = (ordersDb.orders || []).find(o => String(o.id) === id);
  if (!order) return res.status(404).json({ error: 'not_found' });
  // IMPORTANTE: permitir abrir detalhes mesmo quando o pedido está arquivado.
  // Ações que mudam estado (vincular OPs, despachar, etc.) continuam bloqueadas em rotas específicas.

  const fgStock = await ensureStockFile('fg');
  const itemsById = new Map((fgStock.items || []).map(it => [String(it.id), it]));
  const stockMap = computeInventory(fgStock);

  const items = (order.items || []).map(it => {
    const itemId = String(it.itemId);
    const invIt = itemsById.get(itemId);
    const unit = fmtUnit(invIt?.unit || it.unit || '');
    const requested = quantizeQty(it.qty, unit);
    const available = quantizeQty(Number(stockMap.get(itemId) || 0), unit);
    const toProduce = quantizeQty(Math.max(0, requested - available), unit);
    return {
      itemId,
      code: invIt?.code || it.code || '',
      name: invIt?.name || it.name || '',
      unit,
      requested,
      available,
      toProduce,
    };
  }).sort((a,b) => String(a.code||'').localeCompare(String(b.code||''), 'pt-BR'));

  res.json({ order: { id: order.id, number: order.number, type: order.type, status: order.status, archived: !!order.archived }, items });
});

// Desarquivar pedidos/vendas (mantém histórico; não é exclusão)
app.post('/api/sales/orders/unarchive-batch', requireAuth, requirePerm("sales"), async (req, res) => {
  const { ids } = req.body || {};
  if (!Array.isArray(ids) || !ids.length) return res.status(400).json({ error: 'missing_ids' });
  const want = new Set(ids.map(x => String(x)));

  const ordersDb = await ensureSalesOrdersFile();
  const now = new Date().toISOString();
  const unarchived = [];

  for (const o of (ordersDb.orders || [])) {
    if (!want.has(String(o.id))) continue;
    if (o.archived) {
      o.archived = false;
      delete o.archivedAt;
    }
    o.updatedAt = now;
    unarchived.push({ id: o.id, number: o.number, type: o.type });
  }

  await writeSalesOrders(ordersDb);
  res.json({ ok: true, unarchived });
});

app.post('/api/sales/orders/:id/unarchive', requireAuth, requirePerm("sales"), async (req, res) => {
  const id = String(req.params.id || '');
  const ordersDb = await ensureSalesOrdersFile();
  const o = (ordersDb.orders || []).find(x => String(x.id) === id);
  if (!o) return res.status(404).json({ error: 'not_found' });
  const now = new Date().toISOString();
  o.archived = false;
  delete o.archivedAt;
  o.updatedAt = now;
  await writeSalesOrders(ordersDb);
  res.json({ order: o });
});

app.post('/api/sales/orders/:id/link-ops', requireAuth, requirePerm("sales"), async (req, res) => {
  const id = String(req.params.id || '');
  const { ops } = req.body || {};
  if (!Array.isArray(ops) || !ops.length) return res.status(400).json({ error: 'missing_ops' });

  const ordersDb = await ensureSalesOrdersFile();
  const order = (ordersDb.orders || []).find(o => String(o.id) === id);
  if (!order) return res.status(404).json({ error: 'not_found' });
  if (order.archived) return res.status(400).json({ error: 'archived' });

  order.linkedOps = Array.isArray(order.linkedOps) ? order.linkedOps : [];
  for (const op of ops) {
    const opId = String(op.id || '');
    const num = Number(op.number);
    if (!opId) continue;
    if (order.linkedOps.some(x => String(x.id) === opId)) continue;
    order.linkedOps.push({ id: opId, number: Number.isFinite(num) ? num : undefined });
  }
  order.status = 'OPS_CREATED';
  order.updatedAt = new Date().toISOString();
  await writeSalesOrders(ordersDb);
  res.json({ order });
});

app.post('/api/sales/orders/:id/dispatch', requireAuth, requirePerm("sales"), async (req, res) => {
  const id = String(req.params.id || '');
  const ordersDb = await ensureSalesOrdersFile();
  const order = (ordersDb.orders || []).find(o => String(o.id) === id);
  if (!order) return res.status(404).json({ error: 'not_found' });
  if (order.archived) return res.status(400).json({ error: 'archived' });
  if (String(order.status||'').toUpperCase() === 'DISPATCHED') return res.status(400).json({ error: 'already_dispatched' });
  if (String(order.status||'').toUpperCase() === 'CANCELLED') return res.status(400).json({ error: 'cancelled' });

  const fgStock = await ensureStockFile('fg');
  const stockMap = computeInventory(fgStock);
  const itemsById = new Map((fgStock.items || []).map(it => [String(it.id), it]));

  const pointId = String(order.pointId || '');
  const pointsDb = await ensureSalesPointsFile();
  const pt = (pointsDb.points || []).find(x => String(x.id) === pointId);
  if (!pt) return res.status(400).json({ error: 'point_not_found' });

  const user = (await ensureDB()).users.find(u => u.id === req.session.userId);
  const by = user ? { id: user.id, name: user.name } : { id: String(req.session.userId || ''), name: '' };

  // valida estoque central
  const shortages = [];
  for (const it of (order.items || [])) {
    const itemId = String(it.itemId);
    const invIt = itemsById.get(itemId);
    const unit = fmtUnit(invIt?.unit || it.unit || '');
    const q = quantizeQty(it.qty, unit);
    const available = quantizeQty(Number(stockMap.get(itemId) || 0), unit);
    if (available - q < 0) {
      shortages.push({ itemId, code: invIt?.code || it.code || '', name: invIt?.name || it.name || '', unit, available, shortage: quantizeQty(q - available, unit) });
    }
  }
  if (shortages.length) return res.status(400).json({ error: 'insufficient_fg_stock', shortages });

  // aplica baixa no estoque central + entrada no estoque do ponto
  const now = new Date().toISOString();
  const movesDb = await ensureSalesPointMovesFile();
  for (const it of (order.items || [])) {
    const itemId = String(it.itemId);
    const invIt = itemsById.get(itemId);
    const unit = fmtUnit(invIt?.unit || it.unit || '');
    const q = quantizeQty(it.qty, unit);
    const beforeQty = quantizeQty(Number(stockMap.get(itemId) || 0), unit);
    const afterQty = quantizeQty(beforeQty - q, unit);
    stockMap.set(itemId, afterQty);
    const pvCode = formatSalesOrderCode('PV', order.number);

    const fgMvId = newId('mv_');
    fgStock.movements.push({
      id: fgMvId,
      type: 'out',
      itemId,
      qty: q,
      reason: `Despacho ${pvCode} -> ${pt.code}`,
      refType: 'PV',
      refId: order.id,
      refNumber: Number(order.number),
      refCode: pvCode,
      at: now,
      by,
      beforeQty,
      afterQty,
      delta: quantizeQty(afterQty - beforeQty, unit),
    });

    const ptMvId = newId('ptmv_');
    movesDb.moves.push({
      id: ptMvId,
      pointId,
      itemId,
      unit,
      delta: q,
      qty: q,
      type: 'in',
      reason: `Recebido do ${pvCode} (Despacho)`,
      refType: 'PV',
      refId: order.id,
      refNumber: Number(order.number),
      refCode: pvCode,
      at: now,
    });

    order.dispatchMovementIds = Array.isArray(order.dispatchMovementIds) ? order.dispatchMovementIds : [];
    order.dispatchPointMoveIds = Array.isArray(order.dispatchPointMoveIds) ? order.dispatchPointMoveIds : [];
    order.dispatchMovementIds.push(fgMvId);
    order.dispatchPointMoveIds.push(ptMvId);
  }

  await writeStock('fg', fgStock);
  await writeSalesPointMoves(movesDb);

  order.status = 'DISPATCHED';
  order.dispatchedAt = now;
  order.updatedAt = now;
  await writeSalesOrders(ordersDb);
  res.json({ order });
});


app.delete('/api/sales/orders/:id', requireAuth, requirePerm("sales"), requireReauthOnce, async (req, res) => {
  const id = String(req.params.id || '');
  const ordersDb = await ensureSalesOrdersFile();
  const idx = (ordersDb.orders || []).findIndex(o => String(o.id) === id);
  if (idx < 0) return res.status(404).json({ error: 'not_found' });
  const o = ordersDb.orders[idx];
  ordersDb.orders.splice(idx, 1);
  await writeSalesOrders(ordersDb);
  res.json({ ok: true, deleted: { id: o.id, number: o.number } });
});



// -------------------- Custos (Produto Final) --------------------
app.get("/api/costing/recipe/:id", requireAuth, requirePerm("costs"), async (req, res) => {
  const recipeId = String(req.params.id || "").trim();
  const qtyIn = Number(req.query.qty || 1);
  const qty = (Number.isFinite(qtyIn) && qtyIn > 0) ? qtyIn : 1;

  const db = await ensureDB();
  const recipe = (db.mrp.recipes || []).find((r) => String(r.id) === recipeId);
  if (!recipe) return res.status(404).json({ error: "recipe_not_found" });

  // Ensure PF output item exists (used to fetch salePrice + code)
  const fgStock = await ensureStockFile("fg");
  ensureRecipeOutputItem(fgStock, recipe);
  await writeStock("fg", fgStock);

  const outId = String(recipe.productId || recipe.outputItemId || "");
  const outItem = (fgStock.items || []).find((i) => String(i.id) === outId) || null;

  const rawStock = await ensureStockFile("raw");
  const rawById = new Map((rawStock.items || []).map((i) => [String(i.id), i]));
  const yieldQty = parseNumBR(recipe.yieldQty, 1) || 1;
  const yieldUnit = String(recipe.yieldUnit || "un");

  const factor = qty / (yieldQty || 1);

  const lines = (recipe.bom || []).map((line) => {
    const it = rawById.get(String(line.itemId));
    const unit = String(it?.unit || "");
    const lossPercent = parseNumBR(it?.lossPercent ?? 0, 0) || 0;

    const bomQty = parseNumBR(line?.qty ?? line?.qtd ?? line?.quantidade ?? 0, 0) || 0;
    const requiredNet = bomQty * factor;
    const requiredGross = requiredNet * (1 + (lossPercent / 100));
    const required = quantizeQty(requiredGross, unit);

    const unitCost = parseNumBR(it?.cost ?? it?.unitCost ?? it?.custo ?? 0, 0) || 0;
    const lineCost = Number((required * unitCost).toFixed(6));

    return {
      itemId: String(line.itemId),
      itemCode: it?.code || "",
      itemName: it?.name || "—",
      unit,
      lossPercent: Number(lossPercent.toFixed(6)),
      requiredNet: Number(requiredNet.toFixed(6)),
      required: Number(required.toFixed(6)),
      unitCost: Number(unitCost.toFixed(6)),
      lineCost,
    };
  });

  const totalCost = Number(lines.reduce((s, l) => s + (Number(l.lineCost) || 0), 0).toFixed(6));
  const costPerUnit = Number((totalCost / (qty || 1)).toFixed(6));

  const salePrice = parseNumBR(outItem?.salePrice ?? 0, 0) || 0;
  const margin = Number((salePrice - costPerUnit).toFixed(6));
  const marginPct = salePrice > 0 ? Number(((margin / salePrice) * 100).toFixed(6)) : 0;

  res.json({
    recipeId,
    qty: Number(qty.toFixed(6)),
    yieldQty,
    yieldUnit,
    output: {
      id: outItem?.id || "",
      code: outItem?.code || "",
      name: outItem?.name || recipe.name || "",
      unit: outItem?.unit || yieldUnit || "un",
      salePrice: Number(salePrice.toFixed(6)),
    },
    totalCost,
    costPerUnit,
    margin,
    marginPct,
    lines,
  });
});


app.listen(PORT, async () => {
  await ensureDB();
  console.log(`✅ dietON MRP rodando em http://localhost:${PORT} (BUILD ${BUILD})`);
  console.log(`🔐 Login padrão: Felipe  |  Senha padrão: Mestre`);
});

// Fallback: serve SPA
app.get("*", (req, res) => {
  res.sendFile(path.join(__dirname, "public", "index.html"));
});