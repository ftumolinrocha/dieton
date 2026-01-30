
function unescapeHtmlEntities(input) {
  // Some legacy imports/cadastros persisted already-escaped text (e.g. "Frango &amp; Arroz").
  // We normalize it back to plain text before escaping again for innerHTML.
  let s = String(input ?? "");
  if (!s.includes("&")) return s;
  // Handle double-escaped sequences like "&amp;amp;" by decoding a few passes.
  for (let i = 0; i < 3; i++) {
    const prev = s;
    s = s
      .replaceAll("&amp;", "&")
      .replaceAll("&lt;", "<")
      .replaceAll("&gt;", ">")
      .replaceAll("&quot;", '"')
      .replaceAll("&#039;", "'")
      .replaceAll("&#39;", "'");
    if (s === prev) break;
  }
  return s;
}

function escapeHtml(value) {
  const s = unescapeHtmlEntities(value);
  return s
    .replaceAll("&", "&amp;")
    .replaceAll("<", "&lt;")
    .replaceAll(">", "&gt;")
    .replaceAll('"', "&quot;")
    .replaceAll("'", "&#039;");
}

function unitOptionsHtml(selected) {
  const fallback = [
    { v: "kg", l: "kg (kilograma)" },
    { v: "un", l: "un (unidade)" },
    { v: "ml", l: "ml (mililitro)" },
    { v: "l",  l: "l (litro)" }
  ];
  const list = (state?.units && state.units.length) ? state.units : fallback;
  const hasSelected = !!selected && list.some((u) => u.v === selected);
  const prepend = (!hasSelected && selected)
    ? `<option value="${escapeHtml(selected)}" selected>${escapeHtml(selected)}</option>`
    : "";
  const opts = list.map(u => `<option value="${escapeHtml(u.v)}" ${u.v===selected ? "selected" : ""}>${escapeHtml(u.l || u.v)}</option>`).join("");
  // Only real units in the dropdown (custom/manage are accessed in the Unidades screen).
  return prepend + opts;
}





function normalizeItemCode(input, mode){
  const isFg = mode === "fg";
  const prefix = isFg ? "PF" : "MP";
  let c = String(input ?? "").trim().toUpperCase();
  if (!c) return "";

  // Accept just numbers (e.g. 2 -> MP002)
  if (/^\d+$/.test(c)) {
    const n = parseInt(c, 10);
    if (!Number.isFinite(n) || n <= 0) return "";
    return prefix + String(n).padStart(3, "0");
  }

  // Accept MP2, MP002, PF10, etc.
  if (c.startsWith(prefix)) {
    const tail = c.slice(prefix.length).trim();
    if (/^\d+$/.test(tail)) {
      const n = parseInt(tail, 10);
      if (!Number.isFinite(n) || n <= 0) return "";
      return prefix + String(n).padStart(3, "0");
    }
  }

  // If user typed the wrong prefix, keep as-is (server will validate)
  return c;
}

function compareItemCodes(a, b){
  const ac = String(a || "").trim().toUpperCase();
  const bc = String(b || "").trim().toUpperCase();
  const ap = ac.slice(0,2);
  const bp = bc.slice(0,2);
  if ((ap === "MP" || ap === "PF") && ap === bp){
    const an = parseInt(ac.slice(2), 10);
    const bn = parseInt(bc.slice(2), 10);
    const aok = Number.isFinite(an);
    const bok = Number.isFinite(bn);
    if (aok && bok && an !== bn) return an - bn;
  }
  return ac.localeCompare(bc, 'pt-BR');
}

function computeNextCode(items, mode){
  const prefix = (mode === "fg") ? "PF" : "MP";
  const nums = (items || []).map(it => {
    const c = normalizeItemCode(it.code, mode);
    if (!c.startsWith(prefix)) return NaN;
    const n = parseInt(c.slice(prefix.length), 10);
    return Number.isFinite(n) ? n : NaN;
  }).filter(n => Number.isFinite(n) && n > 0);

  const set = new Set(nums);
  const max = nums.length ? Math.max(...nums) : 0;
  let cand = 1;
  while (cand <= max + 1){
    if (!set.has(cand)) break;
    cand += 1;
  }
  return prefix + String(cand).padStart(3, "0");
}

const $ = (sel) => document.querySelector(sel);
const $$ = (sel) => Array.from(document.querySelectorAll(sel));

// Safety: if a previous buggy build submitted forms via GET, it may have left
// sensitive fields (e.g. password) in the URL querystring. Clean it up.
try {
  const sp = new URLSearchParams(location.search || "");
  if (sp.has("password") || sp.has("perm_inventory") || sp.has("perm_admin")) {
    history.replaceState(null, "", location.pathname + (location.hash || ""));
  }
} catch (_) {}

// Helper to switch the main tabs programmatically (used after generating OP/OC).
// Keeps compatibility with code paths that reference `showTab`.
function showTab(tabName){
  try {
    const btn = document.querySelector(`.tab[data-tab="${String(tabName)}"]`);
    if (btn) btn.click();
  } catch (e) {
    // no-op
  }
}

// BUILD (UI): começa com fallback e depois sincroniza com o server (/api/build)
// Versão final
const BUILD_INFO = { version: "1.0", builtAt: "2026-01-30" };

function setBuildBadge(){
  const bb = document.querySelector('#buildBadge');
  if (!bb) return;
  const d = new Date(BUILD_INFO.builtAt);
  const when = isNaN(d.getTime()) ? BUILD_INFO.builtAt : d.toLocaleString('pt-BR');
  bb.textContent = `BUILD ${BUILD_INFO.version} • ${when}`;
}

async function syncBuildFromServer(){
  try {
    const info = await api('/api/build');
    if (info?.build) BUILD_INFO.version = String(info.build);
    if (info?.at) BUILD_INFO.builtAt = String(info.at);
  } catch {
    // mantém fallback
  }
  setBuildBadge();
}

// ---------- Permissões ----------
function getPerms(){
  const p = state && state.me && state.me.permissions;
  return (p && typeof p === "object" && !Array.isArray(p)) ? p : {};
}
function canPerm(key){
  const p = getPerms();
  if (p.admin === true) return true;
  if (p[key] === true) return true;
  // Backward-compat: older builds used a single 'mrp' flag for Receitas/OP/OC/Custos
  if ((key === 'recipes' || key === 'op' || key === 'oc' || key === 'costs') && typeof p[key] !== 'boolean' && p.mrp === true) return true;
  return false;
}
function setElDisplay(el, showIt){
  if (!el) return;
  el.style.display = showIt ? "" : "none";
}
function applyPermissions(){
  // If not logged in, show all tabs (login view hides app anyway)
  if (!state.me) return;

  const tabPermMap = [
    { tab: "estoque", perm: "inventory" },
    { tab: "mrp", perm: "recipes" },
    { tab: "custos", perm: "costs" },
    { tab: "compras", perm: "oc" },
    { tab: "ops", perm: "op" },
    { tab: "vendas", perm: "sales" },
    { tab: "usuarios", perm: "admin" },
  ];

  for (const t of tabPermMap) {
    const allowed = canPerm(t.perm);
    const btn = document.querySelector(`.tab[data-tab="${t.tab}"]`);
    const panel = document.querySelector(`#tab-${t.tab}`);
    setElDisplay(btn, allowed);
    if (panel && !allowed) panel.classList.add("hidden");
  }

  // Disable sensitive buttons by permission
  const canXlsx = canPerm("canImportExport");
  ["#btnImport", "#btnExport", "#btnGeneralImport", "#btnGeneralExport", "#btnInvHistClear", "#btnInvHistExport"].forEach((sel) => {
    const el = document.querySelector(sel);
    if (el) el.disabled = !canXlsx;
  });

  const canReset = canPerm("canReset");
  const btnReset = document.querySelector("#btnResetTotal");
  if (btnReset) btnReset.disabled = !canReset;

  // If current active tab is not allowed, jump to first allowed
  const activeBtn = document.querySelector(".tab.active");
  const activeTab = activeBtn ? String(activeBtn.dataset.tab || "") : "";
  const activeAllowed = !activeTab ? true : canPerm((tabPermMap.find(x => x.tab === activeTab) || {}).perm || "inventory");
  if (!activeAllowed) {
    const firstAllowed = tabPermMap.find(x => canPerm(x.perm));
    if (firstAllowed) {
      try { showTab(firstAllowed.tab); } catch (_) {}
    }
  }
}


// INIT_BUILD_BADGE
(function initBuildBadge(){
  setBuildBadge();
  // tenta sincronizar com server sem bloquear UI
  syncBuildFromServer();
})();


const state = {
  selectedItemId: null,
  cadastroQuery: "",

  me: null,
  users: [],
  usersQuery: "",
  selectedUserId: null,
  items: [],
  movements: [],
  rawItems: [],
  fgItems: [],
  recipes: [],
  selectedPFId: null,
  ops: [],
  purchaseOrders: [],
  selectedRecipeId: null,
  stockMode: "raw",

  // Pedidos de Venda
  salesPoints: [],
  salesPointSelectedIds: [],
  salesOrders: [],
  salesMode: "points", // points | orders | quick
  salesOrdersShowArchived: false,
  salesOrderSelectedIds: [],
  quickSalesShowArchived: false,
  quickSelectedIds: [],

  units: [],

  // Estoque mínimo (MP/PF): status para alerta e botões de geração (OC/OP)
  minStockStatus: null,
};

// Status exibidos (PT-BR) — manter valores internos em EN, mas nunca mostrar para o usuário.
const OC_STATUS_PT = {
  OPEN: "Aberta",
  PARTIAL: "Parcial",
  RECEIVED: "Recebida",
  CANCELLED: "Cancelada",
  CLOSED: "Fechada",
  HOLD: "Emitida",
};

const OP_STATUS_PT = {
  HOLD: "Emitida",
  // Novo fluxo de OP:
  // - ISSUED: gerada (emitida)
  // - IN_PRODUCTION: em produção
  // - CLOSED: encerrada (após executar / dar baixa)
  ISSUED: "Emitida",
  IN_PRODUCTION: "Em produção",
  CLOSED: "Encerrada",
  CANCELLED: "Cancelada",

  // Compat (BD antiga):
  READY: "Emitida",
  EXECUTED: "Encerrada",
};

function showBackendOffline(err){
  try {
    const isLogin = !state.me;
    openModal({
      title: "Servidor local não encontrado",
      subtitle: "O app está aberto, mas o servidor (API) não está respondendo.",
      submitText: "Ok",
      cardClass: "wide",
      bodyHtml: `
        <div class="muted" style="line-height:1.6">
          <b>O que aconteceu:</b> não foi possível conectar em <code>/api</code> (erro: <code>${escapeHtml(err?.message || "failed to fetch")}</code>).
          <br/><br/>
          <b>Como resolver:</b>
          <ol style="margin:8px 0 0 18px">
            <li>Abra um terminal na pasta do projeto</li>
            <li>Rode <code>npm install</code> (só na primeira vez)</li>
            <li>Rode <code>npm start</code></li>
            <li>Abra <code>http://localhost:4000</code> no navegador</li>
          </ol>
          <br/>
          Se o terminal mostrar algum erro, copie e me envie aqui que eu arrumo no pack.
        </div>
      `,
      onSubmit: () => true,
    });
  } catch(_) {}
}

async function api(path, options = {}) {
  try {
    const res = await fetch(path, {
      credentials: "same-origin",
      headers: { "Content-Type": "application/json" },
      ...options,
    });
    const ct = res.headers.get("content-type") || "";
    const data = ct.includes("application/json") ? await res.json() : await res.text();
    if (!res.ok) {
      const err = new Error("API error");
      err.status = res.status;
      err.data = data;
      throw err;
    }
    return data;
  } catch (err) {
    // Network / connection errors
    if (String(err?.message || "").toLowerCase().includes("fetch")) {
      showBackendOffline(err);
    }
    throw err;
  }
}

async function apiUpload(path, formData, options = {}) {
  try {
    const res = await fetch(path, { method: "POST", body: formData, credentials: "same-origin", ...options });
    const ct = res.headers.get("content-type") || "";
    const data = ct.includes("application/json") ? await res.json() : await res.text();
    if (!res.ok) {
      const err = new Error("API upload error");
      err.status = res.status;
      err.data = data;
      throw err;
    }
    return data;
  } catch (err) {
    if (String(err?.message||"").toLowerCase().includes("failed to fetch")) showBackendOffline(err);
    throw err;
  }
}

// ---------- Reauth (Import/Export) ----------
// Para ações sensíveis (importar/exportar XLSX), o servidor exige uma confirmação
// de usuário + senha imediatamente antes da ação.
// Reauth isolado do modal principal. Isso evita que o "Limpar histórico" afete
// os fluxos de Importar/Exportar (que também usam o modal principal).
function requestReauth({ reason = "" } = {}) {
  return new Promise((resolve) => {
    const dlg = document.querySelector('#reauthModal');
    const form = document.querySelector('#reauthForm');
    const subtitle = document.querySelector('#reauthSubtitle');
    const u = document.querySelector('#reauth2User');
    const p = document.querySelector('#reauth2Pass');
    const errBox = document.querySelector('#reauth2Error');
    const btnClose = document.querySelector('#reauthCloseBtn');
    const btnCancel = document.querySelector('#reauthCancelBtn');
    const btnConfirm = document.querySelector('#reauthConfirmBtn');

    if (!dlg || !form || !u || !p) {
      // fallback (não deveria acontecer)
      resolve(false);
      return;
    }

    let settled = false;
    let inflight = false;

    const cleanup = () => {
      try { form.removeEventListener('submit', onSubmit); } catch (_) {}
      try { dlg.removeEventListener('close', onClose); } catch (_) {}
      try { btnClose && btnClose.removeEventListener('click', onCancel); } catch (_) {}
      try { btnCancel && btnCancel.removeEventListener('click', onCancel); } catch (_) {}
    };

    const done = (ok) => {
      if (settled) return;
      settled = true;
      cleanup();
      resolve(!!ok);
    };

    const onClose = () => done(false);
    const onCancel = () => {
      try { dlg.close(); } catch (_) {}
      done(false);
    };

    const onSubmit = async (ev) => {
      ev.preventDefault();
      if (inflight) return;
      inflight = true;
      if (errBox) errBox.style.display = 'none';
      if (btnConfirm) btnConfirm.disabled = true;

      const userId = String(u.value || '').trim();
      const password = String(p.value || '');

      try {
        await api('/api/reauth', { method: 'POST', body: JSON.stringify({ userId, password }) });
        try { dlg.close(); } catch (_) {}
        done(true);
      } catch (e) {
        if (errBox) errBox.style.display = 'block';
        inflight = false;
        if (btnConfirm) btnConfirm.disabled = false;
        try { p.focus(); } catch (_) {}
      }
    };

    // Prep UI
    if (subtitle) subtitle.textContent = reason || 'Para continuar, confirme suas credenciais.';
    // /api/me retorna { user: { id, name } }. Mantém o usuário pré-preenchido no reauth.
    const defaultUser = (state && state.me && state.me.id) ? state.me.id : '';
    u.value = defaultUser;
    p.value = '';
    if (errBox) errBox.style.display = 'none';
    if (btnConfirm) btnConfirm.disabled = false;

    // Wire events (sem acumular listeners entre aberturas)
    form.addEventListener('submit', onSubmit);
    dlg.addEventListener('close', onClose, { once: true });
    if (btnClose) btnClose.addEventListener('click', onCancel);
    if (btnCancel) btnCancel.addEventListener('click', onCancel);

    try {
      if (!dlg.open) dlg.showModal();
      setTimeout(() => { try { p.focus(); } catch (_) {} }, 0);
    } catch (e) {
      // Se o browser bloquear showModal por algum motivo, não quebra o app
      done(false);
    }
  });
}


// Abre um modal único para IMPORTAR XLSX exigindo usuário/senha + arquivo no mesmo fluxo.
// Isso evita problemas de navegadores que bloqueiam abrir o seletor de arquivo depois de awaits.
function openImportXlsxModal({
  title = "Importar XLSX",
  reason = "",
  uploadUrl = "",
  onAfter = async () => {},
} = {}) {
  const defaultUser = (state.me && state.me.id) ? state.me.id : "";

  openModal({
    title,
    subtitle: reason || "Para importar, confirme suas credenciais e selecione o arquivo.",
    submitText: "Importar",
    bodyHtml: `
      <div class="grid2">
        <div class="field">
          <label>Usuário</label>
          <input id="impUser" name="userId" class="input" placeholder="Ex.: Felipe" value="${escapeHtml(defaultUser)}" required />
        </div>
        <div class="field">
          <label>Senha</label>
          <input id="impPass" name="password" type="password" class="input" placeholder="Digite sua senha" required />
        </div>
      </div>

      <div class="field" style="margin-top:10px">
        <label>Arquivo (.xlsx)</label>
        <input id="impFile" name="file" type="file" accept=".xlsx,application/vnd.openxmlformats-officedocument.spreadsheetml.sheet" class="input" required />
        <div class="hint">Selecione um arquivo XLSX no modelo correto.</div>
      </div>

      <div id="impError" class="hint" style="color:var(--danger); margin-top:8px; display:none;"></div>
    `,
    onOpen: () => {
      const u = document.querySelector('#impUser');
      const p = document.querySelector('#impPass');
      const f = document.querySelector('#impFile');
      if (u && !u.value) u.focus();
      else if (p) p.focus();
      else if (f) f.focus();
    },
    onSubmit: async (fd) => {
      const userId = String(fd.get('userId') || '').trim();
      const password = String(fd.get('password') || '');
      const file = fd.get('file');
      const errBox = document.querySelector('#impError');
      if (errBox) { errBox.style.display = 'none'; errBox.textContent = ''; }

      if (!file || typeof file !== 'object') {
        if (errBox) { errBox.textContent = 'Selecione um arquivo .xlsx.'; errBox.style.display = 'block'; }
        return false;
      }
      if (!uploadUrl) {
        if (errBox) { errBox.textContent = 'Rota de importação não configurada.'; errBox.style.display = 'block'; }
        return false;
      }

      // 1) Reauth
      try {
        await api('/api/reauth', { method: 'POST', body: JSON.stringify({ userId, password }) });
      } catch (e) {
        if (errBox) { errBox.textContent = 'Usuário ou senha inválidos.'; errBox.style.display = 'block'; }
        return false;
      }

      // 2) Upload
      try {
        const up = new FormData();
        up.append('file', file);
        await apiUpload(uploadUrl, up);
        await onAfter();
        alert('Importação XLSX concluída.');
        return true;
      } catch (e) {
        console.error(e);
        if (errBox) {
          const msg = (e && e.data && (e.data.error || e.data.message)) ? String(e.data.error || e.data.message) : '';
          errBox.textContent = msg ? ('Falha ao importar: ' + msg) : 'Falha ao importar XLSX. Verifique o arquivo e o formato das colunas.';
          errBox.style.display = 'block';
        }
        return false;
      }
    },
  });
}

function downloadUrl(url) {
  // server envia Content-Disposition.
  // IMPORTANTE: depois de awaits, navegadores podem bloquear downloads via a.click()
  // (perde-se o "user activation"). Para evitar isso, usamos um iframe.
  try {
    const ifr = document.createElement("iframe");
    ifr.style.display = "none";
    ifr.src = url;
    document.body.appendChild(ifr);
    setTimeout(() => {
      try { ifr.remove(); } catch (_) {}
    }, 60 * 1000);
  } catch (e) {
    // fallback
    const a = document.createElement("a");
    a.href = url;
    a.style.display = "none";
    document.body.appendChild(a);
    a.click();
    setTimeout(() => a.remove(), 300);
  }
}

// Baixar XLSX após reauth.
// Importante: não abrir nova aba. Alguns navegadores bloqueiam o download
// quando a ação ocorre após awaits; por isso usamos iframe (downloadUrl).
async function downloadAfterReauth(url, reason){
  const ok = await requestReauth({ reason: reason || '' });
  if (!ok) return;
  // cache-buster para evitar reuso indevido
  const sep = url.includes('?') ? '&' : '?';
  downloadUrl(`${url}${sep}t=${Date.now()}`);
}

// Executar DELETE após reauth (ex.: excluir BOM)
async function deleteAfterReauth(url, reason){
  const ok = await requestReauth({ reason: reason || '' });
  if (!ok) return { ok: false, canceled: true };
  try {
    const r = await api(url, { method: 'DELETE' });
    return r;
  } catch (e) {
    console.error(e);
    throw e;
  }
}

async function pingBackend(){
  try{
    const r = await api("/api/health");
    return !!r?.ok;
  }catch{
    return false;
  }
}


function filteredItems() {
  return state.items.filter((i) => (i.type || "raw") === state.stockMode);
}

function filteredMovements() {
  const allowed = new Set(filteredItems().map((i) => i.id));
  return state.movements.filter((m) => allowed.has(m.itemId));
}


function fmt(n) {
  const x = Number(n);
  if (!Number.isFinite(x)) return "0";
  // show up to 3 decimals, trim zeros
  return x.toLocaleString("pt-BR", { maximumFractionDigits: 3 });
}

function fmtUnit(u) {
  const s = String(u ?? "").trim();
  if (!s) return "";
  return s.toLowerCase();
}

function pad6(n) {
  const x = Number(n);
  if (!Number.isFinite(x) || x <= 0) return "";
  return String(Math.trunc(x)).padStart(6, "0");
}

function salesOrderSeries(o){
  const s = String(o?.series || '').trim().toUpperCase();
  if (s === 'PV' || s === 'PVR') return s;
  const t = String(o?.type || '').trim().toUpperCase();
  return t === 'QUICK' ? 'PVR' : 'PV';
}

function fmtSalesOrderCode(o){
  const ser = salesOrderSeries(o);
  const num = Number(o?.number || 0);
  const p = pad6(num);
  return p ? `${ser}${p}` : `${ser}??????`;
}

function parsePvManualInput(raw){
  const s = String(raw ?? '').trim();
  if (!s) return null;
  const m = s.match(/^(?:PV)?\s*(\d{1,6})$/i);
  if (!m) return null;
  const n = Number(m[1]);
  if (!Number.isFinite(n) || n <= 0) return null;
  return Math.trunc(n);
}


function parseBRNumber(v) {
  // Aceita "1.234,56" ou "1234,56" ou "1234.56"
  if (typeof v === "number") return Number.isFinite(v) ? v : 0;
  let s = String(v ?? "").trim();
  if (!s) return 0;
  s = s.replace(/\s/g, "");
  // Remove símbolos (R$, etc)
  s = s.replace(/[^\d,.\-]/g, "");
  if (s.includes(",") && s.includes(".")) {
    // assume . milhares e , decimal
    s = s.replace(/\./g, "").replace(",", ".");
  } else if (s.includes(",") && !s.includes(".")) {
    s = s.replace(",", ".");
  }
  const n = Number(s);
  return Number.isFinite(n) ? n : 0;
}

// Compat: algumas telas chamam fmtNum(). Mantemos como alias de fmt().
function fmtNum(n){
  const x = Number(n);
  if (!Number.isFinite(x)) return "0";
  return x.toLocaleString("pt-BR", { maximumFractionDigits: 3 });
}

// Alguns trechos usam fmtQty() para exibir quantidades em relatórios/necessidades.
// Mantemos como alias de fmtNum() para evitar ReferenceError.
function fmtQty(n){
  return fmtNum(n);
}

function fmtMoney(n) {
  const x = Number(n);
  if (!Number.isFinite(x)) return "—";
  return x.toLocaleString("pt-BR", { style: "currency", currency: "BRL" });
}

function fmtDate(iso) {
  try {
    const d = new Date(iso);
    return d.toLocaleString("pt-BR");
  } catch {
    return iso;
  }
}

function show(el, on = true) {
  el.classList.toggle("hidden", !on);
}

// Re-render only what the user is currently seeing (avoid jumping tabs)
function renderCurrentView() {
  // Main tabs
  const tabEstoque = $("#tab-estoque");
  const tabMrp = $("#tab-mrp");
  const tabOps = $("#tab-ops");

  const isVisible = (el) => el && !el.classList.contains("hidden");

  // MRP
  if (isVisible(tabMrp)) {
    Promise.resolve(loadMRP())
      .catch(() => null)
      .finally(() => {
        if (typeof renderPFBomList === "function") renderPFBomList();
        if (typeof renderSimulateBox === "function") renderSimulateBox();
      });
    return;
  }

  // OPs
  if (isVisible(tabOps)) {
    if (typeof renderOps === "function") renderOps();
    return;
  }

  // Estoque (default)
  if (isVisible(tabEstoque)) {
    const secGeneral = $("#estoqueCadastroGeneral");
    const secMp = $("#estoqueCadastroMP");
    const secPf = $("#estoqueCadastroPF");
    const secAdjMp = $("#estoqueAdjustMP");
    const secAdjPf = $("#estoqueAdjustPF");
    const secHist = $("#estoqueHistory");

    if (isVisible(secGeneral)) {
      if (typeof renderGeneralTableInline === "function") renderGeneralTableInline();
      return;
    }
    if (isVisible(secHist)) {
      if (typeof renderInvHistTable === "function") renderInvHistTable();
      return;
    }
    if (isVisible(secAdjMp) || isVisible(secAdjPf)) {
      // Ajuste usa a mesma renderização do modo atual
      if (typeof renderAjusteInventario === "function") renderAjusteInventario();
      return;
    }
    if (isVisible(secMp) || isVisible(secPf)) {
      if (typeof renderCadastro === "function") renderCadastro();
      return;
    }
    // fallback
    if (typeof renderCadastro === "function") renderCadastro();
  }
}

// ---------- Login / Session ----------
async function refreshMe() {
  const me = await api("/api/me");
  state.me = me.authenticated ? me.user : null;
  if (state.me && (!state.me.permissions || typeof state.me.permissions !== "object")) state.me.permissions = {};

  show($("#topbar"), true);
  show($("#topbar"), true);
  show($("#loginView"), !state.me);
  show($("#appView"), !!state.me);
show($("#btnLogout"), !!state.me);
  show($("#userBadge"), !!state.me);
  if (state.me) $("#userBadge").textContent = `👤 ${state.me.name} (${state.me.id})`;
  // aplica permissões (tabs e botões)
  try { applyPermissions(); } catch (_) {}
  // mantém o badge consistente mesmo após login/logout
  setBuildBadge();
}

$("#loginForm").addEventListener("submit", async (e) => {
  e.preventDefault();
  show($("#loginError"), false);

  const form = new FormData(e.currentTarget);
  const userId = form.get("userId");
  const password = form.get("password");

  try {
    await api("/api/login", { method: "POST", body: JSON.stringify({ userId, password }) });
    await refreshMe();
    await loadAll();
  } catch (err) {
    show($("#loginError"), true);
    $("#loginError").textContent = "Usuário ou senha inválidos.";
  }
});

$("#btnLogout").addEventListener("click", async () => {
  await api("/api/logout", { method: "POST", body: JSON.stringify({}) });
  state.selectedRecipeId = null;
  await refreshMe();
});

// ---------- Usuários (Admin) ----------
function permSummary(p){
  const pp = (p && typeof p === "object" && !Array.isArray(p)) ? p : {};
  const parts = [];
  if (pp.admin) parts.push("ADMIN");
  if (pp.inventory) parts.push("Estoque");
  // Legacy: older builds used a single 'mrp' flag
  const hasNewMrpFlags = (typeof pp.recipes === 'boolean') || (typeof pp.op === 'boolean') || (typeof pp.oc === 'boolean') || (typeof pp.costs === 'boolean');
  if (!hasNewMrpFlags && pp.mrp) parts.push("MRP");
  if (pp.recipes) parts.push("Receitas");
  if (pp.op) parts.push("OP");
  if (pp.oc) parts.push("OC");
  if (pp.costs) parts.push("Custos");
  if (pp.sales) parts.push("Vendas");
  if (pp.canReset) parts.push("Reset");
  if (pp.canImportExport) parts.push("XLSX");
  return parts.length ? parts.join(" • " ) : "—";
}

function usersEls(){
  return {
    reload: $("#btnUsersReload"),
    form: $("#userForm"),
    formTitle: $("#userFormTitle"),
    saveBtn: $("#btnUserSave"),
    cancelBtn: $("#btnUserCancelEdit"),
    delBtn: $("#btnUserDelete"),
    err: $("#usersError"),
    search: $("#usersSearch"),
    tbody: $("#usersTbody"),
  };
}

function setUsersError(msg){
  const { err } = usersEls();
  if (!err) return;
  if (!msg) { err.classList.add("hidden"); err.textContent = ""; return; }
  err.classList.remove("hidden");
  err.textContent = String(msg);
}

function permsFromForm(fd){
  const p = {
    inventory: !!fd.get("perm_inventory"),
    recipes: !!fd.get("perm_recipes"),
    op: !!fd.get("perm_op"),
    oc: !!fd.get("perm_oc"),
    costs: !!fd.get("perm_costs"),
    sales: !!fd.get("perm_sales"),
    canReset: !!fd.get("perm_canReset"),
    canImportExport: !!fd.get("perm_canImportExport"),
    admin: !!fd.get("perm_admin"),
  };
  if (p.admin) {
    p.inventory = true; p.recipes = true; p.op = true; p.oc = true; p.costs = true; p.sales = true; p.canReset = true; p.canImportExport = true;
  }
  return p;
}

function fillFormFromUser(u){
  const { form, formTitle, saveBtn, cancelBtn, delBtn } = usersEls();
  if (!form) return;
  const idInput = form.querySelector('input[name="id"]');
  const nameInput = form.querySelector('input[name="name"]');
  const passInput = form.querySelector('input[name="password"]');

  const p = u && u.permissions ? u.permissions : {};
  if (idInput) { idInput.value = u ? u.id : ""; idInput.disabled = !!u; }
  if (nameInput) nameInput.value = u ? (u.name || u.id) : "";
  if (passInput) passInput.value = "";

  // checkboxes
  const setCb = (n, v) => {
    const cb = form.querySelector(`input[name="${n}"]`);
    if (cb) cb.checked = !!v;
  };
  setCb("perm_inventory", p.inventory);
  setCb("perm_recipes", p.recipes ?? (p.mrp === true));
  setCb("perm_op", p.op ?? (p.mrp === true));
  setCb("perm_oc", p.oc ?? (p.mrp === true));
  setCb("perm_costs", p.costs ?? (p.mrp === true));
  setCb("perm_sales", p.sales);
  setCb("perm_canReset", p.canReset);
  setCb("perm_canImportExport", p.canImportExport);
  setCb("perm_admin", p.admin);

  if (formTitle) formTitle.textContent = u ? `Editar usuário: ${u.id}` : "+ Novo usuário";
  if (saveBtn) saveBtn.textContent = u ? "Salvar alterações" : "Criar usuário";
  if (cancelBtn) cancelBtn.style.display = u ? "" : "none";
  if (delBtn) delBtn.style.display = u ? "" : "none";
  setUsersError("");
}

function selectedUser(){
  return (state.users || []).find(u => u && u.id === state.selectedUserId) || null;
}

function formatDateMaybe(s){
  if (!s) return "—";
  const d = new Date(String(s));
  return isNaN(d.getTime()) ? String(s) : d.toLocaleString("pt-BR");
}

function renderUsers(){
  const { tbody, search } = usersEls();
  if (!tbody) return;

  const q = String(state.usersQuery || "").trim().toLowerCase();
  const list = Array.isArray(state.users) ? state.users : [];
  const filtered = q ? list.filter(u => String(u.id||"").toLowerCase().includes(q) || String(u.name||"").toLowerCase().includes(q)) : list;

  tbody.innerHTML = "";
  if (!filtered.length) {
    const tr = document.createElement("tr");
    tr.innerHTML = '<td colspan="4" class="muted">Nenhum usuário.</td>';
    tbody.appendChild(tr);
    return;
  }

  for (const u of filtered) {
    const tr = document.createElement("tr");
    if (u.id === state.selectedUserId) tr.classList.add("row-selected");
    tr.style.cursor = "pointer";
    tr.innerHTML = `
      <td><b>${escapeHtml(u.id)}</b></td>
      <td>${escapeHtml(u.name || u.id)}</td>
      <td class="small muted">${escapeHtml(permSummary(u.permissions))}</td>
      <td class="small muted">${escapeHtml(formatDateMaybe(u.updatedAt || u.createdAt))}</td>
    `;
    tr.addEventListener("click", () => {
      state.selectedUserId = u.id;
      fillFormFromUser(u);
      renderUsers();
    });
    tbody.appendChild(tr);
  }

  if (search && !search._wiredUsersSearch) {
    search._wiredUsersSearch = true;
  }
}

async function loadUsers(){
  if (!canPerm("admin")) return;
  try {
    const r = await api("/api/users");
    state.users = Array.isArray(r.users) ? r.users : [];
    renderUsers();
  } catch (e) {
    console.warn("loadUsers failed", e);
    if (e && e.status === 403) setUsersError("Sem permissão para gerenciar usuários (ADMIN).");
    else setUsersError("Não foi possível carregar usuários.");
  }
}

function resetUsersForm(){
  state.selectedUserId = null;
  fillFormFromUser(null);
  renderUsers();
}

(function wireUsersUi(){
  const { reload, form, cancelBtn, delBtn, search } = usersEls();
  if (reload) reload.addEventListener("click", async () => { await loadUsers(); });

  if (search) {
    search.addEventListener("input", () => {
      state.usersQuery = String(search.value || "");
      renderUsers();
    });
  }

  if (cancelBtn) cancelBtn.addEventListener("click", () => resetUsersForm());

  if (delBtn) {
    delBtn.addEventListener("click", async () => {
      const u = selectedUser();
      if (!u) return;
      if (!confirm(`Excluir o usuário "${u.id}"?`)) return;

      const ok = await requestReauth({ reason: "Para excluir usuário, confirme suas credenciais." });
      if (!ok) return;

      try {
        await api(`/api/users/${encodeURIComponent(u.id)}`, { method: "DELETE" });
        await loadUsers();
        resetUsersForm();
        alert("Usuário excluído.");
      } catch (e) {
        const msg = e?.data?.error || e?.data?.message || "";
        if (msg === "cannot_delete_self") alert("Você não pode excluir a si mesmo.");
        else if (msg === "cannot_delete_last_admin") alert("Não é possível excluir o último ADMIN.");
        else alert("Não foi possível excluir o usuário.");
      }
    });
  }

  if (form) {
    form.addEventListener("submit", async (ev) => {
      ev.preventDefault();
      setUsersError("");
      if (!canPerm("admin")) { setUsersError("Sem permissão ADMIN."); return; }

      const fd = new FormData(form);
      const id = String(fd.get("id") || "").trim();
      const name = String(fd.get("name") || "").trim();
      const password = String(fd.get("password") || "");
      const permissions = permsFromForm(fd);

      const editing = !!state.selectedUserId;

      if (!id || !name) { setUsersError("Informe usuário (ID) e nome."); return; }
      if (!editing && password.length < 4) { setUsersError("Senha fraca (mínimo 4 caracteres)."); return; }
      if (editing && password.length > 0 && password.length < 4) { setUsersError("Senha fraca (mínimo 4 caracteres)."); return; }

      const ok = await requestReauth({ reason: editing ? "Para alterar usuário, confirme suas credenciais." : "Para criar usuário, confirme suas credenciais." });
      if (!ok) return;

      try {
        if (!editing) {
          await api("/api/users", { method: "POST", body: JSON.stringify({ id, name, password, permissions }) });
          await loadUsers();
          resetUsersForm();
          // mantém o nome preenchido? limpa
          form.reset();
          fillFormFromUser(null);
          alert("Usuário criado.");
        } else {
          const body = { name, permissions };
          if (password.length > 0) body.password = password;
          await api(`/api/users/${encodeURIComponent(state.selectedUserId)}`, { method: "PUT", body: JSON.stringify(body) });
          await loadUsers();
          // mantém seleção
          const updated = (state.users || []).find(x => x.id === state.selectedUserId) || null;
          fillFormFromUser(updated);
          alert("Usuário atualizado.");
        }
      } catch (e) {
        const code = e?.data?.error || "";
        if (code === "user_exists") setUsersError("Já existe um usuário com esse ID.");
        else if (code === "invalid_user_id") setUsersError("ID inválido. Use letras/números/._- (2 a 32).");
        else if (code === "cannot_remove_last_admin") setUsersError("Não é possível remover o último ADMIN.");
        else if (code === "weak_password") setUsersError("Senha fraca (mín. 4).");
        else setUsersError("Não foi possível salvar. Veja o Console.");
        console.warn("user save error", e);
      }
    });
  }

  // Initial form state
  try { fillFormFromUser(null); } catch (_) {}
})();

// ---------- Tabs ----------
$$('.tab').forEach((btn) => {
  btn.addEventListener('click', async () => {
    $$('.tab').forEach((b) => b.classList.remove('active'));
    btn.classList.add('active');

    const tab = btn.dataset.tab;
    $$('.tabpanel').forEach((p) => p.classList.add('hidden'));
    $(`#tab-${tab}`).classList.remove('hidden');

    try {
      if (tab === 'mrp') {
        await loadInventory();
        await loadMRP();
        if (typeof renderPFBomList === 'function') renderPFBomList();
        if (typeof renderSimulateBox === 'function') renderSimulateBox();
        return;
      }
      if (tab === 'custos') {
        await loadInventory();
        await loadMRP();
        if (typeof renderCostingTab === 'function') renderCostingTab();
        return;
      }
      if (tab === 'compras') {
        await loadPurchaseOrders();
        if (typeof renderPurchaseOrders === 'function') renderPurchaseOrders();
        return;
      }
      if (tab === 'ops') {
        await loadOps();
        if (typeof renderOps === 'function') renderOps();
        return;
      }
      if (tab === 'vendas') {
        await loadInventory();
        await loadSalesAll();
        renderSalesCurrent();
        return;
      }
      if (tab === 'usuarios') {
        await loadUsers();
        renderUsers();
        return;
      }
      // estoque
      await loadInventory();
      // Se o usuário estiver no "Cadastro Geral", recarrega via API (OC/OP podem alterar estoque fora da aba Estoque).
      try {
        const secGeneral = $("#estoqueCadastroGeneral");
        const isVisible = (el) => el && !el.classList.contains("hidden");
        if (isVisible(secGeneral) && typeof refreshGeneralCadastroInline === "function"){
          await refreshGeneralCadastroInline({ preserveUI: true });
          return;
        }
      } catch (e) {
        // fallback para render normal
      }

      // Atualiza indicador/botões de estoque mínimo (MP/PF)
      try { await refreshMinStockStatus({ force: true }); } catch (_) {}

      if (typeof renderCurrentView === 'function') renderCurrentView();
    } catch (e) {
      console.warn('tab refresh error', e);
    }
  });
});

// ---------- Modal ----------
const modal = $("#modal");
const modalTitle = $("#modalTitle");
const modalSubtitle = $("#modalSubtitle");
const modalBody = $("#modalBody");
const modalSubmit = $("#modalSubmit");
const modalExtraBtn = $("#modalExtraBtn");

// Units manager modal (separate dialog, can be opened on top of item modal)
const unitsModal = $("#unitsModal");
const unitsModalTitle = $("#unitsModalTitle");
const unitsModalSubtitle = $("#unitsModalSubtitle");
const unitsModalBody = $("#unitsModalBody");
const unitsModalPickBtn = $("#unitsModalPickBtn");
const unitsModalCloseBtn = $("#unitsModalCloseBtn");
const unitsModalCancelBtn = $("#unitsModalCancelBtn");
if (unitsModalCloseBtn) unitsModalCloseBtn.addEventListener("click", () => unitsModal.close());
if (unitsModalCancelBtn) unitsModalCancelBtn.addEventListener("click", () => unitsModal.close());

// Close/cancel buttons must not trigger form validation
const modalCloseBtn = document.querySelector("#modalCloseBtn");
const modalCancelBtn = document.querySelector("#modalCancelBtn");
if (modalCloseBtn) modalCloseBtn.addEventListener("click", () => modal.close());
if (modalCancelBtn) modalCancelBtn.addEventListener("click", () => modal.close());


let modalOnSubmit = null;
let modalOnExtra = null;

function openModal({ title, subtitle = "", submitText = "Salvar", bodyHtml = "", onSubmit, onOpen, cardClass = "", extraBtn = null }) {
  modalTitle.textContent = title;
  modalSubtitle.textContent = subtitle;
  modalBody.innerHTML = bodyHtml;
  const mf = $("#modalForm");
  if (mf) mf.className = "modal-card" + (cardClass ? (" " + cardClass) : "");
  modalSubmit.textContent = submitText;
  modalOnSubmit = onSubmit;
  // Optional extra action (e.g., Imprimir)
  modalOnExtra = null;
  if (modalExtraBtn) {
    if (extraBtn && extraBtn.text) {
      modalExtraBtn.textContent = extraBtn.text;
      modalExtraBtn.classList.remove("hidden");
      // allow caller to add extra classes
      modalExtraBtn.className = "btn " + (extraBtn.className || "secondary") + (extraBtn.small ? " small" : "");
      modalOnExtra = typeof extraBtn.onClick === "function" ? extraBtn.onClick : null;
    } else {
      modalExtraBtn.classList.add("hidden");
      modalExtraBtn.textContent = "—";
      modalExtraBtn.className = "btn secondary hidden";
    }
  }
  modal.showModal();
  try { if (typeof onOpen === "function") onOpen(); } catch (e) { console.warn("onOpen error", e); }
}

if (modalExtraBtn) {
  modalExtraBtn.addEventListener("click", async () => {
    try {
      if (modalOnExtra) await modalOnExtra();
    } catch (e) {
      console.error("Modal extra action error", e);
      alert("Não foi possível executar esta ação. Veja o Console para detalhes.");
    }
  });
}

$("#modalForm").addEventListener("submit", async (e) => {
  e.preventDefault();
  const form = e.currentTarget;
  const fd = new FormData(form);
  try {
    if (modalOnSubmit) {
      const ok = await modalOnSubmit(fd);
      if (ok !== false) modal.close();
    } else {
      modal.close();
    }
  } catch (err) {
    console.error("Modal submit error", err);
    alert("Ocorreu um erro ao processar esta ação. Veja o Console para detalhes.");
  }
});

// Safety: never leave a stale submit handler around (prevents regressions where one modal
// breaks others after a failed submit / exception).
modal.addEventListener("close", () => {
  modalOnSubmit = null;
  modalOnExtra = null;
  if (modalExtraBtn) {
    modalExtraBtn.classList.add("hidden");
    modalExtraBtn.textContent = "—";
    modalExtraBtn.className = "btn secondary hidden";
  }
});

// ---------- Units (custom) ----------
function setSelectUnit(sel, unitCode) {
  const v = String(unitCode || "").trim();
  const prev = sel.value;
  sel.innerHTML = unitOptionsHtml(v || prev);
  if (v) sel.value = v;
}

function refreshOpenUnitSelects() {
  document.querySelectorAll('select[data-unit-select="1"]').forEach((sel) => {
    const current = sel.value;
    sel.innerHTML = unitOptionsHtml(current);
    sel.value = current;
  });
}

function wireUnitSelect(sel) {
  if (!sel || sel._unitWired) return;
  sel._unitWired = true;
  // Keep for future-proofing; dropdown contains only real units.
  sel.addEventListener("change", () => {});
}

async function openUnitsManager({ pickMode = false, initial = "", onPick, startInAddMode = false } = {}) {
  if (!unitsModal) return;
  unitsModalTitle.textContent = "Unidades";
  unitsModalSubtitle.textContent = "Adicionar / editar / remover unidades do sistema.";
  unitsModalPickBtn.classList.toggle("hidden", !pickMode);

  let selected = String(initial || "");
  let searchQ = "";
  let editingOld = "";

  const u$ = (s) => unitsModal.querySelector(s);

  unitsModalBody.innerHTML = `
    <div class="toolbar">
      <div class="toolbar-left">
        <input id="unitsSearch" class="input" placeholder="Pesquisar unidade..." />
      </div>
      <div class="toolbar-right">
        <button id="btnUnitsNew" type="button" class="btn secondary small">Novo</button>
      </div>
    </div>

    <div class="grid2">
      <div class="field">
        <label>Sigla</label>
        <input id="unitCode" class="input" placeholder="Ex.: g, pct, cx" />
      </div>
      <div class="field">
        <label>Descrição</label>
        <input id="unitLabel" class="input" placeholder="Ex.: g (grama)" />
      </div>
    </div>

    <div class="row" style="justify-content:flex-end; gap:10px">
      <button id="btnUnitDelete" type="button" class="btn danger small">Remover</button>
      <button id="btnUnitSave" type="button" class="btn primary small">Salvar</button>
    </div>

    <div class="table-wrap">
      <table class="table units-table">
        <thead>
          <tr>
            <th style="text-align:center; width:60px">POS</th>
            <th style="text-align:left">COD</th>
          </tr>
        </thead>
        <tbody id="unitsTbody"></tbody>
      </table>
    </div>

    <div id="unitsError" class="error hidden"></div>
  `;

  const render = () => {
    const tbody = u$("#unitsTbody");
    if (!tbody) return;
    tbody.innerHTML = "";

    const qq = searchQ.trim().toLowerCase();
    const list = (state.units || []).slice()
      .sort((a, b) => String(a.v || "").localeCompare(String(b.v || ""), "pt-BR"));

    const filtered = !qq ? list : list.filter((u) => {
      const hay = `${u.v || ""} ${u.l || ""}`.toLowerCase();
      return hay.includes(qq);
    });

    if (filtered.length === 0) {
      const tr = document.createElement("tr");
      tr.innerHTML = `<td colspan="2" class="muted">Nenhuma unidade encontrada.</td>`;
      tbody.appendChild(tr);
      return;
    }

    let pos = 0;
    for (const u of filtered) {
      pos += 1;
      const tr = document.createElement("tr");
      tr.className = (u.v === selected ? "row-selected" : "");
      tr.innerHTML = `
        <td style="text-align:center">${pos}</td>
        <td><b>${escapeHtml(u.v)}</b></td>
      `;
      tr.addEventListener("click", () => {
        selected = u.v;
        editingOld = u.v;
        const codeIn = u$("#unitCode");
        const lblIn = u$("#unitLabel");
        if (codeIn) codeIn.value = u.v;
        if (lblIn) lblIn.value = u.l || u.v;
        render();
      });
      tr.addEventListener("dblclick", () => {
        if (pickMode && typeof onPick === "function") {
          onPick(u.v);
          unitsModal.close();
        }
      });
      tbody.appendChild(tr);
    }
  };

  const showErr = (msg) => {
    const el = u$("#unitsError");
    if (!el) return;
    el.textContent = msg;
    el.classList.remove("hidden");
  };
  const clearErr = () => {
    const el = u$("#unitsError");
    if (!el) return;
    el.classList.add("hidden");
    el.textContent = "";
  };

  u$("#unitsSearch")?.addEventListener("input", (e) => {
    searchQ = e.target.value;
    render();
  });

  u$("#btnUnitsNew")?.addEventListener("click", () => {
    selected = "";
    editingOld = "";
    const codeIn = u$("#unitCode");
    const lblIn = u$("#unitLabel");
    if (codeIn) codeIn.value = "";
    if (lblIn) lblIn.value = "";
    render();
    codeIn?.focus();
  });

  u$("#btnUnitSave")?.addEventListener("click", async () => {
    clearErr();
    const codeIn = u$("#unitCode");
    const lblIn = u$("#unitLabel");
    const code = String(codeIn?.value || "").trim();
    const label = String(lblIn?.value || "").trim();
    if (!code) return showErr("Informe a sigla (ex.: kg, un, g…).");

    try {
      const exists = (state.units || []).some((u) => u.v === editingOld);
      if (exists && editingOld) {
        await api(`/api/units/${encodeURIComponent(editingOld)}`, {
          method: "PUT",
          body: JSON.stringify({ v: code, l: label }),
        });
      } else {
        await api(`/api/units`, {
          method: "POST",
          body: JSON.stringify({ v: code, l: label }),
        });
      }
      await loadUnits();
      refreshOpenUnitSelects();
      selected = code;
      editingOld = code;
      render();
    } catch (err) {
      const code = String(err?.data?.error || "");
      if (code === "unit_in_use") showErr("Não é possível alterar: unidade está em uso por algum item.");
      else if (code === "unit_exists") showErr("Já existe uma unidade com essa sigla.");
      else showErr("Erro ao salvar unidade.");
    }
  });

  u$("#btnUnitDelete")?.addEventListener("click", async () => {
    clearErr();
    const codeIn = u$("#unitCode");
    const code = String(codeIn?.value || "").trim();
    if (!code) return showErr("Selecione uma unidade para remover.");
    if (!confirm(`Remover a unidade "${code}"?`)) return;
    try {
      await api(`/api/units/${encodeURIComponent(code)}`, { method: "DELETE" });
      await loadUnits();
      refreshOpenUnitSelects();
      selected = "";
      editingOld = "";
      const lblIn = u$("#unitLabel");
      if (codeIn) codeIn.value = "";
      if (lblIn) lblIn.value = "";
      render();
    } catch (err) {
      const code = String(err?.data?.error || "");
      if (code === "unit_in_use") showErr("Não é possível remover: unidade está em uso por algum item.");
      else showErr("Erro ao remover unidade.");
    }
  });

  if (unitsModalPickBtn) {
    unitsModalPickBtn.onclick = () => {
      if (!selected) return;
      if (pickMode && typeof onPick === "function") {
        onPick(selected);
        unitsModal.close();
      }
    };
  }

  // Make sure units are loaded
  if (!state.units || state.units.length === 0) {
    try { await loadUnits(); } catch (e) { /* ignore */ }
  }

  // Prefill selection
  if (startInAddMode) {
    u$("#btnUnitsNew")?.click();
  } else if (initial) {
    const found = (state.units || []).find((u) => u.v === initial);
    if (found) {
      selected = found.v;
      editingOld = found.v;
      const codeIn = u$("#unitCode");
      const lblIn = u$("#unitLabel");
      if (codeIn) codeIn.value = found.v;
      if (lblIn) lblIn.value = found.l || found.v;
    }
  }

  render();
  unitsModal.showModal();
}

// ---------- Inventory UI ----------
function renderItems() {
  const tbody = $("#itemsTable tbody");
  if (!tbody) return;
  tbody.innerHTML = "";

  const items = [...filteredItems()].sort((a, b) => a.name.localeCompare(b.name, "pt-BR"));

  for (const it of items) {
    const low = Number(it.currentStock) < Number(it.minStock || 0);
    const tr = document.createElement("tr");
    tr.innerHTML = `
      <td>
        <div style="font-weight:700">${it.name}</div>
        <div class="muted small">${it.sku ? `SKU: ${it.sku}` : ""}</div>
      </td>
      <td>${it.unit}</td>
      <td>${fmt(it.currentStock)} ${it.unit} ${low ? `<span class="pill bad">baixo</span>` : ""}</td>
      <td>${fmt(it.minStock || 0)} ${it.unit}</td>
      <td>${state.stockMode === "fg" ? fmtMoney(it.salePrice || 0) : fmtMoney(it.cost || 0)}</td>
      <td style="text-align:right"><button class="btn secondary small" data-edit="${it.id}">Editar</button></td>
    `;
    tbody.appendChild(tr);
  }

  tbody.querySelectorAll("[data-edit]").forEach((b) => {
    b.addEventListener("click", () => {
      const id = b.getAttribute("data-edit");
      const it = state.items.find((x) => x.id === id);
      openEditItem(it);
    });
  });
}

function renderMovements() {
  const tbody = $("#movementsTable tbody");
  if (!tbody) return;
  tbody.innerHTML = "";
  const itemsById = new Map(((state.rawItems && state.rawItems.length) ? state.rawItems : state.items).map((i) => [i.id, i]));

  for (const mv of filteredMovements()) {
    const it = itemsById.get(mv.itemId);
    const typeLabel = mv.type === "in" ? "Entrada" : mv.type === "out" ? "Saída" : "Ajuste";
    const tr = document.createElement("tr");
    tr.innerHTML = `
      <td>${fmtDate(mv.at)}</td>
      <td>${typeLabel}</td>
      <td>${it?.name || "—"}</td>
      <td>${fmt(mv.qty)} ${it?.unit || ""}</td>
      <td class="muted">${mv.reason || ""}</td>
    `;
    tbody.appendChild(tr);
  }
}

function openNewItem(afterSave) {
  const isRaw = state.stockMode === "raw";
  const isFg = state.stockMode === "fg";
  const suggestedCode = computeNextCode(state.items.filter(i => (i.type||'raw')===state.stockMode), state.stockMode);
  const unitDefault = isFg ? "un" : "kg";
  const subtitle = isFg
    ? "Ex.: Marmita Arroz/Cambotian/Carne Vermelha/Verduras/Ovo 300g"
    : "Ex.: Frango (kg), Arroz integral (kg), Embalagem (un)…";

  openModal({
    title: "Novo item de estoque",
    subtitle,
    submitText: "Salvar",
    bodyHtml: `
      <div class="grid2">\n        <div class="field span2">\n          <label>Código</label>\n          <input id="itCode" class="input" value="${escapeHtml(suggestedCode)}" />\n        </div>\n
        <div class="field span2">
          <label>Nome</label>
          <input id="itName" class="input" required />
        </div>

        <div class="field">
          <div class="labelRow">
            <label>Unidade</label>
            <button id="btnUnitsManage" type="button" class="btn secondary small">Unidades</button>
          </div>
          <select id="itUnit" class="input" required data-unit-select="1">
            ${unitOptionsHtml(unitDefault)}
          </select>
        </div>

        ${isRaw ? `
          <div class="field">
            <div class="labelRow">
              <label>% Perda</label>
              <button id="btnPerdaHelp" type="button" class="btn secondary small">Saiba mais</button>
            </div>
            <input id="itLoss" class="input" type="number" min="0" max="100" step="0.1" value="0" />
          </div>

          <div class="field">
            <div class="labelRow">
              <label>FC</label>
              <button id="btnFcHelp" type="button" class="btn secondary small">Saiba mais</button>
            </div>
            <input id="itFc" class="input" type="number" min="0.01" step="0.01" value="1" placeholder="Ex.: arroz 2,8 | frango 0,80" />
          </div>
        ` : `
          <div class="field">
            <label>Estoque mínimo</label>
            <input id="itMin" class="input" type="number" step="0.01" value="0" />
          </div>
        `}

        ${isRaw ? `
          <div class="field">
            <label>Estoque mínimo</label>
            <input id="itMin" class="input" type="number" step="0.01" value="0" />
          </div>
          <div class="field">
            <label>Custo unitário (R$)</label>
            <input id="itCost" class="input" type="number" step="0.01" value="0" />
          </div>
        ` : `
          <div class="field">
            <label>Valor venda (R$)</label>
            <input id="itSale" class="input" type="number" step="0.01" value="0" />
          </div>
        `}
      </div>
    `,
    onOpen: () => {
      const btn = document.querySelector("#btnPerdaHelp");
      if (btn) btn.addEventListener("click", openPerdaHelp);
      const btnFc = document.querySelector("#btnFcHelp");
      if (btnFc) btnFc.addEventListener("click", openFcHelp);
      const sel = document.querySelector("#itUnit");
      if (sel) wireUnitSelect(sel);
      const btnUnits = document.querySelector("#btnUnitsManage");
      if (btnUnits && sel) btnUnits.addEventListener("click", () => openUnitsManager({ pickMode: true, initial: sel.value, onPick: (u) => setSelectUnit(sel, u) }));
    },
    onSubmit: async () => {
      const name = $("#itName").value.trim();
      const unit = $("#itUnit").value.trim();
      if (!name) { alert('Informe o nome.'); return false; }

      const codeRaw = $("#itCode")?.value || "";
      let code = normalizeItemCode(codeRaw, state.stockMode) || computeNextCode(state.items.filter(i => (i.type||'raw')===state.stockMode), state.stockMode);
      const prefix = (state.stockMode === 'fg') ? 'PF' : 'MP';
      code = String(code || '').trim().toUpperCase();
      if (!code.startsWith(prefix)) {
        alert('Código inválido. Use ' + prefix + '001, ' + prefix + '002...');
        return false;
      }
      const exists = (state.items || []).some(it => String(it.code||'').trim().toUpperCase() === code);
      if (exists) {
        alert('Já existe um item com o código ' + code + '.');
        return false;
      }

      const payload = {
        code,
        name,
        unit,
        minStock: Number($("#itMin")?.value || 0) || 0,
      };

      if (isRaw) {
        payload.lossPercent = Number($("#itLoss")?.value || 0) || 0;
        payload.cookFactor = Number($("#itFc")?.value || 1) || 1;
        payload.cost = Number($("#itCost")?.value || 0) || 0;
      }
      if (isFg) {
        payload.salePrice = Number($("#itSale")?.value || 0) || 0;
      }

      await api(`/api/inventory/items?type=${state.stockMode}`, {
        method: "POST",
        body: JSON.stringify(payload),
      });
      await loadInventory();
      try { await loadMRP(); } catch (e) {}
      // Atualiza qualquer tela que use o cadastro (BOM, receitas, ordens, etc.)
      renderCurrentView();
      if (typeof afterSave === "function") afterSave();
      return true;
    },
  });
}


function openEditItem(it, afterSave) {
  const isRaw = state.stockMode === "raw";
  const isFg = state.stockMode === "fg";
  const subtitle = isFg
    ? "Ex.: Marmita Arroz/Cambotian/Carne Vermelha/Verduras/Ovo 300g"
    : "Ex.: Frango (kg), Arroz integral (kg), Embalagem (un)…";

  openModal({
    title: "Editar item de estoque",
    subtitle,
    submitText: "Salvar",
    bodyHtml: `
      <div class="grid2">\n        <div class="field span2">\n          <label>Código</label>\n          <input id="itCode" class="input" value="${escapeHtml(it?.code || '')}" readonly />\n        </div>\n
        <div class="field span2">
          <label>Nome</label>
          <input id="itName" class="input" required value="${escapeHtml(it?.name || "")}" />
        </div>

        <div class="field">
          <div class="labelRow">
            <label>Unidade</label>
            <button id="btnUnitsManage" type="button" class="btn secondary small">Unidades</button>
          </div>
          <select id="itUnit" class="input" required data-unit-select="1">
            ${unitOptionsHtml(it?.unit || (isFg ? "un" : "kg"))}
          </select>
        </div>

        ${isRaw ? `
          <div class="field">
            <div class="labelRow">
              <label>% Perda</label>
              <button id="btnPerdaHelp" type="button" class="btn secondary small">Saiba mais</button>
            </div>
            <input id="itLoss" class="input" type="number" min="0" max="100" step="0.1" value="${Number(it?.lossPercent ?? 0)}" />
          </div>

          <div class="field">
            <div class="labelRow">
              <label>FC</label>
              <button id="btnFcHelp" type="button" class="btn secondary small">Saiba mais</button>
            </div>
            <input id="itFc" class="input" type="number" min="0.01" step="0.01" value="${Number(it?.cookFactor ?? 1)}" placeholder="Ex.: arroz 2,8 | frango 0,80" />
          </div>
        ` : `
          <div class="field">
            <label>Estoque mínimo</label>
            <input id="itMin" class="input" type="number" step="0.01" value="${Number(it?.minStock ?? 0)}" />
          </div>
        `}

        ${isRaw ? `
          <div class="field">
            <label>Estoque mínimo</label>
            <input id="itMin" class="input" type="number" step="0.01" value="${Number(it?.minStock ?? 0)}" />
          </div>
          <div class="field">
            <label>Custo unitário (R$)</label>
            <input id="itCost" class="input" type="number" step="0.01" value="${Number(it?.cost ?? 0)}" />
          </div>
        ` : `
          <div class="field">
            <label>Valor venda (R$)</label>
            <input id="itSale" class="input" type="number" step="0.01" value="${Number(it?.salePrice ?? 0)}" />
          </div>
        `}
      </div>
    `,
    onOpen: () => {
      const btn = document.querySelector("#btnPerdaHelp");
      if (btn) btn.addEventListener("click", openPerdaHelp);
      const btnFc = document.querySelector("#btnFcHelp");
      if (btnFc) btnFc.addEventListener("click", openFcHelp);
      const sel = document.querySelector("#itUnit");
      if (sel) wireUnitSelect(sel);
      const btnUnits = document.querySelector("#btnUnitsManage");
      if (btnUnits && sel) btnUnits.addEventListener("click", () => openUnitsManager({ pickMode: true, initial: sel.value, onPick: (u) => setSelectUnit(sel, u) }));
    },
    onSubmit: async () => {
      const name = $("#itName").value.trim();
      const rawCode = $("#itCode").value;
      const code = normalizeItemCode(rawCode, state.stockMode) || computeNextCode(state.items.filter(i => (i.type||'raw')===state.stockMode), state.stockMode);
      const prefix = (state.stockMode==="fg") ? "PF" : "MP";
      if (!code || !code.startsWith(prefix)) { alert('Código inválido. Use '+prefix+'001, '+prefix+'002...'); return false; }
      // Ao editar, permita salvar o mesmo código do próprio item.
      // Bloqueia apenas se outro item já usa o mesmo código.
      const normCode = String(code||'').trim().toUpperCase();
      const exists = (state.items||[]).some(i => {
        if (!i) return false;
        if (String(i.id||'') === String(it.id||'')) return false;
        return String(i.code||'').trim().toUpperCase() === normCode;
      });
      if (exists) { alert('Já existe um item com o código '+code+'.'); return false; }

      const unit = $("#itUnit").value.trim();

      const payload = {
        code,
        name,
        unit,
        minStock: Number($("#itMin")?.value || 0) || 0,
      };

      if (isRaw) {
        payload.lossPercent = Number($("#itLoss")?.value || 0) || 0;
        payload.cookFactor = Number($("#itFc")?.value || 1) || 1;
        payload.cost = Number($("#itCost")?.value || 0) || 0;
      }
      if (isFg) {
        payload.salePrice = Number($("#itSale")?.value || 0) || 0;
      }

      await api(`/api/inventory/items/${it.id}?type=${state.stockMode}`, {
        method: "PUT",
        body: JSON.stringify(payload),
      });
      await loadInventory();
      await loadMRP();
      if (typeof state._rerenderRecipeEditor === "function") state._rerenderRecipeEditor();
      // Atualiza qualquer tela que use o cadastro (BOM, receitas, ordens, etc.)
      renderCurrentView();
      if (typeof afterSave === "function") afterSave();
      return true;
    },
  });
}



function openMovement(type, opts = {}) {
  const preselectItemId = String(opts.preselectItemId || "");
  const lockItem = !!opts.lockItem;

  const typeLabel = type === "in" ? "Entrada" : type === "out" ? "Saída" : "Ajuste";
  const note = type === "adjust"
    ? "Ajuste define o estoque ABSOLUTO (ex.: contar e informar o novo valor)."
    : "Movimentação soma/subtrai do estoque.";

  const itemsSorted = filteredItems()
    .slice()
    .sort((a, b) => a.name.localeCompare(b.name, "pt-BR"));

  const options = itemsSorted
    .map((it) => {
      const sel = preselectItemId && String(it.id) === preselectItemId ? "selected" : "";
      return `<option value="${it.id}" ${sel}>${escapeHtml(it.name)} (${fmt(it.currentStock)} ${escapeHtml(it.unit)})</option>`;
    })
    .join("");

  // If select is disabled, it won't submit; add a hidden field for itemId.
  const lockedHidden = lockItem && preselectItemId
    ? `<input type="hidden" name="itemId" value="${escapeHtml(preselectItemId)}" />`
    : "";

  openModal({
    title: `Movimentação • ${typeLabel}`,
    subtitle: note,
    submitText: "Lançar",
    bodyHtml: `
      <label>Item
        <select name="itemId" required ${lockItem ? "disabled" : ""}>
          <option value="" disabled ${preselectItemId ? "" : "selected"}>Selecione…</option>
          ${options}
        </select>
      </label>
      ${lockedHidden}
      <label>${type === "adjust" ? "Novo estoque (absoluto)" : "Quantidade"}
        <input name="qty" type="number" step="0.001" required />
      </label>
      <label>Observação<input name="reason" placeholder="Ex.: compra fornecedor / perda / contagem..." /></label>
    `,
    onSubmit: async (fd) => {
      if (filteredItems().length === 0) {
        alert("Cadastre um item antes.");
        return false;
      }
      const payload = {
        type,
        itemId: fd.get("itemId"),
        qty: Number(fd.get("qty")),
        reason: fd.get("reason"),
      };
      try {
        await api(`/api/inventory/movements?type=${state.stockMode}`, { method: "POST", body: JSON.stringify(payload) });
      } catch (err) {
        if (err?.data?.error === "insufficient_stock") {
          alert(`Estoque insuficiente. Atual: ${fmt(err.data.stock)}.`);
          return false;
        }
        throw err;
      }
      await loadInventory();
      if (typeof afterSave === 'function') { afterSave(); }
      return true;
    },
  });
}

function openAdjustPicker() {
  const modeLabel = state.stockMode === "fg" ? "Produto Final (PF)" : "Matéria-prima & Insumos (MP)";
  // Ordenar por código (ERP-style). Mantém a lista previsível após import/criar/excluir.
  const items = filteredItems()
    .slice()
    .sort((a, b) => String(a.code || "").localeCompare(String(b.code || ""), "pt-BR"));

  if (items.length === 0) {
    alert("Cadastre um item antes de fazer ajuste de inventário.");
    return;
  }

  const rows = items.map(it => {
    const codeRaw = String(it.code || "");
    const nameRaw = String(it.name || "");
    const unitRaw = String(it.unit || "");
    const codeDisplay = codeRaw.length > 6 ? codeRaw.slice(0, 6) : codeRaw;
    const nameDisplay = nameRaw.length > 60 ? (nameRaw.slice(0, 57) + "...") : nameRaw;
    return `
      <tr data-id="${it.id}" data-search="${escapeHtml((codeRaw + " " + nameRaw + " " + unitRaw).toLowerCase())}">
        <td class="inv-code" title="${escapeHtml(codeRaw)}">${escapeHtml(codeDisplay)}</td>
        <td class="inv-desc"><span class="g-desc" title="${escapeHtml(nameRaw)}">${escapeHtml(nameDisplay)}</span></td>
        <td class="inv-stock" title="${escapeHtml(String(it.currentStock ?? ""))}">${fmt(it.currentStock)}</td>
        <td class="inv-unit" title="${escapeHtml(unitRaw)}">${escapeHtml(unitRaw)}</td>
      </tr>
    `;
  }).join("");

  openModal({
    title: `Ajuste de Inventário • ${modeLabel}`,
    subtitle: "Selecione um item para ajustar manualmente. Depois informe o novo estoque (contagem física).",
    submitText: "Selecionar item",
    bodyHtml: `
      <div class="toolbar" style="margin-top:0;">
        <div class="toolbar-left">
          <input id="adjustPickSearch" class="input" placeholder="Pesquisar..." />
        </div>
        <div class="muted small" style="flex: 0 0 auto;">Mostrando: código, descrição e estoque atual</div>
      </div>

      <div class="cadastro-tablewrap inv-pick-wrap" style="margin-top:10px;">
        <table class="table cadastro-table pick-table">
          <thead>
            <tr>
              <th>Código</th>
              <th>Descrição</th>
              <th style="text-align:center;">Estoque atual</th>
              <th style="text-align:center;">UN</th>
            </tr>
          </thead>
          <tbody id="adjustPickTbody">
            ${rows}
          </tbody>
        </table>
      </div>

      <input type="hidden" name="itemId" id="adjustPickItemId" required />
    `,
    // Wide so all columns (Estoque atual + UN) always fit without horizontal scrolling
    // even when descriptions are long.
    cardClass: "wide inv-pick",
    onOpen: () => {
      const tbody = document.querySelector("#adjustPickTbody");
      const hidden = document.querySelector("#adjustPickItemId");
      const search = document.querySelector("#adjustPickSearch");

      modalSubmit.disabled = true;

      function clearSelection() {
        tbody?.querySelectorAll("tr.row-selected").forEach(tr => tr.classList.remove("row-selected"));
        if (hidden) hidden.value = "";
        modalSubmit.disabled = true;
      }

      tbody?.addEventListener("click", (e) => {
        const tr = e.target?.closest?.("tr[data-id]");
        if (!tr) return;
        tbody.querySelectorAll("tr.row-selected").forEach(x => x.classList.remove("row-selected"));
        tr.classList.add("row-selected");
        if (hidden) hidden.value = tr.dataset.id;
        modalSubmit.disabled = false;
      });

      search?.addEventListener("input", () => {
        const q = (search.value || "").toLowerCase().trim();
        let visibleSelected = false;
        tbody?.querySelectorAll("tr[data-id]").forEach(tr => {
          const hay = (tr.dataset.search || "");
          const ok = !q || hay.includes(q);
          tr.style.display = ok ? "" : "none";
          if (ok && tr.classList.contains("row-selected")) visibleSelected = true;
        });
        if (!visibleSelected) clearSelection();
      });
    },
    onSubmit: async (fd) => {
      const itemId = String(fd.get("itemId") || "");
      if (!itemId) {
        alert("Selecione um item.");
        return false;
      }
      setTimeout(() => {
        openMovement("adjust", { preselectItemId: itemId, lockItem: true });
      }, 50);
      return true;
    },
  });
}


// ---------- MRP UI ----------

function renderRecipes() {
  // Legacy (UI moved to PF BOM list). Keep as no-op to avoid breaking older calls.
  return;
}

function recipeForPF(pfId) {
  return (state.recipes || []).find((r) => String(r.productId || "") === String(pfId)) || null;
}

// ---------- Print helpers (no popups) ----------
function printHtmlInIframe(html) {
  // Avoid popup blockers by printing inside a hidden iframe.
  const iframe = document.createElement("iframe");
  iframe.setAttribute("aria-hidden", "true");
  iframe.style.position = "fixed";
  iframe.style.right = "0";
  iframe.style.bottom = "0";
  iframe.style.width = "0";
  iframe.style.height = "0";
  iframe.style.border = "0";
  iframe.style.opacity = "0";
  iframe.style.pointerEvents = "none";

  // Some browsers only support srcdoc; others need document.write.
  // We'll try srcdoc first.
  iframe.srcdoc = html;
  document.body.appendChild(iframe);

  let didPrint = false;
  let didCleanup = false;

  const cleanup = () => {
    if (didCleanup) return;
    didCleanup = true;
    try { iframe.remove(); } catch (e) { /* noop */ }
  };

  const doPrint = () => {
    if (didPrint) return;
    didPrint = true;
    try {
      const w = iframe.contentWindow;
      if (!w) throw new Error("no-window");
      // Ensure focus for some browsers
      w.focus();
      // Give layout a moment
      setTimeout(() => {
        try {
          w.print();
        } catch (e) {
          console.warn("print failed", e);
          alert("Não foi possível abrir a impressão. Verifique as permissões do navegador.");
        }
        // Remove iframe after print dialog
        setTimeout(cleanup, 1500);
      }, 50);
    } catch (e) {
      console.warn("print setup failed", e);
      alert("Não foi possível preparar a impressão.");
      cleanup();
    }
  };

  iframe.addEventListener("load", doPrint, { once: true });
  // Fallback in case load never fires
  setTimeout(doPrint, 800);
}

function printMethodDocument({ pfCode = "", pfName = "", methodText = "" } = {}) {
  const safeTitle = `${pfCode || "PF"} - ${pfName || ""}`.trim();
  const html = `<!doctype html>
  <html lang="pt-BR">
  <head>
    <meta charset="utf-8" />
    <meta name="viewport" content="width=device-width, initial-scale=1" />
    <title>Método - ${escapeHtml(safeTitle)}</title>
    <style>
      /* Force light print */
      html, body { background: #fff !important; color: #000 !important; }
      @page { margin: 12mm; }
      body { font-family: system-ui, -apple-system, Segoe UI, Roboto, Arial, sans-serif; padding: 0; }
      .top { display:flex; align-items:center; gap:16px; }
      .logo { width: 110px; height:auto; }
      h1 { font-size: 24px; margin: 0; }
      .sub { font-size: 13px; margin-top: 4px; color:#111; font-weight: 600; }
      .meta { margin-top: 10px; font-size: 12px; color:#111; }
      .rule { margin-top: 12px; border-top: 1px solid #bbb; }
      .method { margin-top: 10px; white-space: pre-wrap; line-height: 1.38; font-size: 13.2px; }
      .sig { margin-top: 18px; display:flex; gap:18px; flex-wrap: wrap; }
      .sig .line { flex: 1 1 260px; }
      .sig .lbl { font-size: 12px; margin-bottom: 6px; }
      .sig .bar { border-bottom: 1px solid #000; height: 18px; }
      .pagebreak { page-break-before: always; break-before: page; }
      @media print {
        .top { gap: 12px; }
      }
    </style>
  </head>
  <body>
    <div class="top">
      <img class="logo" src="/logo.png" alt="dietON" />
      <div>
        <h1>Método de preparo</h1>
        <div class="sub">${escapeHtml(safeTitle)}</div>
        <div class="meta">Data/Hora: ${new Date().toLocaleString("pt-BR")}</div>
      </div>
    </div>
    <div class="rule"></div>
    <div class="method">${escapeHtml(methodText || "").replace(/\n/g, "\n")}</div>
    <div class="sig">
      <div class="line">
        <div class="lbl">Assinatura</div>
        <div class="bar"></div>
      </div>
      <div class="line" style="max-width:240px;">
        <div class="lbl">Data</div>
        <div class="bar"></div>
      </div>
    </div>
  </body>
  </html>`;

  printHtmlInIframe(html);
}

function printBomAndMethodDocument({ pfCode = "", pfName = "", bomLines = [], itemById = new Map(), methodText = "" } = {}) {
  const safeTitle = `${pfCode || "PF"} - ${pfName || ""}`.trim();
  const rows = (bomLines || [])
    .slice()
    .sort((a, b) => Number(a.pos || 0) - Number(b.pos || 0))
    .map((l, idx) => {
      const it = itemById.get(l.itemId);
      if (!it) return "";
      const du = (function dispUnit(u){ const uu=String(u||"").toLowerCase(); if(uu==="kg")return"g"; if(uu==="l")return"ml"; return uu||""; })(it.unit);
      const fromBaseQty = (baseQty, itemUnit) => {
        const q = Number(baseQty || 0);
        const uu = String(itemUnit||"").toLowerCase();
        if (!Number.isFinite(q)) return 0;
        if (uu === "kg") return q * 1000;
        if (uu === "l") return q * 1000;
        return q;
      };
      const getFC = (it2) => {
        const v = Number(it2?.cookFactor ?? it2?.fc ?? 1);
        return Number.isFinite(v) && v > 0 ? v : 1;
      };
      const fmtLocal = (n) => {
        const x = Number(n);
        if (!Number.isFinite(x)) return "—";
        // keep comma for PT-BR
        return x.toLocaleString("pt-BR", { maximumFractionDigits: 3 });
      };
      const qtyDisp = fromBaseQty(l.qty, it.unit);
      const fcUsed = (l.fc !== null && l.fc !== undefined && Number.isFinite(Number(l.fc)) && Number(l.fc) > 0) ? Number(l.fc) : getFC(it);
      const cookedDisp = qtyDisp * fcUsed;
      return `
        <tr>
          <td style="text-align:center;">${escapeHtml(String(l.pos || (idx+1)))}</td>
          <td><b>${escapeHtml(it.code || "MP")}</b></td>
          <td>${escapeHtml(it.name || "")}</td>
          <td style="text-align:center;">${escapeHtml(it.unit || "")}</td>
          <td style="text-align:center;">${escapeHtml(du || "")}</td>
          <td style="text-align:center;">${fmtLocal(fcUsed)}</td>
          <td style="text-align:right;">${fmtLocal(qtyDisp)} ${escapeHtml(du || "")}</td>
          <td style="text-align:right;">${fmtLocal(cookedDisp)} ${escapeHtml(du || "")}</td>
        </tr>`;
    }).join("");

  const html = `<!doctype html>
  <html lang="pt-BR">
  <head>
    <meta charset="utf-8" />
    <meta name="viewport" content="width=device-width, initial-scale=1" />
    <title>BOM + Método - ${escapeHtml(safeTitle)}</title>
    <style>
      html, body { background:#fff !important; color:#000 !important; }
      @page { margin: 12mm; }
      body { font-family: system-ui, -apple-system, Segoe UI, Roboto, Arial, sans-serif; }
      .top { display:flex; align-items:center; gap:12px; }
      .logo { width: 110px; height:auto; }
      h1 { font-size: 22px; margin: 0; }
      .sub { font-size: 13px; margin-top: 4px; font-weight: 600; }
      .meta { margin-top: 6px; font-size: 12px; }
      h2 { font-size: 15px; margin: 14px 0 8px; }
      table { width: 100%; border-collapse: collapse; }
      thead { display: table-header-group; }
      th, td { border: 1px solid #bbb; padding: 6px 7px; font-size: 11.5px; vertical-align: top; }
      th { background: #f2f2f2; }
      tr { break-inside: avoid; page-break-inside: avoid; }
      .method { white-space: pre-wrap; line-height: 1.38; font-size: 13px; border-top: 1px solid #bbb; padding-top: 10px; }
      .pagebreak { page-break-before: always; break-before: page; }
      .sig { margin-top: 18px; display:flex; gap:18px; flex-wrap: wrap; }
      .sig .line { flex: 1 1 260px; }
      .sig .lbl { font-size: 12px; margin-bottom: 6px; }
      .sig .bar { border-bottom: 1px solid #000; height: 18px; }
    </style>
  </head>
  <body>
    <div class="top">
      <img class="logo" src="/logo.png" alt="dietON" />
      <div>
        <h1>BOM + Método</h1>
        <div class="sub">${escapeHtml(safeTitle)}</div>
        <div class="meta">Data/Hora: ${new Date().toLocaleString("pt-BR")}</div>
      </div>
    </div>

    <h2>Ingredientes (BOM)</h2>
    <table>
      <thead>
        <tr>
          <th style="width:42px; text-align:center;">POS</th>
          <th style="width:68px;">COD</th>
          <th>Descrição</th>
          <th style="width:70px; text-align:center;">UN Compra</th>
          <th style="width:78px; text-align:center;">UN Consumo</th>
          <th style="width:58px; text-align:center;">FC</th>
          <th style="width:92px; text-align:right;">QTE (crua)</th>
          <th style="width:96px; text-align:right;">QTE (cozida)</th>
        </tr>
      </thead>
      <tbody>
        ${rows || `<tr><td colspan="8">Sem itens.</td></tr>`}
      </tbody>
    </table>

    <div class="pagebreak"></div>
    <h2>Método de preparo</h2>
    <div class="method">${escapeHtml(methodText || "")}</div>

    <div class="sig">
      <div class="line">
        <div class="lbl">Assinatura</div>
        <div class="bar"></div>
      </div>
      <div class="line" style="max-width:240px;">
        <div class="lbl">Data</div>
        <div class="bar"></div>
      </div>
    </div>
  </body>
  </html>`;

  printHtmlInIframe(html);
}

// --- Metodo de preparo (novo modulo) ---
function openMethodForPF(pfId) {
  const pf = (state.fgItems || []).find((i) => i.id === pfId);
  const rr = recipeForPF(pfId);
  const hasBom = !!(rr && Array.isArray(rr.bom) && rr.bom.length > 0);
  if (!rr?.id || !hasBom) {
    alert("Crie o BOM deste Produto Final antes de cadastrar o método.");
    return;
  }

  openModal({
    title: `Método de preparo`,
    subtitle: pf ? `Produto Final: ${pf.code || "PF"} — ${pf.name || ""}` : "",
    submitText: "Salvar",
    cardClass: "wide",
    extraBtn: {
      text: "Imprimir",
      className: "secondary",
      onClick: () => {
        const ta = modalBody?.querySelector('textarea[name="method"]');
        const methodText = ta ? ta.value : (rr.method || "");
        printMethodDocument({ pfCode: pf?.code || "", pfName: pf?.name || "", methodText });
      },
    },
    bodyHtml: `
      <label class="tight">Método
        <textarea class="method-textarea" name="method" placeholder="Passo a passo, tempos, ordem de preparo...">${escapeHtml(rr.method || "")}</textarea>
      </label>
    `,
    onSubmit: async (fd) => {
      const method = String(fd.get("method") || "");
      await api(`/api/mrp/recipes/${rr.id}`, { method: "PUT", body: JSON.stringify({ method }) });
      await loadMRP();
      renderPFBomList();
      renderPFPreview();
      return true;
    },
  });
}

function renderPFBomList() {
  const table = document.querySelector("#pfBomTable");
  const searchEl = document.querySelector("#pfBomSearch");
  const tbody = table?.querySelector("tbody");
  if (!tbody) return;

  const q = String(searchEl?.value || "").toLowerCase().trim();

  const pfs = (state.fgItems || [])
    .slice()
    .sort((a, b) => String(a.code || "").localeCompare(String(b.code || ""), "pt-BR"));

  tbody.innerHTML = "";

  const filtered = pfs.filter((pf) => {
    const hay = `${pf.code || ""} ${pf.name || ""} ${pf.unit || ""}`.toLowerCase();
    return !q || hay.includes(q);
  });

  if (filtered.length === 0) {
    const tr = document.createElement("tr");
    tr.innerHTML = `<td colspan="5" class="muted">Nenhum Produto Final encontrado.</td>`;
    tbody.appendChild(tr);
    return;
  }

  for (const pf of filtered) {
    const r = recipeForPF(pf.id);
    const hasBom = !!(r && Array.isArray(r.bom) && r.bom.length > 0);

    const tr = document.createElement("tr");
    tr.dataset.id = pf.id;
    tr.dataset.search = `${pf.code || ""} ${pf.name || ""} ${pf.unit || ""}`.toLowerCase();
    if (state.selectedPFId === pf.id) tr.classList.add("row-selected");

    tr.innerHTML = `
      <td style="text-align:left"><b>${escapeHtml(pf.code || "PF")}</b></td>
      <td style="text-align:left"><span class="pf-desc">${escapeHtml(pf.name || "")}</span></td>
      <td style="text-align:center">${escapeHtml(fmtUnit(pf.unit || ""))}</td>
      <td style="text-align:center">${hasBom ? `<span class="pill ok">cadastrado</span>` : `<span class="pill warn">pendente</span>`}</td>
      <td style="text-align:center">
        <div class="mrp-row-actions">
          <button class="btn secondary small" data-bom="${escapeHtml(pf.id)}">${hasBom ? "Editar BOM" : "Criar BOM"}</button>
          <button class="btn secondary small" data-method="${escapeHtml(pf.id)}" ${hasBom ? "" : "disabled"} title="${hasBom ? "Editar método" : "Crie o BOM antes de cadastrar o método"}">${hasBom ? "Editar Método" : "Criar Método"}</button>
          <button class="btn danger small" data-delbom="${escapeHtml(pf.id)}" ${hasBom ? "" : "disabled"} title="${hasBom ? "Excluir BOM" : "Sem BOM para excluir"}">Excluir BOM</button>
        </div>
      </td>
    `;

    tr.addEventListener("click", (e) => {
      if (e.target?.matches?.("button[data-bom], button[data-method], button[data-delbom]")) return;
      state.selectedPFId = pf.id;
      const rr = recipeForPF(pf.id);
      state.selectedRecipeId = rr?.id || null;
      renderPFBomList();
      renderPFPreview();
    });

    tr.querySelector("button[data-bom]").addEventListener("click", (e) => {
      e.stopPropagation();
      openBOMForPF(pf.id);
    });

    const btnMethod = tr.querySelector("button[data-method]");
    if (btnMethod) {
      btnMethod.addEventListener("click", (e) => {
        e.stopPropagation();
        if (btnMethod.disabled) {
          alert("Crie o BOM deste Produto Final antes de cadastrar o método.");
          return;
        }
        openMethodForPF(pf.id);
      });
    }

    const btnDel = tr.querySelector("button[data-delbom]");
    if (btnDel) {
      btnDel.addEventListener("click", async (e) => {
        e.stopPropagation();
        if (btnDel.disabled) return;
        const rr = recipeForPF(pf.id);
        const has = !!(rr && Array.isArray(rr.bom) && rr.bom.length > 0);
        if (!has || !rr?.id) return;
        const ok = confirm(`Excluir BOM de ${pf.code || "PF"} — ${pf.name || ""}?\n\nIsso remove a estrutura (BOM) e a foto do produto (se existir).`);
        if (!ok) return;
        await deleteAfterReauth(`/api/mrp/recipes/${rr.id}/bom`, "Excluir BOM (confirmar usuário e senha)");
        // mantém seleção no PF atual
        state.selectedPFId = pf.id;
        await loadMRP();
        renderPFBomList();
        renderPFPreview();
      });
    }

    tbody.appendChild(tr);
  }

  // Default selection: first PF
  if (!state.selectedPFId && filtered[0]) {
    state.selectedPFId = filtered[0].id;
    const rr = recipeForPF(filtered[0].id);
    state.selectedRecipeId = rr?.id || null;
  }

  // Atualiza o preview (foto/infos) na aba MRP
  renderPFPreview();
}

// ---------- MRP: preview (lista) vs simulação (tela separada) ----------

function mrpShowSimView(on) {
  const list = document.querySelector('#mrpListView');
  const sim = document.querySelector('#mrpSimView');
  const hList = document.querySelector('#mrpHeaderList');
  const hSim = document.querySelector('#mrpHeaderSim');
  if (on) {
    list?.classList.add('hidden');
    sim?.classList.remove('hidden');
    hList?.classList.add('hidden');
    hSim?.classList.remove('hidden');
  } else {
    sim?.classList.add('hidden');
    list?.classList.remove('hidden');
    hSim?.classList.add('hidden');
    hList?.classList.remove('hidden');
  }
}

function openMrpSimView() {
  // Tela de simulação em lote (seleciona múltiplos PFs com BOM)
  mrpShowSimView(true);
  renderSimulate('#mrpSimulateBox');
}

function openMrpListView() {
  mrpShowSimView(false);
  renderPFPreview();
}

function renderPFPreview() {
  const box = document.querySelector('#pfPreviewBox');
  if (!box) return;

  if (!state.selectedPFId) {
    box.innerHTML = `<div class="muted">Selecione um Produto Final (PF) para ver o preview.</div>`;
    return;
  }

  const pf = (state.fgItems || []).find((x) => x.id === state.selectedPFId);
  const recipe = state.recipes.find((r) => r.id === state.selectedRecipeId);

  if (!recipe) {
    // No preview buttons here: user already has actions in the main table.
    // When there is no BOM yet, show only the selected PF description.
    box.innerHTML = `
      <div class="pf-preview pf-preview--empty">
        <div class="pf-preview-title" style="text-align:center;">${escapeHtml(pf?.code || 'PF')} • ${escapeHtml(pf?.name || '')}</div>
      </div>
    `;
    return;
  }

  box.innerHTML = `
    <div class="pf-preview">
      <div class="pf-preview-title" style="text-align:center;">${escapeHtml(recipe.name)}</div>
      <div class="pf-preview-center">
        <div class="pf-photo-card" style="margin-top:12px;">
          <div id="pfPrevPhotoFrame" class="photo-frame"></div>
        </div>
      </div>
    </div>
  `;

  const frame = document.querySelector('#pfPrevPhotoFrame');
  if (frame) {
    const pfFile = recipe?.photoFile;
    if (pfFile) frame.innerHTML = `<img alt="Foto" src="/photos/${encodeURIComponent(pfFile)}?t=${Date.now()}" />`;
    else frame.innerHTML = `<div class="photo-placeholder">+</div>`;
  }
}

// Compat: chamadas antigas tentavam renderizar a área de simulação na aba MRP.
// Agora, na aba MRP mostramos apenas preview (foto) e a simulação fica em uma tela separada.
function renderSimulateBox() {
  renderPFPreview();
}

function openBOMForPF(pfId, { prefillBomItemId = "" } = {}) {
  state.selectedPFId = pfId;
  const pf = (state.fgItems || []).find((x) => x.id === pfId);
  if (!pf) {
    alert("Produto Final não encontrado.");
    return;
  }

  // Needs MP items to build BOM
  if ((state.rawItems || []).length === 0) {
    alert("Cadastre Matéria-prima & Insumos (MP) antes (ex.: Frango, Arroz, Embalagem).");
    return;
  }

  const existing = recipeForPF(pfId);
  if (existing) {
    // ensure editor is linked
    openRecipeEditor({ mode: "edit", recipe: existing, productId: pfId, prefillBomItemId });
    return;
  }

  // Create new, prefilled with PF info
  const prefill = {
    id: "",
    name: pf.name || `${pf.code || "PF"} - Produto Final`,
    productId: pfId,
    yieldQty: 1,
    yieldUnit: pf.unit || "un",
    notes: "",
    bom: [],
  };

  openRecipeEditor({ mode: "new", recipe: prefill, productId: pfId, prefillBomItemId });
}

function openPFBomPicker({ prefillBomItemId = "" } = {}) {
  const pfs = (state.fgItems || []).slice().sort((a, b) => String(a.code||"").localeCompare(String(b.code||""), "pt-BR"));
  openModal({
    title: "Selecionar Produto Final (PF)",
    subtitle: "Escolha o PF para cadastrar/editar o BOM.",
    submitText: "Abrir BOM",
    bodyHtml: `
      <div class="toolbar">
        <input id="pfPickSearch" class="input" placeholder="Pesquisar PF (código, descrição...)" />
      </div>
      <div class="table-wrap" style="max-height: 340px;">
        <table class="table">
          <thead><tr><th>COD</th><th>Descrição</th><th style="text-align:center">UN</th><th style="text-align:center">BOM</th></tr></thead>
          <tbody id="pfPickTbody">
            ${pfs.map(pf => {
              const r = recipeForPF(pf.id);
              const hasBom = !!(r && r.bom?.length);
              return `<tr data-id="${escapeHtml(pf.id)}" data-search="${escapeHtml((pf.code+" "+pf.name).toLowerCase())}">
                <td><b>${escapeHtml(pf.code||"PF")}</b></td>
                <td>${escapeHtml(pf.name||"")}</td>
                <td style="text-align:center">${escapeHtml(fmtUnit(pf.unit||""))}</td>
                <td style="text-align:center">${hasBom ? `<span class="pill ok">ok</span>` : `<span class="pill warn">—</span>`}</td>
              </tr>`;
            }).join("")}
          </tbody>
        </table>
      </div>
      <input type="hidden" name="pfId" id="pfPickId" required />
    `,
    onOpen: () => {
      const tbody = document.querySelector("#pfPickTbody");
      const hidden = document.querySelector("#pfPickId");
      const search = document.querySelector("#pfPickSearch");
      modalSubmit.disabled = true;

      function clearSel() {
        tbody?.querySelectorAll("tr.row-selected").forEach(tr => tr.classList.remove("row-selected"));
        if (hidden) hidden.value = "";
        modalSubmit.disabled = true;
      }

      tbody?.addEventListener("click", (e) => {
        const tr = e.target?.closest?.("tr[data-id]");
        if (!tr) return;
        tbody.querySelectorAll("tr.row-selected").forEach(x => x.classList.remove("row-selected"));
        tr.classList.add("row-selected");
        if (hidden) hidden.value = tr.dataset.id;
        modalSubmit.disabled = false;
      });

      search?.addEventListener("input", () => {
        const qq = (search.value || "").toLowerCase().trim();
        let visibleSelected = false;
        tbody?.querySelectorAll("tr[data-id]").forEach(tr => {
          const ok = !qq || (tr.dataset.search || "").includes(qq);
          tr.style.display = ok ? "" : "none";
          if (ok && tr.classList.contains("row-selected")) visibleSelected = true;
        });
        if (!visibleSelected) clearSel();
      });
    },
    onSubmit: async (fd) => {
      const pfId = String(fd.get("pfId") || "");
      if (!pfId) return false;
      openBOMForPF(pfId, { prefillBomItemId });
      return true;
    }
  });
}
async function renderSimulate(boxSelector = "#mrpSimulateBox") {
  const box = document.querySelector(boxSelector);
  if (!box) return;
  box.classList.remove('muted');

  const allRecipes = Array.isArray(state.recipes) ? state.recipes : [];
  const fgItems = Array.isArray(state.fgItems) ? state.fgItems : [];

  // Associate each PF item with its recipe (by outputItemId/productId)
  const pfWithRecipe = fgItems
    .map(pf => ({
      pf,
      recipe: allRecipes.find(r => r && (r.outputItemId === pf.id || r.productId === pf.id))
    }))
    .filter(x => x.recipe && Array.isArray(x.recipe.bom) && x.recipe.bom.length > 0);

  const withBom = pfWithRecipe;

  // label: sempre usar PF (nunca UUID de receita)
  const labelByRid = new Map(withBom.map(x => {
    const code = String(x.pf?.code || '').trim();
    const name = String(x.pf?.name || '').trim();
    const label = [code, name].filter(Boolean).join(' - ');
    return [x.recipe.id, label || String(x.recipe.id)];
  }));

  // init maps
  if (!Array.isArray(state.simulateSelected)) state.simulateSelected = [];
  if (!state.simulateQtyById || typeof state.simulateQtyById !== 'object') state.simulateQtyById = {};
  if (typeof state.simulateLocked !== 'boolean') state.simulateLocked = false;
  const isLocked = !!state.simulateLocked;
  // default qty=1 for items missing qty
  for (const x of withBom) {
    const rid = x.recipe.id;
    if (state.simulateQtyById[rid] === undefined) state.simulateQtyById[rid] = 1;
  }

  const isSelected = (rid) => state.simulateSelected.includes(rid);

  const rowsHtml = withBom.map((x) => {
    const rid = x.recipe.id;
    const checked = isSelected(rid) ? 'checked' : '';
    const qty = Number(state.simulateQtyById[rid] ?? 1);
    const safeQty = Number.isFinite(qty) ? qty : 1;
    return `
      <tr data-rid="${rid}">
        <td style="width:34px; text-align:center"><input type="checkbox" data-sel="${rid}" ${checked} /></td>
        <td style="width:70px"><b>${escapeHtml(x.pf.code || '')}</b></td>
        <td>${escapeHtml(x.pf.name || '')}</td>
        <td style="width:120px; text-align:center">
          <input class="input qty-inline" type="number" min="0" step="1" data-qty="${rid}" value="${safeQty}" />
        </td>
      </tr>`;
  }).join('');

  box.innerHTML = `
    <h3 style="margin:0 0 6px 0;">Simular produção</h3>
    <div class="muted small">Selecione os Produtos Finais (PF) com BOM e informe a quantidade de cada um para calcular necessidades e gerar OP/OC.</div>
    ${isLocked ? `<div style="margin-top:10px; padding:10px; border:1px solid #f0c36d; background:#fff8e1; border-radius:10px; font-size:13px; line-height:1.45">
      <b>Simulação travada</b><br/>
      Esta simulação já gerou OP (e OC automaticamente se havia faltas). Para criar outra simulação, clique em <b>Limpar seleção</b>.
    </div>` : ``}

    <div class="row wrap" style="gap:10px; margin-top:10px; align-items:flex-end;">
      <button id="btnSelectAllSim" type="button" class="btn secondary small">Selecionar todos</button>
      <button id="btnClearSim" type="button" class="btn secondary small">Limpar seleção</button>
      <div class="muted small" style="margin-left:auto;">Dica: marque o PF e ajuste a quantidade (ex.: 100, 10, 20...).</div>
    </div>

    <div class="table-wrap" style="margin-top:10px;">
      <table class="table" id="simTable">
        <thead>
          <tr>
            <th style="width:34px;"></th>
            <th style="width:70px">COD</th>
            <th>Descrição</th>
            <th style="width:120px; text-align:center">Qtde (un)</th>
          </tr>
        </thead>
        <tbody>
          ${rowsHtml || `<tr><td colspan="4" class="muted">Nenhum PF com BOM cadastrado.</td></tr>`}
        </tbody>
      </table>
    </div>

    <div class="row wrap" style="gap:10px; margin-top:14px;">
      <button id="btnCalcBatch" type="button" class="btn secondary">Calcular necessidades</button>
      <button id="btnGenOPBatch" type="button" class="btn primary">Gerar Ordem de Produção</button>
      <button id="btnGenOCBatch" type="button" class="btn primary">Gerar Ordem de Compra</button>
    </div>

    <div id="simReqArea" style="margin-top:14px;"></div>
  `;

  const syncSelectedFromDom = () => {
    const cbs = box.querySelectorAll('input[type=checkbox][data-sel]');
    const selected = [];
    cbs.forEach(cb => { if (cb.checked) selected.push(cb.getAttribute('data-sel')); });
    state.simulateSelected = selected;
  };

  const syncQtyFromDom = () => {
    const inputs = box.querySelectorAll('input[data-qty]');
    inputs.forEach(inp => {
      const rid = inp.getAttribute('data-qty');
      let v = parseInt(inp.value, 10);
      if (!Number.isFinite(v) || v < 0) v = 0;
      state.simulateQtyById[rid] = v;
    });
  };

  box.querySelector('#btnSelectAllSim')?.addEventListener('click', () => {
    if (state.simulateLocked) return;
    state.simulateSelected = withBom.map(x => x.recipe.id);
    renderSimulate(boxSelector);
  });

  box.querySelector('#btnClearSim')?.addEventListener('click', () => {
    state.simulateSelected = [];
    state.simulateLocked = false;
    // opcional: não apagamos quantidades por PF; elas voltam ao padrão quando necessário
    renderSimulate(boxSelector);
  });

  // checkbox events
  box.querySelectorAll('input[type=checkbox][data-sel]').forEach(cb => {
    cb.addEventListener('change', () => {
      syncSelectedFromDom();
    });
  });

  // qty events
  box.querySelectorAll('input[data-qty]').forEach(inp => {
    inp.addEventListener('input', () => {
      syncQtyFromDom();
      // convenience: if qty > 0, auto-check
      const rid = inp.getAttribute('data-qty');
      const v = parseInt(inp.value, 10);
      const cb = box.querySelector(`input[type=checkbox][data-sel="${rid}"]`);
      if (cb && Number.isFinite(v) && v > 0) {
        cb.checked = true;
        syncSelectedFromDom();
      }
    });
  });

  const getSelectedPairs = () => {
    syncSelectedFromDom();
    syncQtyFromDom();
    const pairs = [];
    for (const rid of (state.simulateSelected || [])) {
      const q = parseInt(state.simulateQtyById?.[rid] ?? 0, 10);
      if (Number.isFinite(q) && q > 0) pairs.push({ rid, qty: q });
    }
    return pairs;
  };

  const calcBatch = async () => {
    const pairs = getSelectedPairs();
    if (!pairs.length) {
      alert('Selecione pelo menos 1 PF e informe uma quantidade (> 0).');
      return null;
    }

    // aggregate requirements across selected PFs
    const agg = {};
    for (const { rid, qty } of pairs) {
      const data = await api('/api/mrp/requirements', { method: 'POST', body: JSON.stringify({ recipeId: rid, qtyToProduce: qty }) });
      for (const r of (data.requirements || [])) {
        const k = r.itemId;
        if (!agg[k]) {
          agg[k] = {
            itemId: r.itemId,
            code: (r.itemCode ?? r.code ?? ''),
            name: (r.itemName ?? r.name ?? ''),
            unit: r.unit,
            required: 0,
            inStock: r.inStock,
            ok: true,
            shortage: 0,
          };
        }
        agg[k].required += Number(r.required || 0);
      }
    }

    // compute availability with current stock (inStock)
    const list = Object.values(agg).map(x => {
      const inStock = Number(x.inStock || 0);
      const required = Number(x.required || 0);
      const ok = inStock + 1e-9 >= required;
      const shortage = ok ? 0 : Math.max(0, required - inStock);
      return { ...x, ok, shortage };
    }).sort((a,b) => String(a.code||'').localeCompare(String(b.code||'')));

    return { pairs, requirements: list };
  };

  box.querySelector('#btnCalcBatch')?.addEventListener('click', async () => {
    try {
      const res = await calcBatch();
      if (!res) return;
      const { pairs, requirements } = res;
      const hasShortages = requirements.some(r => !r.ok);
      const reqArea = box.querySelector('#simReqArea');
      const header = `<div class="row between wrap" style="margin-bottom:8px;"><div><b>Resultado</b> <span class="muted small">(${pairs.length} PF selecionado(s))</span></div><div class="muted small">${hasShortages ? 'Há itens com falta de estoque.' : 'Estoque OK para todos os itens.'}</div></div>`;

      const rows = requirements.map(r => `
        <tr>
          <td><b>${escapeHtml(r.code||'')}</b></td>
          <td>${escapeHtml(r.name||'')}</td>
          <td style="text-align:center">${escapeHtml(String(r.unit||''))}</td>
          <td style="text-align:right">${fmtQty(r.required)}</td>
          <td style="text-align:right">${fmtQty(r.inStock)}</td>
          <td style="text-align:right">${r.ok ? '-' : `<b>${fmtQty(r.shortage)}</b>`}</td>
        </tr>`).join('');

      reqArea.innerHTML = `
        ${header}
        <div class="table-wrap">
          <table class="table">
            <thead>
              <tr>
                <th style="width:70px">COD</th>
                <th>Descrição</th>
                <th style="width:60px; text-align:center">UN</th>
                <th style="width:110px; text-align:right">Necessário</th>
                <th style="width:110px; text-align:right">Em estoque</th>
                <th style="width:110px; text-align:right">Falta</th>
              </tr>
            </thead>
            <tbody>${rows || `<tr><td colspan="6" class="muted">Sem itens.</td></tr>`}</tbody>
          </table>
        </div>
      `;
    } catch (e) {
      console.error(e);
      alert('Falha ao calcular necessidades. Veja console.');
    }
  });

  let __simInflight = false;
  const setSimDisabled = (v) => {
    ['#btnCalcBatch','#btnGenOPBatch','#btnGenOCBatch','#btnSelectAllSim','#btnClearSim'].forEach(sel => {
      const el = box.querySelector(sel);
      if (el) el.disabled = !!v;
    });
  };

  const applySimLockUI = () => {
    if (!box) return;
    const locked = !!state.simulateLocked;
    // Quando travado, desativamos seleção + quantidades + ações, e mantemos apenas "Limpar seleção" ativo.
    const inputs = box.querySelectorAll('input[type=checkbox][data-sel], input[data-qty]');
    inputs.forEach(el => { el.disabled = locked; });
    ['#btnCalcBatch','#btnGenOPBatch','#btnGenOCBatch','#btnSelectAllSim'].forEach(sel => {
      const el = box.querySelector(sel);
      if (el) el.disabled = locked;
    });
    const btnClear = box.querySelector('#btnClearSim');
    if (btnClear) btnClear.disabled = false;
  };


  box.querySelector('#btnGenOPBatch')?.addEventListener('click', async (ev) => {
    ev.preventDefault();
    ev.stopPropagation();
    try {
      if (__simInflight) return;
      if (state.simulateLocked) {
        alert('Esta simulação está travada para evitar duplicidade. Clique em "Limpar seleção" para destravar.');
        return;
      }
      __simInflight = true;
      setSimDisabled(true);
      const pairs = getSelectedPairs();
      if (!pairs.length) {
        alert('Selecione pelo menos 1 PF e informe uma quantidade (> 0).');
        return;
      }

      // Detecta (antes de gerar) se há faltas. Se houver, o backend criará OC automaticamente.
      let hasShortages = false;
      try {
        const calcRes = await calcBatch();
        if (calcRes && Array.isArray(calcRes.requirements)) {
          hasShortages = calcRes.requirements.some(r => !r.ok && Number(r.shortage || 0) > 0);
        }
      } catch (e) {
        // se falhar o cálculo, seguimos sem bloquear
      }

      const msgPre = hasShortages
        ? 'Existem itens faltantes no estoque. Ao gerar OP, o sistema criará automaticamente OC(s) para os faltantes.\n\nEsta simulação será travada para evitar duplicidade.\n\nContinuar?'
        : 'Ao gerar OP, esta simulação será travada para evitar duplicidade.\n\nContinuar?';
      const okGo = confirm(msgPre);
      if (!okGo) return;

      // create OP per selected PF (maintains current behavior)
      const lines = [];
      for (const { rid, qty } of pairs) {
        lines.push({ recipeId: rid, qtyToProduce: qty });
      }

      // Cria 1 OP por PF selecionado (mantém rastreio simples)
      for (const ln of lines) {
        const label = labelByRid.get(ln.recipeId) || ln.recipeId;
        await api('/api/mrp/production-orders', { method: 'POST', body: JSON.stringify({
          date: new Date().toISOString(),
          recipeId: ln.recipeId,
          qtyToProduce: ln.qtyToProduce,
          note: `Simulação em lote • ${label} • ${ln.qtyToProduce} un`,
          allowInsufficient: true,
        }) });
      }

      // trava a simulação para evitar geração duplicada (OP/OC)
      state.simulateLocked = true;

      await loadAll(); // mantém a aba atual; loadMRP re-renderiza a simulação

      applySimLockUI();
      const msgDone = hasShortages
        ? 'Ordem(ns) de Produção gerada(s). Como há itens faltantes, o sistema criou automaticamente OC(s) para as faltas.\n\nSimulação travada para evitar duplicidade. Use "Limpar seleção" para destravar.'
        : 'Ordem(ns) de Produção gerada(s).\n\nSimulação travada para evitar duplicidade. Use "Limpar seleção" para destravar.';
      alert(msgDone);
    } catch (e) {
      console.error(e);
      alert('Falha ao gerar OP em lote. Veja console.');
    } finally {
      __simInflight = false;
      setSimDisabled(false);
      applySimLockUI();
    }
  });

  box.querySelector('#btnGenOCBatch')?.addEventListener('click', async (ev) => {
    ev.preventDefault();
    ev.stopPropagation();
    try {
      if (__simInflight) return;
      if (state.simulateLocked) {
        alert('Esta simulação está travada para evitar duplicidade. Clique em "Limpar seleção" para destravar.');
        return;
      }
      __simInflight = true;
      setSimDisabled(true);
      const res = await calcBatch();
      if (!res) return;
      const { pairs, requirements } = res;
      const shortages = requirements.filter(r => !r.ok && r.shortage > 0);
      if (!shortages.length) {
        alert('Nenhuma falta de estoque para gerar OC.');
        return;
      }

      const items = shortages.map(s => ({ itemId: s.itemId, qtyOrdered: s.shortage }));
      await api('/api/mrp/purchase-orders', { method: 'POST', body: JSON.stringify({
        date: new Date().toISOString(),
        status: 'open',
        note: `Simulação em lote • ${pairs.length} PF`,
        items,
      }) });

      await loadAll();
      showTab('compras');
      alert('Ordem de Compra gerada com base nas faltas.');
    } catch (e) {
      console.error(e);
      alert('Falha ao gerar OC. Veja console.');
    } finally {
      __simInflight = false;
      setSimDisabled(false);
      applySimLockUI();
    }
  });

  // aplica trava visual (se houver)
  applySimLockUI();
}

function openRecipeEditor({ mode, recipe, productId = "", prefillBomItemId = "" }) {
  // IMPORTANTE:
  // Não cachear a lista de MPs aqui. O usuário pode editar o cadastro de itens e
  // esperamos que o BOM reflita automaticamente a nova descrição/unidade/FC.
  // Por isso, mantemos uma função de refresh que reconstrói rawItems e o índice.
  let rawItems = [];
  let itemById = new Map();
  const refreshRawItems = () => {
    rawItems = (state.rawItems || state.items || [])
      .filter((i) => (i.type || "raw") === "raw")
      .slice()
      .sort((a, b) => String(a.name || "").localeCompare(String(b.name || ""), "pt-BR"));
    itemById = new Map(rawItems.map((it) => [it.id, it]));
  };
  refreshRawItems();

  const buildItemsOptions = () => rawItems
    .map((it) => `<option value="${it.id}">${escapeHtml(it.name)} (${escapeHtml(it.unit)})</option>`)
    .join("");
  const itemsOptions = buildItemsOptions();

  const effectiveProductId = String(productId || recipe?.productId || "");
  const linkedPF = effectiveProductId ? (state.fgItems||[]).find(x=>x.id===effectiveProductId) : null;

  const title = linkedPF ? "BOM do Produto Final" : (mode === "new" ? "Nova receita" : "Editar receita");
  // In BOM do Produto Final, we already show the PF in the form input. Avoid redundant legends.
  const subtitle = linkedPF
    ? ""
    : (mode === "new"
        ? "Monte a lista de ingredientes (BOM) com quantidades na base do rendimento."
        : (recipe?.id || ""));

  // ---- quantity conversion: user inputs g/ml when unit is kg/l ----
  const dispUnit = (u) => {
    const uu = String(u||"").toLowerCase();
    if (uu === "kg") return "g";
    if (uu === "l") return "ml";
    return uu || "";
  };
  const toBaseQty = (inputQty, itemUnit) => {
    const q = Number(inputQty || 0);
    const uu = String(itemUnit||"").toLowerCase();
    if (!Number.isFinite(q)) return 0;
    if (uu === "kg") return q / 1000; // g -> kg
    if (uu === "l") return q / 1000;  // ml -> l
    return q; // un/ml/kg already base or custom
  };
  const fromBaseQty = (baseQty, itemUnit) => {
    const q = Number(baseQty || 0);
    const uu = String(itemUnit||"").toLowerCase();
    if (!Number.isFinite(q)) return 0;
    if (uu === "kg") return q * 1000; // kg -> g
    if (uu === "l") return q * 1000;  // l -> ml
    return q;
  };

  
// editable BOM state (stored in base units: kg/l/un/ml)
// each line: { itemId, qty, fc (override|null), pos (unique int) }
let bomState = Array.isArray(recipe?.bom)
  ? recipe.bom.map((l, idx) => ({
      itemId: l.itemId,
      qty: Number(l.qty || 0),
      fc: (l.fc === undefined ? null : Number(l.fc)),
      pos: (Number.isFinite(Number(l.pos)) && Number(l.pos) > 0) ? Number(l.pos) : (idx + 1),
    }))
  : [];

// helpers for unique POS
const nextFreePos = (excludeIdx = -1, takenAlso = []) => {
  const used = new Set();
  bomState.forEach((l, i) => {
    if (i === excludeIdx) return;
    const p = Number(l.pos);
    if (Number.isFinite(p) && p > 0) used.add(p);
  });
  takenAlso.forEach(p => used.add(Number(p)));
  let p = 1;
  while (used.has(p)) p += 1;
  return p;
};
// allow excluding multiple indices (used for POS conflict handling)
const nextFreePosMulti = (excludeIdxs = [], takenAlso = []) => {
  const ex = new Set(excludeIdxs.filter(x => x !== null && x !== undefined));
  const used = new Set();
  bomState.forEach((l, i) => {
    if (ex.has(i)) return;
    const p = Number(l.pos);
    if (Number.isFinite(p) && p > 0) used.add(p);
  });
  takenAlso.forEach(p => {
    const n = Number(p);
    if (Number.isFinite(n) && n > 0) used.add(n);
  });
  let p = 1;
  while (used.has(p)) p += 1;
  return p;
};


const normalizePositions = () => {
  const used = new Set();
  const ordered = bomState.map((l, idx) => ({ l, idx }))
          .sort((a, b) => Number(a.l.pos||0) - Number(b.l.pos||0));
        for (let i = 0; i < ordered.length; i++) {
    let p = Number(bomState[i].pos);
    if (!Number.isFinite(p) || p <= 0 || used.has(p)) {
      p = nextFreePos(i);
      bomState[i].pos = p;
    }
    used.add(p);
  }
};

// If opening from Cadastro Geral: pre-add MP line (qty 0)
if (prefillBomItemId && !bomState.some(l => l.itemId === prefillBomItemId)) {
  bomState.push({ itemId: prefillBomItemId, qty: 0, fc: null, pos: nextFreePos() });
}
// Não cria linha automaticamente quando o BOM está vazio.
// (Evita "ingrediente fantasma" na 1a abertura e POS começando em 2.)

normalizePositions();

  const yieldUnitDefault = linkedPF ? (linkedPF.unit || "un") : (recipe?.yieldUnit || "un");

  openModal({
    title,
    subtitle,
    submitText: mode === "new" ? "Criar" : "Salvar",
    cardClass: "wide",
    bodyHtml: `
      ${linkedPF ? `
        <div class="bom-header-grid">
          <div class="bom-header-left">
            <div class="pf-line">
              <label class="tight" style="flex:1; min-width:0;">
                <span class="muted small">Produto Final</span>
                <input class="pf-name-input" name="name" readonly required value="${escapeHtml(linkedPF.name || recipe?.name || "")}" />
              </label>
              <div class="pf-inline-actions">
                <button type="button" id="bomExportXlsx" class="btn secondary tiny">Exportar BOM</button>
                <button type="button" id="bomImportXlsx" class="btn secondary tiny">Importar BOM</button>
                <button type="button" id="bomPrintDoc" class="btn secondary tiny">Imprimir</button>
                <button type="button" id="bomPhotoChange" class="btn secondary tiny">Trocar foto</button>
                <button type="button" id="bomPhotoDelete" class="btn danger tiny">Excluir foto</button>
                <input id="bomXlsxFile" type="file" accept=".xlsx,application/vnd.openxmlformats-officedocument.spreadsheetml.sheet" style="display:none" />
                <input id="bomPhotoFile" type="file" accept=".jpg,.jpeg,.png,.webp,image/jpeg,image/png,image/webp" style="display:none" />
              </div>
            </div>

            <div class="row between" style="margin:8px 0 6px">
              <div class="pf-pick-title">Ingredientes (BOM)</div>
              <div class="muted small" style="text-align:right">Digite em <b>g</b> (itens em kg) e <b>ml</b> (itens em l). (un) mantém.</div>
            </div>

            <div class="bom-comp-editor bom-comp-editor-compact">
              <div class="grid bom-comp-grid">
                <div class="field">
                  <label>POS</label>
                  <input id="bomPos" class="input" type="number" min="1" step="1" value="1" />
                </div>
                <div class="field">
                  <label>Item (MP)</label>
                  <select id="bomItemSel" class="input"><option value="">Selecione...</option>${itemsOptions}</select>
                </div>
                <div class="field">
                  <label>UN Compra</label>
                  <input id="bomUnitBuy" class="input" readonly value="" />
                </div>
                <div class="field">
                  <label>UN Consumo</label>
                  <input id="bomUnitUse" class="input" readonly value="" />
                </div>
                <div class="field">
                  <label>FC <button type="button" id="bomFcHelp" class="link-btn">Saiba mais</button></label>
                  <input id="bomFC" class="input" value="" placeholder="Ex.: 2,8" />
                </div>
                <div class="field">
                  <label id="bomQtyLabel">Quantidade</label>
                  <input id="bomQtyInput" class="input" type="number" step="0.001" placeholder="0" />
                </div>
                <div class="field bom-comp-actions">
                  <label>&nbsp;</label>
                  <div class="stack">
                    <button type="button" id="bomSaveLine" class="btn primary small">Salvar linha</button>
                    <button type="button" id="bomClearLine" class="btn secondary small">Limpar</button>
                  </div>
                </div>
              </div>
              <div class="row" style="justify-content:flex-end; margin-top:6px">
                <div id="bomCookPreview" class="muted small"></div>
              </div>
              <div id="bomEditHint" class="muted small" style="margin-top:6px;"></div>
            </div>


      <div class="bom-editor">
<div class="table-wrap bom-comp-table" style="margin-top:8px; max-height: 220px;">
          <table class="table">
            <thead>
              <tr>
                <th style="text-align:center">POS</th>
                <th style="text-align:left">COD</th>
                <th style="text-align:left">Descrição</th>
                <th style="text-align:center">UN Compra</th>
                <th style="text-align:center">UN Consumo</th>
                <th style="text-align:center">FC</th>
                <th style="text-align:center">QTE (crua)</th>
                <th style="text-align:center">QTE (cozida)</th>
                <th style="text-align:center">Ações</th>
              </tr>
            </thead>
            <tbody id="bomCompTbody"></tbody>
          </table>
        </div>

        <div class="muted small" style="margin-top:8px">Dica: inclua também “Embalagem (un)” se você quiser controlar.</div>
      </div>


          </div>

          <div class="bom-header-photo">
            <div id="bomPhotoFrame" class="photo-frame photo-frame-sm"></div>
<div class="muted small bom-photo-help">Foto 1:1 • JPG/PNG/WEBP (até 8MB)</div>
          </div>
        </div>
        <input type="hidden" name="productId" value="${escapeHtml(effectiveProductId)}" />

        
      ` : `
        <label>Nome da receita<input name="name" required value="${escapeHtml(recipe?.name || "")}" /></label>
      `}

      ${linkedPF ? "" : `
      <div class="grid two">
        <label>Rendimento (quantidade)
          <input name="yieldQty" type="number" step="0.001" required value="${Number(recipe?.yieldQty || 1)}" />
        </label>
        <label>Unidade do rendimento (un, kg, porção...)
          <input name="yieldUnit" required value="${escapeHtml(yieldUnitDefault)}" />
        </label>
      </div>
      `}    `,
    onOpen: () => {
      const sel = $("#bomItemSel");

      // Garante que as MPs e o mapa (id -> item) estejam sempre atualizados
      // (ex.: depois de editar a descrição no cadastro de estoque)
      refreshRawItems();
      const prevSel = sel.value;
      sel.innerHTML = `
        <option value="">Selecione...</option>
        ${buildItemsOptions()}
      `;
      if (prevSel) sel.value = prevSel;

      const posInput = $("#bomPos");
      const unitBuy = $("#bomUnitBuy");
      const unitUse = $("#bomUnitUse");
      const fcInput = $("#bomFC");
      const qtyInput = $("#bomQtyInput");
      const cookPrev = $("#bomCookPreview");
      const qtyLabel = $("#bomQtyLabel");
      const saveBtn = $("#bomSaveLine");
      const clearBtn = $("#bomClearLine");
      const tbody = $("#bomCompTbody");

      // --- XLSX Import/Export do BOM (por PF) ---
      const bomExportXlsx = $("#bomExportXlsx");
      const bomImportXlsx = $("#bomImportXlsx");
      const bomPrintDoc = $("#bomPrintDoc");
      const bomXlsxFile = $("#bomXlsxFile");
      // Export precisa de receita (id). Import pode criar/atualizar por Produto Final.
      if (bomExportXlsx) bomExportXlsx.disabled = !recipe?.id;
      if (bomImportXlsx) bomImportXlsx.disabled = false;
      if (bomXlsxFile) bomXlsxFile.disabled = false;

      if (bomExportXlsx && recipe?.id){
        bomExportXlsx.addEventListener("click", async () => {
          await downloadAfterReauth(`/api/mrp/recipes/${recipe.id}/bom.xlsx`, "Exportar BOM (XLSX)");
        });
      }
      if (bomImportXlsx && bomXlsxFile){
        bomImportXlsx.addEventListener("click", () => bomXlsxFile.click());
        bomXlsxFile.addEventListener("change", async () => {
          const file = bomXlsxFile.files && bomXlsxFile.files[0];
          if (!file) return;
          const okAuth = await requestReauth({ reason: "Importar BOM (XLSX)" });
          if (!okAuth) { bomXlsxFile.value = ""; return; }
          try{
            const fd = new FormData();
            fd.append("file", file);
            const uploadUrl = recipe?.id ? `/api/mrp/recipes/${recipe.id}/bom.xlsx` : `/api/mrp/pf/${effectiveProductId}/bom.xlsx`;
            await apiUpload(uploadUrl, fd);
            await loadMRP();
            const updated = recipe?.id
              ? state.recipes.find((r) => r.id === recipe.id)
              : state.recipes.find((r) => String(r.productId||'') === String(effectiveProductId));
            if (updated) {
              // reabrir com bom atualizado (simplifica)
              modal.close();
              openRecipeEditor({ mode: "edit", recipe: updated, productId: updated.productId || effectiveProductId });
            } else {
              alert("BOM importado.");
            }
          } catch(e){
            console.error(e);
            alert("Falha ao importar BOM (.xlsx). Confira as colunas (POS, COD, FC, QTE_CRUA...).");
          } finally {
            bomXlsxFile.value = "";
          }
        });
      }

      // --- Imprimir (BOM + Método) ---
      if (bomPrintDoc){
        bomPrintDoc.addEventListener("click", () => {
          const pf = linkedPF || null;
          const cur = (state.recipes || []).find((r) => r.id === recipe?.id) || recipe;
          const methodText = String(cur?.method || "");
          // Print current editor state (including unsaved changes)
          const bomLines = (bomState || []).map(l => ({
            itemId: l.itemId,
            qty: Number(l.qty || 0),
            pos: Number(l.pos || 0) || 0,
            fc: (l.fc === null || l.fc === undefined) ? null : Number(l.fc),
          })).filter(l => l.itemId && Number.isFinite(l.qty) && l.qty > 0);

          printBomAndMethodDocument({
            pfCode: pf?.code || "",
            pfName: pf?.name || "",
            bomLines,
            itemById,
            methodText,
          });
        });
      }

      // --- Foto 1:1 do PF (salva em /bd/photos) ---
      const bomPhotoFrame = $("#bomPhotoFrame");
      const bomPhotoChange = $("#bomPhotoChange");
      const bomPhotoDelete = $("#bomPhotoDelete");
      const bomPhotoFile = $("#bomPhotoFile");

      const renderBomPhoto = () => {
        if (!bomPhotoFrame) return;
        const cur = state.recipes.find((r) => r.id === recipe.id) || recipe;
        const pf = cur?.photoFile;
        if (pf) {
          bomPhotoFrame.innerHTML = `<img alt="Foto" src="/photos/${encodeURIComponent(pf)}?t=${Date.now()}" />`;
        } else {
          bomPhotoFrame.innerHTML = `<div class="photo-placeholder">+</div>`;
        }
        if (bomPhotoDelete) bomPhotoDelete.disabled = !pf;
        if (bomPhotoChange) bomPhotoChange.textContent = pf ? "Trocar foto" : "Adicionar/Selecionar";
        if (bomPhotoChange) bomPhotoChange.disabled = !recipe?.id;
        if (bomPhotoFrame) bomPhotoFrame.style.opacity = recipe?.id ? '1' : '0.7';

      };
      renderBomPhoto();

      const triggerPickBomPhoto = () => {
        if (!recipe?.id){
          alert('Salve/crie o BOM deste PF antes de adicionar uma foto.');
          return;
        }
        bomPhotoFile && bomPhotoFile.click();
      };
      if (bomPhotoFrame) bomPhotoFrame.addEventListener("click", () => {
        const cur = state.recipes.find((r) => r.id === recipe.id) || recipe;
        if (!cur?.photoFile) triggerPickBomPhoto();
      });
      if (bomPhotoChange) bomPhotoChange.addEventListener("click", triggerPickBomPhoto);

      if (bomPhotoFile){
        bomPhotoFile.addEventListener("change", async () => {
          const file = bomPhotoFile.files && bomPhotoFile.files[0];
          if (!file) return;
          if (!recipe?.id) { alert('Antes de adicionar foto, salve/crie o BOM deste PF.'); bomPhotoFile.value = ''; return; }
          if (!recipe?.id){ alert('Salve/crie o BOM deste PF antes de enviar foto.'); bomPhotoFile.value = ''; return; }
          try{
            if (file.size > 8 * 1024 * 1024){ alert("Arquivo muito grande (máx 8MB)."); return; }
            const extOk = /\.(jpe?g|png|webp)$/i.test(file.name) || ["image/jpeg","image/png","image/webp"].includes(file.type);
            if (!extOk){ alert("Formato inválido. Use JPG/PNG/WEBP."); return; }
            const fd = new FormData();
            fd.append("file", file);
            await apiUpload(`/api/mrp/recipes/${recipe.id}/photo`, fd);
            await loadMRP();
            renderBomPhoto();
          } catch(e){
            console.error(e);
            alert("Falha ao enviar foto.");
          } finally {
            bomPhotoFile.value = "";
          }
        });
      }
      if (bomPhotoDelete){
        bomPhotoDelete.addEventListener("click", async () => {
          if (!recipe?.id){ alert('Salve/crie o BOM deste PF antes de excluir foto.'); return; }
          const ok = confirm("Excluir a foto deste Produto Final?");
          if (!ok) return;
          if (!recipe?.id){ alert('Este PF ainda não tem BOM salvo.'); return; }
          try{
            await api(`/api/mrp/recipes/${recipe.id}/photo`, { method: "DELETE" });
            await loadMRP();
            renderBomPhoto();
          } catch(e){
            console.error(e);
            alert("Falha ao excluir foto.");
          }
        });
      }
      const hint = $("#bomEditHint");

      let editingIndex = -1;
      let posTouched = false; // só considera conflito quando o usuário digita a POS

      // itemById (Map) é mantido/atualizado fora do onOpen via refreshRawItems().

      const getFC = (it) => {
        const v = Number(it?.cookFactor ?? it?.fc ?? 1);
        return Number.isFinite(v) && v > 0 ? v : 1;
      };

      const updateCookPreview = () => {
        const it = itemById.get(sel.value);
        if (!it) return;
        const fcTyped = Number(String(fcInput.value || "").replace(",", "."));
        const fc = (Number.isFinite(fcTyped) && fcTyped > 0) ? fcTyped : getFC(it);
        const qtyIn = Number(qtyInput.value);
        const du = dispUnit(it.unit);
        if (!Number.isFinite(qtyIn) || qtyIn <= 0) {
          cookPrev.textContent = fc !== 1 ? `Após cocção (FC ${fmt(fc)}): —` : "";
          return;
        }
        const cookedDisp = qtyIn * fc;
        cookPrev.textContent = (fc === 1) ? "" : `Após cocção (FC ${fmt(fc)}): ${fmt(cookedDisp)} ${du || ""}`;
      };

      const refreshEditorForItem = (itemId, qtyBase = 0, fcOverride = null) => {
        const it = itemById.get(itemId) || rawItems[0] || null;
        if (!it) return;
        sel.value = it.id;
        unitBuy.value = it.unit || "";
        unitUse.value = dispUnit(it.unit) || (it.unit || "");
        const defFc = getFC(it);
        const fcUse = (fcOverride !== null && fcOverride !== undefined && Number.isFinite(Number(fcOverride)) && Number(fcOverride) > 0) ? Number(fcOverride) : defFc;
        fcInput.value = fmt(fcUse);
        const du = dispUnit(it.unit);
        qtyLabel.textContent = linkedPF ? `Quantidade (${du || "—"}) (por 1 unidade do PF)` : `Quantidade (${du || "—"}) (na base do rendimento)`;
        qtyInput.placeholder = du === "g" ? "Ex.: 150" : (du === "ml" ? "Ex.: 100" : "Ex.: 1");
        qtyInput.value = (qtyBase ? fromBaseQty(qtyBase, it.unit) : "");
        updateCookPreview();
      };

      const setDefaultPos = () => {
        if (!posInput) return;
        if (editingIndex >= 0 && bomState[editingIndex]) {
          posInput.value = String(bomState[editingIndex].pos ?? 1);
        } else {
          // Novo item: sempre vai para a próxima posição disponível
          posInput.value = String(nextFreePos(-1));
        }
        posTouched = false;
      };

      const renderTable = () => {
        tbody.innerHTML = "";
        if (!bomState.length) {
          const tr = document.createElement("tr");
          tr.innerHTML = `<td colspan="9" class="muted">Sem ingredientes ainda. Use o editor acima para adicionar.</td>`;
          tbody.appendChild(tr);
          return;
        }

        const ordered = bomState.map((l, idx) => ({ l, idx }))
          .sort((a, b) => Number(a.l.pos||0) - Number(b.l.pos||0));
        for (let i = 0; i < ordered.length; i++) {
          const l = ordered[i].l;
          const realIdx = ordered[i].idx;
          const it = itemById.get(l.itemId);
          if (!it) continue;

          const du = dispUnit(it.unit);
          const qDisp = fromBaseQty(l.qty, it.unit);

          const tr = document.createElement("tr");
          const fcUsed = (l.fc !== null && l.fc !== undefined && Number.isFinite(Number(l.fc)) && Number(l.fc) > 0) ? Number(l.fc) : getFC(it);
          const cookedDisp = qDisp * fcUsed;

          tr.innerHTML = `
            <td style="text-align:center">${escapeHtml(String(l.pos || (i+1)))}</td>
            <td style="text-align:left"><b>${escapeHtml(it.code || "MP")}</b></td>
            <td style="text-align:left">${escapeHtml(it.name || "")}</td>
            <td style="text-align:center">${escapeHtml(it.unit || "")}</td>
            <td style="text-align:center">${escapeHtml(du || "")}</td>
            <td style="text-align:center">${fmt(fcUsed)}</td>
            <td style="text-align:center">${fmt(qDisp)} ${escapeHtml(du || "")}</td>
            <td style="text-align:center">${fmt(cookedDisp)} ${escapeHtml(du || "")}</td>
            <td style="text-align:center"><div class="bom-row-actions">
              <button type="button" class="btn secondary small" data-edit="${realIdx}">Editar</button>
              <button type="button" class="btn secondary small" data-del="${realIdx}">Excluir</button>
            </div></td>
          `;
          tbody.appendChild(tr);
        }

        tbody.querySelectorAll("button[data-edit]").forEach(btn => {
          btn.addEventListener("click", () => {
            const idx = Number(btn.dataset.edit);
            const l = bomState[idx];
            const it = itemById.get(l.itemId);
            if (!it) return;
            editingIndex = idx;
            posTouched = false;
            if (posInput) posInput.value = String(l.pos || 1);
            refreshEditorForItem(l.itemId, l.qty, l.fc);
            hint.innerHTML = `Editando: <b>${escapeHtml(it.code || "MP")}</b> — ${escapeHtml(it.name || "")}.`;
            saveBtn.textContent = "Atualizar linha";
          });
        });

        tbody.querySelectorAll("button[data-del]").forEach(btn => {
          btn.addEventListener("click", () => {
            const idx = Number(btn.dataset.del);
            bomState.splice(idx, 1);
            // exit edit if needed
            if (editingIndex === idx) {
              editingIndex = -1;
              hint.textContent = "";
              saveBtn.textContent = "Salvar linha";
              qtyInput.value = "";
            }
            renderTable();
          });
        });
      };

      const clearEditor = () => {
        refreshRawItems();
        editingIndex = -1;
        posTouched = false;
        hint.textContent = "";
        saveBtn.textContent = "Salvar linha";

        // POS: ao limpar, já sugere a próxima posição disponível
        setDefaultPos();

        // limpa seleção e campos
        sel.value = "";
        unitBuy.value = "";
        unitUse.value = "";
        fcInput.value = "";
        qtyInput.value = "";
        qtyLabel.textContent = linkedPF ? "Quantidade (—) (por 1 unidade do PF)" : "Quantidade (—) (na base do rendimento)";
        // reset preview (guard: element id is bomCookPreview)
        if (cookPrev) cookPrev.textContent = "Após cocção: —";
      };

      sel.addEventListener("change", () => {
        const it = itemById.get(sel.value);
        if (!it) return;
        unitBuy.value = it.unit || "";
        unitUse.value = dispUnit(it.unit) || (it.unit || "");
        // ao trocar o item, volta para o FC padrão do cadastro
        fcInput.value = fmt(getFC(it));
        const du = dispUnit(it.unit);
        qtyLabel.textContent = linkedPF ? `Quantidade (${du || "—"}) (por 1 unidade do PF)` : `Quantidade (${du || "—"}) (na base do rendimento)`;
        // POS: para novo item, sempre sugere a próxima posição disponível (a não ser que o usuário tenha digitado)
        if (editingIndex < 0 && !posTouched) setDefaultPos();
        updateCookPreview();
      });

      // marcar quando o usuário digita POS (só aí aplicamos regra de substituição)
      posInput.addEventListener("input", () => {
        posTouched = true;
      });

      qtyInput.addEventListener("input", () => updateCookPreview());

      fcInput.addEventListener("input", () => updateCookPreview());

      const fcHelpBtn = $("#bomFcHelp");
      if (fcHelpBtn) fcHelpBtn.addEventListener("click", () => window.open("help/fc.html", "_blank"));

      saveBtn.addEventListener("click", async (e) => {
      e.preventDefault();
        const it = itemById.get(sel.value);
        if (!it) return;

        const qtyIn = Number(qtyInput.value);
        if (!Number.isFinite(qtyIn) || qtyIn <= 0) {
          alert("Informe uma quantidade maior que zero.");
          return;
        }

        const qtyBase = toBaseQty(qtyIn, it.unit);

        const defFc = getFC(it);
        const fcTyped = Number(String(fcInput.value || "").replace(",", "."));
        const fcNew = (Number.isFinite(fcTyped) && fcTyped > 0) ? fcTyped : defFc;

        let fcOverride = null;

        // Se digitou FC diferente do cadastro, pergunta se quer atualizar o cadastro
        if (Math.abs(fcNew - defFc) > 1e-9) {
          const ok = confirm(`Você digitou um FC diferente do cadastro (${fmt(defFc)} → ${fmt(fcNew)}).\n\nDeseja atualizar o FC do item no cadastro?`);
          if (ok) {
            try {
              await api(`/api/inventory/items/${it.id}?type=raw`, { method: "PUT", body: JSON.stringify({ cookFactor: fcNew }) });
              // atualiza cache local
              it.cookFactor = fcNew;
            } catch (e) {
              console.warn("Falha ao atualizar FC no cadastro:", e);
              alert("Não foi possível atualizar o FC no cadastro. Vou manter o FC apenas neste BOM.");
              fcOverride = fcNew;
            }
          } else {
            fcOverride = fcNew;
            hint.innerHTML = `FC salvo apenas neste BOM. Se ficou em dúvida, clique em <b>Saiba mais</b>.`;
          }
        }

        // se igual ao cadastro, não precisa salvar override
        if (fcOverride !== null && Math.abs(fcOverride - getFC(it)) <= 1e-9) {
          fcOverride = null;
        }

        // decide which line we are updating
        let targetIdx = editingIndex;
        if (targetIdx < 0) {
          const existingIdx = bomState.findIndex(x => x.itemId === it.id);
          targetIdx = existingIdx;
        }

        // POS handling (regra: nova linha sempre vai para a próxima POS disponível;
        // só substitui/move quando o usuário DIGITA a POS)
        let desiredPos;
        if (!posTouched) {
          if (targetIdx >= 0 && bomState[targetIdx]) desiredPos = Number(bomState[targetIdx].pos) || 1;
          else desiredPos = nextFreePos(-1);
          if (posInput) posInput.value = String(desiredPos);
        } else {
          desiredPos = posInput ? Number(posInput.value) : NaN;
        }

        if (!Number.isFinite(desiredPos) || desiredPos <= 0) {
          alert("Informe uma POS válida (1, 2, 3...).");
          return;
        }

        // Se há conflito de POS, só resolve se o usuário digitou a POS
        const conflictIdx = bomState.findIndex((l, idx) => idx !== targetIdx && Number(l.pos) === desiredPos);
        if (posTouched && conflictIdx >= 0) {
          const conflictIt = itemById.get(bomState[conflictIdx].itemId);
          const movedTo = nextFreePosMulti([conflictIdx, targetIdx], [desiredPos]);
          const okMove = confirm(`A posição ${desiredPos} já está preenchida com o item ${conflictIt?.code || ""} ${conflictIt?.name || ""}.

Se você confirmar, ele será movido para a próxima posição disponível (${movedTo}).`);
          if (!okMove) return;
          bomState[conflictIdx].pos = movedTo;
        }

        // apply changes to target line
        if (targetIdx >= 0) {
          bomState[targetIdx] = { ...bomState[targetIdx], itemId: it.id, qty: qtyBase, fc: fcOverride, pos: desiredPos };
        } else {
          bomState.push({ itemId: it.id, qty: qtyBase, fc: fcOverride, pos: desiredPos });
        }

        normalizePositions();
        clearEditor();
        renderTable();
      });

      clearBtn.addEventListener("click", (e) => {
      e.preventDefault();
      clearEditor();
    });

  

    // expõe um rerender leve para outras telas (ex: Cadastro) atualizarem rótulos/descrições
    state._rerenderRecipeEditor = () => {
      try {
        refreshRawItems();
        const prev = sel.value;
        sel.innerHTML = `
          <option value="">Selecione...</option>
          ${buildItemsOptions()}
        `;
        if (prev) sel.value = prev;
      } catch (e) { /* noop */ }
      renderTable();
      updateCookPreview();
    };
    // initial editor + table
      renderTable();
      clearEditor();
    },
    onSubmit: async (fd) => {
      const name = fd.get("name");
      const yieldQty = linkedPF ? 1 : Number(fd.get("yieldQty") || 1);
      const yieldUnit = linkedPF ? "un" : String(fd.get("yieldUnit") || "un");
      const notes = fd.get("notes");

      
const lines = bomState
  .map((l) => {
    const o = { itemId: l.itemId, qty: Number(l.qty || 0) };
    const p = Number(l.pos);
    if (Number.isFinite(p) && p > 0) o.pos = p;
    if (l.fc !== null && l.fc !== undefined && Number.isFinite(Number(l.fc)) && Number(l.fc) > 0) o.fc = Number(l.fc);
    return o;
  })
        .filter((l) => l.itemId && Number.isFinite(l.qty) && l.qty > 0);

      if (lines.length === 0) {
        alert("Adicione pelo menos 1 ingrediente (BOM). Use “Salvar linha”.");
        return false;
      }

      const productId2 = fd.get("productId") || "";
      const payload = { name, productId: productId2, yieldQty, yieldUnit, notes, bom: lines };

      if (mode === "new") {
        await api("/api/mrp/recipes", { method: "POST", body: JSON.stringify(payload) });
      } else {
        await api(`/api/mrp/recipes/${recipe.id}`, { method: "PUT", body: JSON.stringify(payload) });
      }

      await loadMRP();
      return true;
    },
  });
}


// ---------- OPs UI ----------
function renderOps() {
  const tbody = $("#opsTable tbody");
  tbody.innerHTML = "";

  if (!Array.isArray(state.opSelectedIds)) state.opSelectedIds = [];
  const selSet = new Set(state.opSelectedIds);

  const itemsById = new Map((state.rawItems || state.items || []).map((i) => [i.id, i]));

  const normOpStatus = (x) => {
    const s = String(x || "ISSUED").toUpperCase();
    if (s === "READY") return "ISSUED";     // compat
    if (s === "HOLD") return "ISSUED";      // remove "Em espera"
    if (s === "EXECUTED") return "CLOSED";   // compat
    return s;
  };

  const statusSelect = (st, id) => {
    const s = normOpStatus(st);
    const isFinal = (s === "CLOSED" || s === "CANCELLED");
    const isIssued = (s === "ISSUED");
    const isInProd = (s === "IN_PRODUCTION");

    // Regras de transição (UX):
    // - ISSUED -> IN_PRODUCTION (baixa no estoque)
    // - IN_PRODUCTION -> CLOSED (entrada do PF no estoque)
    // - CANCELLED somente antes de iniciar produção
    const disabledAttr = isFinal ? "disabled" : "";
    const optIssued = `<option value="ISSUED" ${isIssued ? "selected" : ""} ${(isInProd || isFinal) ? "disabled" : ""}>${OP_STATUS_PT.ISSUED}</option>`;
    const optInProd = `<option value="IN_PRODUCTION" ${isInProd ? "selected" : ""} ${(!(isIssued || isInProd)) ? "disabled" : ""}>${OP_STATUS_PT.IN_PRODUCTION}</option>`;
    const optClosed = `<option value="CLOSED" ${s === "CLOSED" ? "selected" : ""} ${(!(isInProd || s === "CLOSED")) ? "disabled" : ""}>${OP_STATUS_PT.CLOSED}</option>`;
    const optCancelled = `<option value="CANCELLED" ${s === "CANCELLED" ? "selected" : ""} ${((isInProd || s === "CLOSED")) ? "disabled" : ""}>${OP_STATUS_PT.CANCELLED}</option>`;

    return `<select class="input" data-opstatus="${escapeHtml(id)}" style="min-width:160px" ${disabledAttr}>
      ${optIssued}
      ${optInProd}
      ${optClosed}
      ${optCancelled}
    </select>`;
  };


  for (const op of state.ops) {
    const opNo = op.number ? pad6(op.number) : "—";
    const canLabels = normOpStatus(op?.status) === "CLOSED";
    const _lines = (Array.isArray(op.consumed) && op.consumed.length) ? op.consumed : (op.planned?.consumed || []);
    const consumption = (_lines || []).map((c) => {
      const it = itemsById.get(c.itemId);
      return `${escapeHtml(it?.name || "—")}: ${fmt(c.qty)} ${escapeHtml(it?.unit || "")}`;
    }).join("<br/>");

    const missing = (op.shortages || []).map((s) => {
      const it = itemsById.get(s.itemId);
      return `${escapeHtml(it?.name || "—")}: ${fmt(s.shortage)} ${escapeHtml(it?.unit || "")}`;
    }).join("<br/>");

    const st = normOpStatus(op.status);

    const canArchive = (st === "CLOSED" || st === "CANCELLED");

    // UX: quando a OP já está em status final (Encerrada/Cancelada), o select fica desabilitado.
    // Para não parecer "travado", exibimos um pill claro.
    const statusCell = (() => {
      const isFinal = (st === "CLOSED" || st === "CANCELLED");
      if (!isFinal) return statusSelect(st, op.id);
      const cls = (st === "CANCELLED") ? "pill bad" : "pill ok";
      const label = (OP_STATUS_PT && OP_STATUS_PT[st]) ? OP_STATUS_PT[st] : st;
      return `<span class="${cls}" title="Status final (não editável)">${escapeHtml(String(label))}</span>`;
    })();

    // Mini-preview do lote/etiqueta (1 círculo), para rastreabilidade na lista.
    const lotCode = String(op.lotCode || (op.lotNumber ? (`LOTE${String(Math.trunc(Number(op.lotNumber))).padStart(5,"0")}`) : "")).trim();
    const dtIso = op.lotCreatedAt || op.closedAt || op.executedAt || op.createdAt || new Date().toISOString();
    const dt = new Date(dtIso).toLocaleString("pt-BR");
    const recipeCode = String(op.recipeCode || "").trim();
    const recipeName = String(op.recipeName || "").trim();
    const barcodeValue = String(op.barcodeValue || "").trim()
      || (recipeCode && lotCode ? `DIETON-${recipeCode}-${lotCode}` : (lotCode ? `DIETON-${lotCode}` : ""));
    const lotPreview = (canLabels && lotCode && barcodeValue) ? `
      <div class="op-lot-preview" title="${escapeHtml(barcodeValue)}">
        <div class="op-lot-circle">
          <div class="code">${escapeHtml(recipeCode || "PF")}</div>
          <div class="name">${escapeHtml(recipeName || "")}</div>
          <img class="barcode" src="/api/barcode?text=${encodeURIComponent(barcodeValue)}&t=${encodeURIComponent(lotCode)}" alt="barcode" />
          <div class="human">${escapeHtml(barcodeValue)}</div>
          <div class="meta">
            <span>${escapeHtml(lotCode)}</span>
            <span>${escapeHtml(dt)}</span>
          </div>
        </div>
      </div>
    ` : "";

    const actions = `
  <div class="row wrap" style="gap:6px; justify-content:center">
    <button class="btn secondary small" data-act="save" data-id="${escapeHtml(op.id)}">Salvar</button>
    <button class="btn secondary small" data-act="del" data-id="${escapeHtml(op.id)}">Excluir</button>
    <button class="btn secondary small" data-act="print" data-id="${escapeHtml(op.id)}">Imprimir</button>
    <button class="btn secondary small" data-act="labels" data-id="${escapeHtml(op.id)}" ${canLabels ? "" : "disabled"}>Etiquetas</button>
    <button class="btn secondary small" data-act="why" data-id="${escapeHtml(op.id)}">Ver faltas</button>
    ${op.linkedPurchaseOrderId ? `<button class="btn secondary small" data-act="openpo" data-po="${escapeHtml(op.linkedPurchaseOrderId)}">Ver OC</button>` : ""}
    ${canArchive ? `<button class="btn secondary small" data-act="archive" data-id="${escapeHtml(op.id)}">Arquivar</button>` : ""}
  </div>
  ${lotPreview}
`;
    const tr = document.createElement("tr");
    tr.innerHTML = `
      <td style="text-align:center"><input type="checkbox" data-opsel="${escapeHtml(op.id)}" ${selSet.has(op.id)?'checked':''} /></td>
      <td>${fmtDate(op.createdAt)}</td>
      <td><b>${escapeHtml(opNo)}</b></td>
      <td><b>${escapeHtml(op.recipeName || op.recipeId)}</b></td>
      <td>${fmt(op.qtyToProduce)}</td>
      <td>${statusCell}</td>
      <td class="muted">${escapeHtml(op.note || "")}</td>
      <td class="muted small">${consumption || (missing ? `<span class="pill bad">faltando</span><br/>${missing}` : "—")}</td>
      <td style="text-align:center">${actions}</td>
    `;
    tbody.appendChild(tr);
  }

  if (state.ops.length === 0) {
    const tr = document.createElement("tr");
    tr.innerHTML = `<td colspan="9" class="muted">Sem ordens de produção ainda.</td>`;
    tbody.appendChild(tr);
  }

  // select all checkbox
  const selAll = document.querySelector('#opSelectAll');
  if (selAll) {
    const allIds = (state.ops || []).map(o => o.id);
    const allChecked = allIds.length && allIds.every(id => selSet.has(id));
    selAll.checked = !!allChecked;
    selAll.onchange = () => {
      if (selAll.checked) state.opSelectedIds = allIds.slice();
      else state.opSelectedIds = [];
      renderOps();
    };
  }

  // actions
  // selection per row
  tbody.querySelectorAll('input[type=checkbox][data-opsel]').forEach(cb => {
    cb.addEventListener('change', () => {
      const id = cb.getAttribute('data-opsel');
      const set = new Set(state.opSelectedIds || []);
      if (cb.checked) set.add(id);
      else set.delete(id);
      state.opSelectedIds = Array.from(set);
      const selAll = document.querySelector('#opSelectAll');
      if (selAll) {
        const allIds = (state.ops || []).map(o => o.id);
        selAll.checked = allIds.length && allIds.every(x => set.has(x));
      }
    });
  });

  tbody.querySelectorAll("button[data-act]").forEach(btn => {
    btn.addEventListener("click", async () => {
      const act = btn.dataset.act;
      if (act === "labels") {
        const op = state.ops.find(x => x.id === btn.dataset.id);
        if (!op) return;
        if (normOpStatus(op.status) !== "CLOSED") {
          alert("As etiquetas só ficam disponíveis quando a OP estiver ENCERRADA.");
          return;
        }
        printProductionLabels(op);
        return;
      }
      if (act === "why") {

const op = state.ops.find(x => x.id === btn.dataset.id);
if (!op) return;
const itemsById2 = new Map((state.rawItems || []).map(i => [String(i.id), i]));

let shortages = [];
try {
  // Prefer calcular ao vivo (estoque atual) usando o endpoint de requirements
  if (op.recipeId && Number(op.qtyToProduce) > 0) {
    const rr = await api("/api/mrp/requirements", {
      method: "POST",
      body: JSON.stringify({ recipeId: op.recipeId, qtyToProduce: op.qtyToProduce })
    });
    const reqs = Array.isArray(rr.requirements) ? rr.requirements : [];
    shortages = reqs
      .filter(r => Number(r.shortage || 0) > 0)
      .map(r => ({
        itemId: String(r.itemId),
        itemName: r.itemName,
        unit: r.unit,
        required: r.required,
        available: r.available,
        shortage: r.shortage,
      }));
  } else if (Array.isArray(op.shortages)) {
    shortages = op.shortages.map(s => ({
      itemId: String(s.itemId),
      required: s.required,
      available: s.available,
      shortage: s.shortage,
    }));
  }
} catch (e) {
  // fallback: usa o que tiver armazenado na OP
  shortages = Array.isArray(op.shortages) ? op.shortages : [];
}

const rows = (shortages || []).map(s => {
  const it = itemsById2.get(String(s.itemId));
  const name = s.itemName || it?.name || "—";
  const unit = s.unit || it?.unit || "";
  return `<tr>
    <td>${escapeHtml(name)}</td>
    <td>${fmt(s.required)}</td>
    <td>${fmt(s.available)}</td>
    <td>${fmt(s.shortage)} ${escapeHtml(unit)}</td>
  </tr>`;
}).join("");

openModal({
  title: "Faltas da OP",
  subtitle: `OP ${escapeHtml(op.number ? pad6(op.number) : op.id)} • ${escapeHtml(op.recipeName||"")}`,
  submitText: "Fechar",
  bodyHtml: `
    <div class="muted small" style="margin-bottom:8px">Baseado no estoque atual.</div>
    <div class="table-wrap">
      <table class="table">
        <thead><tr><th>Item</th><th>Necessário</th><th>Disponível</th><th>Faltante</th></tr></thead>
        <tbody>${rows || `<tr><td colspan="4" class="muted">Sem faltas.</td></tr>`}</tbody>
      </table>
    </div>
  `,
  onSubmit: () => true
});
      } else if (act === "openpo") {
        const po = state.purchaseOrders.find(x => x.id === btn.dataset.po);
        if (po) openViewPurchaseOrder(po);
      } else if (act === "print") {
        const op = state.ops.find(x => x.id === btn.dataset.id);
        if (op) printProductionOrder(op);
      } else if (act === "archive") {
        const op = state.ops.find(x => x.id === btn.dataset.id);
        if (!op) return;
        const stRaw = String(op.status || "").toUpperCase();
        const st = (stRaw === "EXECUTED") ? "CLOSED" : stRaw;
        if (!(st === "CLOSED" || st === "CANCELLED")) {
          alert("Só é possível arquivar OP Encerrada ou Cancelada.");
          return;
        }
        const opNo = op.number ? pad6(op.number) : "—";
        const ok = confirm(`Arquivar OP ${opNo}? Ela sairá da lista principal e ficará em Arquivadas.`);
        if (!ok) return;
        try {
          await api(`/api/mrp/production-orders/${btn.dataset.id}/archive`, { method: "POST", body: JSON.stringify({}) });
          state.opSelectedIds = (state.opSelectedIds || []).filter(x => x !== btn.dataset.id);
          await loadAll();
        } catch (e) {
          console.error(e);
          alert("Falha ao arquivar OP.");
        }
      } else if (act === 'save') {
        const id = btn.dataset.id;
        const op = (state.ops || []).find(x => x.id === id);
        const sel = tbody.querySelector(`select[data-opstatus="${CSS.escape(id)}"]`);
        const status = sel ? String(sel.value || '') : '';

        const cur = normOpStatus(op?.status);
        const next = String(status || '').toUpperCase();

        // Avisos (conforme regra do usuário)
        if (next === 'IN_PRODUCTION' && cur === 'ISSUED') {
          const ok = confirm(
            'Ao mudar para EM PRODUÇÃO, o sistema vai dar BAIXA no estoque dos itens de consumo desta OP.\n\nDeseja continuar?'
          );
          if (!ok) return;
        }
        if (next === 'CLOSED' && cur === 'IN_PRODUCTION') {
          const ok = confirm(
            'Ao mudar para ENCERRADA, o sistema vai dar ENTRADA no estoque do produto final produzido (PF).\n\nDeseja continuar?'
          );
          if (!ok) return;
        }
        if (next === 'CANCELLED' && cur === 'IN_PRODUCTION') {
          const ok = confirm(
            'Ao CANCELAR uma OP que já estava EM PRODUÇÃO, o sistema vai ESTORNAR a baixa e DEVOLVER os itens ao estoque.\n\nDeseja continuar?'
          );
          if (!ok) return;
        }
        // Proteção extra (UI)
        if (next === 'CLOSED' && cur !== 'IN_PRODUCTION' && cur !== 'CLOSED') {
          alert('Para ENCERRAR a OP, primeiro mude o status para EM PRODUÇÃO e salve.');
          return;
        }

        try {
          const resp = await api(`/api/mrp/production-orders/${id}`, { method: 'PUT', body: JSON.stringify({ status: next }) });

          // Se acabou de ENCERRAR, oferece impressão das etiquetas do lote
          if (next === "CLOSED" && cur === "IN_PRODUCTION") {
            const ord = resp?.order || resp?.op || op;
            const opNo = pad6(ord?.number || op?.number);
            const lot = String(ord?.lotCode || (ord?.lotNumber ? (`LOTE${String(Math.trunc(Number(ord.lotNumber))).padStart(5,"0")}`) : "")).trim();
            const qty = Number(ord?.qtyToProduce || op?.qtyToProduce || 0) || 0;


            const msgLines = [
              `OP ${(opNo||"—")} encerrada.`,
              lot ? `Lote: ${lot}` : null,
              qty ? `Qtd etiquetas: ${qty}` : null,
              "",
              "Deseja imprimir as etiquetas agora?"
            ].filter(x => x !== null);

            const ask = confirm(msgLines.join("\n"));
            if (ask) {
              try { printProductionLabels(ord); } catch (e) {}
            }
          }

          await loadAll();
        } catch (e) {
          console.error(e);
          const code = String(e?.data?.error || '');
          if (code === 'insufficient_stock') {
            const rows = (e?.data?.shortages || []).map(s => {
              const it = itemsById.get(s.itemId);
              const u = it?.unit || '';
              const req = Number(s.required || 0);
              const av = Number(s.available || 0);
              const sh = Number(s.shortage || Math.max(0, req - av));
              return `<tr>
                <td>${escapeHtml(it?.name || '—')}</td>
                <td>${fmt(req)}</td>
                <td>${fmt(av)}</td>
                <td><b>${fmt(sh)}</b> ${escapeHtml(u)}</td>
              </tr>`;
            }).join('');
            openModal({
              title: 'Não dá para iniciar produção',
              subtitle: 'Ainda falta estoque para dar baixa nos itens de consumo.',
              submitText: 'Fechar',
              bodyHtml: `
                <div class="table-wrap">
                  <table class="table">
                    <thead><tr><th>Item</th><th>Necessário</th><th>Disponível</th><th>Faltante</th></tr></thead>
                    <tbody>${rows || `<tr><td colspan="4" class="muted">Sem detalhes.</td></tr>`}</tbody>
                  </table>
                </div>
              `,
              onSubmit: () => true,
            });
            // volta o select para Emitida
            if (sel) sel.value = 'ISSUED';
            return;
          }
          const msg = e?.data?.message ? String(e.data.message) : '';
          alert(msg || 'Falha ao salvar status da OP.');
        }
      } else if (act === 'del') {
        const id = btn.dataset.id;
        const ok = confirm('Excluir esta OP permanentemente? Isso remove também os movimentos de estoque gerados (se houver).');
        if (!ok) return;
        const re = await requestReauth({ reason: 'Excluir Ordem de Produção' });
        if (!re) return;
        try {
          await api(`/api/mrp/production-orders/${id}`, { method: 'DELETE', body: JSON.stringify({}) });
          state.opSelectedIds = (state.opSelectedIds || []).filter(x => x !== id);
          await loadAll();
        } catch (e) {
          console.error(e);
          alert('Falha ao excluir OP.');
        }
      }
    });
  });
}

function openNewProductionOrder() {
  const recipes = (state.recipes || []).slice().sort((a,b)=>String(a.name||"").localeCompare(String(b.name||""),"pt-BR"));
  if (!recipes.length) {
    alert("Cadastre pelo menos 1 Produto Final (PF) com BOM para criar uma OP.");
    return;
  }

  openModal({
    title: "Nova Ordem de Produção",
    subtitle: "Cria OP como Emitida. Ao mudar para Em produção, dá baixa no estoque (com aviso). Ao mudar para Encerrada, dá entrada do PF no estoque. Gera OC automaticamente se faltar estoque.",
    submitText: "Criar",
    cardClass: "wide",
    bodyHtml: `
      <label>Produto Final (Receita)
        <select name="recipeId" class="input" required>
          ${recipes.map(r => `<option value="${escapeHtml(r.id)}">${escapeHtml(r.name || "—")}</option>`).join("")}
        </select>
      </label>
      <div class="row wrap" style="gap:10px">
        <label style="min-width:220px; flex:1">Quantidade a produzir (un)
          <input name="qtyToProduce" class="input" type="number" step="1" min="1" value="1" required />
        </label>
        <label style="min-width:260px; flex:2">Observação
          <input name="note" class="input" value="" placeholder="Opcional" />
        </label>
      </div>
      <label class="row" style="gap:10px; align-items:center; margin-top:8px">
        <input name="createPurchaseOrder" type="checkbox" checked />
        <span>Gerar Ordem de Compra automaticamente quando faltar estoque</span>
      </label>
      <div class="muted small" style="margin-top:6px">Se houver faltas, a OP será criada como <b>Emitida</b>. Ao tentar mudar para <b>Em produção</b>, o sistema dará baixa no estoque; se faltar algum item, ele avisa e mantém em <b>Emitida</b>. Se a opção estiver marcada, uma OC será aberta para os itens faltantes. A entrada do PF ocorre ao salvar como <b>Encerrada</b>.</div>
    `,
    onSubmit: async (fd) => {
      const recipeId = String(fd.get("recipeId") || "");
      const qtyToProduce = Number(fd.get("qtyToProduce") || 0);
      const note = String(fd.get("note") || "");
      const createPurchaseOrder = !!fd.get("createPurchaseOrder");

      if (!recipeId || !Number.isFinite(qtyToProduce) || qtyToProduce <= 0) {
        alert("Preencha Produto Final e Quantidade.");
        return false;
      }

      await api("/api/mrp/production-orders", {
        method: "POST",
        body: JSON.stringify({
          recipeId,
          qtyToProduce,
          note,
          allowInsufficient: true,
          createPurchaseOrder,
        }),
      });
      await loadAll();
      return true;
    }
  });
}

// ---------- Purchase Orders UI ----------
function renderPurchaseOrders() {
  const tbody = $("#poTable tbody");
  if (!tbody) return;
  tbody.innerHTML = "";

  if (!Array.isArray(state.poSelectedIds)) state.poSelectedIds = [];
  const selSet = new Set(state.poSelectedIds);

  const itemsById = new Map((state.rawItems || []).map(i => [i.id, i]));

  const statusSelect = (st, id) => {
    const s = String(st || "OPEN").toUpperCase();
    const opts = ["OPEN","PARTIAL","RECEIVED","CANCELLED"].map(k => {
      const label = OC_STATUS_PT[k] || k;
      return `<option value="${k}" ${k===s?'selected':''}>${escapeHtml(label)}</option>`;
    }).join("");
    return `<select class="input" data-postatus="${escapeHtml(id)}" style="min-width:140px">${opts}</select>`;
  };

  for (const po of (state.purchaseOrders || [])) {
    const itemsTxt = (po.items || []).map(li => {
      const it = itemsById.get(li.itemId);
      const u = it?.unit || "";
      const finalOrd = (Number(li.qtyOrdered || 0) + Number(li.qtyAdjusted || 0));
      const ord = fmt(finalOrd);
      const rec = fmt(li.qtyReceived || 0);
      return `${escapeHtml(it?.name || "—")}: ${ord} ${escapeHtml(u)} <span class="muted small">(rec: ${rec})</span>`;
    }).join("<br/>");

    const ocNo = po.number ? pad6(po.number) : "—";

    const tr = document.createElement("tr");
    tr.innerHTML = `
      <td style="text-align:center"><input type="checkbox" data-posel="${escapeHtml(po.id)}" ${selSet.has(po.id)?'checked':''} /></td>
      <td>${fmtDate(po.createdAt)}</td>
      <td><b>${escapeHtml(ocNo)}</b></td>
      <td>${statusSelect(po.status, po.id)}</td>
      <td class="muted">${escapeHtml(po.note || "")}</td>
      <td class="muted small">${itemsTxt || "—"}</td>
      <td style="text-align:center">
        <div class="row wrap" style="gap:6px; justify-content:center">
          <button type="button" class="btn secondary small" data-act="save" data-id="${escapeHtml(po.id)}">Salvar</button>
          <button type="button" class="btn secondary small" data-act="archive" data-id="${escapeHtml(po.id)}">Arquivar</button>
          <button type="button" class="btn secondary small" data-act="del" data-id="${escapeHtml(po.id)}">Excluir</button>
          <button type="button" class="btn secondary small" data-act="edit" data-id="${escapeHtml(po.id)}">Editar</button>
          <button type="button" class="btn secondary small" data-act="print" data-id="${escapeHtml(po.id)}">Imprimir</button>
          <button type="button" class="btn primary small" data-act="recv" data-id="${escapeHtml(po.id)}">Receber</button>
        </div>
      </td>
    `;
    tbody.appendChild(tr);
  }

  if (!state.purchaseOrders || state.purchaseOrders.length === 0) {
    const tr = document.createElement("tr");
    tr.innerHTML = `<td colspan="7" class="muted">Sem ordens de compra ainda.</td>`;
    tbody.appendChild(tr);
  }

  // select all checkbox
  const selAll = document.querySelector('#poSelectAll');
  if (selAll) {
    const allIds = (state.purchaseOrders || []).map(o => o.id);
    const allChecked = allIds.length && allIds.every(id => selSet.has(id));
    selAll.checked = !!allChecked;
    selAll.onchange = () => {
      if (selAll.checked) state.poSelectedIds = allIds.slice();
      else state.poSelectedIds = [];
      renderPurchaseOrders();
    };
  }

  // selection per row
  tbody.querySelectorAll('input[type=checkbox][data-posel]').forEach(cb => {
    cb.addEventListener('change', () => {
      const id = cb.getAttribute('data-posel');
      const set = new Set(state.poSelectedIds || []);
      if (cb.checked) set.add(id);
      else set.delete(id);
      state.poSelectedIds = Array.from(set);
      const selAll = document.querySelector('#poSelectAll');
      if (selAll) {
        const allIds = (state.purchaseOrders || []).map(o => o.id);
        selAll.checked = allIds.length && allIds.every(x => set.has(x));
      }
    });
  });

  tbody.querySelectorAll("button[data-act]").forEach(btn => {
    btn.addEventListener("click", async (e) => {
      // Evita conflitos de clique (ex.: 1º clique disparar outra ação)
      try {
        e.preventDefault();
        e.stopPropagation();
        e.stopImmediatePropagation();
      } catch {}
      const po = (state.purchaseOrders || []).find(x => x.id === btn.dataset.id);
      if (!po) return;
      if (btn.dataset.act === "edit") return openEditPurchaseOrder(po);
      if (btn.dataset.act === "print") return printPurchaseOrder(po);
      if (btn.dataset.act === "recv") return openReceivePurchaseOrder(po);
      if (btn.dataset.act === "archive") {
        const id = btn.dataset.id;
        const ocNo = po.number ? pad6(po.number) : "—";
        const ok = confirm(`Arquivar a OC ${ocNo}? Ela sairá da lista principal e ficará em "Arquivadas".`);
        if (!ok) return;
        try {
          await api(`/api/mrp/purchase-orders/${id}/archive`, { method: "POST", body: JSON.stringify({}) });
          state.poSelectedIds = (state.poSelectedIds || []).filter(x => x !== id);
          await loadAll();
        } catch (e) {
          console.error(e);
          alert("Falha ao arquivar OC.");
        }
        return;
      }
      if (btn.dataset.act === "save") {
        try {
          const id = btn.dataset.id;
          const sel = tbody.querySelector(`select[data-postatus="${CSS.escape(id)}"]`);
          const status = sel ? String(sel.value || '') : '';
          await api(`/api/mrp/purchase-orders/${id}`, { method: 'PUT', body: JSON.stringify({ status }) });
          await loadPurchaseOrders();
        } catch (e) {
          console.error(e);
          alert('Falha ao salvar status da OC.');
        }
      }
      if (btn.dataset.act === "del") {
        const id = btn.dataset.id;
        const ok = confirm('Excluir esta OC permanentemente? Isso remove também as entradas de estoque (recebimentos) geradas por ela.');
        if (!ok) return;
        const re = await requestReauth({ reason: 'Excluir Ordem de Compra' });
        if (!re) return;
        try {
          await api(`/api/mrp/purchase-orders/${id}`, { method: 'DELETE', body: JSON.stringify({}) });
          state.poSelectedIds = (state.poSelectedIds || []).filter(x => x !== id);
          await loadAll();
        } catch (e) {
          console.error(e);
          alert('Falha ao excluir OC.');
        }
      }
    });
  });
}

async function openArchivedPurchaseOrders() {
  openModal({
    title: "Ordens de Compra Arquivadas",
    subtitle: "Aqui ficam as OCs arquivadas. Você pode imprimir, restaurar ou excluir permanentemente.",
    submitText: "Fechar",
    bodyHtml: `
      <div class="muted small" style="margin-bottom:8px">Dica: arquivar mantém a lista principal limpa, sem apagar dados.</div>
      <div id="poArchivedBox" class="table-wrap" style="max-height:70vh; overflow:auto">
        <table class="table" id="poArchivedTable">
          <thead>
            <tr>
              <th>Data</th>
              <th style="width:90px">OC</th>
              <th>Status</th>
              <th>Observação</th>
              <th>Itens</th>
              <th style="width:240px; text-align:center">Ações</th>
            </tr>
          </thead>
          <tbody>
            <tr><td colspan="6" class="muted">Carregando...</td></tr>
          </tbody>
        </table>
      </div>
    `,
    onSubmit: async () => true,
    onOpen: () => {
      (async () => {
        try {
          const res = await api("/api/mrp/purchase-orders/archived");
          const orders = Array.isArray(res.orders) ? res.orders : [];
          const tbody = document.querySelector('#poArchivedTable tbody');
          if (!tbody) return;
          tbody.innerHTML = "";

          const itemsById = new Map((state.rawItems || []).map(i => [i.id, i]));
          const statusLabel = (st) => OC_STATUS_PT[String(st || "OPEN").toUpperCase()] || String(st || "OPEN");

          for (const po of orders) {
            const itemsTxt = (po.items || []).map(li => {
              const it = itemsById.get(li.itemId);
              const u = it?.unit || "";
              const finalOrd = (Number(li.qtyOrdered || 0) + Number(li.qtyAdjusted || 0));
              const ord = fmt(finalOrd);
              const rec = fmt(li.qtyReceived || 0);
              return `${escapeHtml(it?.name || "—")}: ${ord} ${escapeHtml(u)} <span class="muted small">(rec: ${rec})</span>`;
            }).join("<br/>");

            const ocNo = po.number ? pad6(po.number) : "—";
            const tr = document.createElement('tr');
            tr.innerHTML = `
              <td>${fmtDate(po.archivedAt || po.createdAt)}</td>
              <td><b>${escapeHtml(ocNo)}</b></td>
              <td>${escapeHtml(statusLabel(po.status))}</td>
              <td class="muted">${escapeHtml(po.note || "")}</td>
              <td class="muted small">${itemsTxt || "—"}</td>
              <td style="text-align:center">
                <div class="row wrap" style="gap:6px; justify-content:center">
                  <button type="button" class="btn secondary small" data-poa="print" data-id="${escapeHtml(po.id)}">Imprimir</button>
                  <button type="button" class="btn secondary small" data-poa="restore" data-id="${escapeHtml(po.id)}">Restaurar</button>
                  <button type="button" class="btn secondary small" data-poa="del" data-id="${escapeHtml(po.id)}">Excluir</button>
                </div>
              </td>
            `;
            tbody.appendChild(tr);
          }

          if (!orders.length) {
            const tr = document.createElement('tr');
            tr.innerHTML = `<td colspan="6" class="muted">Sem OCs arquivadas ainda.</td>`;
            tbody.appendChild(tr);
          }

          // actions
          tbody.querySelectorAll('button[data-poa]').forEach((btn) => {
            btn.addEventListener('click', async (e) => {
              try { e.preventDefault(); e.stopPropagation(); e.stopImmediatePropagation(); } catch {}
              const id = btn.getAttribute('data-id');
              const po = orders.find(o => o.id === id);
              if (!po) return;
              const act = btn.getAttribute('data-poa');
              if (act === 'print') return printPurchaseOrder(po);
              if (act === 'restore') {
                const ocNo = po.number ? pad6(po.number) : '—';
                const ok = confirm(`Restaurar a OC ${ocNo} para a lista principal?`);
                if (!ok) return;
                try {
                  await api(`/api/mrp/purchase-orders/archived/${id}/restore`, { method: 'POST', body: JSON.stringify({}) });
                  await loadAll();
                  // refresh modal list
                  openArchivedPurchaseOrders();
                } catch (e) {
                  console.error(e);
                  alert('Falha ao restaurar OC.');
                }
                return;
              }
              if (act === 'del') {
                const ocNo = po.number ? pad6(po.number) : '—';
                const ok = confirm(`Excluir a OC ${ocNo} permanentemente? Isso remove também as entradas de estoque (recebimentos) geradas por ela.`);
                if (!ok) return;
                const re = await requestReauth({ reason: 'Excluir Ordem de Compra Arquivada' });
                if (!re) return;
                try {
                  await api(`/api/mrp/purchase-orders/archived/${id}`, { method: 'DELETE', body: JSON.stringify({}) });
                  await loadAll();
                  openArchivedPurchaseOrders();
                } catch (e) {
                  console.error(e);
                  alert('Falha ao excluir OC arquivada.');
                }
              }
            });
          });
        } catch (e) {
          console.error(e);
          const tbody = document.querySelector('#poArchivedTable tbody');
          if (tbody) tbody.innerHTML = `<tr><td colspan="6" class="muted">Falha ao carregar OCs arquivadas.</td></tr>`;
        }
      })();
    }
  });
}

function printPurchaseOrder(po) {
  const itemsById = new Map((state.rawItems || []).map(i => [i.id, i]));
  const ocNo = po.number ? pad6(po.number) : (po.id ? String(po.id) : "—");
  const createdAt = po.createdAt ? fmtDate(po.createdAt) : "";
  const note = String(po.note || "").trim();

  const lines = (po.items || []).map(li => {
    const it = itemsById.get(li.itemId);
    const name = it?.name || "—";
    const unit = it?.unit || "";
    const finalOrd = (Number(li.qtyOrdered || 0) + Number(li.qtyAdjusted || 0));
    const rec = Number(li.qtyReceived || 0);
    const remaining = Math.max(0, finalOrd - rec);
    return { name, unit, remaining, finalOrd, rec };
  });

  // Para o supermercado: foca no que ainda falta comprar (remaining > 0).
  const toBuy = lines.filter(x => x.remaining > 1e-9).sort((a,b)=>String(a.name).localeCompare(String(b.name),"pt-BR"));

  const rows = toBuy.map((x, i) => {
    return `<tr>
      <td class="chk"><span class="box"></span></td>
      <td class="item">${escapeHtml(x.name)}</td>
      <td class="qty">${escapeHtml(fmt(x.remaining))}</td>
      <td class="unit">${escapeHtml(x.unit)}</td>
    </tr>`;
  }).join("");

  const totalItems = toBuy.length;

  const html = `<!doctype html>
  <html>
  <head>
    <meta charset="utf-8" />
    <meta name="viewport" content="width=device-width, initial-scale=1" />
    <title>OC ${escapeHtml(ocNo)} • Lista de Compras</title>
    <style>
      :root{
        --text:#0f172a;
        --muted:#475569;
        --line:#e2e8f0;
      }
      *{ box-sizing:border-box; }
      body{
        margin:0;
        font-family: ui-sans-serif, system-ui, -apple-system, Segoe UI, Roboto, Helvetica, Arial;
        color: var(--text);
        background:#fff;
      }
      .page{
        padding: 22px 26px 26px;
      }
      .header{
        display:flex;
        align-items:flex-end;
        justify-content:space-between;
        gap:14px;
        padding-bottom: 12px;
        border-bottom: 2px solid #111827;
        margin-bottom: 14px;
      }
      .brand{
        display:flex;
        flex-direction:column;
        gap:4px;
      }
      .brand .kicker{
        font-weight: 900;
        letter-spacing: .12em;
        text-transform: uppercase;
        font-size: 11px;
        color: var(--muted);
      }
      .brand .title{
        font-weight: 900;
        font-size: 22px;
        letter-spacing: .2px;
        line-height: 1.1;
      }
      .brand .meta{
        font-size: 12px;
        color: var(--muted);
      }
      .badge{
        text-align:right;
        font-size: 12px;
        color: var(--muted);
      }
      .badge b{ color: var(--text); }
      .note{
        margin-top: 10px;
        border: 1px solid var(--line);
        border-radius: 12px;
        padding: 10px 12px;
        font-size: 12px;
        color: var(--text);
      }
      .note .lbl{
        color: var(--muted);
        font-weight: 800;
        font-size: 11px;
        letter-spacing: .08em;
        text-transform: uppercase;
        margin-bottom: 6px;
      }

      table{
        width:100%;
        border-collapse:collapse;
        margin-top: 12px;
      }
      thead th{
        font-size: 11px;
        color: var(--muted);
        letter-spacing: .08em;
        text-transform: uppercase;
        text-align:left;
        padding: 8px 8px;
        border-bottom: 1px solid var(--line);
      }
      tbody td{
        padding: 10px 8px;
        border-bottom: 1px solid rgba(226,232,240,.75);
        vertical-align: middle;
        font-size: 13px;
      }
      .chk{ width:34px; }
      .qty{ width:120px; text-align:right; font-weight: 900; }
      .unit{ width:70px; color: var(--muted); font-weight: 800; }
      .box{
        display:inline-block;
        width:18px; height:18px;
        border: 2px solid #111827;
        border-radius: 4px;
      }
      .footer{
        display:flex;
        justify-content:space-between;
        gap: 14px;
        margin-top: 18px;
        font-size: 12px;
        color: var(--muted);
      }
      .sig{
        flex:1;
        border-top: 1px solid var(--line);
        padding-top: 10px;
      }
      .empty{
        margin-top: 14px;
        padding: 14px 12px;
        border: 1px dashed var(--line);
        border-radius: 12px;
        color: var(--muted);
        font-size: 13px;
      }
      @media print{
        .page{ padding: 0; }
        .note{ border-color:#111827; }
        thead{ display: table-header-group; }
      }
    </style>
  </head>
  <body>
    <div class="page">
      <div class="header">
        <div class="brand">
          <div class="kicker">dietON • Lista de compras</div>
          <div class="title">OC ${escapeHtml(ocNo)}</div>
          <div class="meta">${escapeHtml(createdAt)} • Itens: <b>${escapeHtml(String(totalItems))}</b></div>
        </div>
        <div class="badge">
          <div><b>Comprar:</b> pedido final − recebido</div>
          <div>${escapeHtml(new Date().toLocaleString('pt-BR'))}</div>
        </div>
      </div>

      ${note ? `<div class="note"><div class="lbl">Observação</div>${escapeHtml(note)}</div>` : ""}

      ${rows ? `
      <table>
        <thead>
          <tr>
            <th style="width:34px"></th>
            <th>Item</th>
            <th style="text-align:right; width:120px">Qtde</th>
            <th style="width:70px">UN</th>
          </tr>
        </thead>
        <tbody>
          ${rows}
        </tbody>
      </table>
      ` : `<div class="empty">Nada pendente para comprar nesta OC (tudo já foi recebido).</div>`}

      <div class="footer">
        <div class="sig">Comprador: ____________________________</div>
        <div class="sig" style="max-width:220px">Data: ____/____/______</div>
      </div>
    </div>

    <script>
      // imprime automaticamente (com leve delay para garantir render)
      window.addEventListener('load', () => { setTimeout(() => { window.print(); }, 150); });
    </script>
  </body>
  </html>`;

  const w = window.open("", "_blank", "width=980,height=720");
  if (!w) {
    alert("Não foi possível abrir a janela de impressão. Verifique se o navegador bloqueou pop-ups.");
    return;
  }
  w.document.open();
  w.document.write(html);
  w.document.close();
  w.focus();
}

function printProductionOrder(op) {
  const rawById = new Map((state.rawItems || []).map(i => [i.id, i]));
  const fgById = new Map((state.fgItems || []).map(i => [i.id, i]));
  const recipe = (state.recipes || []).find(r => r.id === op.recipeId) || null;
  const methodText = String(recipe?.method || op.method || "").trim();

  const opNo = op.number ? pad6(op.number) : (op.id ? String(op.id) : "—");
  const createdAt = op.createdAt ? fmtDate(op.createdAt) : "";
  const note = String(op.note || "").trim();

  const outId = op.planned?.produced?.itemId || op.produced?.itemId || null;
  const outUnit = fgById.get(outId)?.unit || "un";
  const qtyProd = Number(op.qtyToProduce || op.planned?.produced?.qty || 0);

  const st = OP_STATUS_PT[String(op.status || "ISSUED").toUpperCase()] || String(op.status || "ISSUED");

  const lines = (Array.isArray(op.planned?.consumed) && op.planned.consumed.length)
    ? op.planned.consumed
    : (op.consumed || []);

  // Lista para separação: ordena por nome
  const toPick = (lines || []).map(li => {
    const it = rawById.get(li.itemId);
    return { name: it?.name || "—", unit: it?.unit || "", qty: Number(li.qty || 0) };
  }).sort((a,b)=>String(a.name).localeCompare(String(b.name), "pt-BR"));

  const rows = toPick.map(x => `
    <tr>
      <td class="chk"><span class="box"></span></td>
      <td class="item">${escapeHtml(x.name)}</td>
      <td class="qty">${escapeHtml(fmt(x.qty))}</td>
      <td class="unit">${escapeHtml(x.unit)}</td>
    </tr>
  `).join("");

  const html = `<!doctype html>
  <html>
  <head>
    <meta charset="utf-8" />
    <meta name="viewport" content="width=device-width, initial-scale=1" />
    <title>OP ${escapeHtml(opNo)} • Lista + Método</title>
    <style>
      :root{
        --text:#0f172a;
        --muted:#475569;
        --line:#e2e8f0;
      }
      *{ box-sizing:border-box; }
      body{
        margin:0;
        font-family: ui-sans-serif, system-ui, -apple-system, Segoe UI, Roboto, Helvetica, Arial;
        color: var(--text);
        background:#fff;
      }
      .page{ padding: 22px 26px 26px; }
      .page-break{ page-break-before: always; break-before: page; }
      .header{
        display:flex;
        align-items:flex-end;
        justify-content:space-between;
        gap:14px;
        padding-bottom: 12px;
        border-bottom: 2px solid #111827;
        margin-bottom: 14px;
      }
      .brand{ display:flex; flex-direction:column; gap:4px; }
      .brand .kicker{
        font-weight: 900;
        letter-spacing: .12em;
        text-transform: uppercase;
        font-size: 11px;
        color: var(--muted);
      }
      .brand .title{ font-weight:900; font-size:22px; letter-spacing:.2px; line-height:1.1; }
      .brand .meta{ font-size:12px; color: var(--muted); }
      .right{ text-align:right; font-size:12px; color: var(--muted); line-height:1.45; }
      .right .big{ font-weight:900; font-size:14px; color:#111827; }
      .pill{
        display:inline-block;
        padding:3px 9px;
        border:1px solid #111827;
        border-radius:999px;
        font-weight:700;
        font-size:11px;
        letter-spacing:.04em;
        text-transform: uppercase;
        color:#111827;
        margin-left:8px;
      }
      .note{
        margin: 10px 0 14px;
        padding: 10px 12px;
        border: 1px solid var(--line);
        border-radius: 10px;
        color: var(--muted);
        font-size: 12px;
      }
      table{ width:100%; border-collapse:collapse; margin-top: 10px; }
      thead th{
        text-align:left;
        font-size: 12px;
        padding: 8px 10px;
        border-bottom: 1px solid var(--line);
        color: var(--muted);
      }
      tbody td{
        padding: 10px 10px;
        border-bottom: 1px solid var(--line);
        font-size: 14px;
        vertical-align: top;
      }
      td.qty, td.unit{ text-align:right; white-space:nowrap; font-variant-numeric: tabular-nums; }
      td.chk{ width: 34px; text-align:center; }
      .box{ display:inline-block; width: 18px; height: 18px; border: 2px solid #111827; border-radius: 4px; }
      .empty{ border: 1px dashed var(--line); padding: 16px; border-radius: 12px; color: var(--muted); text-align:center; }
      .section-title{ font-weight: 900; font-size: 16px; letter-spacing:.2px; margin: 6px 0 10px; }
      .method{
        margin-top: 10px;
        white-space: pre-wrap;
        border: 1px solid var(--line);
        border-radius: 12px;
        padding: 12px 14px;
        font-size: 13.2px;
        line-height: 1.42;
        color: var(--text);
      }
      .method-empty{
        color: var(--muted);
        border: 1px dashed var(--line);
        border-radius: 12px;
        padding: 14px 12px;
        font-size: 12.5px;
      }
      .footer{ display:flex; gap:16px; justify-content:space-between; margin-top: 18px; color: var(--muted); font-size: 12px; }
      .sig{ flex:1; border-top: 1px solid var(--line); padding-top: 8px; }
      @media print{ .page{ padding: 14mm 12mm; } }
    </style>
  </head>
  <body>
    <div class="page">
      <div class="header">
        <div class="brand">
          <div class="kicker">dietON • Produção</div>
          <div class="title">Ordem de Produção <span class="pill">${escapeHtml(st)}</span></div>
          <div class="meta">Receita: <b>${escapeHtml(op.recipeName || "—")}</b></div>
          <div class="meta">Quantidade: <b>${escapeHtml(fmt(qtyProd))} ${escapeHtml(outUnit)}</b></div>
        </div>
        <div class="right">
          <div class="big">OP ${escapeHtml(opNo)}</div>
          <div>Data: ${escapeHtml(createdAt)}</div>
        </div>
      </div>

      ${note ? `<div class="note"><b>Observação:</b> ${escapeHtml(note)}</div>` : ""}

      ${toPick.length ? `
        <table>
          <thead>
            <tr>
              <th class="chk"></th>
              <th>Item</th>
              <th style="text-align:right">Qtd</th>
              <th style="text-align:right">UN</th>
            </tr>
          </thead>
          <tbody>${rows}</tbody>
        </table>
      ` : `<div class="empty">Sem consumo registrado para esta OP.</div>`}

      <div class="footer">
        <div class="sig">Responsável: ____________________________</div>
        <div class="sig" style="max-width:220px">Data: ____/____/______</div>
      </div>
    </div>

    <div class="page page-break">
      <div class="header" style="margin-bottom:10px; border-bottom:1px solid var(--line)">
        <div class="brand">
          <div class="kicker">dietON • Produção</div>
          <div class="title">Método de preparo</div>
          <div class="meta">Receita: <b>${escapeHtml(op.recipeName || "—")}</b></div>
          <div class="meta">OP ${escapeHtml(opNo)} • Quantidade: <b>${escapeHtml(fmt(qtyProd))} ${escapeHtml(outUnit)}</b></div>
        </div>
        <div class="right">
          <div>Data: ${escapeHtml(createdAt)}</div>
        </div>
      </div>

      ${methodText
        ? `<div class="method">${escapeHtml(methodText)}</div>`
        : `<div class="method-empty">Nenhum método cadastrado para esta receita ainda. Use <b>MRP → Método</b> para criar/editar.</div>`}

      <div class="footer">
        <div class="sig">Responsável: ____________________________</div>
        <div class="sig" style="max-width:220px">Data: ____/____/______</div>
      </div>
    </div>

    <script>
      window.addEventListener('load', () => { setTimeout(() => { window.print(); }, 150); });
    </script>
  </body>
  </html>`;

  const w = window.open("", "_blank", "width=980,height=720");
  if (!w) {
    alert("Não foi possível abrir a janela de impressão. Verifique se o navegador bloqueou pop-ups.");
    return;
  }
  w.document.open();
  w.document.write(html);
  w.document.close();
  w.focus();
}

function printProductionLabels(op) {
  const recipe = (state.recipes || []).find(r => r.id === op.recipeId) || null;
  const recipeCode = String(op.recipeCode || recipe?.code || "").trim();
  const recipeName = String(op.recipeName || recipe?.name || "").trim();

  const lotCode = String(op.lotCode || (op.lotNumber ? (`LOTE${String(Math.trunc(Number(op.lotNumber))).padStart(5,"0")}`) : "")).trim();
  const dtIso = op.closedAt || op.executedAt || op.createdAt || new Date().toISOString();
  const dt = new Date(dtIso).toLocaleString("pt-BR");

  const barcodeValue = String(op.barcodeValue || "").trim()
    || (recipeCode && lotCode ? `DIETON-${recipeCode}-${lotCode}` : `DIETON-${recipeCode || "PF"}-${lotCode || "LOTE00000"}`);

  const qty = Math.max(0, Math.trunc(Number(op.qtyToProduce || 0)));
  if (!qty) {
    alert("Quantidade da OP é 0 — nada para imprimir.");
    return;
  }

  // Pimaco/A4-6093 (Carta): 24 círculos (4x6). Medidas aproximadas em polegadas.
  // Dica: ao imprimir, usar Escala 100% e margens "Nenhuma" (ou o mínimo possível).
  const DIAM = 1.665;     // diâmetro do círculo (in)
  const PITCH_X = 2.012;  // distância entre colunas (in)
  const PITCH_Y = 1.752;  // distância entre linhas (in)
  const M_LEFT = 0.40;    // margem esquerda (in)
  const M_TOP = 0.28;     // margem superior (in)
  const PER_PAGE = 24;

  const pages = Math.ceil(qty / PER_PAGE);
  const barcodeUrl = `/api/barcode?text=${encodeURIComponent(barcodeValue)}&t=${Date.now()}`;

  const sheets = [];
  for (let p = 0; p < pages; p++) {
    const start = p * PER_PAGE;
    const end = Math.min(qty, start + PER_PAGE);
    const count = end - start;

    const labels = [];
    for (let i = 0; i < PER_PAGE; i++) {
      const globalIndex = start + i;
      if (globalIndex >= qty) {
        labels.push(`<div class="label empty" style="${labelPosStyle(i, DIAM, PITCH_X, PITCH_Y, M_LEFT, M_TOP)}"></div>`);
        continue;
      }

      labels.push(`
        <div class="label" style="${labelPosStyle(i, DIAM, PITCH_X, PITCH_Y, M_LEFT, M_TOP)}">
          <div class="code">${escapeHtml(recipeCode || "PF")}</div>
          <div class="name">${escapeHtml(recipeName || "")}</div>
          <img class="barcode" src="${barcodeUrl}" alt="barcode" />
          <div class="human">${escapeHtml(barcodeValue)}</div>
          <div class="meta">
            <span>${escapeHtml(lotCode || "")}</span>
            <span>${escapeHtml(dt)}</span>
          </div>
        </div>
      `);
    }

    sheets.push(`<div class="sheet">${labels.join("")}</div>`);
  }

  const html = `<!doctype html>
<html>
<head>
  <meta charset="utf-8" />
  <meta name="viewport" content="width=device-width, initial-scale=1" />
  <title>Etiquetas • ${escapeHtml(recipeCode || "PF")} • ${escapeHtml(lotCode || "")}</title>
  <style>
    @page { size: Letter; margin: 0; }
    * { box-sizing: border-box; }
    body { margin: 0; background: #fff; font-family: ui-sans-serif, system-ui, -apple-system, Segoe UI, Roboto, Helvetica, Arial; }
    .screen-note{
      padding: 10px 12px;
      font-size: 12px;
      color:#0f172a;
      border-bottom: 1px solid #e2e8f0;
      background:#f8fafc;
    }
    .screen-note b{ font-weight: 900; }
    .sheet{
      position: relative;
      width: 8.5in;
      height: 11in;
      page-break-after: always;
    }
    .label{
      position: absolute;
      width: ${DIAM}in;
      height: ${DIAM}in;
      border-radius: 999px;
      overflow: hidden;
      display: flex;
      flex-direction: column;
      align-items: center;
      justify-content: center;
      padding: 0.10in 0.10in;
      text-align: center;
      color: #0f172a;
    }
    .label.empty{ }
    .code{ font-size: 9pt; font-weight: 900; line-height: 1.0; margin-bottom: 2px; }
    .name{
      font-size: 7.6pt;
      line-height: 1.05;
      font-weight: 700;
      margin-bottom: 4px;
      max-height: 0.62in;
      overflow: hidden;
      display: -webkit-box;
      -webkit-line-clamp: 3;
      -webkit-box-orient: vertical;
    }
    .barcode{
      width: 1.32in;
      height: auto;
      margin: 0;
    }
    .human{
      font-size: 6.2pt;
      line-height: 1.0;
      color: #334155;
      margin-top: 2px;
      max-width: 100%;
      overflow: hidden;
      text-overflow: ellipsis;
      white-space: nowrap;
    }
    .meta{
      margin-top: 4px;
      width: 100%;
      display: flex;
      flex-direction: column;
      gap: 1px;
      font-size: 6.4pt;
      line-height: 1.0;
      color: #0f172a;
      font-variant-numeric: tabular-nums;
    }
    @media print{
      .screen-note{ display:none; }
      body{ -webkit-print-color-adjust: exact; print-color-adjust: exact; }
    }
  </style>
</head>
<body>
  <div class="screen-note">
    <b>Impressão de etiquetas 6093</b> • Recomendado: <b>Escala 100%</b> e <b>Margens: Nenhuma</b>.
    <span style="opacity:.75">• Se não alinhar na 1ª tentativa, me avise que ajusto as margens/pitch.</span>
  </div>
  ${sheets.join("\n")}
  <script>
    window.addEventListener('load', () => { setTimeout(() => { window.print(); }, 200); });
  </script>
</body>
</html>`;

  const w = window.open("", "_blank", "width=980,height=720");
  if (!w) {
    alert("Não foi possível abrir a janela de impressão. Verifique se o navegador bloqueou pop-ups.");
    return;
  }
  w.document.open();
  w.document.write(html);
  w.document.close();
  w.focus();
}

function labelPosStyle(i, diam, pitchX, pitchY, mLeft, mTop) {
  const col = i % 4;
  const row = Math.floor(i / 4);
  const left = mLeft + (col * pitchX);
  const top = mTop + (row * pitchY);
  return `left:${left}in; top:${top}in;`;
}




function openNewPurchaseOrder() {
  openEditPurchaseOrder({ id: "", status: "OPEN", note: "", items: [], linkedProductionOrderId: null }, { mode: "new" });
}

function openEditPurchaseOrder(po, { mode = "edit" } = {}) {
  const rawItems = (state.rawItems || []).slice().sort((a,b)=>String(a.name||"").localeCompare(String(b.name||""),"pt-BR"));
  const itemsById = new Map(rawItems.map(i => [i.id, i]));

  const unitOf = (itemId) => (itemsById.get(itemId)?.unit || "");
  const isEach = (u) => fmtUnit(u) === "un";

  const clampByUnit = (n, u) => {
    let x = Number(n);
    if (!Number.isFinite(x)) x = 0;
    // Permite ajuste negativo (Ajustar = delta), mas mantém inteiro quando UN = un.
    if (isEach(u)) x = Math.trunc(x);
    return Number(Number(x).toFixed(6));
  };

  const fmtInput = (n, u) => {
    const x = Number(n);
    if (!Number.isFinite(x) || x === 0) return "0";
    if (isEach(u)) return String(Math.trunc(x));
    // input text: mostramos pt-BR (vírgula) e até 3 casas
    return x.toLocaleString("pt-BR", { maximumFractionDigits: 3 });
  };

  let lines = (po.items || []).map(li => ({
    itemId: li.itemId,
    qtyOrdered: Number(li.qtyOrdered || 0),     // Pedida (base do sistema) - NÃO editável
    qtyAdjusted: Number(li.qtyAdjusted || 0),   // Ajustar (valor final) - começa em 0
    qtyReceived: Number(li.qtyReceived || 0),
  }));

  const renderTable = () => {
    const rows = lines.map((li, idx) => {
      const it = itemsById.get(li.itemId);
      const u = it?.unit || "";
      const inputMode = "decimal";
      const finalShow = clampByUnit((Number(li.qtyOrdered || 0) + Number(li.qtyAdjusted || 0)), u);
      return `<tr>
        <td>${escapeHtml(it?.name || "—")}</td>
        <td class="muted" style="width:120px; text-align:right">${fmt(li.qtyOrdered)} ${escapeHtml(u)}</td>
        <td style="width:140px">
          <input class="input poAdj" data-idx="${idx}" data-unit="${escapeHtml(u)}" inputmode="${inputMode}" value="${escapeHtml(fmtInput(li.qtyAdjusted || 0, u))}" />
          <div class="muted small" style="margin-top:2px">Final: ${fmt(finalShow)} ${escapeHtml(u)}</div>
        </td>
        <td class="muted" style="width:120px; text-align:right">${fmt(li.qtyReceived||0)} ${escapeHtml(u)}</td>
        <td class="muted" style="width:70px">${escapeHtml(u)}</td>
        <td style="width:90px; text-align:center">
          <button class="btn danger micro" type="button" data-del="${idx}">Excluir</button>
        </td>
      </tr>`;
    }).join("");

    return `
      <div class="table-wrap">
        <table class="table">
          <thead><tr><th>Item</th><th style="text-align:right">Pedida</th><th>Ajustar</th><th style="text-align:right">Recebida</th><th>UN</th><th></th></tr></thead>
          <tbody>${rows || `<tr><td colspan="6" class="muted">Adicione itens.</td></tr>`}</tbody>
        </table>
      </div>
    `;
  };

  const renderAddRow = () => `
    <div class="po-add-grid">
      <label class="po-add-item">
        Item (MP)
        <select id="poAddItem" class="input">
          <option value="">Selecione...</option>
          ${rawItems.map(it => `<option value="${it.id}">${escapeHtml(it.name)} (${escapeHtml(it.unit)})</option>`).join("")}
        </select>
      </label>

      <label class="po-add-unit">
        UN
        <input id="poAddUnit" class="input" value="" readonly />
      </label>

      <label class="po-add-qty">
        Qtde pedida
        <input id="poAddQty" class="input" inputmode="decimal" value="0" />
      </label>

      <div class="po-add-btn">
        <button id="poAddBtn" class="btn secondary micro" type="button">Adicionar</button>
      </div>
    </div>
  `;

  const ocNo = po.number ? pad6(po.number) : (po.id ? String(po.id) : "—");
  const titleIn = mode === "new" ? "NOVA ORDEM DE COMPRA" : "EDITAR ORDEM DE COMPRA";

  openModal({
    title: mode === "new" ? "Nova Ordem de Compra" : "Editar Ordem de Compra",
    subtitle: po.number ? `OC ${pad6(po.number)}` : (po.id ? `OC ${po.id}` : ""),
    submitText: "Salvar",
    cardClass: "wide po-compact po-edit",
    bodyHtml: `
      <div class="po-split">
        <div class="po-subcard">
          <div class="po-subhead">
            <div class="po-subtitle">${escapeHtml(titleIn)}</div>
            <div class="po-suboc muted">OC ${escapeHtml(ocNo)}</div>
          </div>

          <div class="po-meta-grid">
            <label>Status
              <select name="status" class="input">
                ${["OPEN","HOLD","PARTIAL","RECEIVED","CLOSED","CANCELLED"].map(s => {
                  const label = OC_STATUS_PT[s] || (s[0] + s.slice(1).toLowerCase());
                  return `<option value="${s}" ${String(po.status||"OPEN").toUpperCase()===s?"selected":""}>${escapeHtml(label)}</option>`;
                }).join("")}
              </select>
            </label>

            <label>Observação
              <input name="note" class="input" value="${escapeHtml(po.note||"")}" />
            </label>
          </div>

          ${po.linkedProductionOrderId ? `<div class="muted small">Gerada a partir de uma OP.</div>` : ""}

          <div id="poAddBox"></div>
        </div>

        <div class="po-subcard">
          <div class="row wrap" style="gap:8px; align-items:flex-end; justify-content:space-between">
            <div class="po-subsection">Itens da OC</div>
            <div class="muted small">* <b>Pedida</b> é a base do sistema. Para mudar o pedido final, use <b>Ajustar</b> (começa em 0).</div>
          </div>
          <div id="poLines"></div>
        </div>
      </div>
    `,
    onOpen: () => {
      const addBox = document.querySelector("#poAddBox");
      const linesBox = document.querySelector("#poLines");
      if (!addBox || !linesBox) return;

      addBox.innerHTML = renderAddRow();
      linesBox.innerHTML = renderTable();

      const bindTable = () => {
        // Ajustar (valor final)
        linesBox.querySelectorAll("input.poAdj").forEach(inp => {
          inp.addEventListener("blur", () => {
            const idx = Number(inp.dataset.idx);
            const u = inp.dataset.unit || "";
            const v = parseBRNumber(inp.value);
            const cl = clampByUnit(v, u);
            if (lines[idx]) lines[idx].qtyAdjusted = cl;
            inp.value = fmtInput(cl, u);
          });
        });

        // Excluir linha
        linesBox.querySelectorAll("button[data-del]").forEach(b => {
          b.addEventListener("click", () => {
            const idx = Number(b.dataset.del);
            lines.splice(idx, 1);
            linesBox.innerHTML = renderTable();
            bindTable();
          });
        });
      };

      const bindAddRow = () => {
        const sel = addBox.querySelector("#poAddItem");
        const unitInp = addBox.querySelector("#poAddUnit");
        const qtyInp = addBox.querySelector("#poAddQty");
        const addBtn = addBox.querySelector("#poAddBtn");

        const syncUnit = () => {
          const itemId = sel?.value || "";
          const u = itemId ? unitOf(itemId) : "";
          if (unitInp) unitInp.value = u || "";
          if (qtyInp) {
            qtyInp.setAttribute("inputmode", isEach(u) ? "numeric" : "decimal");
            const v = clampByUnit(parseBRNumber(qtyInp.value), u);
            qtyInp.value = fmtInput(v, u);
          }
        };

        sel?.addEventListener("change", syncUnit);
        qtyInp?.addEventListener("blur", syncUnit);
        syncUnit();

        addBtn?.addEventListener("click", () => {
          const itemId = sel?.value || "";
          if (!itemId) return;
          const u = unitOf(itemId);
          const qty = clampByUnit(parseBRNumber(qtyInp?.value || 0), u);
          if (!Number.isFinite(qty) || qty <= 0) return;

          const existing = lines.find(x => x.itemId === itemId);
          if (existing) {
            existing.qtyAdjusted = Number((Number(existing.qtyAdjusted || 0) + qty).toFixed(6));
          } else {
            lines.push({ itemId, qtyOrdered: qty, qtyAdjusted: 0, qtyReceived: 0 });
          }

          // limpa pra próxima inserção
          if (sel) sel.value = "";
          if (qtyInp) qtyInp.value = "0";
          if (unitInp) unitInp.value = "";
          syncUnit();

          linesBox.innerHTML = renderTable();
          bindTable();
        });
      };

      bindTable();
      bindAddRow();
    },
    onSubmit: async (fd) => {
      if (!lines.length) {
        alert("Adicione pelo menos 1 item na OC.");
        return false;
      }

      // Validação + faltantes
      const missing = [];
      for (const li of lines) {
        const base = Number(li.qtyOrdered || 0);
        const adj = Number(li.qtyAdjusted || 0);
        const u = unitOf(li.itemId);
        const final = clampByUnit(base + adj, u);
        const rec = clampByUnit(Number(li.qtyReceived || 0), u);

        if (!li.itemId) continue;
        if (!Number.isFinite(final) || final <= 0) {
          alert("A OC não pode ter itens com quantidade zero. Remova a linha se não for comprar.");
          return false;
        }
        if (final < rec - 1e-9) {
          alert("Não é possível ajustar o pedido final para menos do que já foi recebido.");
          return false;
        }

        // Se ajustou para menos (Ajustar negativo): faltante = -ajuste
        if (adj < -1e-9) {
          missing.push({ itemId: li.itemId, qty: clampByUnit(-adj, u) });
        }
      }

      const payload = {
        status: fd.get("status"),
        note: fd.get("note") || "",
        linkedProductionOrderId: po.linkedProductionOrderId || null,
        items: lines.map(li => ({
          itemId: li.itemId,
          qtyOrdered: Number(li.qtyOrdered || 0),
          qtyAdjusted: Number(li.qtyAdjusted || 0),
          qtyReceived: Number(li.qtyReceived || 0),
        })).filter(li => li.itemId),
      };

      let saved = null;
      if (mode === "new") {
        const r = await api("/api/mrp/purchase-orders", { method: "POST", body: JSON.stringify(payload) });
        saved = r?.order || null;
      } else {
        const r = await api(`/api/mrp/purchase-orders/${po.id}`, { method: "PUT", body: JSON.stringify(payload) });
        saved = r?.order || null;
      }

      // Se houve redução: oferecer gerar nova OC faltante
      if (missing.length) {
        const ocLabel = saved?.number ? pad6(saved.number) : (po.number ? pad6(po.number) : (saved?.id || po.id || ""));
        const msg = `Você ajustou para menos alguns itens.\n\nDeseja gerar uma nova OC com os faltantes (para não perder a necessidade)?\n\nNova OC será: "Gerado a partir de ajuste parcial faltante da OC ${ocLabel}"`;
        const make = confirm(msg);
        if (make) {
          await api("/api/mrp/purchase-orders", {
            method: "POST",
            body: JSON.stringify({
              note: `Gerado a partir de ajuste parcial faltante da OC ${ocLabel}`,
              linkedProductionOrderId: saved?.linkedProductionOrderId || po.linkedProductionOrderId || null,
              items: missing.map(m => ({ itemId: m.itemId, qtyOrdered: m.qty, qtyAdjusted: 0, qtyReceived: 0 })),
            })
          });
        } else {
          alert("Ajuste salvo. Atenção: ao reduzir a compra, pode faltar para alguma produção futura.");
        }
      }

      await loadAll();
      return true;
    }
  });
}

// --------- Purchase Orders Archived UI ---------
function openArchivedPurchaseOrdersModal() {
  openModal({
    title: "Ordens de Compra Arquivadas",
    subtitle: "As OCs arquivadas ficam fora da lista principal. Você pode imprimir, restaurar ou excluir permanentemente.",
    submitText: "Fechar",
    bodyHtml: `<div id="poArchivedBox" class="muted">Carregando...</div>`,
    onOpen: () => {
      (async () => {
        const box = document.querySelector("#poArchivedBox");
        if (!box) return;

        const itemsById = new Map((state.rawItems || []).map(i => [i.id, i]));

        const render = (orders) => {
          const rows = (orders || []).map(po => {
            const ocNo = po.number ? pad6(po.number) : "—";
            const dt = po.archivedAt || po.createdAt || "";
            const itemsTxt = (po.items || []).map(li => {
              const it = itemsById.get(li.itemId);
              const u = it?.unit || "";
              const finalOrd = (Number(li.qtyOrdered || 0) + Number(li.qtyAdjusted || 0));
              const ord = fmt(finalOrd);
              const rec = fmt(li.qtyReceived || 0);
              return `${escapeHtml(it?.name || "—")}: ${ord} ${escapeHtml(u)} <span class="muted small">(rec: ${rec})</span>`;
            }).join("<br/>");

            const status = OC_STATUS_PT[String(po.status || "OPEN").toUpperCase()] || String(po.status || "OPEN");
            const note = String(po.note || "");
            return `<tr>
              <td>${fmtDate(dt)}</td>
              <td><b>${escapeHtml(ocNo)}</b></td>
              <td>${escapeHtml(status)}</td>
              <td class="muted">${escapeHtml(note)}</td>
              <td class="muted small">${itemsTxt || "—"}</td>
              <td style="text-align:center">
                <div class="row wrap" style="gap:6px; justify-content:center">
                  <button type="button" class="btn secondary small" data-arch-act="print" data-id="${escapeHtml(po.id)}">Imprimir</button>
                  <button type="button" class="btn secondary small" data-arch-act="restore" data-id="${escapeHtml(po.id)}">Restaurar</button>
                  <button type="button" class="btn secondary small" data-arch-act="del" data-id="${escapeHtml(po.id)}">Excluir</button>
                </div>
              </td>
            </tr>`;
          }).join("");

          if (!rows) {
            return `<div class="muted">Nenhuma OC arquivada ainda.</div>`;
          }
          return `
            <div class="table-wrap" style="max-height:60vh; overflow:auto">
              <table class="table">
                <thead>
                  <tr>
                    <th>Data</th>
                    <th style="width:90px">OC</th>
                    <th style="width:120px">Status</th>
                    <th>Observação</th>
                    <th>Itens</th>
                    <th style="width:240px; text-align:center">Ações</th>
                  </tr>
                </thead>
                <tbody>${rows}</tbody>
              </table>
            </div>
          `;
        };

        const refresh = async () => {
          box.innerHTML = `<div class="muted">Carregando...</div>`;
          const res = await api("/api/mrp/purchase-orders/archived");
          const orders = Array.isArray(res.orders) ? res.orders : [];
          box.innerHTML = render(orders);

          box.querySelectorAll("button[data-arch-act]").forEach(btn => {
            btn.addEventListener("click", async (e) => {
              try { e.preventDefault(); e.stopPropagation(); e.stopImmediatePropagation(); } catch {}
              const id = btn.dataset.id;
              const act = btn.dataset.archAct;
              const po = orders.find(x => x.id === id);
              if (!po) return;

              if (act === "print") {
                return printPurchaseOrder(po);
              }
              if (act === "restore") {
                const ocNo = po.number ? pad6(po.number) : "—";
                const ok = confirm(`Restaurar a OC ${ocNo} para a lista principal?`);
                if (!ok) return;
                try {
                  await api(`/api/mrp/purchase-orders/archived/${id}/restore`, { method: "POST", body: JSON.stringify({}) });
                  await loadAll();
                  await refresh();
                } catch (err) {
                  console.error(err);
                  alert("Falha ao restaurar OC.");
                }
                return;
              }
              if (act === "del") {
                const ocNo = po.number ? pad6(po.number) : "—";
                const ok = confirm(`Excluir permanentemente a OC arquivada ${ocNo}? Isso remove também as entradas de estoque (recebimentos) geradas por ela.`);
                if (!ok) return;
                const re = await requestReauth({ reason: "Excluir OC arquivada" });
                if (!re) return;
                try {
                  await api(`/api/mrp/purchase-orders/archived/${id}`, { method: "DELETE", body: JSON.stringify({}) });
                  await loadAll();
                  await refresh();
                } catch (err) {
                  console.error(err);
                  alert("Falha ao excluir OC arquivada.");
                }
              }
            });
          });
        };

        try {
          await refresh();
        } catch (e) {
          console.error(e);
          box.innerHTML = `<div class="muted">Não foi possível carregar as OCs arquivadas.</div>`;
        }
      })();
    },
    onSubmit: async () => true,
  });
}



function openArchivedProductionOrdersModal() {
  openModal({
    title: "Ordens de Produção Arquivadas",
    subtitle: "As OPs arquivadas ficam fora da lista principal. Você pode imprimir, restaurar ou excluir permanentemente.",
    submitText: "Fechar",
    bodyHtml: `
      <div class="muted small" style="margin-bottom:8px">
        <b>Dica:</b> só é possível arquivar OPs com status <b>Encerrada</b> ou <b>Cancelada</b>.
      </div>
      <div class="table-wrap" style="max-height:420px; overflow:auto">
        <table class="table" id="opArchivedTable">
          <thead>
            <tr>
              <th>Data</th>
              <th>OP</th>
              <th>Receita</th>
              <th>Qtd</th>
              <th>Status</th>
              <th>Observação</th>
              <th style="text-align:center">Ações</th>
            </tr>
          </thead>
          <tbody><tr><td colspan="7" class="muted">Carregando...</td></tr></tbody>
        </table>
      </div>
    `,
    onOpen: () => openArchivedProductionOrders(),
  });
}

async function openArchivedProductionOrders() {
  try {
    const res = await api("/api/mrp/production-orders/archived");
    const orders = res.orders || [];
    const tbody = document.querySelector("#opArchivedTable tbody");
    if (!tbody) return;

    const rawById = new Map((state.rawItems || []).map(i => [i.id, i]));
    const fgById = new Map((state.fgItems || []).map(i => [i.id, i]));

    const rows = orders.map(op => {
      const opNo = op.number ? pad6(op.number) : "—";
      const st = OP_STATUS_PT[String(op.status || "ISSUED").toUpperCase()] || String(op.status || "ISSUED");
      const outId = op.planned?.produced?.itemId || op.produced?.itemId || null;
      const outUnit = fgById.get(outId)?.unit || "un";
      const qtyProd = Number(op.qtyToProduce || op.planned?.produced?.qty || 0);

      return `<tr>
        <td>${fmtDate(op.archivedAt || op.createdAt)}</td>
        <td><b>${escapeHtml(opNo)}</b></td>
        <td>${escapeHtml(op.recipeName || op.recipeId || "—")}</td>
        <td>${escapeHtml(fmt(qtyProd))} ${escapeHtml(outUnit)}</td>
        <td>${escapeHtml(st)}</td>
        <td class="muted">${escapeHtml(op.note || "")}</td>
        <td style="text-align:center">
          <div class="row wrap" style="gap:6px; justify-content:center">
            <button class="btn secondary small" data-opa="print" data-id="${escapeHtml(op.id)}">Imprimir</button>
            <button class="btn secondary small" data-opa="restore" data-id="${escapeHtml(op.id)}">Restaurar</button>
            <button class="btn secondary small" data-opa="del" data-id="${escapeHtml(op.id)}">Excluir</button>
          </div>
        </td>
      </tr>`;
    }).join("");

    tbody.innerHTML = rows || `<tr><td colspan="7" class="muted">Nenhuma OP arquivada ainda.</td></tr>`;

    tbody.querySelectorAll("button[data-opa]").forEach(btn => {
      btn.addEventListener("click", async () => {
        const id = btn.dataset.id;
        const act = btn.dataset.opa;
        const op = orders.find(o => o.id === id);
        if (!id || !act) return;

        if (act === "print" && op) return printProductionOrder(op);

        if (act === "restore") {
          const opNo = op?.number ? pad6(op.number) : "—";
          const ok = confirm(`Restaurar a OP ${opNo} para a lista principal?`);
          if (!ok) return;
          try {
            await api(`/api/mrp/production-orders/archived/${id}/restore`, { method: "POST", body: JSON.stringify({}) });
            await loadAll();
            openArchivedProductionOrders();
          } catch (e) {
            console.error(e);
            alert("Falha ao restaurar OP.");
          }
          return;
        }

        if (act === "del") {
          const opNo = op?.number ? pad6(op.number) : "—";
          const ok = confirm(`Excluir a OP ${opNo} permanentemente? Isso remove também os movimentos de estoque gerados (se houver).`);
          if (!ok) return;
          const re = await requestReauth({ reason: "Excluir OP Arquivada" });
          if (!re) return;
          try {
            await api(`/api/mrp/production-orders/archived/${id}`, { method: "DELETE", body: JSON.stringify({}) });
            await loadAll();
            openArchivedProductionOrders();
          } catch (e) {
            console.error(e);
            alert("Falha ao excluir OP arquivada.");
          }
        }
      });
    });
  } catch (e) {
    console.error(e);
    const tbody = document.querySelector("#opArchivedTable tbody");
    if (tbody) tbody.innerHTML = `<tr><td colspan="7" class="muted">Falha ao carregar OPs arquivadas.</td></tr>`;
  }
}




function openViewPurchaseOrder(po) {
  if (!po) return;

  const rawItems = (state.rawItems || []).slice().sort((a,b)=>String(a.name||"").localeCompare(String(b.name||""),"pt-BR"));
  const itemsById = new Map(rawItems.map(i => [i.id, i]));

  const isEach = (u) => fmtUnit(u) === "un";
  const clampByUnit = (n, u) => {
    let x = Number(n);
    if (!Number.isFinite(x)) x = 0;
    if (isEach(u)) x = Math.trunc(x);
    return Number(Number(x).toFixed(6));
  };

  const ocNo = po.number ? pad6(po.number) : (po.id ? String(po.id) : "—");
  const stRaw = String(po.status || "OPEN").toUpperCase();
  const st = stRaw === "CLOSED" ? "CLOSED" : stRaw;
  const stLabel = OC_STATUS_PT[st] || st;

  const lines = (po.items || []).map(li => {
    const it = itemsById.get(li.itemId);
    const u = it?.unit || "";
    const ordered = Number(li.qtyOrdered || 0);
    const adj = Number(li.qtyAdjusted || 0);
    const received = Number(li.qtyReceived || 0);
    const final = clampByUnit(ordered + adj, u);
    return { name: it?.name || "—", unit: u, ordered, adj, final, received };
  });

  const rows = lines.map(li => `
    <tr>
      <td>${escapeHtml(li.name)}</td>
      <td style="text-align:right">${fmt(li.ordered)} ${escapeHtml(li.unit)}</td>
      <td style="text-align:right">${fmt(li.adj)} ${escapeHtml(li.unit)}</td>
      <td style="text-align:right"><b>${fmt(li.final)} ${escapeHtml(li.unit)}</b></td>
      <td style="text-align:right">${fmt(li.received)} ${escapeHtml(li.unit)}</td>
    </tr>
  `).join("");

  const noteHtml = (po.note && String(po.note).trim()) ? escapeHtml(String(po.note)) : `<span class="muted">—</span>`;

  openModal({
    title: "Ordem de Compra",
    subtitle: `OC ${escapeHtml(ocNo)} • ${escapeHtml(stLabel)}`,
    submitText: "Fechar",
    cardClass: "wide po-compact po-view",
    bodyHtml: `
      <div class="muted" style="margin-bottom:8px"><b>Observação:</b> ${noteHtml}</div>
      <div class="muted small" style="margin-bottom:10px">
        <b>Pedido final</b> = <b>Pedida</b> + <b>Ajustar</b>. (Ajustar é um delta e começa em 0.)
      </div>
      <div class="table-wrap" style="max-height:60vh; overflow:auto">
        <table class="table">
          <thead><tr><th>Item</th><th style="text-align:right">Pedida</th><th style="text-align:right">Ajustar</th><th style="text-align:right">Final</th><th style="text-align:right">Recebida</th></tr></thead>
          <tbody>${rows || `<tr><td colspan="5" class="muted">Sem itens.</td></tr>`}</tbody>
        </table>
      </div>
      <div class="muted small" style="margin-top:10px">
        Esta visualização é somente leitura. Para editar a OC, use o módulo <b>Ordem de Compra</b>.
      </div>
    `,
    onSubmit: () => true
  });
}

function openReceivePurchaseOrder(po) {
  const rawItems = (state.rawItems || []).slice();
  const itemsById = new Map(rawItems.map(i => [i.id, i]));

  const isEach = (u) => fmtUnit(u) === "un";
  // IMPORTANT: normalizamos as quantidades que o usuário vê/edita para evitar
  // falsos "faltantes" por diferenças de casas decimais ocultas.
  // - UN = un -> inteiro
  // - demais -> 3 casas decimais
  const normByUnit = (n, u) => {
    let x = Number(n);
    if (!Number.isFinite(x)) x = 0;
    if (x < 0) x = 0;
    if (isEach(u)) return Math.trunc(x);
    // 3 decimais (alinhado ao que exibimos no UI)
    return Number((Math.round(x * 1000) / 1000).toFixed(3));
  };
  const fmtInput = (n, u) => {
    const x = Number(n);
    if (!Number.isFinite(x) || x === 0) return "0";
    if (isEach(u)) return String(Math.trunc(x));
    return x.toLocaleString("pt-BR", { maximumFractionDigits: 3 });
  };

  const rows = (po.items || []).map(li => {
    const it = itemsById.get(li.itemId);
    const u = it?.unit || "";
    const rawFinalOrd = (Number(li.qtyOrdered || 0) + Number(li.qtyAdjusted || 0));
    const finalOrd = normByUnit(rawFinalOrd, u);
    const alreadyRec = normByUnit(Number(li.qtyReceived || 0), u);
    const remaining = normByUnit(Math.max(0, finalOrd - alreadyRec), u);
    const inputMode = isEach(u) ? "numeric" : "decimal";
    return `<tr>
      <td>${escapeHtml(it?.name || "—")}</td>
      <td class="muted" style="text-align:right">${fmt(finalOrd)} ${escapeHtml(u)}<div class="muted small">rec: ${fmt(alreadyRec)} • falta: ${fmt(remaining)}</div></td>
      <td style="width:220px">
        <input class="input recvQty" data-itemid="${escapeHtml(li.itemId)}" data-unit="${escapeHtml(u)}" data-expected="${escapeHtml(String(remaining))}" inputmode="${inputMode}" value="${escapeHtml(fmtInput(remaining, u))}" />
      </td>
      <td class="muted" style="width:70px">${escapeHtml(u)}</td>
    </tr>`;
  }).join("");

  openModal({
    title: "Receber Ordem de Compra",
    subtitle: po.number ? `OC ${pad6(po.number)}` : `OC ${po.id}`,
    submitText: "Receber",
    cardClass: "wide po-compact po-recv",
    bodyHtml: `
      <label>Observação (opcional)
        <input name="note" class="input" placeholder="Ex.: NF 123 / fornecedor..." />
      </label>
      <div class="table-wrap" style="margin-top:10px">
        <table class="table">
          <thead><tr><th>Item</th><th style="text-align:right">Pedido (final)</th><th>Receber agora</th><th>UN</th></tr></thead>
          <tbody>${rows || `<tr><td colspan="4" class="muted">Sem itens.</td></tr>`}</tbody>
        </table>
      </div>
      <div class="muted small" style="margin-top:10px; line-height:1.5">
        Ao receber, o sistema faz <b>entrada no estoque (MP)</b>. Se a OC estiver vinculada a uma OP, ela pode sair de <b>Em espera</b> para pronta automaticamente.
      </div>
    `,
    onOpen: () => {
      // normaliza ao sair do campo (pt-BR vírgula) e aplica regra de inteiro quando UN = un
      document.querySelectorAll(".recvQty").forEach(inp => {
        inp.addEventListener("blur", () => {
          const u = inp.dataset.unit || "";
          const v = normByUnit(parseBRNumber(inp.value), u);
          inp.value = fmtInput(v, u);
        });
      });
    },
    onSubmit: async (fd) => {
      const note = fd.get("note") || "";

      const itemsToReceive = [];
      const missing = [];

      document.querySelectorAll(".recvQty").forEach(inp => {
        const itemId = inp.dataset.itemid;
        const u = inp.dataset.unit || "";
        const expected = normByUnit(Number(inp.dataset.expected || 0), u);
        const qty = normByUnit(parseBRNumber(inp.value), u);
        if (!itemId) return;

        // Recebe (gera entrada) somente se > 0
        if (Number.isFinite(qty) && qty > 0) itemsToReceive.push({ itemId, qty });

        // Se recebeu menos do que o esperado (pedido final ainda pendente), calcula faltante
        if (Number.isFinite(expected) && expected > 0 && Number.isFinite(qty) && qty < expected - 1e-9) {
          missing.push({ itemId, qty: normByUnit(expected - qty, u) });
        }
      });

      if (!itemsToReceive.length) {
        alert('Informe pelo menos 1 quantidade para receber.');
        return false;
      }

      let finalize = false;
      let createMissing = false;

      if (missing.length) {
        const ocLabel = po.number ? pad6(po.number) : (po.id ? String(po.id) : "—");
        const msg = `Você está recebendo menos do que o pedido final em alguns itens.\n\nDeseja gerar uma nova OC com os faltantes?\n\nNova OC será: "Gerado a partir de recebimento parcial faltante da OC ${ocLabel}"`;
        createMissing = confirm(msg);

        // ✅ Regra atual: só fecha (RECEIVED) automaticamente se gerar nova OC faltante.
        // Se NÃO gerar, a OC permanece em aberto como PARCIAL (aguardando chegada do restante).
        finalize = !!createMissing;

        if (!createMissing) {
          alert('Atenção: você recebeu menos do que o pedido final. Isso pode gerar falta de matéria-prima e atrasar a produção. A OC ficará em aberto como PARCIAL.');
        }
      }

      try {
        await api(`/api/mrp/purchase-orders/${po.id}/receive`, {
          method: "POST",
          body: JSON.stringify({ items: itemsToReceive, note, finalize }),
        });
      } catch (e) {
        console.error(e);
        alert('Falha ao receber a OC. Veja console.');
        return false;
      }

      if (missing.length && createMissing) {
        const ocLabel = po.number ? pad6(po.number) : (po.id ? String(po.id) : "—");
        await api("/api/mrp/purchase-orders", {
          method: "POST",
          body: JSON.stringify({
            note: `Gerado a partir de recebimento parcial faltante da OC ${ocLabel}`,
            linkedProductionOrderId: po.linkedProductionOrderId || null,
            items: missing.map(m => ({ itemId: m.itemId, qtyOrdered: m.qty, qtyAdjusted: 0, qtyReceived: 0 })),
          })
        });
      }

      await loadAll();
      return true;
    }
  });
}

// ---------- Loaders ----------
async function loadInventory() {
  const itemsRes = await api(`/api/inventory/items?type=${state.stockMode}`);
  state.items = itemsRes.items;

  // Mantém caches por tipo para o painel de estoque mínimo
  if (state.stockMode === "raw") state.rawItems = itemsRes.items;
  if (state.stockMode === "fg") state.fgItems = itemsRes.items;
  try { setMinStockStatus(state.rawItems, state.fgItems); } catch (_) {}

  const mvRes = await api(`/api/inventory/movements?type=${state.stockMode}&limit=80`);
  state.movements = mvRes.movements;

  renderItems();
  renderMovements();
}

async function loadMRP() {
  const [recRes, rawRes, fgRes] = await Promise.all([
    api("/api/mrp/recipes"),
    api("/api/inventory/items?type=raw"),
    api("/api/inventory/items?type=fg"),
  ]);

  state.recipes = recRes.recipes;
  state.rawItems = rawRes.items;
  state.fgItems = fgRes.items;

  // Atualiza status de estoque mínimo (usa os dados recém carregados)
  try { setMinStockStatus(state.rawItems, state.fgItems); } catch (_) {}

  // ensure selection still exists
  if (state.selectedRecipeId && !state.recipes.find((r) => r.id === state.selectedRecipeId)) {
    state.selectedRecipeId = state.recipes[0]?.id || null;
  }
  renderPFBomList();
  renderSimulate();
}

async function loadOps() {
  const res = await api("/api/mrp/production-orders");
  state.ops = res.orders;
  renderOps();
}

async function loadPurchaseOrders() {
  const res = await api("/api/mrp/purchase-orders");
  state.purchaseOrders = res.orders;
  renderPurchaseOrders();
}

async function loadUnits() {
  const res = await api("/api/units");
  state.units = Array.isArray(res.units) ? res.units : [];
}

// ---------- Estoque mínimo (MP/PF) ----------
const _minStockMsg = document.querySelector('#minStockMsg');
const _btnGenOcMin = document.querySelector('#btnGenOcMin');
const _btnGenOpMin = document.querySelector('#btnGenOpMin');

function _num(v){
  const n = Number(v);
  return Number.isFinite(n) ? n : 0;
}

function computeMinStockBuckets(items){
  const list = Array.isArray(items) ? items : [];
  const tracked = list.filter(it => _num(it.minStock) > 0);
  const atOrBelow = tracked.filter(it => _num(it.currentStock) <= _num(it.minStock) + 1e-9);
  const below = tracked.filter(it => _num(it.currentStock) < _num(it.minStock) - 1e-9);
  return { tracked, atOrBelow, below };
}

function setMinStockStatus(rawItems, fgItems){
  const raw = computeMinStockBuckets(rawItems);
  const fg = computeMinStockBuckets(fgItems);
  state.minStockStatus = {
    rawAtOrBelow: raw.atOrBelow,
    rawBelow: raw.below,
    fgAtOrBelow: fg.atOrBelow,
    fgBelow: fg.below,
  };
  updateMinStockUI();
}

function updateMinStockUI(){
  const s = state.minStockStatus || { rawAtOrBelow: [], rawBelow: [], fgAtOrBelow: [], fgBelow: [] };
  const rawAt = s.rawAtOrBelow || [];
  const rawBelow = s.rawBelow || [];
  const fgAt = s.fgAtOrBelow || [];
  const fgBelow = s.fgBelow || [];

  if (_btnGenOcMin) _btnGenOcMin.disabled = (rawBelow.length === 0) || !canPerm('oc');
  if (_btnGenOpMin) _btnGenOpMin.disabled = (fgBelow.length === 0) || !canPerm('op');

  if (_minStockMsg){
    const mpPart = rawAt.length
      ? `MP: ${rawAt.length} no mínimo (${rawBelow.length} abaixo)`
      : 'MP: ok';
    const pfPart = fgAt.length
      ? `PF: ${fgAt.length} no mínimo (${fgBelow.length} abaixo)`
      : 'PF: ok';

    const extras = [];
    if (rawBelow.length){
      const codes = rawBelow.slice(0, 5).map(it => it.code || it.name || '').filter(Boolean);
      if (codes.length) extras.push(`MP abaixo: ${codes.join(', ')}${rawBelow.length > codes.length ? '…' : ''}`);
    }
    if (fgBelow.length){
      const codes = fgBelow.slice(0, 5).map(it => it.code || it.name || '').filter(Boolean);
      if (codes.length) extras.push(`PF abaixo: ${codes.join(', ')}${fgBelow.length > codes.length ? '…' : ''}`);
    }
    _minStockMsg.textContent = [mpPart, pfPart, ...extras].join(' • ');
  }
}

async function refreshMinStockStatus({ force = true } = {}){
  // Se não for forçado e já existe cache, só renderiza.
  if (!force && Array.isArray(state.rawItems) && Array.isArray(state.fgItems) && (state.rawItems.length || state.fgItems.length)) {
    setMinStockStatus(state.rawItems, state.fgItems);
    return;
  }
  const [rawRes, fgRes] = await Promise.all([
    api('/api/inventory/items?type=raw'),
    api('/api/inventory/items?type=fg'),
  ]);
  // atualiza caches globais para uso em outros módulos
  state.rawItems = Array.isArray(rawRes.items) ? rawRes.items : [];
  state.fgItems = Array.isArray(fgRes.items) ? fgRes.items : [];
  setMinStockStatus(state.rawItems, state.fgItems);
}

async function generateOCFromMinStock(){
  if (!canPerm('oc')) { alert('Sem permissão para gerar OC (Ordens de Compra).'); return; }
  await refreshMinStockStatus({ force: true });
  const below = (state.minStockStatus?.rawBelow || []).slice();
  if (!below.length){
    alert('Nenhuma Matéria Prima abaixo do estoque mínimo.');
    return;
  }
  const items = below
    .map(it => ({
      itemId: String(it.id),
      qtyOrdered: Math.max(0, _num(it.minStock) - _num(it.currentStock)),
    }))
    .filter(x => x.itemId && x.qtyOrdered > 0);

  if (!items.length){
    alert('Nenhuma Matéria Prima abaixo do estoque mínimo.');
    return;
  }

  const ok = confirm(`Gerar uma OC de estoque mínimo com ${items.length} item(ns), para repor até o mínimo?`);
  if (!ok) return;

  try {
    const note = `Estoque mínimo: OC gerada automaticamente (${new Date().toLocaleString('pt-BR')}).`;
    const res = await api('/api/mrp/purchase-orders', {
      method: 'POST',
      body: JSON.stringify({ note, items }),
    });
    await loadPurchaseOrders();
    await refreshMinStockStatus({ force: true });
    showTab('compras');
    const n = res?.order?.number;
    alert(`OC ${pad6(n || 0)} gerada com sucesso.`);
  } catch (e) {
    console.error(e);
    alert('Falha ao gerar OC de estoque mínimo.');
  }
}

async function generateOPFromMinStock(){
  if (!canPerm('op')) { alert('Sem permissão para gerar OP (Ordens de Produção).'); return; }
  await refreshMinStockStatus({ force: true });
  const below = (state.minStockStatus?.fgBelow || []).slice();
  if (!below.length){
    alert('Nenhum Produto Final abaixo do estoque mínimo.');
    return;
  }

  // Carrega receitas (para mapear PF -> Receita)
  let recipes = Array.isArray(state.recipes) && state.recipes.length ? state.recipes : [];
  try {
    const recRes = await api('/api/mrp/recipes');
    recipes = Array.isArray(recRes.recipes) ? recRes.recipes : recipes;
    state.recipes = recipes;
  } catch (_) {}

  // Agrupa por receita (um OP por receita)
  const byRecipe = new Map();
  const missing = [];
  for (const pf of below){
    const pfId = String(pf.id);
    const pfCode = String(pf.code || '').trim().toUpperCase();
    const deficit = Math.max(0, _num(pf.minStock) - _num(pf.currentStock));
    if (deficit <= 0) continue;

    const r = recipes.find(x => String(x.outputItemId || x.productId || '') === pfId)
      || (pfCode ? recipes.find(x => String(x.code || '').trim().toUpperCase() === pfCode) : null);

    if (!r){
      missing.push(pfCode || pf.name || pfId);
      continue;
    }

    const key = String(r.id);
    const cur = byRecipe.get(key) || { recipeId: key, recipeName: r.name, totalQty: 0 };
    cur.totalQty += deficit;
    byRecipe.set(key, cur);
  }

  const list = Array.from(byRecipe.values()).filter(x => x.totalQty > 0);
  if (!list.length){
    alert('Não foi possível mapear os PFs abaixo do mínimo para uma Receita (BOM).');
    return;
  }

  const ok = confirm(`Gerar ${list.length} OP(s) de estoque mínimo para repor PF até o mínimo?\n\nObs: a OP será criada SEM gerar OC automática (mesmo que falte MP).`);
  if (!ok) return;

  const created = [];
  try {
    for (const x of list){
      const note = `Estoque mínimo: OP gerada automaticamente (${new Date().toLocaleString('pt-BR')}).`;
      const res = await api('/api/mrp/production-orders', {
        method: 'POST',
        body: JSON.stringify({
          recipeId: x.recipeId,
          qtyToProduce: x.totalQty,
          note,
          allowInsufficient: true,
          createPurchaseOrder: false,
        }),
      });
      if (res?.order?.number) created.push(res.order.number);
    }

    await loadOps();
    await refreshMinStockStatus({ force: true });
    showTab('ops');

    const msg = [
      created.length ? `OPs geradas: ${created.map(n => pad6(n)).join(', ')}` : 'OPs geradas.',
      missing.length ? `Sem receita/BOM: ${missing.join(', ')}` : '',
    ].filter(Boolean).join('\n');
    alert(msg);
  } catch (e) {
    console.error(e);
    alert('Falha ao gerar OP de estoque mínimo.');
  }
}

if (_btnGenOcMin) _btnGenOcMin.addEventListener('click', generateOCFromMinStock);
if (_btnGenOpMin) _btnGenOpMin.addEventListener('click', generateOPFromMinStock);

// ---------- Pedidos de Venda (Pontos/Freezers + Venda rápida) ----------
function salesPointName(pointId){
  const p = (state.salesPoints || []).find(x => String(x.id) === String(pointId));
  return p ? (p.name || p.code || String(pointId)) : String(pointId || "—");
}

function salesStatusPt(s){
  const v = String(s || "").toUpperCase();
  if (v === "DISPATCHED") return "Despachado";
  if (v === "OPS_CREATED") return "OPs geradas";
  if (v === "CANCELLED") return "Cancelado";
  if (v === "DONE") return "Concluída";
  return "Aberto";
}

async function loadSalesAll(){
  // garante PFs atualizados para selects
  try {
    const fgRes = await api("/api/inventory/items?type=fg");
    state.fgItems = Array.isArray(fgRes.items) ? fgRes.items : [];
  } catch (_) {}

  const [pRes, oRes] = await Promise.all([
    api("/api/sales/points"),
    api("/api/sales/orders"),
  ]);
  state.salesPoints = Array.isArray(pRes.points) ? pRes.points : [];
  state.salesOrders = Array.isArray(oRes.orders) ? oRes.orders : [];
  renderQuickSalePfOptions();
}

function setSalesMode(mode){
  const m = (mode === "orders" || mode === "quick") ? mode : "points";
  state.salesMode = m;
  renderSalesCurrent();
}

function renderSalesCurrent(){
  const vPoints = $("#salesPointsView");
  const vOrders = $("#salesOrdersView");
  const vQuick = $("#salesQuickView");
  if (!vPoints || !vOrders || !vQuick) return;

  show(vPoints, state.salesMode === "points");
  show(vOrders, state.salesMode === "orders");
  show(vQuick, state.salesMode === "quick");

  const bP = $("#salesModePoints");
  const bO = $("#salesModeOrders");
  const bQ = $("#salesModeQuick");
  const setBtn = (btn, active) => {
    if (!btn) return;
    btn.className = active ? "btn primary" : "btn secondary";
  };
  setBtn(bP, state.salesMode === "points");
  setBtn(bO, state.salesMode === "orders");
  setBtn(bQ, state.salesMode === "quick");

  if (state.salesMode === "points") renderSalesPoints();
  if (state.salesMode === "orders") renderSalesOrders();
  if (state.salesMode === "quick") { renderQuickSaleHint(); renderQuickSalesList(); }
}

function renderSalesPoints(){
  const tb = $("#salesPointsTable tbody");
  if (!tb) return;

  const selAll = $("#salesPointsSelectAll");
  if (selAll) selAll.checked = false;

  const q = String($("#salesPointSearch")?.value || "").trim().toLowerCase();
  const rows = (state.salesPoints || [])
    .filter(p => {
      if (!q) return true;
      const hay = `${p.code||""} ${p.name||""} ${p.address||""} ${p.note||""}`.toLowerCase();
      return hay.includes(q);
    })
    .sort((a,b) => String(a.code||"").localeCompare(String(b.code||""), 'pt-BR'));

  // sane selection (apenas o que está visível nesta lista)
  let selected = new Set(Array.isArray(state.salesPointSelectedIds) ? state.salesPointSelectedIds.map(x => String(x)) : []);
  const visibleIds = new Set(rows.map(r => String(r.id)));
  selected = new Set([...selected].filter(id => visibleIds.has(id)));
  state.salesPointSelectedIds = [...selected];

  if (!rows.length) {
    tb.innerHTML = `<tr><td colspan="6" class="muted">Nenhum ponto cadastrado.</td></tr>`;
    return;
  }

  tb.innerHTML = rows.map(p => {
    const code = escapeHtml(p.code || "—");
    const name = escapeHtml(p.name || "—");
    const addr = escapeHtml(p.address || "");
    const note = escapeHtml(p.note || "");
    const checked = selected.has(String(p.id));
    return `
      <tr>
        <td style="text-align:center"><input type="checkbox" data-sp="1" data-id="${escapeHtml(p.id)}" ${checked?"checked":""} /></td>
        <td><b>${code}</b></td>
        <td>${name}</td>
        <td>${addr}</td>
        <td class="muted">${note}</td>
        <td style="text-align:center">
          <button class="btn secondary small" data-act="stock" data-id="${escapeHtml(p.id)}">Estoque</button>
          <button class="btn secondary small" data-act="edit" data-id="${escapeHtml(p.id)}">Editar</button>
          <button class="btn danger small" data-act="del" data-id="${escapeHtml(p.id)}">Excluir</button>
        </td>
      </tr>
    `;
  }).join("");

  tb.querySelectorAll('button[data-act]').forEach(btn => {
    btn.addEventListener('click', async () => {
      const id = btn.dataset.id;
      const act = btn.dataset.act;
      const p = (state.salesPoints || []).find(x => String(x.id) === String(id));
      if (!p) return;
      if (act === 'stock') return openSalesPointStock(p);
      if (act === 'edit') return openEditSalesPoint(p);
      if (act === 'del') return deleteSalesPoint(p);
    });
  });

  // checkbox selection
  tb.querySelectorAll('input[data-sp="1"]').forEach(ch => {
    ch.addEventListener('change', () => {
      const id = String(ch.dataset.id||"");
      let set = new Set(Array.isArray(state.salesPointSelectedIds) ? state.salesPointSelectedIds.map(x => String(x)) : []);
      if (ch.checked) set.add(id); else set.delete(id);
      state.salesPointSelectedIds = [...set];
      if (selAll) {
        const visible = rows.map(r => String(r.id));
        selAll.checked = visible.length > 0 && visible.every(v => set.has(v));
      }
    });
  });

  if (selAll) {
    const visible = rows.map(r => String(r.id));
    const set = new Set(Array.isArray(state.salesPointSelectedIds) ? state.salesPointSelectedIds.map(x => String(x)) : []);
    selAll.checked = visible.length > 0 && visible.every(v => set.has(v));
  }
}

async function printSalesPointsReport(pointIds){
  const ids = (pointIds || []).map(x => String(x)).filter(Boolean);
  if (!ids.length) return;

  const points = ids
    .map(id => (state.salesPoints || []).find(p => String(p.id) === String(id)))
    .filter(Boolean);

  if (!points.length) return;

  const blocks = [];
  try {
    for (const p of points) {
      const r = await api(`/api/sales/points/${encodeURIComponent(p.id)}/stock`);
      blocks.push({ point: p, items: (r.items || []) });
    }
  } catch (e) {
    console.error(e);
    alert('Falha ao gerar relatório do ponto.');
    return;
  }

  const now = fmtDate(new Date().toISOString());
  const docHtml = `
<!doctype html>
<html>
<head>
  <meta charset="utf-8" />
  <title>Relatório por Ponto</title>
  <style>
    @page { size: Letter landscape; margin: 10mm; }
    body { font-family: Arial, sans-serif; font-size: 12px; color: #111; }
    h1 { font-size: 16px; margin: 0 0 6px 0; }
    .top { display:flex; justify-content:space-between; align-items:flex-end; gap: 12px; }
    .muted { color:#555; font-size: 11px; }
    .point { margin-top: 14px; page-break-after: always; }
    .point:last-child { page-break-after: auto; }
    .ph { font-size: 14px; font-weight: 700; margin: 0 0 6px 0; }
    .meta { margin: 0 0 8px 0; }
    table { width: 100%; border-collapse: collapse; }
    th, td { border: 1px solid #bbb; padding: 4px 6px; }
    th { background: #f3f3f3; text-align: left; }
    td.num { text-align: right; white-space: nowrap; }
  </style>
</head>
<body>
  <div class="top">
    <h1>Relatório de Estoque por Ponto (Freezer)</h1>
    <div class="muted">Gerado em: ${escapeHtml(now)}</div>
  </div>

  ${blocks.map(b => {
    const p = b.point || {};
    const items = b.items || [];
    const addr = escapeHtml(p.address || '');
    const note = escapeHtml(p.note || '');
    const title = `${p.code || ''} - ${p.name || ''}`;
    return `
      <div class="point">
        <div class="ph">${escapeHtml(title)}</div>
        <div class="meta muted">
          ${addr ? `Endereço: ${addr}<br/>` : ''}
          ${note ? `Obs: ${note}<br/>` : ''}
        </div>

        <table>
          <thead>
            <tr>
              <th style="width:110px">PF</th>
              <th>Descrição</th>
              <th style="width:90px">UN</th>
              <th style="width:120px">Qtd</th>
            </tr>
          </thead>
          <tbody>
            ${items.length ? items.map(it => `
              <tr>
                <td><b>${escapeHtml(it.code || '')}</b></td>
                <td>${escapeHtml(it.name || '')}</td>
                <td>${escapeHtml(it.unit || '')}</td>
                <td class="num">${escapeHtml(String(it.qty ?? ''))}</td>
              </tr>
            `).join('') : `<tr><td colspan="4" class="muted">Sem itens em estoque neste ponto.</td></tr>`}
          </tbody>
        </table>
      </div>
    `;
  }).join('')}
</body>
</html>`;

  const w = window.open("", "_blank");
  if (!w) {
    alert('Bloqueio de pop-up: permita pop-ups para imprimir.');
    return;
  }
  w.document.open();
  w.document.write(docHtml);
  w.document.close();
  w.focus();
  setTimeout(() => {
    try { w.print(); } catch(e) {}
  }, 250);
}

function printSingleSelectedSalesPoint(){
  const ids = Array.isArray(state.salesPointSelectedIds) ? state.salesPointSelectedIds.map(x => String(x)) : [];
  if (ids.length !== 1) {
    alert('Selecione exatamente 1 ponto para imprimir.');
    return;
  }
  printSalesPointsReport(ids);
}

function printSelectedSalesPoints(){
  const ids = Array.isArray(state.salesPointSelectedIds) ? state.salesPointSelectedIds.map(x => String(x)) : [];
  if (!ids.length) {
    alert('Selecione 1 ou mais pontos para imprimir.');
    return;
  }
  printSalesPointsReport(ids);
}

function getFilteredQuickSalesRows(){
  const q = String(document.querySelector('#quickSalesSearch')?.value || '').trim().toLowerCase();
  const rows = (state.salesOrders || [])
    .filter(o => String(o.type||'') === 'QUICK')
    .filter(o => {
      const isArch = !!o.archived;
      return state.quickSalesShowArchived ? isArch : !isArch;
    })
    .filter(o => {
      if (!q) return true;
      const items = (o.items || []).map(it => `${it.code||''} ${it.name||''}`).join(' ');
      const hay = `${fmtSalesOrderCode(o)} ${o.channel||''} ${items}`.toLowerCase();
      return hay.includes(q);
    })
    .sort((a,b) => String(b.createdAt||'').localeCompare(String(a.createdAt||'')));
  return rows;
}

function formatQuickSaleItemsInline(o){
  const items = Array.isArray(o?.items) ? o.items : [];
  if (!items.length) return '—';
  return items.map(it => {
    const code = String(it.code || '').trim();
    const qty = fmtNum(it.qty);
    const unit = String(it.unit || '').trim();
    return `${code}${qty ? ` x ${qty}` : ''}${unit ? ` ${unit}` : ''}`.trim();
  }).join(' • ');
}

function printQuickSalesLandscape(rows, { title = 'Relatório de Venda Rápida (PVR)' } = {}){
  rows = Array.isArray(rows) ? rows : [];
  if (!rows.length) {
    alert('Nada para imprimir.');
    return;
  }

  const now = new Date();
  const headerDate = now.toLocaleString('pt-BR');

  const docHtml = `<!doctype html>
<html>
<head>
<meta charset="utf-8" />
<title>${escapeHtml(title)}</title>
<style>
  @page { size: letter landscape; margin: 12mm; }
  body { font-family: Arial, sans-serif; font-size: 12px; color: #111; }
  h1 { font-size: 16px; margin: 0 0 4px 0; }
  .meta { color: #555; margin-bottom: 10px; }
  table { width: 100%; border-collapse: collapse; }
  th, td { border: 1px solid #ccc; padding: 6px 8px; vertical-align: top; }
  th { background: #f4f4f4; text-align: left; }
  .col-date { width: 165px; white-space: nowrap; }
  .col-code { width: 110px; white-space: nowrap; }
  .col-channel { width: 120px; }
  .items { font-size: 11px; line-height: 1.25; }
</style>
</head>
<body>
  <h1>${escapeHtml(title)}</h1>
  <div class="meta">Gerado em: ${escapeHtml(headerDate)} • Total: ${rows.length}</div>
  <table>
    <thead>
      <tr>
        <th class="col-date">Data/Hora</th>
        <th class="col-code">Código</th>
        <th class="col-channel">Canal</th>
        <th>Itens</th>
      </tr>
    </thead>
    <tbody>
      ${rows.map(o => `
        <tr>
          <td class="col-date">${escapeHtml(fmtDate(o.createdAt))}</td>
          <td class="col-code"><b>${escapeHtml(fmtSalesOrderCode(o))}</b></td>
          <td class="col-channel">${escapeHtml(String(o.channel || '—'))}</td>
          <td class="items">${escapeHtml(formatQuickSaleItemsInline(o))}</td>
        </tr>
      `).join('')}
    </tbody>
  </table>
</body>
</html>`;

  const w = window.open("", "_blank");
  if (!w) {
    alert('Bloqueio de pop-up: permita pop-ups para imprimir.');
    return;
  }
  w.document.open();
  w.document.write(docHtml);
  w.document.close();
  w.focus();
  setTimeout(() => {
    try { w.print(); } catch(e) {}
  }, 250);
}

function printQuickSalesVisible(){
  const rows = getFilteredQuickSalesRows();
  printQuickSalesLandscape(rows, { title: 'Relatório de Venda Rápida (PVR) — Visível' });
}

function printQuickSalesSelected(){
  const selectedIds = new Set(Array.isArray(state.quickSelectedIds) ? state.quickSelectedIds.map(x => String(x)) : []);
  if (!selectedIds.size) {
    alert('Selecione pelo menos 1 venda para imprimir.');
    return;
  }
  const visible = getFilteredQuickSalesRows();
  const rows = visible.filter(o => selectedIds.has(String(o.id)));
  if (!rows.length) {
    alert('Nenhuma venda selecionada está visível no filtro atual.');
    return;
  }
  printQuickSalesLandscape(rows, { title: 'Relatório de Venda Rápida (PVR) — Selecionadas' });
}




async function openSalesPointStock(point){
  try {
    const res = await api(`/api/sales/points/${encodeURIComponent(point.id)}/stock`);
    const items = Array.isArray(res.items) ? res.items : [];
    const body = `
      <div class="muted" style="margin-bottom:8px">${escapeHtml(point.name || point.code || "Ponto")}</div>
      <div class="table-wrap" style="max-height:420px; overflow:auto;">
        <table class="table">
          <thead>
            <tr>
              <th style="width:70px">COD</th>
              <th>Produto Final</th>
              <th style="width:110px; text-align:right">Qtde</th>
              <th style="width:60px">UN</th>
            </tr>
          </thead>
          <tbody>
            ${items.length ? items.map(it => `
              <tr>
                <td><b>${escapeHtml(it.code||"")}</b></td>
                <td>${escapeHtml(it.name||"")}</td>
                <td style="text-align:right">${escapeHtml(fmtNum(it.qty))}</td>
                <td>${escapeHtml(String(it.unit||""))}</td>
              </tr>
            `).join("") : `<tr><td colspan="4" class="muted">Sem movimentações.</td></tr>`}
          </tbody>
        </table>
      </div>
    `;
    openModal({ title: "Estoque do ponto", subtitle: "Controle local (por despacho).", submitText: "Fechar", bodyHtml: body, onSubmit: () => true });
  } catch (e) {
    console.error(e);
    alert('Falha ao carregar estoque do ponto.');
  }
}

function openNewSalesPoint(){
  openModal({
    title: 'Novo Ponto (Freezer)',
    subtitle: 'Ex.: Academia Mega Fitness 3',
    submitText: 'Criar',
    bodyHtml: `
      <label>Código (opcional)
        <input class="input" name="code" placeholder="P001" />
      </label>
      <label>Nome
        <input class="input" name="name" required placeholder="Academia Mega Fitness 3" />
      </label>
      <label>Endereço
        <input class="input" name="address" placeholder="Rua..., Bairro..." />
      </label>
      <label>Observação
        <input class="input" name="note" placeholder="Freezer 1 • chave..." />
      </label>
    `,
    onSubmit: async (fd) => {
      const code = String(fd.get('code')||"").trim();
      const name = String(fd.get('name')||"").trim();
      const address = String(fd.get('address')||"").trim();
      const note = String(fd.get('note')||"").trim();
      if (!name) return false;

      if (code && !/^P?\s*\d{1,6}$/i.test(code)) {
        alert('Código inválido. Ex.: P001');
        return false;
      }

      try {
        await api('/api/sales/points', { method:'POST', body: JSON.stringify({ code, name, address, note }) });
      } catch (e) {
        if (e && e.status === 409) {
          alert('Esse código já existe em outro ponto.');
          return false;
        }
        console.error(e);
        alert('Falha ao criar ponto.');
        return false;
      }

      await loadSalesAll();
      renderSalesCurrent();
      return true;
    }
  });
}


function openEditSalesPoint(point){
  openModal({
    title: `Editar ponto ${point.code || ''}`,
    subtitle: point.name || '',
    submitText: 'Salvar',
    bodyHtml: `
      <label>Nome
        <input class="input" name="name" required value="${escapeHtml(point.name||"")}" />
      </label>
      <label>Endereço
        <input class="input" name="address" value="${escapeHtml(point.address||"")}" />
      </label>
      <label>Observação
        <input class="input" name="note" value="${escapeHtml(point.note||"")}" />
      </label>
    `,
    onSubmit: async (fd) => {
      const name = String(fd.get('name')||"").trim();
      const address = String(fd.get('address')||"").trim();
      const note = String(fd.get('note')||"").trim();
      if (!name) return false;
      await api(`/api/sales/points/${encodeURIComponent(point.id)}`, { method:'PUT', body: JSON.stringify({ name, address, note }) });
      await loadSalesAll();
      renderSalesCurrent();
      return true;
    }
  });
}

async function deleteSalesPoint(point){
  const ok = confirm(`Excluir o ponto ${point.code || ''} permanentemente?`);
  if (!ok) return;
  const re = await requestReauth({ reason: `Excluir ponto ${point.code || ''}` });
  if (!re) return;
  try {
    await api(`/api/sales/points/${encodeURIComponent(point.id)}`, { method:'DELETE', body: JSON.stringify({}) });
    await loadSalesAll();
    renderSalesCurrent();
  } catch (e) {
    console.error(e);
    alert('Falha ao excluir ponto.');
  }
}

function renderSalesOrders(){
  const tb = $("#salesOrdersTable tbody");
  if (!tb) return;

  // Toggle botão Arquivados/Voltar
  const btnArch = $("#btnSalesOrdersArchived");
  if (btnArch) {
    const on = !!state.salesOrdersShowArchived;
    btnArch.textContent = on ? "Voltar" : "Arquivados";
    btnArch.className = on ? "btn primary" : "btn secondary";
  }

  // Botões batch (arquivar / desarquivar)
  const btnArchSel = $("#btnSalesOrdersArchiveSelected");
  const btnUnarchSel = $("#btnSalesOrdersUnarchiveSelected");
  if (btnArchSel) show(btnArchSel, !state.salesOrdersShowArchived);
  if (btnUnarchSel) show(btnUnarchSel, !!state.salesOrdersShowArchived);

  const selAll = $("#salesOrdersSelectAll");
  if (selAll) {
    selAll.checked = false;
  }

  const q = String($("#salesOrderSearch")?.value || "").trim().toLowerCase();
  const rows = (state.salesOrders || [])
    .filter(o => String(o.type||"") === "POINT")
    .filter(o => {
      const isArch = !!o.archived;
      return state.salesOrdersShowArchived ? isArch : !isArch;
    })
    .filter(o => {
      if (!q) return true;
      const pName = salesPointName(o.pointId);
      const items = (o.items || []).map(it => `${it.code||""} ${it.name||""}`).join(' ');
      const hay = `${fmtSalesOrderCode(o)} ${pName} ${items} ${o.status||""}`.toLowerCase();
      return hay.includes(q);
    })
    .sort((a,b) => String(b.createdAt||"").localeCompare(String(a.createdAt||"")));

  // sane selection
  let selected = new Set(Array.isArray(state.salesOrderSelectedIds) ? state.salesOrderSelectedIds.map(x => String(x)) : []);
  // intersect selection with visible rows
  const visibleIds = new Set(rows.map(r => String(r.id)));
  selected = new Set([...selected].filter(id => visibleIds.has(id)));
  state.salesOrderSelectedIds = [...selected];

  if (!rows.length) {
    const msg = state.salesOrdersShowArchived ? "Nenhum pedido arquivado." : "Nenhum pedido criado.";
    tb.innerHTML = `<tr><td colspan="7" class="muted">${msg}</td></tr>`;
    return;
  }

  tb.innerHTML = rows.map(o => {
    const pName = escapeHtml(salesPointName(o.pointId));
    const items = (o.items || []).map(it => `${it.code||""} • ${it.qty} ${it.unit||""}`).join(' / ');
    const st = salesStatusPt(o.status);
    const isArch = !!o.archived;
    const canDispatch = !isArch && String(o.status||"").toUpperCase() !== 'DISPATCHED' && String(o.status||"").toUpperCase() !== 'CANCELLED';
    const checked = selected.has(String(o.id));

    const actions = isArch ? `
      <button class="btn secondary small" data-act="det" data-id="${escapeHtml(o.id)}">Detalhes</button>
      <button class="btn secondary small" data-act="unarch" data-id="${escapeHtml(o.id)}">Desarquivar</button>
      <button class="btn danger small" data-act="del" data-id="${escapeHtml(o.id)}">Excluir</button>
    ` : `
      <button class="btn secondary small" data-act="det" data-id="${escapeHtml(o.id)}">Detalhes</button>
      <button class="btn secondary small" data-act="op" data-id="${escapeHtml(o.id)}">Gerar OP</button>
      <button class="btn secondary small" ${canDispatch ? "" : "disabled"} data-act="dispatch" data-id="${escapeHtml(o.id)}">Despachar</button>
      <button class="btn secondary small" data-act="arch" data-id="${escapeHtml(o.id)}">Arquivar</button>
      <button class="btn danger small" data-act="del" data-id="${escapeHtml(o.id)}">Excluir</button>
    `;

    return `
      <tr>
        <td style="text-align:center"><input type="checkbox" data-so="1" data-id="${escapeHtml(o.id)}" ${checked?"checked":""} /></td>
        <td>${escapeHtml(fmtDate(o.createdAt))}</td>
        <td><b>${escapeHtml(fmtSalesOrderCode(o))}</b></td>
        <td>${pName}</td>
        <td class="muted">${escapeHtml(items || '—')}</td>
        <td><b>${escapeHtml(st)}</b></td>
        <td style="text-align:center">${actions}</td>
      </tr>
    `;
  }).join("");

  // checkbox selection
  tb.querySelectorAll('input[data-so="1"]').forEach(ch => {
    ch.addEventListener('change', () => {
      const id = String(ch.dataset.id||"");
      let set = new Set(Array.isArray(state.salesOrderSelectedIds) ? state.salesOrderSelectedIds.map(x => String(x)) : []);
      if (ch.checked) set.add(id); else set.delete(id);
      state.salesOrderSelectedIds = [...set];
      if (selAll) {
        const visible = rows.map(r => String(r.id));
        selAll.checked = visible.length > 0 && visible.every(v => set.has(v));
      }
    });
  });

  if (selAll) {
    const visible = rows.map(r => String(r.id));
    const set = new Set(Array.isArray(state.salesOrderSelectedIds) ? state.salesOrderSelectedIds.map(x => String(x)) : []);
    selAll.checked = visible.length > 0 && visible.every(v => set.has(v));
  }

  tb.querySelectorAll('button[data-act]').forEach(btn => {
    btn.addEventListener('click', async () => {
      const id = btn.dataset.id;
      const act = btn.dataset.act;
      const o = (state.salesOrders || []).find(x => String(x.id) === String(id));
      if (!o) return;
      if (act === 'det') return openSalesOrderDetails(o);
      if (act === 'op') return generateOpsFromSalesOrder(o);
      if (act === 'dispatch') return dispatchSalesOrder(o);
      if (act === 'arch') return archiveSalesOrder(o);
      if (act === 'unarch') return unarchiveSalesOrder(o);
      if (act === 'del') return deleteSalesOrder(o);
    });
  });
}

async function archiveSelectedSalesOrders(){
  const ids = Array.isArray(state.salesOrderSelectedIds) ? state.salesOrderSelectedIds.map(x => String(x)).filter(Boolean) : [];
  if (!ids.length) return alert('Selecione pelo menos 1 pedido para arquivar.');
  const ok = confirm(`Arquivar ${ids.length} pedido(s) (Pontos)?`);
  if (!ok) return;
  try {
    await api('/api/sales/orders/archive-batch', { method:'POST', body: JSON.stringify({ ids }) });
    state.salesOrderSelectedIds = [];
    await loadSalesAll();
    renderSalesCurrent();
  } catch (e) {
    console.error(e);
    const msg = e?.data?.error || e?.data?.message || '';
    alert('Falha ao arquivar pedidos.' + (msg ? ` (${msg})` : ''));
  }
}

async function unarchiveSelectedSalesOrders(){
  const ids = Array.isArray(state.salesOrderSelectedIds) ? state.salesOrderSelectedIds.map(x => String(x)).filter(Boolean) : [];
  if (!ids.length) return alert('Selecione pelo menos 1 pedido para desarquivar.');
  const ok = confirm(`Desarquivar ${ids.length} pedido(s) (Pontos)?`);
  if (!ok) return;
  try {
    await api('/api/sales/orders/unarchive-batch', { method:'POST', body: JSON.stringify({ ids }) });
    state.salesOrderSelectedIds = [];
    await loadSalesAll();
    renderSalesCurrent();
  } catch (e) {
    console.error(e);
    const msg = e?.data?.error || e?.data?.message || '';
    alert('Falha ao desarquivar pedidos.' + (msg ? ` (${msg})` : ''));
  }
}

async function unarchiveSalesOrder(order){
  try {
    await api(`/api/sales/orders/${encodeURIComponent(order.id)}/unarchive`, { method:'POST', body: JSON.stringify({}) });
    await loadSalesAll();
    renderSalesCurrent();
  } catch (e) {
    console.error(e);
    alert('Falha ao desarquivar pedido.');
  }
}


async function openSalesOrderDetails(order){
  try {
    const plan = await api(`/api/sales/orders/${encodeURIComponent(order.id)}/plan`);
    const items = Array.isArray(plan.items) ? plan.items : [];
    const linkedOps = Array.isArray(order.linkedOps) ? order.linkedOps : [];
    const body = `
      <div style="margin-bottom:8px">
        <b>${escapeHtml(fmtSalesOrderCode(order))}</b> • ${escapeHtml(salesPointName(order.pointId))}<br/>
        <span class="muted">Status: <b>${escapeHtml(salesStatusPt(order.status))}</b></span>
      </div>
      ${linkedOps.length ? `<div class="muted" style="margin-bottom:10px">OPs vinculadas: ${escapeHtml(linkedOps.map(x => pad6(x.number)).join(', '))}</div>` : ''}
      <div class="table-wrap" style="max-height:420px; overflow:auto;">
        <table class="table">
          <thead>
            <tr>
              <th style="width:70px">COD</th>
              <th>Produto Final</th>
              <th style="width:90px; text-align:right">Pedido</th>
              <th style="width:90px; text-align:right">Disp.</th>
              <th style="width:110px; text-align:right">A produzir</th>
              <th style="width:60px">UN</th>
            </tr>
          </thead>
          <tbody>
            ${items.map(it => `
              <tr>
                <td><b>${escapeHtml(it.code||"")}</b></td>
                <td>${escapeHtml(it.name||"")}</td>
                <td style="text-align:right">${escapeHtml(fmtNum(it.requested))}</td>
                <td style="text-align:right">${escapeHtml(fmtNum(it.available))}</td>
                <td style="text-align:right"><b>${escapeHtml(fmtNum(it.toProduce))}</b></td>
                <td>${escapeHtml(it.unit||"")}</td>
              </tr>
            `).join('')}
          </tbody>
        </table>
      </div>
      <div class="muted small" style="margin-top:10px">“A produzir” = max(0, Pedido − Estoque disponível). “Gerar OP” cria OPs apenas do faltante.</div>
    `;
    openModal({ title: 'Detalhes do Pedido', subtitle: 'Planejamento (estoque central de PF).', submitText: 'Fechar', bodyHtml: body, onSubmit: () => true });
  } catch (e) {
    console.error(e);
    alert('Falha ao abrir detalhes.');
  }
}

async function ensureRecipesLoaded(){
  if (Array.isArray(state.recipes) && state.recipes.length) return;
  try {
    const recRes = await api('/api/mrp/recipes');
    state.recipes = Array.isArray(recRes.recipes) ? recRes.recipes : [];
  } catch (_) {}
}

function findRecipeForPfItemId(pfItemId, pfName){
  const id = String(pfItemId);
  let r = (state.recipes || []).find(x => String(x.outputItemId||"") === id);
  if (r) return r;
  const name = String(pfName||"").trim().toLowerCase();
  if (!name) return null;
  r = (state.recipes || []).find(x => String(x.name||"").trim().toLowerCase() === name);
  return r || null;
}

async function generateOpsFromSalesOrder(order){
  try {
    await ensureRecipesLoaded();
    const plan = await api(`/api/sales/orders/${encodeURIComponent(order.id)}/plan`);
    const items = Array.isArray(plan.items) ? plan.items : [];
    const need = items.filter(it => Number(it.toProduce) > 0);
    if (!need.length) {
      alert('O estoque central de PF já cobre o pedido. Não há OP para gerar.');
      return;
    }

    const pointLabel = salesPointName(order.pointId);
    const missingRecipes = [];
    for (const it of need) {
      const r = findRecipeForPfItemId(it.itemId, it.name);
      if (!r) missingRecipes.push(`${it.code||''} - ${it.name||''}`);
    }
    if (missingRecipes.length) {
      alert('Não encontrei Receita/BOM para:\n\n' + missingRecipes.join('\n') + '\n\nCrie a BOM na aba MRP e tente novamente.');
      return;
    }

    const ok = confirm(`Gerar OPs do faltante para o ${fmtSalesOrderCode(order)} (${pointLabel})?`);
    if (!ok) return;

    const created = [];
    for (const it of need) {
      const r = findRecipeForPfItemId(it.itemId, it.name);
      const note = `Pedido de ponto • ${fmtSalesOrderCode(order)} • ${pointLabel} • ${it.code||''} - ${it.name||''} • ${fmtNum(it.toProduce)} ${it.unit||''}`;
      const res = await api('/api/mrp/production-orders', {
        method: 'POST',
        body: JSON.stringify({
          recipeId: r.id,
          qtyToProduce: Number(it.toProduce),
          note,
          allowInsufficient: true,
          createPurchaseOrder: true,
        })
      });
      if (res?.order) created.push({ id: res.order.id, number: res.order.number });
    }

    if (created.length) {
      await api(`/api/sales/orders/${encodeURIComponent(order.id)}/link-ops`, {
        method: 'POST',
        body: JSON.stringify({ ops: created })
      });
    }

    await Promise.all([loadOps(), loadPurchaseOrders()]);
    await loadSalesAll();
    renderSalesCurrent();
    alert(`OPs geradas: ${created.map(x => pad6(x.number)).join(', ')}`);
  } catch (e) {
    console.error(e);
    const msg = e?.data?.message || e?.data?.error || '';
    alert('Falha ao gerar OPs. ' + (msg ? `(${msg})` : ''));
  }
}

async function dispatchSalesOrder(order){
  const status = String(order.status||"").toUpperCase();
  if (status === 'DISPATCHED') return alert('Este pedido já foi despachado.');
  if (status === 'CANCELLED') return alert('Este pedido está cancelado.');

  const ok = confirm(`Despachar ${fmtSalesOrderCode(order)} para ${salesPointName(order.pointId)}?\n\nIsso baixa do estoque central de PF.`);
  if (!ok) return;
  try {
    await api(`/api/sales/orders/${encodeURIComponent(order.id)}/dispatch`, { method:'POST', body: JSON.stringify({}) });
    await Promise.all([loadInventory(), loadSalesAll()]);
    renderSalesCurrent();
    alert('Despacho realizado e estoque do ponto atualizado.');
  } catch (e) {
    console.error(e);
    if (e?.data?.shortages?.length) {
      const lines = e.data.shortages.map(s => `- ${s.code} ${s.name}: faltam ${fmtNum(s.shortage)} ${s.unit}`);
      alert('Não dá para despachar — falta estoque de PF:\n\n' + lines.join('\n'));
      return;
    }
    alert('Falha ao despachar.');
  }
}

async function deleteSalesOrder(order){
  const ok = confirm(`Excluir o ${fmtSalesOrderCode(order)} permanentemente?`);
  if (!ok) return;
  const re = await requestReauth({ reason: `Excluir ${fmtSalesOrderCode(order)}` });
  if (!re) return;
  try {
    await api(`/api/sales/orders/${encodeURIComponent(order.id)}`, { method:'DELETE', body: JSON.stringify({}) });
    await loadSalesAll();
    renderSalesCurrent();
  } catch (e) {
    console.error(e);
    alert('Falha ao excluir PV.');
  }
}

async function archiveSalesOrder(order){
  const ok = confirm(`Arquivar ${fmtSalesOrderCode(order)}?`);
  if (!ok) return;
  try {
    await api(`/api/sales/orders/${encodeURIComponent(order.id)}/archive`, { method: 'POST', body: JSON.stringify({}) });
    await loadSalesAll();
    renderSalesCurrent();
  } catch (e) {
    console.error(e);
    const msg = e?.data?.error || e?.data?.message || '';
    alert('Falha ao arquivar PV.' + (msg ? ` (${msg})` : ''));
  }
}

function renderQuickSalePfOptions(){
  const sel = $("#quickSalePf");
  if (!sel) return;
  const list = (state.fgItems || []).slice().sort((a,b) => compareItemCodes(a.code, b.code));
  sel.innerHTML = list.map(it => `<option value="${escapeHtml(it.id)}">${escapeHtml(it.code||"")} - ${escapeHtml(it.name||"")}</option>`).join('');
  if (!list.length) sel.innerHTML = `<option value="">(Nenhum PF cadastrado)</option>`;
  renderQuickSaleHint();
}

function renderQuickSaleHint(){
  const hint = $("#quickSaleHint");
  const sel = $("#quickSalePf");
  if (!hint || !sel) return;
  const it = (state.fgItems || []).find(x => String(x.id) === String(sel.value));
  if (!it) {
    hint.textContent = 'Cadastre Produtos Finais (PF) antes de usar venda rápida.';
    return;
  }
  hint.textContent = `Selecionado: ${it.code || ''} • ${it.name || ''} (UN: ${fmtUnit(it.unit) || '—'})`;
}



function renderQuickSalesList(){
  const tb = $("#quickSalesTable tbody");
  if (!tb) return;

  // Botões / toggle
  const btnArch = $("#btnQuickArchived");
  if (btnArch) {
    const on = !!state.quickSalesShowArchived;
    btnArch.textContent = on ? "Voltar" : "Arquivados";
    btnArch.className = on ? "btn primary" : "btn secondary";
  }
  const btnArchSel = $("#btnQuickArchiveSelected");
  if (btnArchSel) show(btnArchSel, !state.quickSalesShowArchived);

  const selAll = $("#quickSalesSelectAll");
  if (selAll) {
    selAll.disabled = !!state.quickSalesShowArchived;
    if (state.quickSalesShowArchived) selAll.checked = false;
  }

  const q = String($("#quickSalesSearch")?.value || "").trim().toLowerCase();
  const rows = (state.salesOrders || [])
    .filter(o => String(o.type||"") === "QUICK")
    .filter(o => {
      const isArch = !!o.archived;
      return state.quickSalesShowArchived ? isArch : !isArch;
    })
    .filter(o => {
      if (!q) return true;
      const items = (o.items || []).map(it => `${it.code||""} ${it.name||""}`).join(' ');
      const hay = `${fmtSalesOrderCode(o)} ${o.channel||""} ${items}`.toLowerCase();
      return hay.includes(q);
    })
    .sort((a,b) => String(b.createdAt||"").localeCompare(String(a.createdAt||"")));

  // sane selection
  let selected = new Set(Array.isArray(state.quickSelectedIds) ? state.quickSelectedIds.map(x => String(x)) : []);
  if (state.quickSalesShowArchived) selected = new Set();

  // intersect selection with visible rows
  const visibleIds = new Set(rows.map(r => String(r.id)));
  selected = new Set([...selected].filter(id => visibleIds.has(id)));
  state.quickSelectedIds = [...selected];

  if (!rows.length) {
    const msg = state.quickSalesShowArchived ? "Nenhuma venda arquivada." : "Nenhuma venda registrada ainda.";
    tb.innerHTML = `<tr><td colspan="5" class="muted">${msg}</td></tr>`;
    return;
  }

  tb.innerHTML = rows.map(o => {
    const items = (o.items || []).map(it => `${it.code||""} • ${it.qty} ${it.unit||""}`).join(' / ');
    const sumQty = (o.items || []).reduce((acc,it) => acc + (Number(it.qty)||0), 0);
    const checked = selected.has(String(o.id));
    const dis = state.quickSalesShowArchived ? "disabled" : "";
    return `
      <tr>
        <td style="text-align:center"><input type="checkbox" data-qs="1" data-id="${escapeHtml(o.id)}" ${dis} ${checked?"checked":""} /></td>
        <td>${escapeHtml(fmtDate(o.createdAt))}</td>
        <td><b>${escapeHtml(fmtSalesOrderCode(o))}</b></td>
        <td>${escapeHtml(o.channel || '—')}</td>
        <td class="muted">${escapeHtml(items || '—')}<span class="muted">${sumQty ? ` • total ${fmtNum(sumQty)}` : ''}</span></td>
      </tr>
    `;
  }).join('');

  tb.querySelectorAll('input[data-qs="1"]').forEach(ch => {
    ch.addEventListener('change', () => {
      const id = String(ch.dataset.id||"");
      let set = new Set(Array.isArray(state.quickSelectedIds) ? state.quickSelectedIds.map(x => String(x)) : []);
      if (ch.checked) set.add(id); else set.delete(id);
      state.quickSelectedIds = [...set];
      if (selAll) {
        // mark select-all if all visible checked
        const visible = rows.map(r => String(r.id));
        selAll.checked = visible.length > 0 && visible.every(v => set.has(v));
      }
    });
  });

  if (selAll) {
    const visible = rows.map(r => String(r.id));
    const set = new Set(Array.isArray(state.quickSelectedIds) ? state.quickSelectedIds.map(x => String(x)) : []);
    selAll.checked = visible.length > 0 && visible.every(v => set.has(v));
  }
}

async function archiveSelectedQuickSales(){
  const ids = Array.isArray(state.quickSelectedIds) ? state.quickSelectedIds.map(x => String(x)).filter(Boolean) : [];
  if (!ids.length) return alert('Selecione pelo menos 1 venda para arquivar.');
  const ok = confirm(`Arquivar ${ids.length} venda(s) de Venda Rápida?`);
  if (!ok) return;
  try {
    await api('/api/sales/orders/archive-batch', { method:'POST', body: JSON.stringify({ ids }) });
    state.quickSelectedIds = [];
    await loadSalesAll();
    renderSalesCurrent();
  } catch (e) {
    console.error(e);
    const msg = e?.data?.error || e?.data?.message || '';
    alert('Falha ao arquivar vendas.' + (msg ? ` (${msg})` : ''));
  }
}
async function doQuickSale(){
  const sel = $("#quickSalePf");
  const qtyEl = $("#quickSaleQty");
  const ch = $("#quickSaleChannel");
  if (!sel || !qtyEl || !ch) return;
  const itemId = String(sel.value || "");
  const qty = Number(qtyEl.value);
  if (!itemId) return alert('Selecione um PF.');
  if (!Number.isFinite(qty) || qty <= 0) return alert('Quantidade inválida.');
  const channel = String(ch.value || 'Delivery');
  try {
    await api('/api/sales/orders/quick', { method:'POST', body: JSON.stringify({ channel, items: [{ itemId, qty }] }) });
    await Promise.all([loadInventory(), loadSalesAll()]);
    renderSalesCurrent();
    alert('Venda registrada e estoque atualizado.');
  } catch (e) {
    console.error(e);
    if (e?.data?.shortages?.length) {
      const s = e.data.shortages[0];
      alert(`Sem estoque suficiente: ${s.code} ${s.name}. Disponível: ${fmtNum(s.available)} ${s.unit}.`);
      return;
    }
    alert('Falha ao registrar venda.');
  }
}

function openNewSalesOrderPoint(){
  if (!(state.salesPoints || []).length) return alert('Cadastre pelo menos 1 ponto antes de criar pedidos.');
  const points = (state.salesPoints || []).slice().sort((a,b) => String(a.code||"").localeCompare(String(b.code||""), 'pt-BR'));
  const pfList = (state.fgItems || []).slice().sort((a,b) => compareItemCodes(a.code, b.code));
  if (!pfList.length) return alert('Cadastre Produtos Finais (PF) antes de criar pedidos.');

  let lines = [];
  const renderLines = () => {
    const tb = document.querySelector('#soLinesTbody');
    if (!tb) return;
    if (!lines.length) {
      tb.innerHTML = `<tr><td colspan="5" class="muted">Nenhum item adicionado.</td></tr>`;
      return;
    }
    tb.innerHTML = lines.map((ln, idx) => {
      return `
        <tr>
          <td><b>${escapeHtml(ln.code||"")}</b></td>
          <td>${escapeHtml(ln.name||"")}</td>
          <td style="text-align:right">${escapeHtml(fmtNum(ln.qty))}</td>
          <td>${escapeHtml(ln.unit||"")}</td>
          <td style="text-align:center"><button type="button" class="btn danger small" data-rm="${idx}">Excluir</button></td>
        </tr>
      `;
    }).join('');
    tb.querySelectorAll('button[data-rm]').forEach(b => {
      b.addEventListener('click', () => {
        const idx = Number(b.dataset.rm);
        lines = lines.filter((_, i) => i !== idx);
        renderLines();
      });
    });
  };

  openModal({
    title: 'Novo Pedido (Ponto)',
    subtitle: 'Monte a lista de PF para reposição do freezer.',
    submitText: 'Criar',
    bodyHtml: `
      <div class="row wrap" style="gap:10px; align-items:flex-end;">
        <label style="flex:1; min-width:240px;">Ponto
          <select class="input" name="pointId" id="soPointId"></select>
        </label>
        <label style="width:160px;">PV (opcional)
          <input class="input" name="number" id="soPvNumber" placeholder="ex.: 1, 001, PV001" />
        </label>
      </div>
      <div class="row wrap" style="gap:10px; align-items:flex-end; margin-top:10px;">
        <label style="flex:1; min-width:240px;">Produto Final (PF)
          <select class="input" id="soPf"></select>
        </label>
        <label style="width:120px;">Qtde
          <input class="input" id="soQty" type="number" step="1" min="1" value="1" />
        </label>
        <button type="button" class="btn secondary" id="soAdd">Adicionar</button>
      </div>
      <div class="table-wrap" style="margin-top:12px; max-height:280px; overflow:auto;">
        <table class="table">
          <thead>
            <tr>
              <th style="width:70px">COD</th>
              <th>Produto</th>
              <th style="width:90px; text-align:right">Qtde</th>
              <th style="width:60px">UN</th>
              <th style="width:110px; text-align:center"></th>
            </tr>
          </thead>
          <tbody id="soLinesTbody"></tbody>
        </table>
      </div>
      <div class="muted small" style="margin-top:10px">Dica: “Gerar OP” cria OPs apenas do faltante de PF no estoque central.</div>
    `,
    onOpen: () => {
      const selPoint = document.querySelector('#soPointId');
      const selPf = document.querySelector('#soPf');
      if (selPoint) selPoint.innerHTML = points.map(p => `<option value="${escapeHtml(p.id)}">${escapeHtml(p.code||"")} - ${escapeHtml(p.name||"")}</option>`).join('');
      if (selPf) selPf.innerHTML = pfList.map(it => `<option value="${escapeHtml(it.id)}">${escapeHtml(it.code||"")} - ${escapeHtml(it.name||"")}</option>`).join('');
      const btnAdd = document.querySelector('#soAdd');
      if (btnAdd) btnAdd.addEventListener('click', () => {
        const pfId = String(selPf?.value || "");
        const it = pfList.find(x => String(x.id) === pfId);
        const qty = Number(document.querySelector('#soQty')?.value);
        if (!it) return;
        if (!Number.isFinite(qty) || qty <= 0) return alert('Quantidade inválida.');
        const unit = fmtUnit(it.unit);
        const qtz = unit === 'un' ? Math.round(qty) : qty;
        const existing = lines.find(x => String(x.itemId) === pfId);
        if (existing) existing.qty = (Number(existing.qty) || 0) + qtz;
        else lines.push({ itemId: pfId, code: it.code, name: it.name, unit: unit || it.unit, qty: qtz });
        renderLines();
      });
      renderLines();
    },
    onSubmit: async (fd) => {
      const pointId = String(fd.get('pointId')||"");
      if (!pointId) return false;
      if (!lines.length) { alert('Adicione pelo menos 1 item.'); return false; }
      const numberRaw = String(fd.get('number')||"").trim();
      let number = null;
      if (numberRaw) {
        const n = parsePvManualInput(numberRaw);
        if (!n) { alert('PV inválido. Use 1, 001, PV001 ou PV000001.'); return false; }
        number = n;
      }
      const payload = { pointId, items: lines.map(x => ({ itemId: x.itemId, qty: x.qty })) };
      if (number) payload.number = number;
      await api('/api/sales/orders/point', { method:'POST', body: JSON.stringify(payload) });
      await loadSalesAll();
      setSalesMode('orders');
      return true;
    }
  });
}

async function loadAll() {
  await Promise.all([loadInventory(), loadMRP(), loadOps(), loadPurchaseOrders(), loadUnits()]);
  // Atualiza alerta/botões de estoque mínimo na aba Estoque
  try { await refreshMinStockStatus({ force: false }); } catch (_) {}
}

// ---------- Buttons ----------
const _btnNewItem = document.querySelector("#btnNewItem");
if (_btnNewItem) _btnNewItem.addEventListener("click", openNewItem);
const _btnIn = document.querySelector("#btnIn");
if (_btnIn) _btnIn.addEventListener("click", () => openMovement("in"));
const _btnOut = document.querySelector("#btnOut");
if (_btnOut) _btnOut.addEventListener("click", () => openMovement("out"));
const _btnAdjust = document.querySelector("#btnAdjust");
if (_btnAdjust) _btnAdjust.addEventListener("click", () => openMovement("adjust"));

const _btnNewRecipe = document.querySelector("#btnNewRecipe");
if (_btnNewRecipe) _btnNewRecipe.addEventListener("click", openNewRecipe);

const _btnNewPO = document.querySelector("#btnNewPO");
if (_btnNewPO) _btnNewPO.addEventListener("click", openNewPurchaseOrder);

const _btnPOArchived = document.querySelector("#btnPOArchived");
if (_btnPOArchived) _btnPOArchived.addEventListener("click", openArchivedPurchaseOrdersModal);

const _btnOPArchived = document.querySelector("#btnOPArchived");
if (_btnOPArchived) _btnOPArchived.addEventListener("click", openArchivedProductionOrdersModal);

// ---------- Vendas (UI) ----------
const _salesModePoints = document.querySelector('#salesModePoints');
if (_salesModePoints) _salesModePoints.addEventListener('click', () => setSalesMode('points'));
const _salesModeOrders = document.querySelector('#salesModeOrders');
if (_salesModeOrders) _salesModeOrders.addEventListener('click', () => setSalesMode('orders'));
const _salesModeQuick = document.querySelector('#salesModeQuick');
if (_salesModeQuick) _salesModeQuick.addEventListener('click', () => setSalesMode('quick'));

const _btnNewSalesPoint = document.querySelector('#btnNewSalesPoint');
if (_btnNewSalesPoint) _btnNewSalesPoint.addEventListener('click', openNewSalesPoint);

const _btnPrintSalesPoint = document.querySelector('#btnPrintSalesPoint');
if (_btnPrintSalesPoint) _btnPrintSalesPoint.addEventListener('click', printSingleSelectedSalesPoint);

const _btnPrintSalesPointsSelected = document.querySelector('#btnPrintSalesPointsSelected');
if (_btnPrintSalesPointsSelected) _btnPrintSalesPointsSelected.addEventListener('click', printSelectedSalesPoints);

const _salesPointsSelectAll = document.querySelector('#salesPointsSelectAll');
if (_salesPointsSelectAll) _salesPointsSelectAll.addEventListener('change', () => {
  const q = String($("#salesPointSearch")?.value || "").trim().toLowerCase();
  const rows = (state.salesPoints || [])
    .filter(p => {
      if (!q) return true;
      const hay = `${p.code||""} ${p.name||""} ${p.address||""} ${p.note||""}`.toLowerCase();
      return hay.includes(q);
    })
    .sort((a,b) => String(a.code||"").localeCompare(String(b.code||""), 'pt-BR'));
  const visible = rows.map(r => String(r.id));
  let set = new Set(Array.isArray(state.salesPointSelectedIds) ? state.salesPointSelectedIds.map(x => String(x)) : []);
  if (_salesPointsSelectAll.checked) visible.forEach(v => set.add(v));
  else visible.forEach(v => set.delete(v));
  state.salesPointSelectedIds = [...set];
  if (state.salesMode === 'points') renderSalesPoints();
});

const _btnNewSalesOrderPoint = document.querySelector('#btnNewSalesOrderPoint');
if (_btnNewSalesOrderPoint) _btnNewSalesOrderPoint.addEventListener('click', openNewSalesOrderPoint);

const _salesPointSearch = document.querySelector('#salesPointSearch');
if (_salesPointSearch) _salesPointSearch.addEventListener('input', () => {
  if (state.salesMode === 'points') renderSalesPoints();
});
const _salesOrderSearch = document.querySelector('#salesOrderSearch');
if (_salesOrderSearch) _salesOrderSearch.addEventListener('input', () => {
  if (state.salesMode === 'orders') renderSalesOrders();
});

const _btnSalesOrdersArchived = document.querySelector('#btnSalesOrdersArchived');
if (_btnSalesOrdersArchived) _btnSalesOrdersArchived.addEventListener('click', () => {
  state.salesOrdersShowArchived = !state.salesOrdersShowArchived;
  state.salesOrderSelectedIds = [];
  if (state.salesMode === 'orders') renderSalesOrders();
});

const _btnSalesOrdersArchiveSelected = document.querySelector('#btnSalesOrdersArchiveSelected');
if (_btnSalesOrdersArchiveSelected) _btnSalesOrdersArchiveSelected.addEventListener('click', archiveSelectedSalesOrders);

const _btnSalesOrdersUnarchiveSelected = document.querySelector('#btnSalesOrdersUnarchiveSelected');
if (_btnSalesOrdersUnarchiveSelected) _btnSalesOrdersUnarchiveSelected.addEventListener('click', unarchiveSelectedSalesOrders);

const _salesOrdersSelectAll = document.querySelector('#salesOrdersSelectAll');
if (_salesOrdersSelectAll) _salesOrdersSelectAll.addEventListener('change', () => {
  const q = String(document.querySelector('#salesOrderSearch')?.value || '').trim().toLowerCase();
  const rows = (state.salesOrders || [])
    .filter(o => String(o.type||'') === 'POINT')
    .filter(o => {
      const isArch = !!o.archived;
      return state.salesOrdersShowArchived ? isArch : !isArch;
    })
    .filter(o => {
      if (!q) return true;
      const pName = salesPointName(o.pointId);
      const items = (o.items || []).map(it => `${it.code||''} ${it.name||''}`).join(' ');
      const hay = `${fmtSalesOrderCode(o)} ${pName} ${items} ${o.status||''}`.toLowerCase();
      return hay.includes(q);
    })
    .sort((a,b) => String(b.createdAt||'').localeCompare(String(a.createdAt||'')));
  const visible = rows.map(r => String(r.id));
  let set = new Set(Array.isArray(state.salesOrderSelectedIds) ? state.salesOrderSelectedIds.map(x => String(x)) : []);
  if (_salesOrdersSelectAll.checked) visible.forEach(v => set.add(v));
  else visible.forEach(v => set.delete(v));
  state.salesOrderSelectedIds = [...set];
  if (state.salesMode === 'orders') renderSalesOrders();
});

const _quickSalesSearch = document.querySelector('#quickSalesSearch');
if (_quickSalesSearch) _quickSalesSearch.addEventListener('input', () => {
  if (state.salesMode === 'quick') renderQuickSalesList();
});

const _btnQuickArchived = document.querySelector('#btnQuickArchived');
if (_btnQuickArchived) _btnQuickArchived.addEventListener('click', () => {
  state.quickSalesShowArchived = !state.quickSalesShowArchived;
  state.quickSelectedIds = [];
  if (state.salesMode === 'quick') renderQuickSalesList();
});

const _btnQuickArchiveSelected = document.querySelector('#btnQuickArchiveSelected');
if (_btnQuickArchiveSelected) _btnQuickArchiveSelected.addEventListener('click', archiveSelectedQuickSales);

const _btnQuickPrintVisible = document.querySelector('#btnQuickPrintVisible');
if (_btnQuickPrintVisible) _btnQuickPrintVisible.addEventListener('click', printQuickSalesVisible);

const _btnQuickPrintSelected = document.querySelector('#btnQuickPrintSelected');
if (_btnQuickPrintSelected) _btnQuickPrintSelected.addEventListener('click', printQuickSalesSelected);


const _quickSalesSelectAll = document.querySelector('#quickSalesSelectAll');
if (_quickSalesSelectAll) _quickSalesSelectAll.addEventListener('change', () => {
  if (state.quickSalesShowArchived) { _quickSalesSelectAll.checked = false; return; }
  const q = String(document.querySelector('#quickSalesSearch')?.value || '').trim().toLowerCase();
  const rows = (state.salesOrders || [])
    .filter(o => String(o.type||'') === 'QUICK')
    .filter(o => !o.archived)
    .filter(o => {
      if (!q) return true;
      const items = (o.items || []).map(it => `${it.code||''} ${it.name||''}`).join(' ');
      const hay = `${fmtSalesOrderCode(o)} ${o.channel||''} ${items}`.toLowerCase();
      return hay.includes(q);
    })
    .sort((a,b) => String(b.createdAt||'').localeCompare(String(a.createdAt||'')));
  const ids = rows.map(r => String(r.id));
  state.quickSelectedIds = _quickSalesSelectAll.checked ? ids : [];
  renderQuickSalesList();
});

const _quickSalePf = document.querySelector('#quickSalePf');
if (_quickSalePf) _quickSalePf.addEventListener('change', renderQuickSaleHint);

const _btnQuickSale = document.querySelector('#btnQuickSale');
if (_btnQuickSale) _btnQuickSale.addEventListener('click', doQuickSale);

const _btnDelPOSelected = document.querySelector('#btnDeletePOSelected');
if (_btnDelPOSelected) _btnDelPOSelected.addEventListener('click', async () => {
  const ids = Array.isArray(state.poSelectedIds) ? state.poSelectedIds.slice() : [];
  if (!ids.length) return alert('Selecione pelo menos 1 OC para excluir.');
  const ok = confirm(`Excluir ${ids.length} OC(s) permanentemente? Isso remove também os recebimentos/entradas de estoque vinculados.`);
  if (!ok) return;
  const re = await requestReauth({ reason: `Excluir ${ids.length} Ordem(ns) de Compra` });
  if (!re) return;
  try {
    await api('/api/mrp/purchase-orders/batch', { method: 'DELETE', body: JSON.stringify({ ids }) });
    state.poSelectedIds = [];
    await loadAll();
  } catch (e) {
    console.error(e);
    alert('Falha ao excluir OCs selecionadas.');
  }
});

const _btnNewOP = document.querySelector("#btnNewOP");
if (_btnNewOP) _btnNewOP.addEventListener("click", openNewProductionOrder);

const _btnDelOPSelected = document.querySelector('#btnDeleteOPSelected');
if (_btnDelOPSelected) _btnDelOPSelected.addEventListener('click', async () => {
  const ids = Array.isArray(state.opSelectedIds) ? state.opSelectedIds.slice() : [];
  if (!ids.length) return alert('Selecione pelo menos 1 OP para excluir.');
  const ok = confirm(`Excluir ${ids.length} OP(s) permanentemente? Isso remove também os movimentos de estoque gerados (se houver).`);
  if (!ok) return;
  const re = await requestReauth({ reason: `Excluir ${ids.length} Ordem(ns) de Produção` });
  if (!re) return;
  try {
    await api('/api/mrp/production-orders/batch', { method: 'DELETE', body: JSON.stringify({ ids }) });
    state.opSelectedIds = [];
    await loadAll();
  } catch (e) {
    console.error(e);
    alert('Falha ao excluir OPs selecionadas.');
  }
});

// Reset OP/OC — ação perigosa: apaga TODAS as OPs/OCs (inclui arquivadas) e remove movimentos de estoque gerados por elas.
async function resetOrdersFlow() {
  const ok = confirm(
    'RESET OP/OC vai APAGAR TODAS as Ordens de Produção e Ordens de Compra (inclusive Arquivadas), remover os movimentos de estoque gerados por elas e também ZERAR o contador de LOTES (etiquetas).\n\nIsso NÃO pode ser desfeito.\n\nDeseja continuar?'
  );
  if (!ok) return;

  const re = await requestReauth({ reason: 'Reset OP/OC (apagar todas OPs e OCs)' });
  if (!re) return;

  try {
    await api('/api/mrp/reset-orders', { method: 'POST', body: JSON.stringify({}) });
    state.opSelectedIds = [];
    state.poSelectedIds = [];
    await loadAll();
    alert('OPs/OCs resetadas com sucesso.');
  } catch (e) {
    console.error(e);
    alert('Falha ao resetar OP/OC.');
  }
}

const _btnResetOrdersOP = document.querySelector('#btnResetOrdersOP');
if (_btnResetOrdersOP) _btnResetOrdersOP.addEventListener('click', resetOrdersFlow);

const _btnResetOrdersOC = document.querySelector('#btnResetOrdersOC');
if (_btnResetOrdersOC) _btnResetOrdersOC.addEventListener('click', resetOrdersFlow);


// Reset TOTAL — ação extremamente perigosa: apaga OP/OC + LOTES + PV/PVR + PONTOS + movimentos relacionados.
// NÃO mexe em cadastros MP/PF/BOM.
async function resetTotalFlow() {
  const ok = confirm(`RESET TOTAL vai APAGAR:
• OP/OC (incluindo Arquivadas)
• LOTES (etiquetas)
• Pedidos para Ponto (PV)
• Venda Rápida (PVR)
• Pontos (Freezers)
• Movimentos gerados por OP/OC/PV/PVR (estoque central e pontos)

Isso NÃO pode ser desfeito.

NÃO mexe em MP/PF/BOM.

Deseja continuar?`);
  if (!ok) return;

  const re = await requestReauth({ reason: 'Reset TOTAL (apagar OP/OC/PV/PVR/Pontos e movimentos)' });
  if (!re) return;

  try {
    await api('/api/mrp/reset-orders', { method: 'POST', body: JSON.stringify({ scope: 'all' }) });

    // limpa seleções locais
    state.opSelectedIds = [];
    state.poSelectedIds = [];
    state.salesOrderSelectedIds = [];
    state.quickSelectedIds = [];
    state.salesPointSelectedIds = [];

    await Promise.all([loadAll(), loadSalesAll()]);

    // re-render do que estiver aberto
    const tabVendas = document.querySelector('#tab-vendas');
    const isVisible = (el) => el && !el.classList.contains('hidden');
    if (isVisible(tabVendas)) {
      renderSalesCurrent();
    } else if (typeof renderCurrentView === 'function') {
      renderCurrentView();
    }
    alert('Reset TOTAL concluído com sucesso.');
  } catch (e) {
    console.error(e);
    alert('Falha ao executar Reset TOTAL.');
  }
}

const _btnResetTotal = document.querySelector('#btnResetTotal');
if (_btnResetTotal) _btnResetTotal.addEventListener('click', resetTotalFlow);


const _btnNewPFBom = document.querySelector("#btnNewPFBom");
if (_btnNewPFBom) _btnNewPFBom.addEventListener("click", async () => {
  // ensure MRP data is loaded
  await loadMRP();
  openPFBomPicker();
});

// MRP: botão "Simular produção" abre uma tela separada (calcular necessidades + OP/OC)
const _btnSimulateFocus = document.querySelector("#btnSimulateFocus");
if (_btnSimulateFocus) _btnSimulateFocus.addEventListener("click", async () => {
  try { await loadMRP(); } catch (e) {}
  // garante que tenha um PF selecionado
  if (!state.selectedPFId) {
    const first = (state.fgItems || [])[0];
    if (first) state.selectedPFId = first.id;
  }
  openMrpSimView();
});

const _btnMrpBackToListTop = document.querySelector('#btnMrpBackToListTop');
if (_btnMrpBackToListTop) _btnMrpBackToListTop.addEventListener('click', () => openMrpListView());

const _pfBomSearch = document.querySelector("#pfBomSearch");
if (_pfBomSearch) _pfBomSearch.addEventListener("input", () => renderPFBomList());

const _btnPFBomExport = document.querySelector("#btnPFBomExport");
if (_btnPFBomExport) _btnPFBomExport.addEventListener("click", () => {
  const out = (state.fgItems || []).slice().sort((a,b)=>String(a.code||"").localeCompare(String(b.code||""),"pt-BR")).map(pf => {
    const r = recipeForPF(pf.id);
    return {
      pf: { id: pf.id, code: pf.code, name: pf.name, unit: pf.unit },
      bom: r?.bom || [],
      recipeId: r?.id || null,
      notes: r?.notes || ""
    };
  });
  downloadText("bom_produto_final.json", JSON.stringify({ exportedAt: new Date().toISOString(), items: out }, null, 2));
});



// ---------- Estoque: tipo (Matéria-prima vs Produto final) ----------
const stockModeRawBtn = document.querySelector("#stockModeRaw");
const stockModeFgBtn = document.querySelector("#stockModeFg");
const stockModeAllBtn = document.querySelector("#stockModeAll");
const stockAdjustRawBtn = document.querySelector("#stockAdjustRaw");
const stockAdjustFgBtn = document.querySelector("#stockAdjustFg");
const stockInvHistoryBtn = document.querySelector("#stockInvHistory");


function openCadastroItensModal(){
  const typeLabel = state.stockMode === "fg" ? "Produto Final" : "Matéria-prima & Insumos";
  const dbLabel = state.stockMode === "fg" ? "bd/estoque_pf.json" : "bd/estoque_mp.json";
  const nextCode = computeNextCode(state.items.filter(i => (i.type||'raw')===state.stockMode), state.stockMode);

  let selectedId = null;
  let q = "";

  const sortItems = (arr) => [...arr].sort((a,b) => {
    const ac = compareItemCodes(a.code || "", b.code || "");
    if (ac !== 0) return ac;
    return String(a.name || "").localeCompare(String(b.name || ""), 'pt-BR');
  });

  const filterItems = () => {
    const qq = q.trim().toLowerCase();
    if (!qq) return sortItems(state.items);
    return sortItems(state.items).filter((it) => {
      const hay = `${it.code||""} ${it.name||""} ${it.sku||""} ${it.unit||""}`.toLowerCase();
      return hay.includes(qq);
    });
  };

  const body = `
    <div class="cadastro-sub muted">
      Cadastro de itens — você pode reutilizar códigos excluídos (ex.: MP002). Sugestão automática: <b>${nextCode}</b>
    </div>

    <div class="cadastro-toolbar">
      <div class="cadastro-search">
        <input id="cadSearch" class="input" placeholder="Pesquisar..." />
      </div>
      <div class="cadastro-actions">
        <button id="cadImport" type="button" class="btn secondary small">Importar</button>
        <button id="cadExport" type="button" class="btn secondary small">Exportar</button>
        <span class="cadastro-sep"></span>
        <button id="cadNew" type="button" class="btn primary small">+ Novo</button>
        <button id="cadEdit" type="button" class="btn secondary small" disabled>Editar</button>
        <button id="cadDup" type="button" class="btn secondary small" disabled>Duplicar</button>
        <button id="cadDel" type="button" class="btn danger small" disabled>Excluir</button>
      </div>
      <input id="cadFile" type="file" accept=".xlsx,application/vnd.openxmlformats-officedocument.spreadsheetml.sheet" style="display:none" />
    </div>

    <div class="table-wrap cadastro-tablewrap">
      <table class="table cadastro-table">
        <thead class="cad-head">
          <tr>
            ${state.stockMode === "fg" ? `
              <th style="width:90px">CÓDIGO</th>
              <th>DESCRIÇÃO</th>
              <th class="cad-num" style="width:80px">UN</th>
              <th class="cad-num col-sale">V.VENDA (R$)</th>
              <th class="cad-num" style="width:160px">ESTOQUE MÍNIMO</th>
            ` : `
              <th style="width:90px">CÓDIGO</th>
              <th>DESCRIÇÃO</th>
              <th class="cad-num" style="width:80px">UN</th>
              <th class="cad-num" style="width:120px">CUSTO (R$)</th>
              <th class="cad-num" style="width:90px">% PERDA</th>
              <th class="cad-num" style="width:160px">ESTOQUE MÍNIMO</th>
            `}
          </tr>
        </thead>
        <tbody id="cadTbody"></tbody>
      </table>
    </div>

    <div class="muted small" style="margin-top:10px">
      Dica: clique em uma linha para selecionar. Editar/Duplicar/Excluir usam o item selecionado.
    </div>
  `;

  openModal({
    title: `Cadastro • ${typeLabel}`,
    subtitle: "",
    submitText: "Fechar",
    bodyHtml: body,
    onSubmit: async () => true,
  });

  const elSearch = document.querySelector("#cadSearch");
  const elTbody = document.querySelector("#cadTbody");
  const elEdit = document.querySelector("#cadEdit");
  const elDup = document.querySelector("#cadDup");
  const elDel = document.querySelector("#cadDel");
  const elNew = document.querySelector("#cadNew");
  const elImport = document.querySelector("#cadImport");
  const elExport = document.querySelector("#cadExport");
  const elFile = document.querySelector("#cadFile");

  const setButtons = () => {
    const ok = !!selectedId;
    if (elEdit) elEdit.disabled = !ok;
    if (elDup) elDup.disabled = !ok;
    if (elDel) elDel.disabled = !ok;
  };

  const rowHtml = (it) => {
    const isFg = state.stockMode === "fg";
    const cols = isFg ? `
      <td><b>${escapeHtml(normalizeItemCode(it.code||"", state.stockMode) || it.code || "")}</b></td>
      <td><b>${escapeHtml(it.name)}</b></td>
      <td class="cad-num">${escapeHtml(it.unit)}</td>
      <td class="cad-num col-sale">${fmtMoney(it.salePrice || 0)}</td>
      <td class="cad-num">${fmt(it.minStock || 0)} ${escapeHtml(it.unit)}</td>
    ` : `
      <td><b>${escapeHtml(normalizeItemCode(it.code||"", state.stockMode) || it.code || "")}</b></td>
      <td><b>${escapeHtml(it.name)}</b></td>
      <td class="cad-num">${escapeHtml(it.unit)}</td>
      <td class="cad-num">${fmtMoney(it.cost || 0)}</td>
      <td class="cad-num">${fmt(it.lossPercent || 0)}%</td>
      <td class="cad-num">${fmt(it.minStock || 0)} ${escapeHtml(it.unit)}</td>
    `;
    return `
      <tr class="${it.id === selectedId ? "row-selected" : ""}" data-id="${it.id}">
        ${cols}
      </tr>
    `;
  };

  const renderRows = () => {
    if (!elTbody) return;
    const items = filterItems();
    elTbody.innerHTML = items.length ? items.map(rowHtml).join("") : `<tr><td colspan="${state.stockMode==='fg'?6:6}" class="muted">Nenhum item cadastrado.</td></tr>`;
    elTbody.querySelectorAll("tr[data-id]").forEach((tr) => {
      tr.addEventListener("click", () => {
        selectedId = tr.getAttribute("data-id");
        renderRows();
        setButtons();
      });
    });
  };

  if (elSearch) elSearch.addEventListener("input", () => {
    q = elSearch.value || "";
    renderRows();
  });

  if (elNew) elNew.addEventListener("click", () => {
    modal.close();
    openNewItem(() => openCadastroItensModal());
  });

  if (elEdit) elEdit.addEventListener("click", () => {
    const it = state.items.find((x) => x.id === selectedId);
    if (!it) return;
    modal.close();
    openEditItem(it, () => openCadastroItensModal());
  });

  if (elDup) elDup.addEventListener("click", async () => {
    const it = state.items.find((x) => x.id === selectedId);
    if (!it) return;
    const isFg = state.stockMode === "fg";
    const payload = isFg
      ? { name: `${it.name} (cópia)`, unit: it.unit, minStock: it.minStock || 0, salePrice: it.salePrice || 0 }
      : { name: `${it.name} (cópia)`, unit: it.unit, minStock: it.minStock || 0, cost: it.cost || 0, lossPercent: it.lossPercent || 0 };
    await api(`/api/inventory/items?type=${state.stockMode}`, { method: "POST", body: JSON.stringify(payload) });
    await loadInventory();
    selectedId = null;
    setButtons();
    renderRows();
  });

  if (elDel) elDel.addEventListener("click", async () => {
    const it = state.items.find((x) => x.id === selectedId);
    if (!it) return;
    const ok0 = confirm(`Excluir o item "${it.name}"?`);
    if (!ok0) return;
    try{
      await api(`/api/inventory/items/${it.id}?type=${state.stockMode}`, { method: "DELETE" });
    } catch (err){
      if (err?.data?.error === "has_links"){
        const links = Array.isArray(err.data.links) ? err.data.links : [];
        const parts = [];
        if (err.data.hasMovements) parts.push("• Movimentações de estoque");
        for (const l of links.slice(0, 15)) parts.push(`• ${l.label}`);
        const msg = [
          `O item "${it.name}" possui vínculos e/ou movimentações.`,
          parts.length ? ("\n\nOnde está amarrado:\n" + parts.join("\n")) : "",
          "\n\nDeseja EXCLUIR mesmo assim? (isso removerá o item e limpará referências/movimentações dele)"
        ].join("");
        const ok = confirm(msg);
        if (!ok) return;
        await api(`/api/inventory/items/${it.id}?type=${state.stockMode}&force=1`, { method: "DELETE" });
      } else if (err?.data?.error === "has_movements"){
        const ok = confirm(`O item "${it.name}" possui movimentações. Excluir mesmo assim?`);
        if (!ok) return;
        await api(`/api/inventory/items/${it.id}?type=${state.stockMode}&force=1`, { method: "DELETE" });
      } else {
        throw err;
      }
    }
    await loadInventory();
    selectedId = null;
    setButtons();
    renderRows();
  });

  if (elExport) elExport.addEventListener("click", async () => {
    await downloadAfterReauth(`/api/inventory/export.xlsx?type=${state.stockMode === "fg" ? "fg" : "raw"}`, "Exportar cadastro (XLSX)");
  });

  if (elImport) elImport.addEventListener("click", () => {
    // Importação com reauth + arquivo no mesmo modal (evita bloqueio do file picker após await)
    const typeLabel = (state.stockMode === "fg") ? "Produto Final" : "Matéria-prima & Insumos";
    openImportXlsxModal({
      title: "Importar Cadastro (.xlsx)",
      reason: `Importar cadastro — ${typeLabel} (XLSX)`,
      uploadUrl: `/api/inventory/import.xlsx?type=${state.stockMode === "fg" ? "fg" : "raw"}`,
      onAfter: async () => {
        await loadInventory();
        selectedId = null;
        setButtons();
        renderRows();
      },
    });
  });

  setButtons();
  renderRows();
}


async function setStockMode(mode){
  state.stockMode = (mode === "fg") ? "fg" : "raw";
  if (stockModeRawBtn && stockModeFgBtn){
    stockModeRawBtn.classList.toggle("active", state.stockMode === "raw");
    stockModeFgBtn.classList.toggle("active", state.stockMode === "fg");
  }
  if (stockModeAllBtn) stockModeAllBtn.classList.remove("active");
  if (stockInvHistoryBtn) stockInvHistoryBtn.classList.remove("active");
  // volta para a área padrão (cadastro MP/PF)
  if (generalBox) generalBox.classList.add("hidden");
  if (invHistBox) invHistBox.classList.add("hidden");
  if (cadastroBox) cadastroBox.classList.remove("hidden");
  state.selectedItemId = null;
  await loadInventory();
  try { renderCadastro(); } catch(e) {}
}


if (stockModeRawBtn) stockModeRawBtn.addEventListener("click", async () => { await setStockMode("raw"); });
if (stockModeFgBtn) stockModeFgBtn.addEventListener("click", async () => { await setStockMode("fg"); });

// Cadastro Geral (MP + PF) abre na mesma área (sem modal)
if (stockModeAllBtn) stockModeAllBtn.addEventListener("click", async () => { await openGeneralCadastroInline(); });

if (stockInvHistoryBtn) stockInvHistoryBtn.addEventListener("click", async () => { await openInventoryHistoryInline(); });

if (stockAdjustRawBtn) stockAdjustRawBtn.addEventListener("click", async () => { await setStockMode("raw"); openAdjustPicker(); });
if (stockAdjustFgBtn) stockAdjustFgBtn.addEventListener("click", async () => { await setStockMode("fg"); openAdjustPicker(); });



// ---------- Cadastro Geral (MP + PF) ----------
const generalState = { items: [], rawItems: [], fgItems: [], filter: "", sortKey: "code", sortDir: "asc", selectedId: null };

const generalBox = document.querySelector("#estoqueCadastroGeneral");
const invHistBox = document.querySelector("#estoqueInventoryHistory");
const generalSearchInline = document.querySelector("#generalSearchInline");
const generalTbodyInline = document.querySelector("#generalTbodyInline");
const generalTableInline = document.querySelector("#generalTableInline");
const btnGeneralImport = document.querySelector("#btnGeneralImport");
const btnGeneralExport = document.querySelector("#btnGeneralExport");
const btnGeneralEdit = document.querySelector("#btnGeneralEdit");
const btnGeneralInv = document.querySelector("#btnGeneralInv");
const btnGeneralRecipe = document.querySelector("#btnGeneralRecipe");

// ---------- Histórico de Inventário (ajustes) ----------
const invHistState = { rows: [], filter: "", sortKey: "at", sortDir: "desc" };
const invHistSearch = document.querySelector("#invHistSearch");
const invHistTbody = document.querySelector("#invHistTbody");

if (invHistTbody){
  invHistTbody.addEventListener("click", (e) => {
    const btn = e.target.closest(".obs-btn");
    if (!btn) return;
    e.preventDefault();
    e.stopPropagation();
    const idx = Number(btn.getAttribute("data-idx"));
    if (!Number.isFinite(idx)) return;
    openInvHistObs(idx);
  });
}

const invHistTable = document.querySelector("#invHistTable");
const btnInvHistExport = document.querySelector("#btnInvHistExport");
const btnInvHistClear = document.querySelector("#btnInvHistClear");

function normalizeHistRow(r){
  const st = r.stockType === "fg" ? "PF" : "MP";
  const byName = String(r.by?.name || r.byName || "").trim() || String(r.by?.id || "").trim() || "";
  return {
    id: String(r.id || ""),
    at: String(r.at || ""),
    stockType: st,
    code: String(r.code || ""),
    name: String(r.name || ""),
    unit: String(r.unit || ""),
    beforeQty: Number(r.beforeQty ?? 0),
    afterQty: Number(r.afterQty ?? 0),
    delta: Number(r.delta ?? (Number(r.afterQty ?? 0) - Number(r.beforeQty ?? 0))),
    byName,
    reason: String(r.reason || ""),
  };
}

function sortInvHistRows(rows){
  const key = invHistState.sortKey;
  const dir = invHistState.sortDir;
  const cmp = (a, b) => {
    const va = a[key];
    const vb = b[key];
    if (key === "at") return String(va).localeCompare(String(vb));
    if (typeof va === "number" && typeof vb === "number") return va - vb;
    return String(va ?? "").localeCompare(String(vb ?? ""), "pt-BR", { numeric: true, sensitivity: "base" });
  };
  const out = rows.slice().sort(cmp);
  return dir === "desc" ? out.reverse() : out;
}

function updateInvHistSortIndicators(){
  const keys = ["at","stockType","code","name","unit","beforeQty","afterQty","delta","byName","reason"];
  keys.forEach((k) => {
    const el = document.querySelector("#hSort_" + k);
    if (!el) return;
    if (invHistState.sortKey !== k) { el.textContent = ""; return; }
    el.textContent = invHistState.sortDir === "asc" ? "▲" : "▼";
  });
}

function renderInvHistTable(){
  if (!invHistTbody) return;
  const q = String(invHistState.filter || "").trim().toLowerCase();
  let rows = invHistState.rows.slice();
  if (q){
    rows = rows.filter((r) => (
      r.code.toLowerCase().includes(q) ||
      r.name.toLowerCase().includes(q) ||
      r.unit.toLowerCase().includes(q) ||
      r.stockType.toLowerCase().includes(q) ||
      r.byName.toLowerCase().includes(q) ||
      r.reason.toLowerCase().includes(q)
    ));
  }
  rows = sortInvHistRows(rows);

  if (rows.length === 0){
    invHistTbody.innerHTML = `<tr><td colspan="10" class="muted">Nenhum registro encontrado.</td></tr>`;
    updateInvHistSortIndicators();
    return;
  }

  invHistState.viewRows = rows;

  invHistTbody.innerHTML = rows.map((r, i) => {
    return `
      <tr>
        <td>${escapeHtml(fmtDate(r.at))}</td>
        <td><b>${escapeHtml(r.stockType)}</b></td>
        <td><b>${escapeHtml(r.code)}</b></td>
        <td><span class="g-desc" title="${escapeHtml(r.name)}">${escapeHtml(r.name)}</span></td>
        <td>${escapeHtml(r.unit)}</td>
        <td>${escapeHtml(fmt(r.beforeQty))}</td>
        <td>${escapeHtml(fmt(r.afterQty))}</td>
        <td>${escapeHtml(fmt(r.delta))}</td>
        <td>${escapeHtml(r.byName)}</td>
        <td>
          <div class="obs-cell">
            <button type="button" class="btn micro obs-btn" data-idx="${i}" title="Ver observação">Obs</button>
          </div>
        </td>
      </tr>
    `;
  }).join("");

  updateInvHistSortIndicators();
}


function openInvHistObs(idx){
  const rows = invHistState.viewRows || [];
  const r = rows[idx];
  if (!r) return;
  const reason = (r.reason || "").trim();
  openModal({
    title: "Observação",
    subtitle: `${r.code} • ${r.name} • ${fmtDate(r.at)} • ${r.byName}`,
    submitText: "Fechar",
    bodyHtml: `<div class="obs-full">${escapeHtml(reason || "(sem observação)")}</div>`
  });
}

function normalizeGeneralItem(it){
  return {
    id: it.id,
    code: String(it.code || ""),
    name: String(it.name || ""),
    unit: String(it.unit || ""),
    minStock: Number(it.minStock || 0),
    cost: Number(it.cost || 0),
    salePrice: Number(it.salePrice || 0),
    lossPercent: Number(it.lossPercent || 0),
    cookFactor: Number(it.cookFactor || 0),
    currentStock: Number(it.currentStock || 0),
    type: it.type === "fg" ? "fg" : "raw",
  };
}

function sortGeneralItems(items){
  const key = generalState.sortKey;
  const dir = generalState.sortDir;
  const cmp = (a, b) => {
    const va = a[key];
    const vb = b[key];
    if (typeof va === "number" && typeof vb === "number") return va - vb;
    return String(va ?? "").localeCompare(String(vb ?? ""), "pt-BR", { numeric: true, sensitivity: "base" });
  };
  const out = items.slice().sort(cmp);
  return dir === "desc" ? out.reverse() : out;
}

function formatCellMoney(n){ return (Number(n) ? fmtMoney(n) : ""); }
function formatCellNum(n){ return (Number(n) || Number(n) === 0) ? fmtNum(n) : ""; }

function closeAllRowMenus(exceptId=""){
  document.querySelectorAll(".row-menu.open").forEach((m) => {
    if (exceptId && m.getAttribute("data-menu") === exceptId) return;
    m.classList.remove("open");
  });
}

function updateGeneralSortIndicators(){
  const keys = ["code","name","unit","cost","salePrice","lossPercent","cookFactor","minStock","currentStock"];
  keys.forEach((k) => {
    const el = document.querySelector("#gSort_" + k);
    if (!el) return;
    if (generalState.sortKey !== k) { el.textContent = ""; return; }
    el.textContent = generalState.sortDir === "asc" ? "▲" : "▼";
  });
}

function updateGeneralActionButtons(){
  const id = generalState.selectedId;
  const has = !!id && generalState.items.some((x) => x.id === id);
  if (btnGeneralEdit) btnGeneralEdit.disabled = !has;
  if (btnGeneralInv) btnGeneralInv.disabled = !has;
  if (btnGeneralRecipe) btnGeneralRecipe.disabled = !has;
}

function renderGeneralTableInline(){
  if (!generalTbodyInline) return;

  const q = (generalState.filter || "").trim().toLowerCase();
  let rows = generalState.items.slice();
  if (q){
    rows = rows.filter((it) => (
      it.code.toLowerCase().includes(q) ||
      it.name.toLowerCase().includes(q) ||
      it.unit.toLowerCase().includes(q)
    ));
  }
  rows = sortGeneralItems(rows);

  if (rows.length === 0){
    generalTbodyInline.innerHTML = `<tr><td colspan="9" class="muted">Nenhum item encontrado.</td></tr>`;
    return;
  }

  generalTbodyInline.innerHTML = rows.map((it) => {
    const selected = (generalState.selectedId === it.id) ? "selected" : "";
    // destaque de estoque mínimo
    let minCls = "";
    const minS = Number(it.minStock);
    const curS = Number(it.currentStock);
    if (Number.isFinite(minS) && minS > 0 && Number.isFinite(curS)) {
      if (curS < minS - 1e-9) minCls = "min-below";
      else if (curS <= minS + 1e-9) minCls = "min-at";
    }
    const cls = [selected, minCls].filter(Boolean).join(" ");
    return `
      <tr class="${cls}" data-id="${escapeHtml(it.id)}">
        <td><b>${escapeHtml(it.code)}</b></td>
        <td><span class="g-desc" title="${escapeHtml(it.name)}">${escapeHtml(it.name)}</span></td>
        <td>${escapeHtml(it.unit)}</td>
        <td>${escapeHtml(formatCellMoney(it.cost))}</td>
        <td>${escapeHtml(formatCellMoney(it.salePrice))}</td>
        <td>${escapeHtml(formatCellNum(it.lossPercent))}</td>
        <td>${escapeHtml(formatCellNum(it.cookFactor))}</td>
        <td>${escapeHtml(formatCellNum(it.minStock))}</td>
        <td>${escapeHtml(formatCellNum(it.currentStock))}</td>
      </tr>
    `;
  }).join("");

  // row select
  generalTbodyInline.querySelectorAll("tr[data-id]").forEach((tr) => {
    tr.addEventListener("click", (e) => {
      generalState.selectedId = tr.getAttribute("data-id");
      renderGeneralTableInline();
    });
  });
  updateGeneralActionButtons();
  updateGeneralSortIndicators();
}


async function handleGeneralAction(action, item){
  // IMPORTANTE:
  // As ações do Cadastro Geral já tinham regras próprias. O comportamento correto é:
  // - NÃO navegar para o cadastro MP/PF (manter o usuário no Cadastro Geral)
  // - Abrir os modais/fluxos existentes e, ao fechar, recarregar o Cadastro Geral

  const returnToGeneralAfterClose = async () => {
    try {
      // preserva seleção atual (se existir)
      const keepId = generalState.selectedId;
      await openGeneralCadastroInline();
      if (keepId){
        generalState.selectedId = keepId;
        renderGeneralTableInline();
      }
    } catch(_) {}
  };

  const bindReturnOnce = () => {
    try {
      const onClose = async () => {
        try { modal.removeEventListener("close", onClose); } catch(_) {}
        await returnToGeneralAfterClose();
      };
      modal.addEventListener("close", onClose);
    } catch(_) {}
  };

  // Carrega o cadastro do tipo correto SEM mexer na UI (não alterna abas/boxes).
  const loadTypeSilent = async (type) => {
    state.stockMode = (type === "fg") ? "fg" : "raw";
    await loadInventory();
  };

  if (action === "edit"){
    await loadTypeSilent(item.type);
    const it = state.items.find((x) => x.id === item.id);
    if (!it){ alert("Item não encontrado no cadastro."); return; }
    bindReturnOnce();
    openEditItem(it);
    return;
  }

  if (action === "inv"){
    await loadTypeSilent(item.type);
    bindReturnOnce();
    openMovement("adjust", { preselectItemId: item.id, lockItem: true });
    return;
  }

  if (action === "recipe"){
    // garante dados do MRP carregados
    await loadMRP();
    if (item.type === "fg"){
      bindReturnOnce();
      await openRecipeForPF(item.id);
    } else {
      const pfId = await pickPFForMP();
      if (!pfId) return;
      bindReturnOnce();
      await openRecipeForPF(pfId, { addRawItemId: item.id });
    }
  }
}

async function pickPFForMP(){
  const pfItems = (state.fgItems || []).slice().sort((a,b) => String(a.code||"").localeCompare(String(b.code||""), "pt-BR", { numeric:true }));
  if (pfItems.length === 0){ alert("Cadastre pelo menos 1 Produto Final (PF) antes."); return null; }

  return await new Promise((resolve) => {
    let selected = pfItems[0]?.id || "";
    let q = "";

    const renderList = () => {
      const box = document.querySelector("#pfPickList");
      const hint = document.querySelector("#pfPickHint");
      if (!box) return;
      const filtered = q
        ? pfItems.filter((it) => (String(it.code||"")+" "+String(it.name||"")).toLowerCase().includes(q))
        : pfItems;
      if (hint) hint.textContent = filtered.length ? `${filtered.length} encontrado(s)` : "Nenhum resultado";
      box.innerHTML = filtered.map((it) => {
        const active = it.id === selected ? "active" : "";
        return `<div class="list-item ${active}" data-pf="${escapeHtml(it.id)}" style="cursor:pointer">
          <div class="pf-pick-title">${escapeHtml(normalizeItemCode(it.code||"", state.stockMode) || it.code || "")} • ${escapeHtml(it.name || "")}</div>
          <div class="muted small">UN: ${escapeHtml(it.unit || "")}</div>
        </div>`;
      }).join("");
      box.querySelectorAll("[data-pf]").forEach((div) => {
        div.addEventListener("click", () => {
          selected = div.getAttribute("data-pf");
          renderList();
        });
      });
    };

    const onClose = () => {
      modal.removeEventListener("close", onClose);
      resolve(null);
    };
    modal.addEventListener("close", onClose);

    openModal({
      title: "Receita • escolher Produto Final",
      subtitle: "Você clicou “Receita” em um item MP. Selecione o PF onde ele será incluído.",
      submitText: "Selecionar",
      cardClass: "wide",
      bodyHtml: `
        <input id="pfPickSearch" class="input search" placeholder="Pesquisar PF (código / nome)..." />
        <div class="muted small" id="pfPickHint" style="margin-top:6px"></div>
        <div id="pfPickList" class="list pf-pick-list" style="margin-top:10px"></div>
      `,
      onOpen: () => {
        const s = document.querySelector("#pfPickSearch");
        if (s){
          s.addEventListener("input", () => {
            q = String(s.value||"").trim().toLowerCase();
            renderList();
          });
        }
        renderList();
      },
      onSubmit: () => {
        modal.removeEventListener("close", onClose);
        resolve(selected || null);
        return true;
      },
    });
  });
}

async function openRecipeForPF(pfId, opts = {}){
  await loadMRP();
  const pfItem = (state.fgItems || []).find((x) => x.id === pfId);
  if (!pfItem){ alert("Produto Final não encontrado."); return; }

  const existing = (state.recipes || []).find((r) => r.productId === pfId);
  if (existing){
    openRecipeEditor({ mode: "edit", recipe: existing, productId: pfId, prefillBomItemId: opts.addRawItemId });
  } else {
    openRecipeEditor({
      mode: "new",
      recipe: { name: pfItem.name, yieldQty: 1, yieldUnit: pfItem.unit, notes: "", bom: [] },
      productId: pfId,
      prefillBomItemId: opts.addRawItemId,
    });
  }
}

let generalInlineBound = false;

// Recarrega os dados do Cadastro Geral (MP + PF).
// Por padrão, preserva filtro/ordem/seleção atuais para evitar "pulos" na UI.
async function refreshGeneralCadastroInline({ preserveUI = true } = {}){
  const prev = {
    filter: generalState.filter,
    sortKey: generalState.sortKey,
    sortDir: generalState.sortDir,
    selectedId: generalState.selectedId,
    searchValue: generalSearchInline ? String(generalSearchInline.value || "") : "",
  };

  // carrega itens MP e PF (sempre atual, pois OC/OP podem alterar estoque fora da aba Estoque)
  const [rawRes, fgRes] = await Promise.all([
    api("/api/inventory/items?type=raw"),
    api("/api/inventory/items?type=fg"),
  ]);
  generalState.rawItems = (rawRes.items || []).map((x) => ({ ...x, type: "raw" }));
  generalState.fgItems = (fgRes.items || []).map((x) => ({ ...x, type: "fg" }));
  // Atualiza alerta/botões de estoque mínimo com dados mais atuais
  try { setMinStockStatus(rawRes.items || [], fgRes.items || []); } catch (_) {}
  const merged = [...generalState.rawItems, ...generalState.fgItems].map(normalizeGeneralItem);
  generalState.items = merged;

  if (preserveUI){
    generalState.filter = prev.filter;
    generalState.sortKey = prev.sortKey;
    generalState.sortDir = prev.sortDir;
    // mantém seleção se ainda existir
    generalState.selectedId = merged.some((x) => x.id === prev.selectedId) ? prev.selectedId : null;
    if (generalSearchInline) generalSearchInline.value = prev.searchValue;
  } else {
    generalState.filter = "";
    generalState.sortKey = "code";
    generalState.sortDir = "asc";
    generalState.selectedId = null;
    if (generalSearchInline) generalSearchInline.value = "";
  }

  renderGeneralTableInline();
}

async function openGeneralCadastroInline(){
  if (stockModeAllBtn) stockModeAllBtn.classList.add("active");
  if (stockInvHistoryBtn) stockInvHistoryBtn.classList.remove("active");
  if (stockModeRawBtn) stockModeRawBtn.classList.remove("active");
  if (stockModeFgBtn) stockModeFgBtn.classList.remove("active");
  if (cadastroBox) cadastroBox.classList.add("hidden");
  if (invHistBox) invHistBox.classList.add("hidden");
  if (generalBox) generalBox.classList.remove("hidden");

  // (re)carrega os itens e reseta UI (abertura do modo)
  await refreshGeneralCadastroInline({ preserveUI: false });

  // bind listeners once
  if (!generalInlineBound){
    generalInlineBound = true;

    if (generalSearchInline){
      generalSearchInline.addEventListener("input", () => {
        generalState.filter = String(generalSearchInline.value || "");
        renderGeneralTableInline();
      });
    }

    if (generalTableInline){
      generalTableInline.querySelectorAll("th.sortable").forEach((th) => {
        th.addEventListener("click", () => {
          const key = th.getAttribute("data-key");
          if (!key) return;
          if (generalState.sortKey === key) {
            generalState.sortDir = generalState.sortDir === "asc" ? "desc" : "asc";
          } else {
            generalState.sortKey = key;
            generalState.sortDir = "asc";
          }
          renderGeneralTableInline();
        });
      });
    }


    if (btnGeneralImport){
      btnGeneralImport.addEventListener("click", async () => {
        openImportXlsxModal({
          title: "Importar Cadastro Geral (XLSX)",
          reason: "Importar Cadastro Geral (MP + PF)",
          uploadUrl: "/api/inventory/import-all.xlsx",
          onAfter: async () => { await loadInventory(); await openGeneralCadastroInline(); },
        });
      });
    }

    const runGeneralTopAction = async (action) => {
      const id = generalState.selectedId;
      if (!id){ alert("Selecione um item na tabela."); return; }
      const item = generalState.items.find((x) => x.id === id);
      if (!item){ alert("Item não encontrado."); return; }
      await handleGeneralAction(action, item);
    };

    if (btnGeneralEdit) btnGeneralEdit.addEventListener("click", async () => { await runGeneralTopAction("edit"); });
    if (btnGeneralInv) btnGeneralInv.addEventListener("click", async () => { await runGeneralTopAction("inv"); });
    if (btnGeneralRecipe) btnGeneralRecipe.addEventListener("click", async () => { await runGeneralTopAction("recipe"); });

    if (btnGeneralExport){
      btnGeneralExport.addEventListener("click", async () => {
        await downloadAfterReauth("/api/inventory/export.xlsx?type=all", "Exportar Cadastro Geral (XLSX)");
      });
    }
  }
  // render já é chamado pelo refresh
}


// ---------- Histórico de Inventário (inline) ----------
let invHistInlineBound = false;

async function loadInvHistInline(){
  const res = await api("/api/inventory/inventory-history?limit=800");
  invHistState.rows = (res.history || []).map(normalizeHistRow);
  invHistState.filter = invHistState.filter || "";
  if (!invHistState.sortKey) { invHistState.sortKey = "at"; invHistState.sortDir = "desc"; }
  renderInvHistTable();
}

async function openInventoryHistoryInline(){
  if (stockInvHistoryBtn) stockInvHistoryBtn.classList.add("active");
  if (stockModeAllBtn) stockModeAllBtn.classList.remove("active");
  if (stockModeRawBtn) stockModeRawBtn.classList.remove("active");
  if (stockModeFgBtn) stockModeFgBtn.classList.remove("active");

  if (cadastroBox) cadastroBox.classList.add("hidden");
  if (generalBox) generalBox.classList.add("hidden");
  if (invHistBox) invHistBox.classList.remove("hidden");

  // carrega ajustes (MP + PF)
  await loadInvHistInline();
  invHistState.filter = "";
  invHistState.sortKey = "at";
  invHistState.sortDir = "desc";
  if (invHistSearch) invHistSearch.value = "";

  if (!invHistInlineBound){
    invHistInlineBound = true;
    if (invHistSearch){
      invHistSearch.addEventListener("input", () => {
        invHistState.filter = String(invHistSearch.value || "");
        renderInvHistTable();
      });
    }
    if (invHistTable){
      invHistTable.querySelectorAll("th.sortable").forEach((th) => {
        th.addEventListener("click", () => {
          const key = th.getAttribute("data-key");
          if (!key) return;
          if (invHistState.sortKey === key) {
            invHistState.sortDir = invHistState.sortDir === "asc" ? "desc" : "asc";
          } else {
            invHistState.sortKey = key;
            invHistState.sortDir = "asc";
          }
          renderInvHistTable();
        });
      });
    }
    if (btnInvHistClear && !btnInvHistClear.__bound){
      btnInvHistClear.__bound = true;
      btnInvHistClear.addEventListener("click", async () => {
        const okConfirm = confirm(
          "Tem certeza que deseja limpar o histórico de inventário?\n\n" +
          "Isso não altera o estoque atual, apenas oculta os registros de ajustes no histórico."
        );
        if (!okConfirm) return;
        const ok = await requestReauth({ reason: "Limpar Histórico de Inventário" });
        if (!ok) return;
        try {
          await api("/api/inventory/inventory-history/clear", { method: "POST" });
          await loadInvHistInline();
          // Re-render imediato para o usuário ver a limpeza sem precisar recarregar.
          renderInvHistTable();
          alert("Histórico de inventário limpo.");
        } catch (e) {
          console.error(e);
          if (e?.data?.error === "reauth_required") alert("Precisa confirmar usuário e senha novamente.");
          else alert("Falha ao limpar o histórico. Veja o console.");
        }
      });
    }

    if (btnInvHistExport){
      btnInvHistExport.addEventListener("click", async () => {
        await downloadAfterReauth("/api/inventory/inventory-history.xlsx?limit=5000", "Exportar Histórico de Inventário (XLSX)");
      });
    }
  }

  renderInvHistTable();
}


// ---------- Cadastro inline (Estoque) ----------
const cadastroBox = document.querySelector("#estoqueCadastro");
const cadastroTitle = document.querySelector("#cadastroTitle");
const cadastroCodeHint = document.querySelector("#cadastroCodeHint");
const cadastroSearch = document.querySelector("#cadastroSearch");
const cadastroTbody = document.querySelector("#cadastroTbody");
const cadastroPriceTh = document.querySelector("#cadastroPriceTh");
const cadastroMidTh = document.querySelector("#cadastroMidTh");
const cadastroFcTh = document.querySelector("#cadastroFcTh");
const cadastroTable = document.querySelector("#cadastroTable");
const btnImport = document.querySelector("#btnImport");
const btnExport = document.querySelector("#btnExport");
const btnCadNew = document.querySelector("#btnCadNew");
const btnCadEdit = document.querySelector("#btnCadEdit");
const btnCadDup = document.querySelector("#btnCadDup");
const btnCadDel = document.querySelector("#btnCadDel");

function currentDbLabel(){
  return state.stockMode === "fg" ? "bd/estoque_pf.json" : "bd/estoque_mp.json";
}
function currentTitle(){
  return state.stockMode === "fg" ? "Cadastro • Produto Final" : "Cadastro • Matéria Prima & Insumos";
}
function currentCodeHint(){
  return state.stockMode === "fg" ? "PF001" : "MP001";
}

function applyQuery(items){
  const q = (state.cadastroQuery || "").trim().toLowerCase();
  if (!q) return items;
  return items.filter(it => {
    const code = (it.code || it.sku || "").toLowerCase();
    return (it.name||"").toLowerCase().includes(q) || code.includes(q) || (it.unit||"").toLowerCase().includes(q);
  });
}

// Cadastros (MP/PF): sempre ordenar por código (MP001, MP002... / PF001, PF002...)
// Isso evita que itens recém-importados/recadastrados "caiam no final" da lista.
function sortCadastroItems(items){
  const arr = (items || []).slice();
  arr.sort((a, b) => {
    const ca = String(a?.code || a?.sku || "");
    const cb = String(b?.code || b?.sku || "");
    const cmp = ca.localeCompare(cb, "pt-BR", { numeric: true, sensitivity: "base" });
    if (cmp !== 0) return cmp;
    return String(a?.name || "").localeCompare(String(b?.name || ""), "pt-BR", { sensitivity: "base" });
  });
  return arr;
}

function renderCadastro(){
  if (!cadastroBox) return;
  cadastroBox.classList.remove("hidden");
  if (cadastroTitle) cadastroTitle.textContent = currentTitle();
    if (cadastroCodeHint) cadastroCodeHint.textContent = currentCodeHint();
  if (cadastroPriceTh) cadastroPriceTh.textContent = (state.stockMode === "fg") ? "V.VENDA (R$)" : "CUSTO (R$)";
  if (cadastroPriceTh) cadastroPriceTh.style.width = (state.stockMode === "fg") ? "140px" : "110px";
  // Produto Final não exibe colunas % Perda e FC
  if (cadastroTable) cadastroTable.classList.toggle("hide-loss", state.stockMode === "fg");
  if (cadastroMidTh) cadastroMidTh.textContent = "% PERDA";
  if (cadastroFcTh) cadastroFcTh.textContent = "FC";


  const items = applyQuery(sortCadastroItems(state.items || []));
  if (!cadastroTbody) return;

  if (items.length === 0){
    cadastroTbody.innerHTML = `<tr><td colspan="9" class="muted">Nenhum item cadastrado.</td></tr>`;
    return;
  }

  cadastroTbody.innerHTML = items.map(it => {
    const selected = (state.selectedItemId === it.id) ? "selected" : "";
    const code = escapeHtml(it.code || it.sku || "");
    if (state.stockMode === "fg") {
      // PF: não exibe coluna % Perda (mantemos a célula para manter a estrutura de 6 colunas e escondemos via CSS)
      return `
        <tr class="${selected}" data-item-id="${it.id}">
          <td><b>${code}</b></td>
          <td>${escapeHtml(it.name)}</td>
          <td>${escapeHtml(it.unit)}</td>
          <td class="col-sale">${fmtMoney(it.salePrice || 0)}</td>
          <td class="col-loss"></td>
          <td class="col-fc"></td>
          <td class="col-min">${fmt(it.minStock || 0)}</td>
        </tr>
      `;
    }
    return `
      <tr class="${selected}" data-item-id="${it.id}">
        <td><b>${code}</b></td>
        <td>${escapeHtml(it.name)}</td>
        <td>${escapeHtml(it.unit)}</td>
        <td>${fmtMoney(it.cost || 0)}</td>
        <td>${fmt(it.lossPercent || 0)}</td>
        <td>${fmt(it.cookFactor ?? 1)}</td>
        <td>${fmt(it.minStock || 0)}</td>
      </tr>
    `;
  }).join("");

  // row click selection
  cadastroTbody.querySelectorAll("tr[data-item-id]").forEach(tr => {
    tr.addEventListener("click", () => {
      state.selectedItemId = tr.getAttribute("data-item-id");
      renderCadastro();
    });
  });
}

function getSelectedItem(){
  return (state.items || []).find(it => it.id === state.selectedItemId) || null;
}

function requireSelection(){
  const it = getSelectedItem();
  if (!it){
    alert("Selecione um item na tabela primeiro.");
    return null;
  }
  return it;
}

if (cadastroSearch){
  cadastroSearch.addEventListener("input", () => {
    state.cadastroQuery = cadastroSearch.value || "";
    renderCadastro();
  });
}

if (btnCadNew){
  btnCadNew.addEventListener("click", () => {
    openNewItem(() => { renderCadastro(); });
  });
}
if (btnCadEdit){
  btnCadEdit.addEventListener("click", () => {
    const it = requireSelection();
    if (!it) return;
    openEditItem(it, () => { renderCadastro(); });
  });
}
if (btnCadDup){
  btnCadDup.addEventListener("click", async () => {
    const it = requireSelection();
    if (!it) return;
    const isFg = state.stockMode === "fg";
    const payload = isFg
      ? { name: `${it.name} (cópia)`, unit: it.unit, minStock: it.minStock || 0, salePrice: it.salePrice || 0 }
      : { name: `${it.name} (cópia)`, unit: it.unit, minStock: it.minStock || 0, cost: it.cost || 0, lossPercent: it.lossPercent || 0 };
    await api(`/api/inventory/items?type=${state.stockMode}`, { method: "POST", body: JSON.stringify(payload) });
    await loadInventory();
    renderCadastro();
  });
}
if (btnCadDel){
  btnCadDel.addEventListener("click", async () => {
    const it = requireSelection();
    if (!it) return;
    const ok0 = confirm(`Excluir o item "${it.name}"?`);
    if (!ok0) return;
    try{
      await api(`/api/inventory/items/${it.id}?type=${state.stockMode}`, { method: "DELETE" });
    } catch (err){
      if (err?.data?.error === "has_links"){
        const links = Array.isArray(err.data.links) ? err.data.links : [];
        const parts = [];
        if (err.data.hasMovements) parts.push("• Movimentações de estoque");
        for (const l of links.slice(0, 15)) parts.push(`• ${l.label}`);
        const msg = [
          `O item "${it.name}" possui vínculos e/ou movimentações.`,
          parts.length ? ("\n\nOnde está amarrado:\n" + parts.join("\n")) : "",
          "\n\nDeseja EXCLUIR mesmo assim? (isso removerá o item e limpará referências/movimentações dele)"
        ].join("");
        const ok = confirm(msg);
        if (!ok) return;
        await api(`/api/inventory/items/${it.id}?type=${state.stockMode}&force=1`, { method: "DELETE" });
      } else if (err?.data?.error === "has_movements") {
        // compat
        const ok = confirm(`O item "${it.name}" possui movimentações. Excluir mesmo assim?`);
        if (!ok) return;
        await api(`/api/inventory/items/${it.id}?type=${state.stockMode}&force=1`, { method: "DELETE" });
      } else {
        throw err;
      }
    }
    state.selectedItemId = null;
    await loadInventory();
    renderCadastro();
  });
}

function openPerdaHelp(){ window.open('/help/perda.html', '_blank', 'noopener,noreferrer'); }
function openFcHelp(){ window.open('/help/fc.html', '_blank', 'noopener,noreferrer'); }

function downloadText(filename, text){
  const blob = new Blob([text], { type: "application/json;charset=utf-8" });
  const url = URL.createObjectURL(blob);
  const a = document.createElement("a");
  a.href = url;
  a.download = filename;
  document.body.appendChild(a);
  a.click();
  a.remove();
  URL.revokeObjectURL(url);
}

if (btnExport){
  btnExport.addEventListener("click", async () => {
    await downloadAfterReauth(`/api/inventory/export.xlsx?type=${state.stockMode === "fg" ? "fg" : "raw"}`, "Exportar cadastro (XLSX)");
  });
}

if (btnImport){
  btnImport.addEventListener("click", () => {
    const typeLabel = (state.stockMode === "fg") ? "Produto Final" : "Matéria-prima & Insumos";
    openImportXlsxModal({
      title: "Importar Cadastro (.xlsx)",
      reason: `Importar cadastro — ${typeLabel} (XLSX)`,
      uploadUrl: `/api/inventory/import.xlsx?type=${state.stockMode === "fg" ? "fg" : "raw"}`,
      onAfter: async () => {
        await loadInventory();
        renderCadastro();
      },
    });
  });
}



// -------------------- Custos (Produto Final) --------------------
state.costing = state.costing || { last: null };

function labelForCostRecipe(r){
  const pfId = String(r?.productId || r?.outputItemId || "");
  const pf = (state.fgItems || []).find((i) => String(i.id) === pfId) || null;
  const code = pf?.code ? `${pf.code} - ` : "";
  return `${code}${r?.name || "—"}`;
}

function renderCostingTab(){
  const sel = $("#costRecipeSelect");
  if (!sel) return;
  const recipes = (state.recipes || []).filter((r) => Array.isArray(r.bom) && r.bom.length);
  const prev = sel.value;

  sel.innerHTML = recipes.map((r) => {
    const label = labelForCostRecipe(r);
    return `<option value="${escapeHtml(String(r.id))}">${escapeHtml(label)}</option>`;
  }).join("");

  if (prev && recipes.some((r) => String(r.id) === String(prev))) sel.value = prev;
  else sel.value = recipes[0]?.id || "";

  const sum = $("#costSummary");
  const tb = $("#costTbody");
  const warn = $("#costWarn");
  if (tb) tb.innerHTML = "";
  if (warn) warn.textContent = "";
  if (sum) {
    sum.innerHTML = recipes.length
      ? `<div class="muted">Selecione um Produto Final e clique em <b>Calcular</b>. (Usa o campo <b>Custo</b> de cada MP no cadastro.)</div>`
      : `<div class="muted">Nenhuma receita com BOM cadastrada.</div>`;
  }
}

function renderCostingResult(data){
  const sum = $("#costSummary");
  const tb = $("#costTbody");
  const warn = $("#costWarn");
  if (!data) return;

  const out = data.output || {};
  const totalCost = Number(data.totalCost || 0);
  const costPerUnit = Number(data.costPerUnit || 0);
  const salePrice = Number(out.salePrice || 0);
  const margin = Number(data.margin || 0);
  const marginPct = Number(data.marginPct || 0);
  const qty = Number(data.qty || 1);

  if (sum) {
    const saleTxt = salePrice > 0 ? fmtMoney(salePrice) : "—";
    const marginTxt = salePrice > 0 ? `${fmtMoney(margin)} (${fmtNum(marginPct)}%)` : "—";
    sum.innerHTML = `
      <div class="row between wrap">
        <div>
          <div><b>${escapeHtml(out.code || "")}${out.code ? " — " : ""}${escapeHtml(out.name || "")}</b></div>
          <div class="muted">Qtd analisada: <b>${escapeHtml(fmtQty(qty))}</b> ${escapeHtml(out.unit || "")} • Rendimento receita: ${escapeHtml(fmtQty(data.yieldQty || 1))} ${escapeHtml(data.yieldUnit || "un")}</div>
        </div>
        <div class="right">
          <div><span class="muted">Custo total:</span> <b>${escapeHtml(fmtMoney(totalCost))}</b></div>
          <div><span class="muted">Custo por unidade:</span> <b>${escapeHtml(fmtMoney(costPerUnit))}</b></div>
          <div><span class="muted">Preço de venda:</span> <b>${escapeHtml(saleTxt)}</b></div>
          <div><span class="muted">Margem:</span> <b>${escapeHtml(marginTxt)}</b></div>
        </div>
      </div>
    `;
  }

  const lines = Array.isArray(data.lines) ? data.lines : [];
  // Se não vier nenhum item, mostre aviso claro (evita tela vazia com custo 0)
  if (!Array.isArray(lines) || lines.length === 0) {
    if (tb) tb.innerHTML = "";
    if (warn) warn.textContent = "⚠️ Não encontrei itens no BOM dessa receita. Abra MRP • Receitas, edite a receita e salve o BOM.";
    return;
  }

  const missing = [];

  if (tb) {
    tb.innerHTML = lines.map((l) => {
      const req = Number(l.required || 0);
      const unit = l.unit || "";
      const unitCost = Number(l.unitCost || 0);
      const lineCost = Number(l.lineCost || 0);
      const pct = totalCost > 0 ? (lineCost / totalCost) * 100 : 0;
      if (unitCost <= 0) missing.push(`${l.itemCode || ""} ${l.itemName || ""}`.trim());
      return `
        <tr>
          <td>${escapeHtml(l.itemCode || "")}</td>
          <td>${escapeHtml(l.itemName || "")}</td>
          <td class="right">${escapeHtml(fmtQty(req))}</td>
          <td>${escapeHtml(unit)}</td>
          <td class="right">${escapeHtml(fmtMoney(unitCost))}</td>
          <td class="right">${escapeHtml(fmtMoney(lineCost))}</td>
          <td class="right">${escapeHtml(fmtNum(pct))}%</td>
        </tr>
      `;
    }).join("");
  }

  if (warn) {
    if (missing.length) {
      warn.textContent = `⚠️ MPs sem custo (Custo = 0): ${missing.slice(0, 10).join(" • ")}${missing.length > 10 ? " ..." : ""}.`;
    } else {
      warn.textContent = "";
    }
  }
}

async function calcCosting(){
  const sel = $("#costRecipeSelect");
  if (!sel) return;
  const recipeId = String(sel.value || "");
  if (!recipeId) return;

  const qtyEl = $("#costQty");
  const qtyIn = Number(qtyEl?.value || 1);
  const qty = (Number.isFinite(qtyIn) && qtyIn > 0) ? qtyIn : 1;
  if (qtyEl) qtyEl.value = String(qty);

  try {
    const data = await api(`/api/costing/recipe/${encodeURIComponent(recipeId)}?qty=${encodeURIComponent(qty)}`);
    state.costing.last = data;
    renderCostingResult(data);
  } catch (err) {
    showBackendOffline(err);
  }
}

function printCosting(){
  const data = state.costing?.last;
  if (!data) return;
  const out = data.output || {};
  const title = `Custo • ${out.code ? out.code + " - " : ""}${out.name || ""}`;

  const lines = Array.isArray(data.lines) ? data.lines : [];
  const totalCost = Number(data.totalCost || 0);
  const costPerUnit = Number(data.costPerUnit || 0);

  const rows = lines.map((l) => {
    const req = Number(l.required || 0);
    const unit = l.unit || "";
    const unitCost = Number(l.unitCost || 0);
    const lineCost = Number(l.lineCost || 0);
    const pct = totalCost > 0 ? (lineCost / totalCost) * 100 : 0;
    return `
      <tr>
        <td>${escapeHtml(l.itemCode || "")}</td>
        <td>${escapeHtml(l.itemName || "")}</td>
        <td style="text-align:right">${escapeHtml(fmtQty(req))}</td>
        <td>${escapeHtml(unit)}</td>
        <td style="text-align:right">${escapeHtml(fmtMoney(unitCost))}</td>
        <td style="text-align:right">${escapeHtml(fmtMoney(lineCost))}</td>
        <td style="text-align:right">${escapeHtml(fmtNum(pct))}%</td>
      </tr>
    `;
  }).join("");

  const html = `
  <html>
    <head>
      <meta charset="utf-8"/>
      <title>${escapeHtml(title)}</title>
      <style>
        body{ font-family: Arial, sans-serif; padding: 18px; }
        h1{ font-size: 18px; margin: 0 0 6px 0; }
        .muted{ color:#666; font-size:12px; }
        table{ width:100%; border-collapse: collapse; margin-top: 12px; }
        th, td{ border: 1px solid #ddd; padding: 6px; font-size: 12px; }
        th{ background:#f5f5f5; text-align:left; }
        .sum{ margin-top: 10px; font-size: 13px; }
      </style>
    </head>
    <body>
      <h1>${escapeHtml(title)}</h1>
      <div class="muted">Gerado em: ${escapeHtml(new Date().toLocaleString("pt-BR"))}</div>
      <div class="sum">
        <div><b>Custo total:</b> ${escapeHtml(fmtMoney(totalCost))}</div>
        <div><b>Custo por unidade:</b> ${escapeHtml(fmtMoney(costPerUnit))}</div>
      </div>
      <table>
        <thead>
          <tr>
            <th>MP</th><th>Descrição</th><th style="text-align:right">Qtd</th><th>UN</th><th style="text-align:right">Custo UN</th><th style="text-align:right">Custo</th><th style="text-align:right">%</th>
          </tr>
        </thead>
        <tbody>${rows}</tbody>
      </table>
    </body>
  </html>`;
  printHtmlInIframe(html);
}

// Hook UI events (if tab exists)
const _btnCostCalc = document.querySelector("#btnCostCalc");
if (_btnCostCalc) _btnCostCalc.addEventListener("click", calcCosting);
const _btnCostPrint = document.querySelector("#btnCostPrint");
if (_btnCostPrint) _btnCostPrint.addEventListener("click", printCosting);
const _selCostRecipe = document.querySelector("#costRecipeSelect");
if (_selCostRecipe) _selCostRecipe.addEventListener("change", () => {
  state.costing.last = null;
  const sum = $("#costSummary");
  if (sum) sum.innerHTML = `<div class="muted">Selecione um Produto Final e clique em <b>Calcular</b>.</div>`;
  const tb = $("#costTbody"); if (tb) tb.innerHTML = "";
  const warn = $("#costWarn"); if (warn) warn.textContent = "";
});
