const API_BASE =
  (location.hostname === "127.0.0.1" || location.hostname === "localhost")
    ? "http://127.0.0.1:8000/api"
    : `${location.origin}/api`;

const el = (id) => document.getElementById(id);

const state = {
  lastCalc: null,
  loading: false,
  appBooted: false,
};

function setLoginMsg(text) {
  const n = el("loginMsg");
  if (n) n.textContent = text || "";
}

function getToken() { return localStorage.getItem("token"); }
function getRole() { return localStorage.getItem("role"); }

function clearAuth() {
  localStorage.removeItem("token");
  localStorage.removeItem("role");
  localStorage.removeItem("user");
}

function isAllowedRole(role) {
  return role === "admin" || role === "engenharia";
}

function showLogin(msg = "") {
  const app = el("app");
  const login = el("login");
  if (app) app.style.display = "none";
  if (login) login.style.display = "flex";
  setLoginMsg(msg);
}

function showApp() {
  const login = el("login");
  const app = el("app");
  if (login) login.style.display = "none";
  if (app) app.style.display = "block";

  const user = localStorage.getItem("user") || "";
  const role = getRole() || "";
  const who = el("whoami");
  if (who) who.textContent = user ? `Logado: ${user} (${role})` : "";
}

function escapeHtml(value) {
  return String(value ?? "")
    .replace(/&/g, "&")
    .replace(/</g, "<")
    .replace(/>/g, ">")
    .replace(/"/g, "&quot;")
    .replace(/'/g, "&#39;");
}

function setMsg(text, kind = "") {
  const box = el("msg");
  if (!box) return;
  box.className = "msg" + (kind ? ` ${kind}` : "");
  box.textContent = text || "";
}

function setLoading(isLoading) {
  state.loading = isLoading;
  const b1 = el("btnCalcular");
  const b2 = el("btnLimpar");
  if (b1) b1.disabled = isLoading;
  if (b2) b2.disabled = isLoading;
}

function fmt(n, digits = 2) {
  if (n === null || n === undefined || Number.isNaN(Number(n))) return "—";
  return Number(n).toFixed(digits);
}

function badge(status) {
  const s = String(status || "");
  if (s === "APROVADO") return `<span class="badge ok">APROVADO</span>`;
  if (s === "ATENÇÃO") return `<span class="badge warn">ATENÇÃO</span>`;
  return `<span class="badge bad">REPROVADO</span>`;
}

async function apiGet(path) {
  const token = getToken();
  const r = await fetch(`${API_BASE}${path}`, {
    headers: token ? { Authorization: `Bearer ${token}` } : {}
  });
  if (!r.ok) throw new Error(`GET ${path} -> ${r.status}`);
  return await r.json();
}

async function apiPost(path, body) {
  const token = getToken();
  const r = await fetch(`${API_BASE}${path}`, {
    method: "POST",
    headers: {
      "Content-Type": "application/json",
      ...(token ? { Authorization: `Bearer ${token}` } : {}),
    },
    body: JSON.stringify(body),
  });
  const data = await r.json().catch(() => ({}));
  if (!r.ok) throw new Error(data?.detail || `POST ${path} -> ${r.status}`);
  return data;
}

async function login(username, password) {
  const form = new URLSearchParams();
  form.set("username", username);
  form.set("password", password);

  const r = await fetch(`${API_BASE}/auth/token`, {
    method: "POST",
    headers: { "Content-Type": "application/x-www-form-urlencoded" },
    body: form.toString(),
  });

  const data = await r.json().catch(() => ({}));
  if (!r.ok) throw new Error(data?.detail || "Falha no login");

  localStorage.setItem("token", data.access_token);
  localStorage.setItem("role", data.role || "");
  localStorage.setItem("user", username);

  if (!isAllowedRole(data.role)) {
    clearAuth();
    throw new Error("Seu perfil ainda não está liberado para uso.");
  }
}

async function checkHealth() {
  try {
    const data = await apiGet("/health");
    const dot = el("apiDot");
    const st = el("apiStatus");
    if (dot) dot.style.background = data.last_error ? "#f59e0b" : "#16a34a";
    if (st) st.textContent = data.last_error ? "API: ok (com erro no Excel)" : "API: ok";
    if (data.last_error) setMsg(`API conectou, mas o Excel não carregou:\n${data.last_error}`, "warn");
  } catch (e) {
    const dot = el("apiDot");
    const st = el("apiStatus");
    if (dot) dot.style.background = "#dc2626";
    if (st) st.textContent = "API: offline";
    setMsg(String(e.message || e), "bad");
  }
}

function readInputs() {
  const kwpSis = Number(el("kwpSis").value);
  const qtdInv = Number(el("qtdInv").value);
  const tmin = Number(el("tmin").value);
  const tmax = Number(el("tmax").value);

  const fabMod = el("fabMod").value || "";
  const modeloMod = el("modeloMod").value || "";
  const fabInv = el("fabInv").value || "";
  const modeloInv = el("modeloInv").value || "";

  const tamStringRaw = el("tamString").value.trim();
  const entradasRaw = el("entradasUsadas").value.trim();

  return {
    kwp_sis: kwpSis,
    qtd_inv: Math.max(1, Math.floor(qtdInv || 1)),
    fabricante_mod: fabMod,
    modelo_mod: modeloMod,
    fabricante_inv: fabInv,
    modelo_inv: modeloInv,
    tmin,
    tmax,
    tam_string: tamStringRaw ? Number(tamStringRaw) : null,
    entradas_usadas: entradasRaw ? Number(entradasRaw) : null,
  };
}

function validatePayload(p) {
  if (!p.fabricante_mod || !p.modelo_mod) return "Selecione fabricante/modelo do módulo.";
  if (!p.fabricante_inv || !p.modelo_inv) return "Selecione fabricante/modelo do inversor.";
  if (!(p.kwp_sis > 0)) return "Potência (kWp) deve ser > 0.";
  if (!(p.qtd_inv >= 1)) return "Qtd inversores deve ser >= 1.";
  if (!(p.tmin < p.tmax)) return "Temperaturas inválidas: Tmin deve ser menor que Tmax.";
  return null;
}

function renderCards(data) {
  const set = (id, value) => {
    const node = el(id);
    if (node) node.textContent = value;
  };

  if (!data?.ok) {
    ["mPotInv","mPotSis","mDif","mTam","mQtd","mTotal","mOv","mOvMax"].forEach(id => set(id, "—"));
    return;
  }

  const m = data.melhor;
  set("mPotInv", fmt(m.pot, 2));
  set("mPotSis", fmt(data.pot_sis_melhor, 2));
  set("mDif", fmt(m.dif, 2));
  set("mTam", String(Math.trunc(m.tamanho)));
  set("mQtd", String(Math.trunc(m.qtd)));
  set("mTotal", String(Math.trunc(m.total)));
  set("mOv", fmt(m.overload * 100, 1));
  set("mOvMax", fmt(data.ovmax * 100, 0));
}

function renderMelhor(data) {
  const tbody = el("tblMelhor")?.querySelector("tbody");
  if (!tbody) return;
  tbody.innerHTML = "";

  if (!data?.ok) {
    tbody.innerHTML = `<tr><td colspan="5" class="muted">—</td></tr>`;
    return;
  }

  const m = data.melhor;
  tbody.innerHTML = `
    <tr class="best">
      <td>${Math.trunc(m.tamanho)}</td>
      <td>${Math.trunc(m.qtd)}</td>
      <td>${Math.trunc(m.total)}</td>
      <td>${fmt(m.pot, 2)}</td>
      <td>${fmt(m.dif, 2)}</td>
    </tr>
  `;
}

function renderCombos(data) {
  const tbody = el("tblCombos")?.querySelector("tbody");
  if (!tbody) return;
  tbody.innerHTML = "";

  if (!data?.ok) {
    tbody.innerHTML = `<tr><td colspan="7" class="muted">—</td></tr>`;
    return;
  }

  const rows = data.combos || [];
  if (!rows.length) {
    tbody.innerHTML = `<tr><td colspan="7" class="muted">Nenhuma combinação retornada.</td></tr>`;
    return;
  }

  rows.forEach((r, idx) => {
    const tr = document.createElement("tr");
    tr.className = (idx === 0) ? "best" : "";
    tr.innerHTML = `
      <td>${idx + 1}</td>
      <td>${Math.trunc(r.tamanho)}</td>
      <td>${Math.trunc(r.qtd)}</td>
      <td>${Math.trunc(r.total)}</td>
      <td>${fmt(r.pot, 2)}</td>
      <td>${fmt(r.dif, 2)}</td>
      <td>${fmt(r.overload * 100, 1)}%</td>
    `;
    tr.addEventListener("click", async () => {
      el("tamString").value = String(Math.trunc(r.tamanho));
      el("entradasUsadas").value = String(Math.trunc(r.qtd));
      await doCalcular(true);
    });
    tbody.appendChild(tr);
  });
}

function renderCriterios(data) {
  const tbody = el("tblCriterios")?.querySelector("tbody");
  if (!tbody) return;
  tbody.innerHTML = "";

  if (!data?.ok) {
    tbody.innerHTML = `<tr><td colspan="2" class="muted">—</td></tr>`;
    return;
  }

  const crit = data.criterios_melhor || [];
  if (!crit.length) {
    tbody.innerHTML = `<tr><td colspan="2" class="muted">Sem critérios retornados.</td></tr>`;
    return;
  }

  crit.forEach(([nome, status]) => {
    if (nome === "Corrente de operação" && status === "REPROVADO") status = "ATENÇÃO";
    const tr = document.createElement("tr");
    tr.innerHTML = `<td style="text-align:left">${escapeHtml(nome)}</td><td>${badge(status)}</td>`;
    tbody.appendChild(tr);
  });
}

function v(obj, key, fallback = "—") {
  const val = obj?.[key];
  if (val === null || val === undefined) return fallback;
  const s = String(val).trim();
  if (!s || s.toLowerCase() === "nan") return fallback;
  return s;
}

function vAny(obj, keys, fallback = "—") {
  for (const k of keys) {
    const val = v(obj, k, null);
    if (val !== null) return val;
  }
  return fallback;
}

function renderResumo(data) {
  const wrap = el("resumo");
  const raw = el("raw");
  if (!wrap || !raw) return;

  if (!data?.ok) {
    wrap.innerHTML = `<div class="muted">—</div>`;
    raw.textContent = data ? JSON.stringify(data, null, 2) : "—";
    return;
  }

  const mod = data.mod || {};
  const inv = data.inv || {};
  const corr = data.correcoes || {};
  const lim = data.intervalo_string || {};

  const invCorrenteMaxCA = vAny(inv, [
    "CORRENTE MÁX. SAÍDA (A)",
    "CORRENTE MÁX. CA (A)",
    "CORRENTE MÁXIMA CA (A)",
    "CORRENTE NOMINAL CA (A)",
  ]);

  const mpptFaixa = (() => {
    const vstart = v(inv, "TENSÃO MÍN PARTIDA. (V)");
    const vmppMax = v(inv, "TENSÃO MÁX. MPP (V)");
    if (vstart === "—" || vmppMax === "—") return "—";
    return `${vstart} – ${vmppMax} V`;
  })();

  const intervaloString = (lim?.n_min && lim?.n_max)
    ? `${lim.n_min} – ${lim.n_max} módulos/string`
    : "—";

  wrap.innerHTML = `
    <div class="box">
      <h4>Módulo</h4>
      <table class="kv">
        <tr><td class="k">Fabricante</td><td class="v">${escapeHtml(v(mod, "FABRICANTE"))}</td></tr>
        <tr><td class="k">Modelo</td><td class="v">${escapeHtml(v(mod, "MODELO"))}</td></tr>
        <tr><td class="k">Potência</td><td class="v">${escapeHtml(v(mod, "POTÊNCIA (KWP)"))}</td></tr>
        <tr><td class="k">Tipo de célula</td><td class="v">${escapeHtml(v(mod, "TIPO CÉLULA"))}</td></tr>
        <tr><td class="k">Tensão Voc</td><td class="v">${escapeHtml(v(mod, "TENSÃO ABERTO (Voc)"))} V</td></tr>
        <tr><td class="k">Corrente Isc</td><td class="v">${escapeHtml(v(mod, "CORRENTE CURTO (A)"))} A</td></tr>
        <tr><td class="k">Tensão Vmp</td><td class="v">${escapeHtml(v(mod, "TENSÃO OP. STC (Vmp)"))} V</td></tr>
        <tr><td class="k">Corrente Imp</td><td class="v">${escapeHtml(v(mod, "CORRENTE (A)"))} A</td></tr>
        <tr><td class="k">Coef. temperatura Voc</td><td class="v">${escapeHtml(v(mod, "COEF. TEMP. TENSÃO ABERTO (%/°C)"))}</td></tr>
        <tr><td class="k">Coef. temperatura Isc</td><td class="v">${escapeHtml(v(mod, "COEF. TEMP. CORR. CURTO (%/°C)"))}</td></tr>
      </table>
    </div>

    <div class="box">
      <h4>Inversor</h4>
      <table class="kv">
        <tr><td class="k">Fabricante</td><td class="v">${escapeHtml(v(inv, "FABRICANTE"))}</td></tr>
        <tr><td class="k">Modelo</td><td class="v">${escapeHtml(v(inv, "MODELO"))}</td></tr>
        <tr><td class="k">Potência nominal</td><td class="v">${escapeHtml(v(inv, "MAX. POT. SAÍDA (KW)"))}</td></tr>
        <tr><td class="k">Potência Max. de entrada (CC)</td><td class="v">${escapeHtml(v(inv, "MAX. POT. ENTRADA (KWP)"))}</td></tr>
        <tr><td class="k">Tensão Max CC</td><td class="v">${escapeHtml(v(inv, "TENSÃO MÁX. ENTRADA (V)"))} V</td></tr>
        <tr><td class="k">Tensão de Partida</td><td class="v">${escapeHtml(v(inv, "TENSÃO MÍN PARTIDA. (V)"))} V</td></tr>
        <tr><td class="k">Faixa de tensão MPPT</td><td class="v">${escapeHtml(mpptFaixa)}</td></tr>
        <tr><td class="k">Corrente de Entrada</td><td class="v">${escapeHtml(v(inv, "CORRENTE MÁX. ENTRADA FV (A)"))} A</td></tr>
        <tr><td class="k">Corrente de Curto</td><td class="v">${escapeHtml(v(inv, "CORRENTE MÁX. CURTO (A)"))} A</td></tr>
        <tr><td class="k">N° de Entradas</td><td class="v">${escapeHtml(v(inv, "Nº ENTRADAS INVERSOR"))}</td></tr>
        <tr><td class="k">N° de MPPTS</td><td class="v">${escapeHtml(v(inv, "Nº MPPTS"))}</td></tr>
        <tr><td class="k">Corrente Max CA</td><td class="v">${escapeHtml(invCorrenteMaxCA)}</td></tr>
        <tr><td class="k">Tensão Nominal CA</td><td class="v">${escapeHtml(v(inv, "TENSÃO SAÍDA (V)"))} V</td></tr>
      </table>
    </div>

    <div class="box">
      <h4>Cálculo de Correção</h4>
      <table class="kv">
        <tr><td class="k">Tensão Voc corrigida</td><td class="v">${fmt(corr.voc_corrigida, 2)} V</td></tr>
        <tr><td class="k">Corrente Isc corrigida</td><td class="v">${fmt(corr.isc_corrigida, 2)} A</td></tr>
        <tr><td class="k">Tensão Vmp corrigida</td><td class="v">${fmt(corr.vmp_corrigida, 2)} V</td></tr>
      </table>
    </div>

    <div class="box">
      <h4>Distribuição Max</h4>
      <table class="kv">
        <tr><td class="k">Intervalo de String</td><td class="v">${escapeHtml(intervaloString)}</td></tr>
      </table>
    </div>
  `;

  raw.textContent = JSON.stringify(data, null, 2);
}

async function doCalcular(fromCombo = false) {
  const payload = readInputs();
  const err = validatePayload(payload);
  if (err) { setMsg(err, "warn"); return; }

  try {
    setLoading(true);
    setMsg(fromCombo ? "Recalculando…" : "Calculando…");

    const data = await apiPost("/calcular", payload);
    state.lastCalc = data;

    renderCards(data);
    renderMelhor(data);
    renderCombos(data);
    renderCriterios(data);
    renderResumo(data);

    setMsg("Cálculo concluído.", "ok");
  } catch (e) {
    renderCards(null);
    renderMelhor(null);
    renderCombos(null);
    renderCriterios(null);
    renderResumo(null);
    setMsg(String(e.message || e), "bad");
  } finally {
    setLoading(false);
  }
}

function initTabs() {
  document.querySelectorAll(".tab").forEach(btn => {
    btn.addEventListener("click", () => {
      document.querySelectorAll(".tab").forEach(b => b.classList.remove("active"));
      document.querySelectorAll(".pane").forEach(p => p.classList.remove("active"));
      btn.classList.add("active");
      const name = btn.dataset.tab;
      const pane = el(`pane-${name}`);
      if (pane) pane.classList.add("active");
    });
  });
}

function limpar() {
  el("kwpSis").value = "10";
  el("qtdInv").value = "1";
  el("tmin").value = "0";
  el("tmax").value = "50";
  el("tamString").value = "";
  el("entradasUsadas").value = "";
  setMsg("Campos limpos.");
}

async function loadSelects() {
  const fabsMod = await apiGet("/modulos/fabricantes");
  el("fabMod").innerHTML = fabsMod.map(f => `<option value="${escapeHtml(f)}">${escapeHtml(f)}</option>`).join("");
  await onFabModChange();

  const fabsInv = await apiGet("/inversores/fabricantes");
  el("fabInv").innerHTML = fabsInv.map(f => `<option value="${escapeHtml(f)}">${escapeHtml(f)}</option>`).join("");
  await onFabInvChange();
}

async function onFabModChange() {
  const fab = el("fabMod").value;
  const modelos = await apiGet(`/modulos?fabricante=${encodeURIComponent(fab)}`);
  el("modeloMod").innerHTML = modelos.map(m => `<option value="${escapeHtml(m)}">${escapeHtml(m)}</option>`).join("");
}

async function onFabInvChange() {
  const fab = el("fabInv").value;
  const modelos = await apiGet(`/inversores?fabricante=${encodeURIComponent(fab)}`);
  el("modeloInv").innerHTML = modelos.map(m => `<option value="${escapeHtml(m)}">${escapeHtml(m)}</option>`).join("");
}

/* ===== Cadastro de Usuário (admin) ===== */

async function listUsers() {
  return await apiGet("/users");
}

async function createUser(username, password, role) {
  return await apiPost("/users", { username, password, role, is_active: true });
}

function renderUsersTable(users) {
  const tbody = el("tblUsers")?.querySelector("tbody");
  if (!tbody) return;

  if (!users?.length) {
    tbody.innerHTML = `<tr><td colspan="4" class="muted">Nenhum usuário.</td></tr>`;
    return;
  }

  tbody.innerHTML = users.map(u => `
    <tr>
      <td>${u.id}</td>
      <td style="text-align:left">${escapeHtml(u.username)}</td>
      <td>${escapeHtml(u.role)}</td>
      <td>${u.is_active ? "Sim" : "Não"}</td>
    </tr>
  `).join("");
}

async function adminUsersBoot() {
  const role = getRole();
  const gate = el("cadastroGate");
  const box = el("cadastroBox");
  const msg = el("usersMsg");

  if (role !== "admin") {
    if (gate) gate.style.display = "block";
    if (box) box.style.display = "none";
    return;
  }

  if (gate) gate.style.display = "none";
  if (box) box.style.display = "block";

  async function reload() {
    try {
      if (msg) msg.textContent = "Carregando usuários…";
      const users = await listUsers();
      renderUsersTable(users);
      if (msg) msg.textContent = "";
    } catch (e) {
      if (msg) msg.textContent = String(e.message || e);
    }
  }

  el("btnReloadUsers")?.addEventListener("click", reload);

  el("btnCreateUser")?.addEventListener("click", async () => {
    const u = el("newUserName").value.trim();
    const p = el("newUserPass").value;
    const r = el("newUserRole").value;

    try {
      if (msg) msg.textContent = "Criando…";
      await createUser(u, p, r);
      if (msg) msg.textContent = `Usuário "${u}" criado com role "${r}".`;
      el("newUserPass").value = "";
      await reload();
    } catch (e) {
      if (msg) msg.textContent = String(e.message || e);
    }
  });

  await reload();
}

/* ===== Boot ===== */

async function afterAuthBoot() {
  showApp();
  if (state.appBooted) return;
  state.appBooted = true;

  initTabs();
  await checkHealth();

  el("fabMod").addEventListener("change", () => onFabModChange().catch(e => setMsg(String(e.message || e), "bad")));
  el("fabInv").addEventListener("change", () => onFabInvChange().catch(e => setMsg(String(e.message || e), "bad")));
  el("btnCalcular").addEventListener("click", () => doCalcular(false));
  el("btnLimpar").addEventListener("click", limpar);

  setLoading(true);
  setMsg("Carregando lista de fabricantes/modelos…");
  try {
    await loadSelects();
    setMsg("Pronto. Selecione os itens e clique em Calcular.");
  } catch (e) {
    setMsg("Falha carregando listas:\n" + String(e.message || e), "bad");
  } finally {
    setLoading(false);
  }

  await adminUsersBoot();
}

async function onLoginClick() {
  try {
    const u = el("loginUser").value.trim();
    const p = el("loginPass").value;
    setLoginMsg("Entrando…");
    await login(u, p);
    await afterAuthBoot();
    setLoginMsg("");
  } catch (e) {
    clearAuth();
    showLogin(String(e.message || e));
  }
}

async function main() {
  showLogin("Faça login para usar o sistema.");

  el("btnLogin").addEventListener("click", (ev) => { ev.preventDefault(); onLoginClick(); });
  el("loginPass").addEventListener("keydown", (ev) => { if (ev.key === "Enter") onLoginClick(); });

  el("btnLogout").addEventListener("click", () => {
    clearAuth();
    showLogin("Sessão encerrada.");
  });

  const token = getToken();
  const role = getRole();

  if (token && isAllowedRole(role)) {
    try {
      await apiGet("/auth/me");
      await afterAuthBoot();
      return;
    } catch {
      clearAuth();
      showLogin("Sessão expirada. Faça login novamente.");
      return;
    }
  }
}

window.addEventListener("DOMContentLoaded", main);