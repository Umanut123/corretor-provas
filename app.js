import { createClient } from "https://cdn.jsdelivr.net/npm/@supabase/supabase-js@2.49.1/+esm";

const SUPABASE_URL = "https://ukuagyzpxklygbxjkrwt.supabase.co";
const SUPABASE_KEY = "sb_publishable_cyqAdvhO2vI58GnBE_1Ydw_Qnjct_h9";
const sb = createClient(SUPABASE_URL, SUPABASE_KEY, { auth: { persistSession: true, autoRefreshToken: true } });

// ================= utilidades =================
const $ = (s, r = document) => r.querySelector(s);
const $$ = (s, r = document) => [...r.querySelectorAll(s)];
const esc = s => String(s ?? "").replace(/[&<>"']/g, c => ({ "&": "&amp;", "<": "&lt;", ">": "&gt;", '"': "&quot;", "'": "&#39;" }[c]));
const fmt = (n, d = 2) => (n == null || isNaN(n)) ? "–" : Number(n).toLocaleString("pt-BR", { minimumFractionDigits: 0, maximumFractionDigits: d });
const num = v => { const n = parseFloat(String(v ?? "").replace(",", ".")); return isNaN(n) ? null : n; };
const r2 = n => Math.round(n * 100) / 100;
const clamp = (v, a, b) => Math.max(a, Math.min(b, v));
const pctClass = p => p == null ? "" : p >= 70 ? "ok" : p >= 50 ? "warn" : "bad";
const dataBR = d => d ? new Date(d + (String(d).length === 10 ? "T12:00:00" : "")).toLocaleDateString("pt-BR") : "";
function toast(m) { const t = $("#toast"); t.textContent = m; t.classList.add("on"); clearTimeout(toast.t); toast.t = setTimeout(() => t.classList.remove("on"), 3000); }
function confirmar(msg, ok = "Confirmar") {
  return new Promise(res => {
    const m = document.createElement("div"); m.className = "modal";
    m.innerHTML = `<div class="box"><p style="margin-top:0">${esc(msg)}</p><div class="row" style="justify-content:flex-end"><button data-a="0">Cancelar</button><button class="pri" data-a="1">${esc(ok)}</button></div></div>`;
    m.onclick = e => { const a = e.target.dataset.a; if (a != null) { m.remove(); res(a === "1"); } };
    document.body.append(m);
  });
}
function perguntar(msg, valor = "") {
  return new Promise(res => {
    const m = document.createElement("div"); m.className = "modal";
    m.innerHTML = `<div class="box"><p style="margin-top:0">${esc(msg)}</p><input type="text" dir="auto" value="${esc(valor)}"><div class="row" style="justify-content:flex-end;margin-top:12px"><button data-a="0">Cancelar</button><button class="pri" data-a="1">OK</button></div></div>`;
    const inp = $("input", m);
    m.onclick = e => { const a = e.target.dataset.a; if (a != null) { m.remove(); res(a === "1" ? inp.value.trim() : null); } };
    document.body.append(m); setTimeout(() => inp.focus(), 50);
  });
}
function zoom(src) { const m = document.createElement("div"); m.className = "modal"; m.style.padding = "0"; m.innerHTML = `<img src="${src}" style="max-width:100vw;max-height:100vh">`; m.onclick = () => m.remove(); document.body.append(m); }
function falha(e) { console.error(e); toast(e?.message?.includes("Failed to fetch") ? "Sem conexão com a internet." : (e?.message || "Algo deu errado.")); }
async function q(promise) { const { data, error } = await promise; if (error) throw error; return data; }
function baixarArquivo(nome, conteudo, tipo = "text/csv;charset=utf-8") {
  const a = document.createElement("a"); a.href = URL.createObjectURL(new Blob([conteudo], { type: tipo })); a.download = nome; a.click();
  setTimeout(() => URL.revokeObjectURL(a.href), 5000);
}
const slug = s => (s || "arquivo").normalize("NFD").replace(/[̀-ͯ]/g, "").replace(/[^a-zA-Z0-9]+/g, "-").replace(/^-|-$/g, "").toLowerCase();
const csvLinha = arr => arr.map(v => `"${String(v ?? "").replace(/"/g, '""')}"`).join(";");

// ================= Excel (.xlsx) =================
let xlsxPromise = null;
function xlsxLib() {
  if (!xlsxPromise) xlsxPromise = carregarScript("https://cdn.sheetjs.com/xlsx-0.20.3/package/dist/xlsx.full.min.js")
    .catch(() => carregarScript("https://cdn.jsdelivr.net/npm/xlsx@0.18.5/dist/xlsx.full.min.js"))
    .then(() => window.XLSX).catch(e => { xlsxPromise = null; throw e; });
  return xlsxPromise;
}
function nomeAba(s, usados) {
  const base = String(s || "Planilha").replace(/[\[\]*?/\\:]/g, " ").replace(/\s+/g, " ").trim().slice(0, 31) || "Planilha";
  let n = base, i = 2;
  while (usados.has(n.toLowerCase())) { const suf = ` (${i++})`; n = base.slice(0, 31 - suf.length) + suf; }
  usados.add(n.toLowerCase()); return n;
}
// abas: [{ nome, linhas: [[cabeçalho...], [valores...], ...] }]
async function baixarExcel(nomeArq, abas) {
  const av = aviso("Gerando o arquivo Excel…");
  try {
    const X = await xlsxLib(); const wb = X.utils.book_new(); const usados = new Set();
    for (const a of abas) {
      const ws = X.utils.aoa_to_sheet(a.linhas);
      ws["!cols"] = (a.linhas[0] || []).map((_, j) => ({ wch: Math.min(60, Math.max(8, ...a.linhas.map(l => String(l[j] ?? "").length + 2))) }));
      X.utils.book_append_sheet(wb, ws, nomeAba(a.nome, usados));
    }
    X.writeFile(wb, nomeArq);
    toast("Arquivo Excel baixado");
  } catch (e) { falha(e); toast("Não consegui gerar o Excel. Verifique a internet e tente de novo."); }
  av.fechar();
}
const n2 = v => (v == null || v === "" || isNaN(v)) ? "" : r2(Number(v));
// linhas de uma prova: um aluno por linha, com a nota de cada questão
function linhasProva(cs, am) {
  const corr = cs.filter(c => c.status === "corrigido").sort((a, b) => nomeCorr(a, am).localeCompare(nomeCorr(b, am), "pt-BR"));
  const labels = []; corr.forEach(c => (c.questoes || []).forEach(x => { if (!labels.includes(x.q)) labels.push(x.q); }));
  const linhas = [["Aluno", ...labels.map(l => `Q${l}`), "Nota", "Máximo", "%", "Revisado", "Comentário"]];
  corr.forEach(c => linhas.push([nomeCorr(c, am), ...labels.map(l => { const x = (c.questoes || []).find(y => y.q === l); return contaQ(x) ? n2(x.nota) : ""; }),
    n2(c.nota), n2(c.nota_max), c.nota_max ? Math.round(c.nota / c.nota_max * 100) : "", c.revisado ? "sim" : "não", c.comentario || ""]));
  if (corr.length) {
    const media = arr => arr.length ? r2(arr.reduce((s, v) => s + v, 0) / arr.length) : "";
    linhas.push([]);
    linhas.push(["Média da turma", ...labels.map(l => media(corr.map(c => (c.questoes || []).find(y => y.q === l)).filter(contaQ).map(x => Number(x.nota)))),
      media(corr.map(c => Number(c.nota))), "", Math.round(media(corr.map(c => c.nota_max ? c.nota / c.nota_max * 100 : 0)))]);
  }
  return linhas;
}

// ================= imagens =================
async function reduzir(src, MAX = 2400) {
  const w = src.width || src.videoWidth, h = src.height || src.videoHeight, k = Math.min(1, MAX / Math.max(w, h));
  const c = document.createElement("canvas"); c.width = Math.round(w * k); c.height = Math.round(h * k);
  const ctx = c.getContext("2d"); ctx.fillStyle = "#fff"; ctx.fillRect(0, 0, c.width, c.height); ctx.drawImage(src, 0, 0, c.width, c.height);
  return await new Promise(r => c.toBlob(r, "image/jpeg", 0.9));
}
// Mede a nitidez da foto (variância do Laplaciano numa versão reduzida, em tons de cinza).
// Foto de texto bem focada dá valor alto; foto tremida/desfocada dá valor baixo.
const NITIDEZ_MIN = 60;
async function nitidez(blob) {
  try {
    const bmp = await createImageBitmap(blob), k = Math.min(1, 1000 / Math.max(bmp.width, bmp.height));
    const w = Math.round(bmp.width * k), h = Math.round(bmp.height * k);
    const c = document.createElement("canvas"); c.width = w; c.height = h;
    const ctx = c.getContext("2d", { willReadFrequently: true }); ctx.drawImage(bmp, 0, 0, w, h);
    const d = ctx.getImageData(0, 0, w, h).data, g = new Float32Array(w * h);
    for (let i = 0; i < w * h; i++) g[i] = 0.299 * d[i * 4] + 0.587 * d[i * 4 + 1] + 0.114 * d[i * 4 + 2];
    let n = 0, soma = 0, soma2 = 0;
    for (let y = 1; y < h - 1; y += 1) for (let x = 1; x < w - 1; x += 1) {
      const i = y * w + x, v = 4 * g[i] - g[i - 1] - g[i + 1] - g[i - w] - g[i + w];
      soma += v; soma2 += v * v; n++;
    }
    const m = soma / n; return soma2 / n - m * m;
  } catch { return 999; }
}
// Pergunta se quer tirar de novo quando a foto parece borrada. Devolve true se é para FICAR com a foto.
async function aceitarFoto(blob, pagina = "") {
  const v = await nitidez(blob);
  if (v >= NITIDEZ_MIN) return true;
  return !(await confirmar(`A foto${pagina ? " da " + pagina : ""} parece borrada ou tremida, e a IA pode não conseguir ler as respostas. Quer tirar de novo?`, "Tirar de novo"));
}

let pdfjsPromise = null;
function carregarScript(src) { return new Promise((ok, no) => { const s = document.createElement("script"); s.src = src; s.onload = ok; s.onerror = no; document.head.append(s); }); }
async function pdfjs() {
  if (!pdfjsPromise) pdfjsPromise = (async () => {
    await carregarScript("https://cdnjs.cloudflare.com/ajax/libs/pdf.js/3.11.174/pdf.min.js");
    window.pdfjsLib.GlobalWorkerOptions.workerSrc = "https://cdnjs.cloudflare.com/ajax/libs/pdf.js/3.11.174/pdf.worker.min.js";
    return window.pdfjsLib;
  })();
  return pdfjsPromise;
}
async function pdfParaImagens(file) {
  const lib = await pdfjs(); const doc = await lib.getDocument({ data: await file.arrayBuffer() }).promise; const out = [];
  for (let i = 1; i <= doc.numPages; i++) {
    const p = await doc.getPage(i), vp = p.getViewport({ scale: 1 }), v = p.getViewport({ scale: Math.min(3, 2400 / Math.max(vp.width, vp.height)) });
    const c = document.createElement("canvas"); c.width = v.width; c.height = v.height;
    await p.render({ canvasContext: c.getContext("2d"), viewport: v }).promise; out.push(await reduzir(c));
  }
  return out;
}
async function pdfTexto(file) {
  const lib = await pdfjs(); const doc = await lib.getDocument({ data: await file.arrayBuffer() }).promise; let t = "";
  for (let i = 1; i <= doc.numPages; i++) { const tc = await (await doc.getPage(i)).getTextContent(); t += "\n" + tc.items.map(it => it.str + (it.hasEOL ? "\n" : " ")).join(""); }
  return t.replace(/\s/g, "").length > 80 * doc.numPages ? t.trim() : "";
}
async function arquivoParaImagens(f) {
  if (f.type === "application/pdf" || /\.pdf$/i.test(f.name)) return await pdfParaImagens(f);
  return [await reduzir(await createImageBitmap(f))];
}
function escolherArquivos(accept, multiple = true) {
  return new Promise(res => {
    const i = document.createElement("input"); i.type = "file"; i.accept = accept; i.multiple = multiple;
    i.onchange = () => res([...i.files]); i.click();
  });
}

// ================= câmera =================
// Abre a câmera em tela cheia. Cada aluno pode ter várias páginas.
// onAluno(blobs) é chamado sempre que um aluno é concluído (para já enviar e corrigir).
async function abrirCamera({ titulo = "Fotografar", umSo = false, onAluno, onAnexar, rotuloConcluir }) {
  // Usa SEMPRE a câmera normal do celular (foco e nitidez melhores). Esta tela só organiza as páginas de cada aluno.
  // Abre a câmera do celular; as fotos chegam em "onchange" (pode demorar, por isso não trava o botão).
  const fotografar = () => {
    const i = document.createElement("input"); i.type = "file"; i.accept = "image/*"; i.capture = "environment";
    i.onchange = () => receber([...i.files]); i.click();
  };
  const el = document.createElement("div"); el.className = "cam";
  el.innerHTML = `
    <div class="topbar"><button class="ico" data-a="fechar">✕</button><div class="grow" id="camInfo"></div>${onAnexar ? `<button class="side" data-a="anexar" style="background:rgba(255,255,255,.18);color:#fff;border:none">📎 Anexar arquivo</button>` : ""}</div>
    <div class="camvazio" id="camPrev"></div>
    <div class="bottom">
      <div class="strip" id="camStrip"></div>
      <div class="ctrl">
        <button class="side" data-a="desfazer">↶ Apagar última</button>
        <button class="shutter" data-a="foto" aria-label="Tirar foto"></button>
        <button class="side" data-a="proximo">${umSo ? (rotuloConcluir || "Concluir ✓") : "Próximo aluno ▸"}</button>
      </div>
    </div>`;
  document.body.append(el);
  let paginas = [], alunosFeitos = 0, ocupado = false;
  const info = () => {
    $("#camInfo", el).innerHTML = `<b>${esc(titulo)}</b><br><span class="small">${umSo ? "" : `Aluno ${alunosFeitos + 1} · `}${paginas.length} página(s)${alunosFeitos ? ` · ${alunosFeitos} aluno(s) enviados` : ""}</span>`;
    $("#camStrip", el).innerHTML = paginas.map(p => `<img src="${p.url}">`).join("");
    const ult = paginas[paginas.length - 1];
    $("#camPrev", el).innerHTML = ult ? `<img src="${ult.url}">` : `<div>Toque no botão redondo para abrir a câmera do celular.<br><br>Fotografe uma página de cada vez, com boa luz e a folha inteira aparecendo.${umSo ? "" : "<br><br>Quando terminar as páginas de um aluno, toque em “Próximo aluno”."}</div>`;
  };
  info();
  const concluirAluno = () => { if (!paginas.length) return false; const blobs = paginas.map(p => p.blob); paginas.forEach(p => URL.revokeObjectURL(p.url)); paginas = []; alunosFeitos++; onAluno(blobs); return true; };
  const fechar = () => { paginas.forEach(p => URL.revokeObjectURL(p.url)); el.remove(); };
  const receber = async fs => {
    ocupado = true;
    try {
      for (const f of fs) { const blob = await reduzir(await createImageBitmap(f)); if (await aceitarFoto(blob)) { paginas.push({ blob, url: URL.createObjectURL(blob) }); navigator.vibrate?.(30); } }
    } catch (err) { falha(err); }
    ocupado = false; info();
  };
  const tirar = () => { if (!ocupado) fotografar(); };
  el.onclick = async e => {
    const a = e.target.closest("[data-a]")?.dataset.a; if (!a || ocupado) return;
    if (a === "foto") tirar();
    else if (a === "desfazer") { const p = paginas.pop(); if (p) URL.revokeObjectURL(p.url); info(); }
    else if (a === "proximo") {
      if (!paginas.length) { toast("Tire pelo menos uma foto."); return; }
      concluirAluno(); if (umSo) fechar(); else { info(); toast("Enviado. Pode fotografar o próximo aluno."); }
    } else if (a === "anexar") {
      fechar(); onAnexar();
    } else if (a === "fechar") {
      if (paginas.length && await confirmar(`Enviar as ${paginas.length} página(s) já tiradas antes de fechar?`, "Enviar")) concluirAluno();
      fechar();
    }
  };
  if (navigator.userActivation?.isActive !== false) tirar();   // já abre a câmera do celular (se o celular deixar)
}

// ================= estado / dados =================
const S = { user: null, cache: {}, fila: [], filaAtiva: 0, rodandoIds: new Set() };
const C = (tab, id) => S.cache[tab + ":" + id];
const put = (tab, rows) => { (Array.isArray(rows) ? rows : [rows]).forEach(r => r && (S.cache[tab + ":" + r.id] = r)); return rows; };
const contaQ = x => x && x.conta !== false;   // questão entra na nota (quando o aluno escolhe N de M, as outras não contam)
const nomeCorr = (c, alunosMap) => (c.aluno_id && alunosMap?.[c.aluno_id]?.nome) || c.nome_lido || "(nome não identificado)";

async function enviarFotos(provaId, correcaoId, blobs) {
  const paths = [];
  for (let i = 0; i < blobs.length; i++) {
    const path = `${S.user.id}/${provaId}/${correcaoId}/${Date.now()}-${i}.jpg`;
    await q(sb.storage.from("provas").upload(path, blobs[i], { contentType: "image/jpeg", upsert: false }));
    paths.push(path);
  }
  return paths;
}
async function novaCorrecao(prova, blobs) {
  const id = crypto.randomUUID();
  try {
    const fotos = await enviarFotos(prova.id, id, blobs);
    put("correcoes", await q(sb.from("correcoes").insert({ id, prova_id: prova.id, fotos, status: "pendente" }).select().single()));
    lerNome(id);
    if ((prova.gabarito || "").trim()) enfileirar(id);
  } catch (e) { falha(e); toast("Falha ao enviar uma prova. Verifique a internet e tente de novo."); }
  if (rotaAtual().startsWith("#/p/" + prova.id)) render();
}
// Lê só o nome do aluno logo após o envio (rápido), para o nome aparecer antes de a correção terminar.
async function lerNome(id) {
  try {
    const { data } = await sb.functions.invoke("corrigir", { body: { acao: "ler_nome", correcao_id: id } });
    const atual = C("correcoes", id);
    if (data?.correcao && atual && atual.status !== "corrigido") { put("correcoes", { ...atual, ...data.correcao, status: atual.status }); redesenharSe(id); }
  } catch (e) { console.warn("ler nome", e); }
}
function enfileirar(id) { if (!S.fila.includes(id) && !S.rodandoIds.has(id)) S.fila.push(id); processarFila(); }
async function processarFila() {
  while (S.filaAtiva < 2 && S.fila.length) {
    const id = S.fila.shift(); S.filaAtiva++; S.rodandoIds.add(id);
    const c = C("correcoes", id); if (c) c.status = "corrigindo";
    redesenharSe(id);
    (async () => {
      try {
        const { data, error } = await sb.functions.invoke("corrigir", { body: { acao: "corrigir", correcao_id: id } });
        if (error) { let msg = error.message; try { msg = (await error.context.json()).erro || msg; } catch { } throw new Error(msg); }
        if (data?.correcao) put("correcoes", data.correcao);
      } catch (e) {
        console.error(e); const c = C("correcoes", id); if (c) { c.status = "erro"; c.erro = e.message; }
        if (/autorizado|chave|créditos|limite/i.test(e.message)) { toast(e.message); S.fila = []; }
      } finally { S.filaAtiva--; S.rodandoIds.delete(id); redesenharSe(id); processarFila(); }
    })();
  }
}
function redesenharSe(id) {
  const c = C("correcoes", id); const r = rotaAtual();
  if (c && (r.startsWith("#/p/" + c.prova_id) || r === "#/c/" + id)) { if (!edicaoAtiva()) render(true); }
}
const edicaoAtiva = () => { const a = document.activeElement; return a && /INPUT|TEXTAREA|SELECT/.test(a.tagName) && $("#main").contains(a); };

// ================= roteamento =================
const rotaAtual = () => location.hash || "#/";
const ir = h => { if (location.hash === h) render(); else location.hash = h; };
window.addEventListener("hashchange", () => render());
$("#voltar").onclick = () => { const p = S.pai; if (p) ir(p); else history.back(); };
$("#menu").onclick = async () => { if (await confirmar("Sair da sua conta neste celular?", "Sair")) { await sb.auth.signOut(); } };

function cabecalho(titulo, pai) {
  $("#titulo").textContent = titulo; S.pai = pai;
  $("#voltar").hidden = !pai; $("#menu").hidden = !!pai;
}
let renderSeq = 0;
async function render(silencioso) {
  const seq = ++renderSeq;
  const main = $("#main");
  if (!S.user) { cabecalho("Corretor de Provas"); main.innerHTML = telaLogin(); bindLogin(); return; }
  const r = rotaAtual().slice(2).split("/");
  if (!silencioso) main.innerHTML = `<div class="empty">Carregando…</div>`;
  try {
    let html, bind;
    if (!r[0]) [html, bind] = await telaInicio();
    else if (r[0] === "d") [html, bind] = await telaDisciplina(r[1], r[2] || "provas");
    else if (r[0] === "p") [html, bind] = await telaProva(r[1], r[2] || "corrigir");
    else if (r[0] === "c") [html, bind] = await telaCorrecao(r[1]);
    else if (r[0] === "a") [html, bind] = await telaAluno(r[1]);
    else { ir("#/"); return; }
    if (seq !== renderSeq) return;
    const y = window.scrollY; main.innerHTML = html; bind?.(); if (silencioso) window.scrollTo(0, y);
  } catch (e) { if (seq === renderSeq) { main.innerHTML = `<div class="err">${esc(e.message || "Erro ao carregar.")}</div><button onclick="location.reload()">Tentar de novo</button>`; } console.error(e); }
}

// ================= login =================
let modoLogin = "entrar";
function telaLogin() {
  return `<div class="card" style="margin-top:20px">
    <h2>${modoLogin === "entrar" ? "Entrar" : "Criar conta"}</h2>
    <p class="sub">${modoLogin === "entrar" ? "Use seu e-mail e senha." : "Depois de criar, confirme pelo link que chega no seu e-mail e volte aqui para entrar."}</p>
    <label class="f">E-mail<input type="email" id="lEmail" autocomplete="email"></label>
    <label class="f">Senha<input type="password" id="lSenha" autocomplete="${modoLogin === "entrar" ? "current-password" : "new-password"}"></label>
    <div id="lMsg"></div>
    <button class="pri big" id="lOk">${modoLogin === "entrar" ? "Entrar" : "Criar conta"}</button>
    <button class="ghost big" id="lTroca" style="margin-top:6px">${modoLogin === "entrar" ? "Primeira vez? Criar conta" : "Já tenho conta"}</button>
  </div>`;
}
function bindLogin() {
  $("#lTroca").onclick = () => { modoLogin = modoLogin === "entrar" ? "criar" : "entrar"; render(); };
  $("#lOk").onclick = async () => {
    const email = $("#lEmail").value.trim(), senha = $("#lSenha").value, msg = $("#lMsg");
    if (!email || senha.length < 6) { msg.innerHTML = `<div class="err">Informe o e-mail e uma senha de pelo menos 6 caracteres.</div>`; return; }
    $("#lOk").disabled = true;
    try {
      if (modoLogin === "entrar") {
        const { error } = await sb.auth.signInWithPassword({ email, password: senha });
        if (error) throw new Error(/confirm/i.test(error.message) ? "Confirme seu e-mail pelo link que enviamos antes de entrar." : "E-mail ou senha incorretos.");
      } else {
        const { error } = await sb.auth.signUp({ email, password: senha, options: { emailRedirectTo: location.origin + location.pathname } });
        if (error) throw error;
        modoLogin = "entrar"; render();
        $("#lMsg").innerHTML = `<div class="info">Conta criada. Abra o e-mail de confirmação, clique no link e depois entre aqui.</div>`;
        $("#lEmail").value = email; return;
      }
    } catch (e) { msg.innerHTML = `<div class="err">${esc(e.message)}</div>`; }
    const b = $("#lOk"); if (b) b.disabled = false;
  };
}

// ================= escolha em botões =================
function escolher(titulo, opcoes, sub = "") {
  return new Promise(res => {
    const m = document.createElement("div"); m.className = "modal"; m.style.alignItems = "flex-end";
    m.innerHTML = `<div class="box" style="max-height:85vh;overflow:auto"><h3 style="font-size:17px">${esc(titulo)}</h3>${sub ? `<p class="small muted" style="margin-top:-4px">${esc(sub)}</p>` : ""}
      ${opcoes.map((o, i) => `<button class="big ${o.pri ? "pri" : ""}" data-i="${i}" style="margin-bottom:8px;text-align:left" dir="auto">${esc(o.t)}${o.s ? `<div class="small" style="font-weight:400;opacity:.75">${esc(o.s)}</div>` : ""}</button>`).join("")}
      <button class="ghost big" data-i="-1">Cancelar</button></div>`;
    m.onclick = e => { const b = e.target.closest("[data-i]"); if (!b) return; m.remove(); const i = +b.dataset.i; res(i < 0 ? null : opcoes[i].v); };
    document.body.append(m);
  });
}
function aviso(msg) { const m = document.createElement("div"); m.className = "modal"; m.innerHTML = `<div class="box"><div id="avTxt">${msg}</div></div>`; document.body.append(m); return { set: h => { $("#avTxt", m).innerHTML = h; }, fechar: () => m.remove() }; }

async function transcreverFotos(prova, tipo, blobs) {
  const paths = await enviarFotos(prova.id, tipo, blobs);
  const { data, error } = await sb.functions.invoke("corrigir", { body: { acao: tipo === "modelo" ? "transcrever_modelo" : "transcrever_gabarito", fotos: paths } });
  if (error) { let m = error.message; try { m = (await error.context.json()).erro || m; } catch { } throw new Error(m); }
  const campo = tipo === "modelo" ? "modelo_fotos" : "gabarito_fotos";
  try { await q(sb.from("provas").update({ [campo]: [...(prova[campo] || []), ...paths] }).eq("id", prova.id)); } catch { }
  return data.texto;
}
async function corrigirPendentes(provaId) {
  const ps = await q(sb.from("correcoes").select("id").eq("prova_id", provaId).in("status", ["pendente", "erro"]));
  ps.forEach(c => enfileirar(c.id));
  if (ps.length) toast(`Corrigindo ${ps.length} prova(s) que estavam esperando o gabarito.`);
}
// lê arquivos anexados: imagens (por arquivo) e textos (Word/TXT/PDF digital)
async function lerAnexos(fs) {
  const grupos = [], textos = [];
  for (const f of fs) {
    try {
      if (/\.docx$/i.test(f.name)) { await carregarScript("https://cdn.jsdelivr.net/npm/mammoth@1.8.0/mammoth.browser.min.js"); textos.push((await window.mammoth.extractRawText({ arrayBuffer: await f.arrayBuffer() })).value.trim()); }
      else if (/\.txt$/i.test(f.name)) textos.push((await f.text()).trim());
      else if (/\.pdf$/i.test(f.name) || f.type === "application/pdf") { const imgs = await pdfParaImagens(f); grupos.push(imgs); const t = await pdfTexto(f).catch(() => ""); if (t) textos.push(t); }
      else grupos.push(await arquivoParaImagens(f));
    } catch (e) { falha(e); }
  }
  return { grupos, textos };
}

// Fluxo principal: foto/arquivo → o que é → disciplina → prova
async function fluxoEntrada(entrada) {
  const tipo = await escolher("O que é isso?", [
    { v: "feita", t: "📝 Prova feita", s: "Prova respondida por um aluno — vai para a correção", pri: true },
    { v: "gabarito", t: "✅ Gabarito", s: "As respostas certas (ou a prova preenchida por você)" },
    { v: "modelo", t: "📄 Exemplo de prova", s: "A prova em branco, com as perguntas" },
  ]);
  if (!tipo) return;
  const ds = await q(sb.from("disciplinas").select("*").order("nome").order("turma"));
  if (!ds.length) { toast("Cadastre uma disciplina primeiro."); return; }
  const discId = await escolher("De qual disciplina?", ds.map(d => ({ v: d.id, t: `${d.nome} — ${d.turma || ""}` })));
  if (!discId) return;
  const disc = ds.find(d => d.id === discId);
  const provas = await q(sb.from("provas").select("*").eq("disciplina_id", discId).order("created_at", { ascending: false }));
  const novaProva = async () => {
    const titulo = await perguntar("Nome da prova", `Prova ${new Date().toLocaleDateString("pt-BR")}`); if (!titulo) return null;
    return put("provas", await q(sb.from("provas").insert({ disciplina_id: discId, titulo }).select().single()));
  };
  const escolherProva = async (titulo) => {
    if (!provas.length) return await novaProva();
    const v = await escolher(titulo, [...provas.map(p => ({ v: p.id, t: p.titulo, s: `${dataBR(p.data)} · ${p.gabarito.trim() ? "com gabarito" : "sem gabarito"}` })), { v: "__nova", t: "+ Nova prova" }], `${disc.nome} — ${disc.turma || ""}`);
    if (!v) return null; if (v === "__nova") return await novaProva(); return provas.find(p => p.id === v);
  };

  if (tipo === "feita") {
    if (!entrada.grupos.length) { toast("Para corrigir preciso de foto ou PDF da prova."); return; }
    let prova = provas.length === 1 ? provas[0] : await escolherProva("Qual prova é essa?");
    if (!prova) return;
    let lotes = entrada.grupos;
    if (lotes.length > 1) {
      const m = await escolher(`Você enviou ${lotes.length} arquivos.`, [{ v: "cada", t: "Cada arquivo é de um aluno diferente", pri: true }, { v: "um", t: "Todos são do mesmo aluno" }]);
      if (!m) return; if (m === "um") lotes = [lotes.flat()];
    }
    const av = aviso(`Enviando ${lotes.length} prova(s)…`);
    for (const l of lotes) await novaCorrecao(prova, l);
    av.fechar();
    const semGab = !prova.gabarito.trim();
    const prox = await escolher(semGab ? "Prova guardada ✓" : "Enviada para correção ✓",
      [{ v: "prox", t: "📷 Fotografar o próximo aluno", s: `${disc.nome} — ${disc.turma || ""} · ${prova.titulo}`, pri: true }, { v: "ver", t: "Ver as correções desta prova" }, { v: "fim", t: "Concluir" }],
      semGab ? "Esta prova ainda não tem gabarito. Assim que você enviar o gabarito, eu corrijo todas." : "A IA identifica o nome do aluno sozinha.");
    if (prox === "prox") abrirCamera({ titulo: `${disc.nome} — ${disc.turma || ""}`, onAluno: blobs => novaCorrecao(prova, blobs) });
    else if (prox === "ver") ir(`#/p/${prova.id}/corrigir`);
    else render(true);
    return;
  }

  // gabarito ou modelo
  const prova = await escolherProva(tipo === "gabarito" ? "Gabarito de qual prova?" : "Modelo de qual prova?");
  if (!prova) return;
  const campo = tipo === "gabarito" ? "gabarito" : "modelo";
  let texto = entrada.textos.join("\n\n");
  const av = aviso(tipo === "gabarito" ? "Lendo o gabarito… (até 1 minuto)" : "Lendo o modelo da prova… (até 1 minuto)");
  try {
    if (!texto && entrada.grupos.length) texto = await transcreverFotos(prova, tipo, entrada.grupos.flat());
    if (!texto) throw new Error("Não consegui ler o conteúdo.");
    let final = texto;
    if ((prova[campo] || "").trim()) {
      av.fechar();
      const m = await escolher(`Esta prova já tem ${tipo === "gabarito" ? "gabarito" : "modelo"}.`, [{ v: "sub", t: "Substituir pelo novo", pri: true }, { v: "junta", t: "Juntar com o que já existe" }]);
      if (!m) return; if (m === "junta") final = prova[campo].trim() + "\n\n" + texto;
    } else av.fechar();
    await q(sb.from("provas").update({ [campo]: final }).eq("id", prova.id));
    if (tipo === "gabarito") await corrigirPendentes(prova.id);
    toast(tipo === "gabarito" ? "Gabarito salvo. Confira o texto." : "Modelo salvo. Confira o texto.");
    ir(`#/p/${prova.id}/gabarito`);
  } catch (e) { av.fechar(); falha(e); }
}
async function capturar() {
  abrirCamera({ titulo: "Nova foto", umSo: true, rotuloConcluir: "Pronto ✓", onAluno: blobs => fluxoEntrada({ grupos: [blobs], textos: [] }), onAnexar: anexar });
}
async function anexar() {
  const fs = await escolherArquivos("image/*,application/pdf,.pdf,.docx,.txt"); if (!fs.length) return;
  const av = aviso("Preparando arquivos…"); const e = await lerAnexos(fs); av.fechar();
  if (!e.grupos.length && !e.textos.length) return;
  fluxoEntrada(e);
}

// ================= início =================
async function telaInicio() {
  cabecalho("Corretor de Provas");
  const ds = put("disciplinas", await q(sb.from("disciplinas").select("*").order("nome").order("turma")));
  const html = `
    <div class="card" style="text-align:center">
      <button class="pri big" id="iFoto" style="font-size:19px;padding:22px">📷 Tirar foto</button>
      <button class="big" id="iAnexo" style="margin-top:10px">📎 Anexar arquivo</button>
      <p class="small muted" style="margin:10px 0 0">Prova feita, gabarito ou exemplo de prova — eu pergunto depois.</p>
    </div>
    <h3 style="margin:18px 2px 8px">Minhas disciplinas</h3>
    <div class="card">${ds.length ? ds.map(d => `
      <div class="item" data-go="#/d/${d.id}"><div class="grow"><div class="t" dir="auto">${esc(d.nome)}</div><div class="s">${esc(d.turma || "")}</div></div><span class="chev">›</span></div>`).join("")
      : `<div class="empty">Nenhuma disciplina ainda.</div>`}
      <button class="ghost" id="novaDisc" style="width:100%;margin-top:6px">+ Nova disciplina</button></div>`;
  return [html, () => {
    $$("[data-go]").forEach(el => el.onclick = () => ir(el.dataset.go));
    $("#iFoto").onclick = capturar; $("#iAnexo").onclick = anexar;
    $("#novaDisc").onclick = async () => {
      const nome = await perguntar("Nome da disciplina (ex.: Rashi)"); if (!nome) return;
      const turma = await perguntar("Turma (ex.: Shiur Alef)") ?? "";
      try { const d = await q(sb.from("disciplinas").insert({ nome, turma }).select().single()); put("disciplinas", d); ir("#/d/" + d.id); } catch (e) { falha(e); }
    };
    if (!S.cameraInicialAberta) { S.cameraInicialAberta = true; capturar(); }
  }];
}

// ================= disciplinas =================
async function telaDisciplinas() {
  cabecalho("Minhas disciplinas");
  const ds = put("disciplinas", await q(sb.from("disciplinas").select("*").order("nome")));
  const html = `
    <div class="card">${ds.length ? ds.map(d => `
      <div class="item" data-go="#/d/${d.id}"><div class="grow"><div class="t" dir="auto">${esc(d.nome)}</div><div class="s">${esc(d.turma || "")}</div></div><span class="chev">›</span></div>`).join("")
      : `<div class="empty">Nenhuma disciplina ainda.<br>Crie a primeira no botão abaixo.</div>`}</div>
    <button class="pri big fab" id="novaDisc">+ Nova disciplina</button>`;
  return [html, () => {
    $$("[data-go]").forEach(el => el.onclick = () => ir(el.dataset.go));
    $("#novaDisc").onclick = async () => {
      const nome = await perguntar("Nome da disciplina (ex.: Guemará – 2º ano)"); if (!nome) return;
      const turma = await perguntar("Turma (opcional)") ?? "";
      try { const d = await q(sb.from("disciplinas").insert({ nome, turma }).select().single()); put("disciplinas", d); ir("#/d/" + d.id); } catch (e) { falha(e); }
    };
  }];
}

async function telaDisciplina(id, aba) {
  const d = put("disciplinas", await q(sb.from("disciplinas").select("*").eq("id", id).single()));
  cabecalho(d.nome, "#/");
  const tabs = `<div class="tabs">${[["provas", "Provas"], ["alunos", "Alunos"], ["boletim", "Boletim"]].map(([k, l]) => `<button class="${aba === k ? "on" : ""}" data-go="#/d/${id}/${k}">${l}</button>`).join("")}</div>`;
  let html = "", bind = () => { };
  if (aba === "provas") {
    const ps = put("provas", await q(sb.from("provas").select("*, correcoes(count)").eq("disciplina_id", id).order("data", { ascending: false })));
    html = `<div class="card">${ps.length ? ps.map(p => `
      <div class="item" data-go="#/p/${p.id}"><div class="grow"><div class="t" dir="auto">${esc(p.titulo)}</div>
        <div class="s">${dataBR(p.data)} · ${p.correcoes?.[0]?.count || 0} prova(s) · ${p.gabarito?.trim() ? "gabarito ok" : "sem gabarito"}</div></div><span class="chev">›</span></div>`).join("")
      : `<div class="empty">Nenhuma prova nesta disciplina.</div>`}</div>
      <button class="ghost danger" id="exclDisc" style="width:100%">Excluir disciplina</button>
      <button class="pri big fab" id="novaProva">+ Nova prova</button>`;
    bind = () => {
      $("#novaProva").onclick = async () => {
        const titulo = await perguntar("Nome da prova (ex.: Prova 1 – Beitzá)"); if (!titulo) return;
        try { const p = await q(sb.from("provas").insert({ disciplina_id: id, titulo }).select().single()); put("provas", p); ir(`#/p/${p.id}/gabarito`); } catch (e) { falha(e); }
      };
      $("#exclDisc").onclick = async () => { if (await confirmar(`Excluir "${d.nome}" com TODOS os alunos, provas e notas? Não dá para desfazer.`, "Excluir")) { try { await q(sb.from("disciplinas").delete().eq("id", id)); ir("#/"); } catch (e) { falha(e); } } };
    };
  } else if (aba === "alunos") {
    const as = put("alunos", await q(sb.from("alunos").select("*").eq("disciplina_id", id).order("nome")));
    html = `<p class="sub">A IA usa esta lista para reconhecer o nome escrito na prova (em hebraico ou português).</p>
      <div class="card">${as.length ? as.map(a => `
        <div class="item" data-go="#/a/${a.id}"><div class="grow t" dir="auto">${esc(a.nome)}</div><span class="chev">›</span></div>`).join("") : `<div class="empty">Nenhum aluno cadastrado.</div>`}</div>
      <div class="card"><h3>Adicionar alunos</h3>
        <label class="f">Um nome por linha (pode colar a lista inteira)<textarea id="novosAlunos" rows="4" dir="auto"></textarea></label>
        <button class="pri" id="addAlunos">Adicionar</button></div>`;
    bind = () => {
      $("#addAlunos").onclick = async () => {
        const nomes = $("#novosAlunos").value.split("\n").map(s => s.trim()).filter(Boolean); if (!nomes.length) return;
        try { await q(sb.from("alunos").insert(nomes.map(nome => ({ disciplina_id: id, nome })))); toast(`${nomes.length} aluno(s) adicionado(s)`); render(); } catch (e) { falha(e); }
      };
    };
  } else {
    const [as, ps, cs] = await Promise.all([
      q(sb.from("alunos").select("*").eq("disciplina_id", id).order("nome")),
      q(sb.from("provas").select("id,titulo,data").eq("disciplina_id", id).order("data")),
      q(sb.from("correcoes").select("id,prova_id,aluno_id,nome_lido,nota,nota_max,status,questoes,revisado,comentario,provas!inner(disciplina_id)").eq("provas.disciplina_id", id).eq("status", "corrigido")),
    ]);
    const cel = (a, p) => cs.find(c => c.aluno_id === a.id && c.prova_id === p.id);
    const linhas = as.map(a => { const v = ps.map(p => cel(a, p)); const soma = v.reduce((s, c) => s + (c ? Number(c.nota) : 0), 0), max = v.reduce((s, c) => s + (c ? Number(c.nota_max) : 0), 0); return { a, v, soma, max }; });
    html = !as.length || !ps.length ? `<div class="empty">O boletim aparece quando houver alunos e provas corrigidas.</div>` : `
      <div class="card scroll"><table>
        <thead><tr><th>Aluno</th>${ps.map(p => `<th class="num" dir="auto">${esc(p.titulo)}</th>`).join("")}<th class="num">Total</th></tr></thead>
        <tbody>${linhas.map(({ a, v, soma, max }) => `<tr data-go="#/a/${a.id}"><td dir="auto">${esc(a.nome)}</td>${v.map(c => `<td class="num">${c ? `${fmt(c.nota)}<span class="small muted">/${fmt(c.nota_max)}</span>` : "–"}</td>`).join("")}<td class="num"><b>${max ? fmt(soma) : "–"}</b>${max ? `<span class="small muted">/${fmt(max)}</span>` : ""}</td></tr>`).join("")}</tbody>
      </table></div>
      <button class="pri big" id="expBol">📊 Baixar Excel com as notas</button>
      <p class="small muted" style="margin:6px 2px 0">O arquivo tem uma aba com todos os alunos e provas, e uma aba para cada prova com a nota de cada questão.</p>`;
    bind = () => {
      const b = $("#expBol"); if (b) b.onclick = () => {
        const am = Object.fromEntries(as.map(a => [a.id, a]));
        const boletim = [["Aluno", ...ps.map(p => p.titulo + (p.data ? ` (${dataBR(p.data)})` : "")), "Total", "Máximo", "%"],
          ...linhas.map(({ a, v, soma, max }) => [a.nome, ...v.map(c => c ? n2(c.nota) : ""), max ? n2(soma) : "", max ? n2(max) : "", max ? Math.round(soma / max * 100) : ""])];
        const abas = [{ nome: "Boletim", linhas: boletim }];
        ps.forEach(p => { const cp = cs.filter(c => c.prova_id === p.id); if (cp.length) abas.push({ nome: p.titulo, linhas: linhasProva(cp, am) }); });
        baixarExcel(`notas-${slug(d.nome)}-${slug(d.turma || "")}.xlsx`.replace("-.xlsx", ".xlsx"), abas);
      };
    };
  }
  return [tabs + html, () => { $$("[data-go]").forEach(el => el.onclick = () => ir(el.dataset.go)); bind(); }];
}

async function telaAluno(id) {
  const a = put("alunos", await q(sb.from("alunos").select("*, disciplinas(nome)").eq("id", id).single()));
  cabecalho(a.nome, `#/d/${a.disciplina_id}/alunos`);
  const cs = await q(sb.from("correcoes").select("id,nota,nota_max,status,revisado,provas(titulo,data)").eq("aluno_id", id).order("created_at"));
  const corr = cs.filter(c => c.status === "corrigido");
  const soma = corr.reduce((s, c) => s + Number(c.nota), 0), max = corr.reduce((s, c) => s + Number(c.nota_max), 0);
  const html = `
    <div class="stats"><div class="stat"><div class="v">${corr.length}</div><div class="l">provas corrigidas</div></div>
      <div class="stat"><div class="v">${max ? fmt(soma / max * 100, 0) + "%" : "–"}</div><div class="l">aproveitamento geral</div></div></div>
    <div class="card">${cs.length ? cs.map(c => { const p = c.nota_max ? c.nota / c.nota_max * 100 : null; return `
      <div class="item" data-go="#/c/${c.id}"><div class="grow"><div class="t" dir="auto">${esc(c.provas?.titulo)}</div><div class="s">${dataBR(c.provas?.data)}${c.status === "corrigido" && !c.revisado ? " · a revisar" : ""}</div></div>
      ${c.status === "corrigido" ? `<b>${fmt(c.nota)}</b><span class="small muted">/${fmt(c.nota_max)}</span> <span class="pill ${pctClass(p)}">${fmt(p, 0)}%</span>` : `<span class="pill">${esc(c.status)}</span>`}<span class="chev">›</span></div>`; }).join("") : `<div class="empty">Nenhuma prova deste aluno ainda.</div>`}</div>
    <div class="row"><button id="renAluno">Renomear</button><span class="grow"></span><button class="ghost danger" id="exclAluno">Excluir aluno</button></div>`;
  return [html, () => {
    $$("[data-go]").forEach(el => el.onclick = () => ir(el.dataset.go));
    $("#renAluno").onclick = async () => { const n = await perguntar("Nome do aluno", a.nome); if (!n) return; try { await q(sb.from("alunos").update({ nome: n }).eq("id", id)); render(); } catch (e) { falha(e); } };
    $("#exclAluno").onclick = async () => { if (await confirmar(`Excluir ${a.nome}? As provas dele ficam, mas sem vínculo com o aluno.`, "Excluir")) { try { await q(sb.from("alunos").delete().eq("id", id)); ir(`#/d/${a.disciplina_id}/alunos`); } catch (e) { falha(e); } } };
  }];
}

// ================= prova =================
let modoGaleria = "cada";
async function telaProva(id, aba) {
  const p = put("provas", await q(sb.from("provas").select("*, disciplinas(nome)").eq("id", id).single()));
  cabecalho(p.titulo, `#/d/${p.disciplina_id}`);
  if (!p.gabarito.trim() && aba !== "gabarito") aba = "gabarito";
  const tabs = `<div class="tabs">${[["gabarito", "Gabarito"], ["corrigir", "Corrigir"], ["resultados", "Resultados"]].map(([k, l]) => `<button class="${aba === k ? "on" : ""}" data-go="#/p/${id}/${k}">${l}</button>`).join("")}</div>`;
  let r;
  if (aba === "gabarito") r = abaGabarito(p);
  else {
    const [cs, as] = await Promise.all([
      q(sb.from("correcoes").select("*").eq("prova_id", id).order("created_at")),
      q(sb.from("alunos").select("id,nome").eq("disciplina_id", p.disciplina_id).order("nome")),
    ]);
    put("correcoes", cs);
    cs.forEach(c => { if (S.rodandoIds.has(c.id)) c.status = "corrigindo"; });
    const am = Object.fromEntries(as.map(a => [a.id, a]));
    r = aba === "corrigir" ? abaCorrigir(p, cs, am) : abaResultados(p, cs, am);
  }
  return [tabs + r[0], () => { $$(".tabs [data-go]").forEach(el => el.onclick = () => ir(el.dataset.go)); r[1](); }];
}

function abaGabarito(p) {
  const html = `
    ${!p.gabarito.trim() ? `<div class="info">Primeiro cadastre o gabarito. Depois é só fotografar as provas.</div>` : ""}
    <div class="card">
      <label class="f">Nome da prova<input type="text" id="gTit" dir="auto" value="${esc(p.titulo)}"></label>
      <div class="row"><label class="f grow">Data<input type="date" id="gData" value="${esc(p.data || "")}"></label>
        <label class="f grow">Valor padrão de cada questão<input type="text" inputmode="decimal" id="gVal" value="${esc(fmt(p.valor_padrao))}"></label></div>
    </div>
    <div class="card">
      <h3>Gabarito</h3>
      <p class="small muted" style="margin-top:-4px">Escreva o valor de cada questão no próprio gabarito, ex.: "1) (10 pontos) …". Questões sem valor usam o valor padrão.</p>
      <div class="row" style="margin-bottom:10px">
        <button class="grow" id="gFoto">📷 Fotografar gabarito</button>
        <button class="grow" id="gArq">📎 Enviar arquivo</button>
      </div>
      <div id="gMsg"></div>
      <textarea id="gTxt" dir="auto" style="min-height:240px" placeholder="1) (10 pontos) …&#10;2) (15 pontos) …&#10;3) (5 pontos) …">${esc(p.gabarito)}</textarea>
      <h3 style="margin-top:16px">Exemplo da prova (modelo em branco) <span class="small muted" style="font-weight:400">— opcional, ajuda a IA a entender as perguntas</span></h3>
      <div class="row" style="margin-bottom:8px"><button class="grow" id="mFoto">📷 Fotografar modelo</button></div>
      <textarea id="mTxt" dir="auto" style="min-height:120px" placeholder="Enunciados das questões…">${esc(p.modelo || "")}</textarea>
      <label class="f" style="margin-top:10px">Orientações de correção (opcional)
        <textarea id="gInst" rows="2" placeholder="Ex.: o aluno escolhe 10 de 15 questões; aceitar resposta em hebraico ou português; não descontar ortografia.">${esc(p.instrucoes)}</textarea></label>
    </div>
    <button class="ghost danger" id="exclProva" style="width:100%">Excluir prova</button>
    <button class="pri big fab" id="gSalvar">Salvar gabarito</button>`;
  const transcrever = async blobs => {
    const msg = $("#gMsg"); msg.innerHTML = `<div class="info">Enviando e lendo o gabarito… (pode levar até 1 minuto)</div>`;
    try {
      const paths = await enviarFotos(p.id, "gabarito", blobs);
      const { data, error } = await sb.functions.invoke("corrigir", { body: { acao: "transcrever_gabarito", fotos: paths } });
      if (error) { let m = error.message; try { m = (await error.context.json()).erro || m; } catch { } throw new Error(m); }
      const t = $("#gTxt"); t.value = (t.value.trim() ? t.value.trim() + "\n\n" : "") + data.texto;
      msg.innerHTML = `<div class="info">Gabarito transcrito. Confira (principalmente o hebraico) e toque em Salvar.</div>`;
    } catch (e) { msg.innerHTML = `<div class="err">${esc(e.message)}</div>`; }
  };
  return [html, () => {
    $("#gFoto").onclick = () => abrirCamera({ titulo: "Gabarito", umSo: true, onAluno: transcrever });
    $("#mFoto").onclick = () => abrirCamera({ titulo: "Modelo da prova", umSo: true, onAluno: async blobs => {
      const msg = $("#gMsg"); msg.innerHTML = `<div class="info">Lendo o modelo da prova…</div>`;
      try { const t = await transcreverFotos(p, "modelo", blobs); const m = $("#mTxt"); m.value = (m.value.trim() ? m.value.trim() + "\n\n" : "") + t; msg.innerHTML = `<div class="info">Modelo transcrito. Confira e toque em Salvar.</div>`; }
      catch (e) { msg.innerHTML = `<div class="err">${esc(e.message)}</div>`; }
    } });
    $("#gArq").onclick = async () => {
      const fs = await escolherArquivos("image/*,application/pdf,.pdf,.docx,.txt"); if (!fs.length) return;
      const imgs = [];
      for (const f of fs) {
        try {
          if (/\.docx$/i.test(f.name)) { await carregarScript("https://cdn.jsdelivr.net/npm/mammoth@1.8.0/mammoth.browser.min.js"); const r = await window.mammoth.extractRawText({ arrayBuffer: await f.arrayBuffer() }); const t = $("#gTxt"); t.value = (t.value.trim() ? t.value.trim() + "\n\n" : "") + r.value.trim(); }
          else if (/\.txt$/i.test(f.name)) { const t = $("#gTxt"); t.value = (t.value.trim() ? t.value.trim() + "\n\n" : "") + (await f.text()).trim(); }
          else if (/\.pdf$/i.test(f.name) || f.type === "application/pdf") { const tx = await pdfTexto(f); if (tx) { const t = $("#gTxt"); t.value = (t.value.trim() ? t.value.trim() + "\n\n" : "") + tx; } else imgs.push(...await pdfParaImagens(f)); }
          else imgs.push(...await arquivoParaImagens(f));
        } catch (e) { falha(e); }
      }
      if (imgs.length) transcrever(imgs); else $("#gMsg").innerHTML = `<div class="info">Texto colocado no gabarito. Confira e toque em Salvar.</div>`;
    };
    $("#gSalvar").onclick = async () => {
      const upd = { titulo: $("#gTit").value.trim() || p.titulo, data: $("#gData").value || null, valor_padrao: num($("#gVal").value) || 1, gabarito: $("#gTxt").value, modelo: $("#mTxt").value, instrucoes: $("#gInst").value };
      try { await q(sb.from("provas").update(upd).eq("id", p.id)); toast("Gabarito salvo"); if (upd.gabarito.trim()) await corrigirPendentes(p.id); ir(`#/p/${p.id}/corrigir`); } catch (e) { falha(e); }
    };
    $("#exclProva").onclick = async () => { if (await confirmar(`Excluir a prova "${p.titulo}" e todas as correções dela?`, "Excluir")) { try { await q(sb.from("provas").delete().eq("id", p.id)); ir(`#/d/${p.disciplina_id}`); } catch (e) { falha(e); } } };
  }];
}

function statusPill(c) {
  if (S.rodandoIds.has(c.id) || c.status === "corrigindo") return `<span class="pill run">Corrigindo…</span>`;
  if (S.fila.includes(c.id)) return `<span class="pill run">Na fila</span>`;
  if (c.status === "erro") return `<span class="pill bad">Erro</span>`;
  if (c.status === "pendente") return `<span class="pill">Pendente</span>`;
  const p = c.nota_max ? c.nota / c.nota_max * 100 : null;
  return `<b>${fmt(c.nota)}</b><span class="small muted">/${fmt(c.nota_max)}</span> <span class="pill ${pctClass(p)}">${fmt(p, 0)}%</span>`;
}
function abaCorrigir(p, cs, am) {
  const pend = cs.filter(c => (c.status === "pendente" || c.status === "erro") && !S.rodandoIds.has(c.id) && !S.fila.includes(c.id));
  const contagem = {}; cs.forEach(c => { if (c.aluno_id) contagem[c.aluno_id] = (contagem[c.aluno_id] || 0) + 1; });
  const html = `
    <div class="card">
      <button class="pri big" id="cCam">📷 Fotografar provas</button>
      <p class="small muted" style="margin:8px 0 0">Fotografe todas as páginas de um aluno e toque em <b>Próximo aluno</b>. A correção começa enquanto você fotografa os próximos.</p>
      <div class="row" style="margin-top:12px">
        <div class="seg grow"><button data-mg="cada" class="${modoGaleria === "cada" ? "on" : ""}">1 aluno por arquivo</button><button data-mg="um" class="${modoGaleria === "um" ? "on" : ""}">Todos do mesmo aluno</button></div>
        <button id="cGal">🖼 Da galeria</button>
      </div>
    </div>
    ${pend.length ? `<button class="big" id="cPend" style="margin-bottom:12px">Corrigir pendentes (${pend.length})</button>` : ""}
    <div class="card">${cs.length ? cs.map(c => `
      <div class="item" data-go="#/c/${c.id}"><div class="grow"><div class="t" dir="auto">${esc(nomeCorr(c, am))}</div>
        <div class="s">${c.status === "corrigido" ? (c.revisado ? "revisado" : "a revisar") : c.status === "erro" ? esc(c.erro || "erro") : `${c.fotos.length} página(s)`}
        ${c.status === "corrigido" && !c.aluno_id ? ` · <span style="color:var(--warn)">aluno não vinculado</span>` : ""}${c.aluno_id && contagem[c.aluno_id] > 1 ? ` · <span style="color:var(--bad)">duplicado</span>` : ""}</div></div>
        ${statusPill(c)}<span class="chev">›</span></div>`).join("") : `<div class="empty">Nenhuma prova enviada ainda.</div>`}</div>`;
  return [html, () => {
    $$("[data-go]").forEach(el => el.onclick = () => ir(el.dataset.go));
    $$("[data-mg]").forEach(b => b.onclick = () => { modoGaleria = b.dataset.mg; render(true); });
    $("#cCam").onclick = () => abrirCamera({ titulo: p.titulo, onAluno: blobs => novaCorrecao(p, blobs) });
    $("#cGal").onclick = async () => {
      const fs = await escolherArquivos("image/*,application/pdf,.pdf"); if (!fs.length) return;
      toast("Preparando imagens…");
      const grupos = [];
      for (const f of fs) { try { grupos.push(await arquivoParaImagens(f)); } catch (e) { falha(e); } }
      const lotes = modoGaleria === "cada" ? grupos : [grupos.flat()];
      toast(`Enviando ${lotes.length} prova(s)…`);
      for (const l of lotes) if (l.length) await novaCorrecao(p, l);
    };
    const pb = $("#cPend"); if (pb) pb.onclick = () => { pend.forEach(c => enfileirar(c.id)); render(true); };
  }];
}

function abaResultados(p, cs, am) {
  const corr = cs.filter(c => c.status === "corrigido");
  if (!corr.length) return [`<div class="empty">Nenhuma prova corrigida ainda.</div>`, () => { }];
  const pcts = corr.map(c => c.nota_max ? c.nota / c.nota_max * 100 : 0);
  const media = corr.reduce((s, c) => s + Number(c.nota), 0) / corr.length;
  const maxv = Math.max(...corr.map(c => Number(c.nota_max)));
  const labels = []; corr.forEach(c => c.questoes.forEach(x => { if (!labels.includes(x.q)) labels.push(x.q); }));
  const porQ = labels.map(l => { const xs = corr.map(c => c.questoes.find(x => x.q === l)).filter(contaQ); return { l, valor: xs[0]?.valor, media: xs.reduce((s, x) => s + x.nota, 0) / xs.length, pct: xs.reduce((s, x) => s + x.pct, 0) / xs.length, n: xs.length }; }).filter(x => x.n).sort((a, b) => a.pct - b.pct);
  const ord = [...corr].sort((a, b) => nomeCorr(a, am).localeCompare(nomeCorr(b, am), "pt-BR"));
  const naoRev = corr.filter(c => !c.revisado).length;
  const html = `
    ${naoRev ? `<div class="note">${naoRev} correção(ões) ainda não revisada(s).</div>` : ""}
    <div class="stats">
      <div class="stat"><div class="v">${corr.length}</div><div class="l">provas corrigidas</div></div>
      <div class="stat"><div class="v">${fmt(media)}</div><div class="l">nota média (de ${fmt(maxv)})</div></div>
      <div class="stat"><div class="v">${fmt(Math.max(...pcts), 0)}% / ${fmt(Math.min(...pcts), 0)}%</div><div class="l">maior / menor</div></div>
      <div class="stat"><div class="v">${pcts.filter(x => x >= 60).length}</div><div class="l">com 60% ou mais</div></div>
    </div>
    <div class="card scroll"><table>
      <thead><tr><th>Aluno</th>${labels.map(l => `<th class="num">Q${esc(l)}</th>`).join("")}<th class="num">Nota</th></tr></thead>
      <tbody>${ord.map(c => `<tr data-go="#/c/${c.id}"><td dir="auto">${esc(nomeCorr(c, am))}${c.revisado ? "" : " <span class='pill warn'>revisar</span>"}</td>${labels.map(l => { const x = c.questoes.find(y => y.q === l); return `<td class="num">${contaQ(x) ? fmt(x.nota) : "–"}</td>`; }).join("")}<td class="num"><b>${fmt(c.nota)}</b><span class="small muted">/${fmt(c.nota_max)}</span></td></tr>`).join("")}</tbody>
    </table></div>
    <div class="row" style="margin-bottom:12px"><button class="grow pri" id="rCsv">📊 Baixar Excel</button><button class="grow" id="rDev">Baixar devolutivas</button></div>
    <div class="card"><h3>Por questão (da mais difícil)</h3>
      ${porQ.map(x => `<div style="margin:10px 0"><div class="row"><b class="grow">Questão ${esc(x.l)}</b><span class="small">${x.n < corr.length ? `${x.n} aluno(s) · ` : ""}média ${fmt(x.media)} de ${fmt(x.valor)} · ${fmt(x.pct, 0)}%</span></div>
        <div class="bar ${x.pct < 50 ? "lo" : x.pct < 70 ? "mid" : ""}"><i style="width:${x.pct}%"></i></div></div>`).join("")}</div>
    <div class="card"><div class="row"><h3 class="grow" style="margin:0">Análise da turma</h3><button id="rAn">${p.analise ? "Gerar de novo" : "Gerar com IA"}</button></div>
      <div id="rAnTxt" style="white-space:pre-wrap;margin-top:10px">${p.analise ? esc(p.analise) : `<span class="muted small">Pontos fortes, dificuldades e sugestões para a próxima aula.</span>`}</div></div>`;
  return [html, () => {
    $$("[data-go]").forEach(el => el.onclick = () => ir(el.dataset.go));
    $("#rCsv").onclick = () => baixarExcel(`${slug(p.titulo)}-notas.xlsx`, [{ nome: p.titulo, linhas: linhasProva(cs, am) }]);
    $("#rDev").onclick = () => baixarArquivo(`${slug(p.titulo)}-devolutivas.txt`, ord.map(c => devolutiva(c, am)).join("\n\n------------------------------\n\n"), "text/plain;charset=utf-8");
    $("#rAn").onclick = async () => {
      const b = $("#rAn"); b.disabled = true; $("#rAnTxt").textContent = "Analisando…";
      try { const { data, error } = await sb.functions.invoke("corrigir", { body: { acao: "analise", prova_id: p.id } }); if (error) { let m = error.message; try { m = (await error.context.json()).erro || m; } catch { } throw new Error(m); } $("#rAnTxt").textContent = data.texto; }
      catch (e) { $("#rAnTxt").textContent = ""; falha(e); }
      b.disabled = false;
    };
  }];
}
function devolutiva(c, am) {
  return `${nomeCorr(c, am)} — ${fmt(c.nota)} de ${fmt(c.nota_max)}\n\n${c.comentario || ""}\n\n` + (c.escolher ? `(Você devia escolher ${c.escolher} questões; contaram as ${c.escolher} melhores.)\n\n` : "") +
    c.questoes.filter(contaQ).map(x => `• Questão ${x.q}: ${fmt(x.nota)} de ${fmt(x.valor)}${x.obs ? ` — ${x.obs}` : ""}`).join("\n");
}

// ================= correção (detalhe) =================
async function telaCorrecao(id) {
  const c = put("correcoes", await q(sb.from("correcoes").select("*").eq("id", id).single()));
  if (S.rodandoIds.has(id)) c.status = "corrigindo";
  const p = put("provas", await q(sb.from("provas").select("id,titulo,disciplina_id").eq("id", c.prova_id).single()));
  const as = await q(sb.from("alunos").select("id,nome").eq("disciplina_id", p.disciplina_id).order("nome"));
  const am = Object.fromEntries(as.map(a => [a.id, a]));
  cabecalho(nomeCorr(c, am), `#/p/${p.id}/corrigir`);
  let urls = [];
  if (c.fotos.length) { const { data } = await sb.storage.from("provas").createSignedUrls(c.fotos, 3600); urls = (data || []).map(x => x.signedUrl).filter(Boolean); }
  const pctT = c.nota_max ? c.nota / c.nota_max * 100 : null;
  const html = `
    <div class="card">
      <div class="thumbs">${urls.map(u => `<img src="${u}" data-zoom>`).join("")}</div>
      <label class="f" style="margin-top:10px">Aluno ${c.nome_lido ? `<span>· escrito na prova: <b dir="auto">${esc(c.nome_lido)}</b>${c.nome_incerto ? " <span class='pill warn'>conferir</span>" : ""}</span>` : ""}
        <select id="cAluno"><option value="">— escolher aluno —</option>${as.map(a => `<option value="${a.id}" ${a.id === c.aluno_id ? "selected" : ""}>${esc(a.nome)}</option>`).join("")}<option value="__novo">+ Cadastrar novo aluno…</option></select></label>
      ${c.status === "erro" ? `<div class="err">${esc(c.erro || "Erro na correção.")}</div>` : ""}
      ${c.status === "corrigindo" ? `<div class="info">Corrigindo… a tela atualiza sozinha.</div>` : ""}
      ${c.status === "pendente" ? `<div class="info">Ainda não corrigida.</div>` : ""}
      <div class="row">
        <button id="cRefazer" ${c.status === "corrigindo" ? "disabled" : ""}>${c.status === "corrigido" ? "Corrigir de novo" : "Corrigir agora"}</button>
        <button id="cMaisPag">+ Página</button><span class="grow"></span>
        <button class="ghost danger" id="cExcl">Excluir</button>
      </div>
    </div>
    ${c.status === "corrigido" ? `
    <div class="card">
      <div class="row" style="margin-bottom:4px"><h3 class="grow" style="margin:0">Nota: ${fmt(c.nota)} de ${fmt(c.nota_max)}</h3><span class="pill ${pctClass(pctT)}">${fmt(pctT, 0)}%</span></div>
      ${c.escolher ? `<div class="info small">O aluno escolhe <b>${c.escolher}</b> de ${c.questoes.length} questões: contam as ${c.escolher} respondidas com maior nota. Toque em “Contar / Não contar” para trocar.</div>` : ""}
      ${c.questoes.map((x, i) => `<div class="q" ${contaQ(x) ? "" : `style="opacity:.55"`}>
        <div class="row"><b class="grow">Questão ${esc(x.q)} <span class="small muted">· vale ${fmt(x.valor)}</span>${contaQ(x) ? "" : ` <span class="pill">não conta</span>`}</b>
          <input class="notaIn" inputmode="decimal" data-n="${i}" value="${esc(fmt(x.nota))}"><span class="small muted">/ ${fmt(x.valor)}</span></div>
        <div class="ans" dir="auto">${esc(x.resposta || "(em branco)")}</div>
        <div class="row small"><span class="pill ${pctClass(x.pct)}">atendeu ${fmt(x.pct, 0)}%</span>${x.legivel === false ? `<span class="pill warn">leitura incerta</span>` : ""}<span class="grow"></span>${c.escolher ? `<button class="ghost" data-conta="${i}">${contaQ(x) ? "Não contar" : "Contar"}</button>` : ""}</div>
        ${x.obs ? `<div class="obs" style="margin-top:6px" dir="auto">${esc(x.obs)}</div>` : ""}
      </div>`).join("")}
    </div>
    <div class="card"><label class="f">Comentário para o aluno<textarea id="cCom" rows="4" dir="auto">${esc(c.comentario)}</textarea></label>
      <button id="cCopiar">Copiar devolutiva</button></div>
    <button class="pri big fab" id="cRev" ${c.revisado ? "disabled" : ""}>${c.revisado ? "Revisado ✓" : "Marcar como revisado"}</button>` : ""}`;
  return [html, () => {
    $$("[data-zoom]").forEach(i => i.onclick = () => zoom(i.src));
    $("#cAluno").onchange = async e => {
      let v = e.target.value;
      if (v === "__novo") { const n = await perguntar("Nome do novo aluno", c.nome_lido || ""); if (!n) { e.target.value = c.aluno_id || ""; return; } try { v = (await q(sb.from("alunos").insert({ disciplina_id: p.disciplina_id, nome: n }).select().single())).id; } catch (er) { falha(er); return; } }
      try { await q(sb.from("correcoes").update({ aluno_id: v || null, nome_incerto: false }).eq("id", id)); toast("Aluno vinculado"); render(true); } catch (er) { falha(er); }
    };
    $("#cRefazer").onclick = () => { enfileirar(id); render(true); };
    $("#cMaisPag").onclick = () => abrirCamera({ titulo: "Página extra", umSo: true, onAluno: async blobs => { try { const novos = await enviarFotos(p.id, id, blobs); await q(sb.from("correcoes").update({ fotos: [...c.fotos, ...novos] }).eq("id", id)); toast("Página adicionada. Toque em Corrigir de novo."); render(true); } catch (e) { falha(e); } } });
    $("#cExcl").onclick = async () => { if (await confirmar("Excluir esta prova corrigida e as fotos?", "Excluir")) { try { if (c.fotos.length) await sb.storage.from("provas").remove(c.fotos); await q(sb.from("correcoes").delete().eq("id", id)); ir(`#/p/${p.id}/corrigir`); } catch (e) { falha(e); } } };
    if (c.status !== "corrigido") return;
    $$("[data-n]").forEach(inp => inp.onchange = async () => {
      const x = c.questoes[+inp.dataset.n]; x.nota = r2(clamp(num(inp.value) ?? 0, 0, Number(x.valor))); x.pct = x.valor ? Math.round(x.nota / x.valor * 100) : 0;
      await salvarQuestoes();
    });
    async function salvarQuestoes() {
      const cont = c.questoes.filter(contaQ);
      const nota = r2(cont.reduce((s, y) => s + Number(y.nota), 0)), nota_max = r2(cont.reduce((s, y) => s + Number(y.valor), 0));
      try { await q(sb.from("correcoes").update({ questoes: c.questoes, nota, nota_max, revisado: false }).eq("id", id)); render(true); } catch (e) { falha(e); }
    }
    $$("[data-conta]").forEach(b => b.onclick = async () => {
      const x = c.questoes[+b.dataset.conta]; x.conta = !contaQ(x);
      const n = c.questoes.filter(contaQ).length;
      if (c.escolher && n !== c.escolher) toast(`Atenção: agora contam ${n} questões (o certo são ${c.escolher}).`);
      await salvarQuestoes();
    });
    $("#cCom").onchange = async e => { try { await q(sb.from("correcoes").update({ comentario: e.target.value }).eq("id", id)); } catch (er) { falha(er); } };
    $("#cCopiar").onclick = async () => { try { await navigator.clipboard.writeText(devolutiva({ ...c, comentario: $("#cCom").value }, am)); toast("Copiado"); } catch { toast("Não foi possível copiar"); } };
    $("#cRev").onclick = async () => { try { await q(sb.from("correcoes").update({ revisado: true }).eq("id", id)); toast("Revisado"); ir(`#/p/${p.id}/corrigir`); } catch (e) { falha(e); } };
  }];
}

// ================= início =================
sb.auth.onAuthStateChange((_ev, session) => {
  const antes = S.user?.id; S.user = session?.user || null;
  if (antes !== S.user?.id) render();
});
(async () => { const { data } = await sb.auth.getSession(); S.user = data.session?.user || null; render(); })();
if ("serviceWorker" in navigator) navigator.serviceWorker.register("sw.js").catch(() => { });
