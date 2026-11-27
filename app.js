/* ===== PDF no 11A e na Previsão: DESLIGADO até validar com PDFs reais =====
   A extração por grade (pdfParaGrade / lerArquivoTabela) continua no código e foi testada com PDFs gerados pelo Excel.
   Para reativar: trocar a constante abaixo para true (ela libera .pdf nos campos de arquivo, ajusta os textos e permite a leitura).
   O 16B aceita PDF independentemente desta chave (já validado com PDF real). */
const PDF_GRADE_HABILITADO = false;
(function aplicarPdfGrade() {
  if (!PDF_GRADE_HABILITADO) return;
  [["file11a", "dz11a", "Arraste o relatório 11A (.xlsx ou .pdf) aqui, ou clique para selecionar"], ["filepo", "dzpo", "Arraste a previsão orçamentária (.xlsx ou .pdf) aqui, ou clique para selecionar"]].forEach(([inp, dz, txt]) => {
    const i = document.getElementById(inp), d = document.getElementById(dz);
    if (!i || !d) return;
    i.accept = ".xlsx,.xls,.pdf";
    d.childNodes.forEach((n) => { if (n.nodeType === 3 && n.textContent.trim()) n.textContent = " " + txt + " "; });
  });
})();
/* ===================== Busca de condomínios (API Superlógica) ===================== */
(function(){
  const API_BASE = ["localhost", "127.0.0.1"].includes(location.hostname) ? "" : "https://jetro-superlogica-api-agencia-lumno.vercel.app";
  const input = document.getElementById("f_condominio");
  const hiddenId = document.getElementById("f_condominio_id");
  const dropdown = document.getElementById("condomDropdown");
  const statusEl = document.getElementById("condomStatus");

  let condomList = null;
  let loadingPromise = null;
  let activeIndex = -1;
  let currentItems = [];

  function normalize(s){
    return (s || "").toString().normalize("NFD").replace(/[\u0300-\u036f]/g, "").toLowerCase();
  }

  function escapeHtml(s){
    const d = document.createElement("div");
    d.textContent = s;
    return d.innerHTML;
  }

  function loadList(){
    if (condomList) return Promise.resolve(condomList);
    if (loadingPromise) return loadingPromise;
    statusEl.textContent = "Carregando carteira de condomínios...";
    statusEl.className = "ac-status";
    loadingPromise = fetch(API_BASE + "/api/superlogica/condominios/get?id=-1&apenasColunasPrincipais=1&somenteCondominiosAtivos=1&ignorarCondominioModelo=1")
      .then(function(res){ return res.json(); })
      .then(function(data){
        if (!Array.isArray(data)) throw new Error((data && data.error) || "resposta inesperada");
        condomList = data;
        statusEl.textContent = condomList.length + " condomínios carregados da carteira Jetro";
        statusEl.className = "ac-status ok";
        return condomList;
      })
      .catch(function(){
        statusEl.textContent = "Não foi possível carregar a carteira automaticamente — digite o nome manualmente.";
        statusEl.className = "ac-status err";
        condomList = [];
        return condomList;
      })
      .finally(function(){ loadingPromise = null; });
    return loadingPromise;
  }

  function renderDropdown(items){
    currentItems = items;
    activeIndex = -1;
    if (!items.length) {
      dropdown.innerHTML = '<div class="ac-empty">Nenhum condomínio encontrado na carteira — pode digitar o nome manualmente</div>';
      dropdown.classList.remove("hidden");
      return;
    }
    dropdown.innerHTML = items.map(function(c, i){
      const nome = c.st_fantasia_cond || c.st_nome_cond || "";
      const sub = [c.st_bairro_cond, c.st_cidade_cond].filter(Boolean).join(" — ");
      return '<div class="ac-item" data-idx="' + i + '">' +
        '<span class="ac-id">#' + c.id_condominio_cond + '</span> ' + escapeHtml(nome) +
        (sub ? '<div class="ac-sub">' + escapeHtml(sub) + '</div>' : '') +
        '</div>';
    }).join("");
    dropdown.classList.remove("hidden");
  }

  function closeDropdown(){
    dropdown.classList.add("hidden");
    dropdown.innerHTML = "";
  }

  function selectItem(c){
    input.value = c.st_fantasia_cond || c.st_nome_cond || "";
    hiddenId.value = c.id_condominio_cond || "";
    closeDropdown();
    if (typeof window.onCondominioSelected === "function") window.onCondominioSelected(c.id_condominio_cond);
  }

  input.addEventListener("focus", function(){ loadList(); });

  input.addEventListener("input", function(){
    hiddenId.value = "";
    const q = normalize(input.value.trim());
    if (!q) { closeDropdown(); return; }
    loadList().then(function(list){
      const matches = list.filter(function(c){
        return normalize(c.st_fantasia_cond).includes(q) || normalize(c.st_nome_cond).includes(q);
      }).slice(0, 15);
      renderDropdown(matches);
    });
  });

  dropdown.addEventListener("click", function(e){
    const item = e.target.closest(".ac-item");
    if (!item) return;
    const idx = parseInt(item.getAttribute("data-idx"), 10);
    if (currentItems[idx]) selectItem(currentItems[idx]);
  });

  input.addEventListener("keydown", function(e){
    if (dropdown.classList.contains("hidden") || !currentItems.length) return;
    const items = dropdown.querySelectorAll(".ac-item");
    if (e.key === "ArrowDown") {
      e.preventDefault();
      activeIndex = Math.min(activeIndex + 1, items.length - 1);
      items.forEach(function(it, i){ it.classList.toggle("active", i === activeIndex); });
    } else if (e.key === "ArrowUp") {
      e.preventDefault();
      activeIndex = Math.max(activeIndex - 1, 0);
      items.forEach(function(it, i){ it.classList.toggle("active", i === activeIndex); });
    } else if (e.key === "Enter") {
      if (activeIndex >= 0 && currentItems[activeIndex]) {
        e.preventDefault();
        selectItem(currentItems[activeIndex]);
      }
    } else if (e.key === "Escape") {
      closeDropdown();
    }
  });

  document.addEventListener("click", function(e){
    if (!e.target.closest(".autocomplete-wrap")) closeDropdown();
  });
})();

/* ===================== Gerador de Apresentações Jetro ===================== */
/* Identidade visual: #1d1339 (texto/dark), #452e86 (botões), #d08e02 (dourado),
   #f7f7f7 (fundo), cards brancos — mesma identidade do sistema de apuração de votos. */

/* ---------- Helpers de número em formato BR ---------- */
function parseBRNumber(val) {
  if (val === null || val === undefined) return 0;
  if (typeof val === "number") return val;
  if (typeof val !== "string") return 0;
  let s = val.trim();
  if (s === "" || /conf(orme|\.?)\s*(o\s*)?consumo/i.test(s)) return null; // valor não numérico
  s = s.replace(/[R$\s]/g, "");
  // remove separador de milhar '.' e troca decimal ',' por '.'
  const negative = /^\(.*\)$/.test(s);
  s = s.replace(/[()]/g, "");
  s = s.replace(/\./g, "").replace(",", ".");
  const n = parseFloat(s);
  if (isNaN(n)) return null;
  return negative ? -n : n;
}

// Converte "HH:MM" (campo type=time) para o formato falado "19h" ou "19h30"
function fmtHoraBR(hhmm) {
  const m = /^(\d{1,2}):(\d{2})/.exec(hhmm || "");
  if (!m) return ""; // vazio ou inválido: nunca devolve resíduo de formato
  const h = parseInt(m[1], 10);
  return m[2] === "00" ? h + "h" : h + "h" + m[2];
}

// Linhas do cabeçalho do edital (mesma regra de linhas_edital em api/generate_pptx.py):
// só entra o trecho que foi preenchido; sem nenhum dado a linha inteira é omitida.
function linhasEdital(dataExtenso, local, conv1, conv2) {
  const horaOk = (h) => /^\d{1,2}h(\d{2})?$/.test(h || "");
  const linha1 = [dataExtenso ? "Data: " + dataExtenso : "", local ? "Local: " + local : ""].filter(Boolean);
  const linha2 = [horaOk(conv1) ? "1ª convocação: " + conv1 : "", horaOk(conv2) ? "2ª convocação: " + conv2 : ""].filter(Boolean);
  return [linha1, linha2].filter((l) => l.length).map((l) => l.join("   •   "));
}

function fmtBRL(n, withSymbol) {
  if (n === null || n === undefined || isNaN(n)) return "—";
  const s = Math.abs(n).toLocaleString("pt-BR", { minimumFractionDigits: 2, maximumFractionDigits: 2 });
  return (n < 0 ? "-" : "") + (withSymbol ? "R$ " : "") + s;
}

/* ===================== PARSER: Relatório 11A ===================== */
/* Linhas = categorias (coluna A), colunas seguintes = meses.
   Mantém consciência de seção (Receitas / Despesas) pra resolver rótulos
   duplicados como "Agua e esgoto" e "Caixa Pequeno". */
function parse11A(workbook) {
  const sheetName = workbook.SheetNames[0];
  const sheet = workbook.Sheets[sheetName];
  const rows = XLSX.utils.sheet_to_json(sheet, { header: 1, raw: true, defval: null });

  if (!rows.length) throw new Error("Planilha vazia.");

  let headerRowIdx = -1;
  let months = [];
  const monthRe = /^[A-Za-zçÇ]{3}\/\d{4}$/;
  for (let i = 0; i < Math.min(rows.length, 12); i++) {
    const r = rows[i];
    const found = r.filter((c) => typeof c === "string" && monthRe.test(c.trim()));
    if (found.length >= 2) {
      headerRowIdx = i;
      months = r.map((c) => (typeof c === "string" ? c.trim() : null));
      break;
    }
  }
  if (headerRowIdx === -1) {
    throw new Error("Não encontrei a linha com os meses (ex: Dez/2025) no arquivo 11A.");
  }

  const monthCols = [];
  for (let c = 1; c < months.length; c++) {
    if (months[c] && monthRe.test(months[c])) monthCols.push(c);
  }
  const monthsClean = monthCols.map((c) => months[c]);
  const nMeses = monthsClean.length;

  const data = {};
  let section = null;
  const rootRows = {};
  const contasPlano = [];
  let grupoAtual = null;

  function normLabel(x) { return String(x || "").normalize("NFD").replace(/[\u0300-\u036f]/g, "").toLowerCase().trim(); }

  for (let i = headerRowIdx + 1; i < rows.length; i++) {
    const r = rows[i];
    const label = r[0];
    if (label === null || label === undefined || String(label).trim() === "") continue;
    const labelTrim = String(label).trim();

    if (/^receitas$/i.test(labelTrim)) { section = "receitas"; grupoAtual = null; continue; }
    if (/^despesas$/i.test(labelTrim)) { section = "despesas"; grupoAtual = null; continue; }
    if (/^mov\.?\s*l[íi]quido/i.test(labelTrim)) { section = null; grupoAtual = null; continue; }

    const vals = monthCols.map((c) => parseBRNumber(r[c]));
    const todosVazios = vals.every((v) => v === null || v === undefined);

    if (/^saldo (anterior|final)$/i.test(labelTrim)) {
      rootRows[/^saldo anterior$/i.test(labelTrim) ? "Saldo anterior" : "Saldo Final"] = vals;
      continue;
    }

    if (section === null) {
      rootRows[labelTrim] = vals;
      continue;
    }

    if (/^total de (receitas|despesas)$/i.test(labelTrim)) {
      data[section + ":" + (/receitas/i.test(labelTrim) ? "Total de Receitas" : "Total de Despesas")] = vals;
      grupoAtual = null;
      continue;
    }

    // Linha de subtotal de grupo (ex: "Total de Com Pessoal")
    if (/^total de /i.test(labelTrim)) {
      if (grupoAtual) {
        grupoAtual.valores = vals.map((v) => v || 0);
        grupoAtual = null;
      }
      continue;
    }

    // Cabeçalho de grupo sem valores na linha (ex: "Com Pessoal", "Mensais", "IMPOSTOS")
    if (todosVazios) {
      grupoAtual = {
        conta: (section === "receitas" ? "1." : "2.") + (contasPlano.length + 1),
        nome: labelTrim,
        nivel: 0,
        natureza: "0",
        secao: section,
        valores: new Array(nMeses).fill(0),
      };
      contasPlano.push(grupoAtual);
      continue;
    }

    const key = section + ":" + labelTrim;
    data[key] = vals;

    const valsNum = vals.map((v) => v || 0);
    if (grupoAtual) {
      for (let m = 0; m < nMeses; m++) grupoAtual.valores[m] += valsNum[m];
    }
    contasPlano.push({
      conta: (grupoAtual ? grupoAtual.conta + "." : (section === "receitas" ? "1." : "2.")) + (contasPlano.length + 1),
      nome: labelTrim,
      nivel: grupoAtual ? 1 : 0,
      natureza: section === "receitas" ? "1" : "2",
      secao: section,
      valores: valsNum,
    });
  }

  function getSeries(sectionKey, ...labels) {
    for (const label of labels) {
      const alvo = sectionKey + ":" + normLabel(label);
      const key = Object.keys(data).find((k) => {
        const idx = k.indexOf(":");
        return k.slice(0, idx) + ":" + normLabel(k.slice(idx + 1)) === alvo;
      });
      if (key) return data[key];
    }
    return null;
  }

  const categoriasDetalhadas = Object.keys(data)
    .filter((key) => !/:Total de (Receitas|Despesas)$/i.test(key))
    .map((key) => {
      const idx = key.indexOf(":");
      return { secao: key.slice(0, idx), label: key.slice(idx + 1), valores: data[key] };
    })
    .sort((a, b) => a.label.localeCompare(b.label, "pt-BR"));

  const totalReceitas = getSeries("receitas", "Total de Receitas");
  const totalDespesas = getSeries("despesas", "Total de Despesas");

  if (!totalReceitas || !totalDespesas) {
    throw new Error('Não encontrei as linhas "Total de Receitas" / "Total de Despesas" no arquivo.');
  }

  return {
    months: monthsClean,
    contasPlano: contasPlano.filter((c) => c.valores.some((v) => v)),
    saldoAnterior: rootRows["Saldo anterior"] || new Array(nMeses).fill(null),
    saldoFinal: rootRows["Saldo Final"] || new Array(nMeses).fill(null),
    totalReceitas,
    totalDespesas,
    energiaEletrica: getSeries("despesas", "Energia Elétrica"),
    aguaEsgotoDespesa: getSeries("despesas", "Agua e esgoto", "Agua/Esgoto", "Água"),
    categoriasDetalhadas,
  };
}

/* ===================== PARSER: Previsão Orçamentária ===================== */
// Lê a planilha sem depender de posição fixa: acha a linha de cabeçalho (Atual/Previsto) e as
// colunas por texto, grupos por numeração/itens, e os blocos de rateio por título + linha "Fração".
function prevNorm(v) {
  return String(v === null || v === undefined ? "" : v).trim();
}
function prevHeaderInfo(rows) {
  for (let i = 0; i < Math.min(rows.length, 15); i++) {
    const r = rows[i] || [];
    let atual = -1, previsto = -1, desc = -1;
    for (let c = 0; c < r.length; c++) {
      const t = prevNorm(r[c]);
      if (/^atual$/i.test(t) && atual < 0) atual = c;
      else if (/^previs[ãa]o$|^previsto$/i.test(t) && previsto < 0) previsto = c;
      else if (/^descri[çc][ãa]o$/i.test(t) && desc < 0) desc = c;
    }
    if (previsto >= 0) return { row: i, atual, previsto, desc: desc >= 0 ? desc : 1 };
  }
  return null;
}
// Modelo de 3 colunas: Previsto (última previsão aprovada), Realizado e Futuro (nova previsão).
//  - cabeçalhos "Previsto" + "Realizado" [+ "Futuro"]  => papéis literais (sem Futuro: Futuro = Realizado);
//  - cabeçalhos "Última previsão" + "Atual" + "Previsto" => previsto, realizado e futuro respectivamente.
// Qualquer outra combinação (ex.: só "Atual"/"Previsto") é o modelo antigo, tratado pelo parsePrevisao.
function cabecalho3Col(rows) {
  for (let i = 0; i < Math.min(rows.length, 15); i++) {
    const r = rows[i] || [];
    const achados = {};
    let desc = -1;
    for (let c = 0; c < r.length; c++) {
      const t = prevNorm(r[c]);
      if (/^descri[çc][ãa]o$/i.test(t) && desc < 0) desc = c;
      else if (/^[úu]ltima previs[ãa]o$/i.test(t) && achados.ultima === undefined) achados.ultima = c;
      else if (/^atual$/i.test(t) && achados.atual === undefined) achados.atual = c;
      else if (/^previsto$|^previs[ãa]o$/i.test(t) && achados.previsto === undefined) achados.previsto = c;
      else if (/^realizado$/i.test(t) && achados.realizado === undefined) achados.realizado = c;
      else if (/^futuro$/i.test(t) && achados.futuro === undefined) achados.futuro = c;
    }
    let map = null;
    if (achados.previsto !== undefined && achados.realizado !== undefined) map = { previsto: achados.previsto, realizado: achados.realizado, futuro: achados.futuro };
    else if (achados.ultima !== undefined && achados.atual !== undefined && achados.previsto !== undefined) map = { previsto: achados.ultima, realizado: achados.atual, futuro: achados.previsto };
    if (map) return { row: i, desc: desc >= 0 ? desc : 1, map };
  }
  return null;
}
function abaEh3Col(wb, nome) {
  const rows = XLSX.utils.sheet_to_json(wb.Sheets[nome], { header: 1, raw: true, defval: null });
  return !!cabecalho3Col(rows);
}
const fmtNumPO = (n) => n.toLocaleString("pt-BR", { minimumFractionDigits: 2, maximumFractionDigits: 4 });
// Lê uma aba de 3 colunas e PREENCHE a tabela manual: grupos, itens, sub-totais, totais e percentuais (6% da inadimplência,
// 5% do fundo de reserva...). Futuro vem da planilha quando preenchido; vazio/ausente => Futuro = Realizado (sugestão editável).
// Percentuais lidos sem fórmula (PDF): descobre se o valor de cada coluna é "% x total anterior" ou "por dentro" (total x %/(1-%));
// se nenhum bater, o valor fica fixo (como veio no arquivo).
function poInferirModos() {
  const calc = poCalcular();
  let ult = null;
  PO_MANUAL.linhas.forEach((l, i) => {
    if (l.tipo === "total") { ult = calc[i]; return; }
    if (l.tipo !== "percentual" || !l.modos || l.faixaSheet) return;
    const f = (poNumero(l.pct) || 0) / 100;
    l.modos = l.modos.map((m, c) => {
      if (m !== "fixo" || !l.sem_formula) return m;
      const v = poNumero(l.fixos && l.fixos[c]), base = ult ? ult[c] : null;
      if (v === null || base === null) return m;
      if (Math.abs(base * f - v) <= 0.05) return "sobre";
      if (f < 1 && Math.abs(base * f / (1 - f) - v) <= 0.05) return "dentro";
      return m;
    });
  });
}
function poImportar3Col(wb, nomeAba) {
  const ws = wb.Sheets[nomeAba];
  const rows = XLSX.utils.sheet_to_json(ws, { header: 1, raw: true, defval: null });
  const h = cabecalho3Col(rows);
  if (!h) throw new Error("Não encontrei as colunas Previsto/Realizado nessa aba.");
  const cols = [h.map.previsto, h.map.realizado, h.map.futuro === undefined ? -1 : h.map.futuro];
  const vazio = (v) => v === null || v === undefined || (typeof v === "string" && v.trim() === "");
  const txt = (v) => (vazio(v) ? "" : typeof v === "number" ? fmtNumPO(v) : String(v).trim());
  const formula = (r, c) => { const cel = c >= 0 ? ws[XLSX.utils.encode_cell({ r, c })] : null; return cel && cel.f ? String(cel.f) : ""; };
  const linhas = [];
  let grupoAberto = false, temFuturoPreenchido = false;
  let unidades = "";
  const hdr = rows[h.row] || [];
  const cUnid = hdr.findIndex((x) => /^unid/i.test(prevNorm(x)));
  for (let i = h.row + 1; i < rows.length; i++) {
    const r = rows[i] || [];
    const label = prevNorm(r[h.desc]);
    if (!label) continue;
    if (/^rateio\b/i.test(label) || /^previs[ãa]o\b.*(-|rateio)/i.test(label) || /^previs[ãa]o de rateio/i.test(label)) break;
    const vals = cols.map((c) => (c >= 0 ? r[c] : null));
    const semValor = vals.every(vazio);
    const colA = r[0];
    let calc = r[h.desc + 1];
    if (typeof calc !== "number") { // PDF/colunas deslocadas: procura o percentual entre a descrição e a 1ª coluna de valor
      const primeira = Math.min(...cols.filter((c) => c >= 0));
      for (let c = h.desc + 1; c < primeira; c++) if (typeof r[c] === "number" && r[c] > 0 && r[c] < 1) { calc = r[c]; break; }
    }
    const ehCodGrupo = (typeof colA === "number" && Number.isInteger(colA)) || /^\d+$/.test(prevNorm(colA));
    let prox = false;
    for (let j = i + 1; j < Math.min(i + 3, rows.length); j++) {
      if (!vazio((rows[j] || [])[h.desc])) {
        // a linha seguinte é item numerado (1.1) OU tem valores (ex.: "FUNDO DE RESERVA" sem código, seguido do 5%): esta é um cabeçalho de grupo
        prox = /^\d+\.\d+/.test(prevNorm((rows[j] || [])[0])) || (!/^sub-?total$|^total/i.test(prevNorm((rows[j] || [])[h.desc])) && cols.some((c) => c >= 0 && !vazio((rows[j] || [])[c])));
        break;
      }
    }
    if (semValor && label.length > 2 && !/^sub-?total$|^total/i.test(label) && (ehCodGrupo || (vazio(colA) && prox))) {
      linhas.push({ tipo: "grupo", nome: label, vals: [] }); grupoAberto = true;
      if (cUnid >= 0 && !unidades && typeof r[cUnid] === "number") unidades = String(r[cUnid]);
      continue;
    }
    if (/^sub-?total$/i.test(label)) { linhas.push({ tipo: grupoAberto ? "subtotal" : "total", nome: "Sub-Total", vals: [] }); grupoAberto = false; continue; }
    if (/^total\b/i.test(label)) { linhas.push({ tipo: "total", nome: label.replace(/\s+/g, " "), vals: [] }); grupoAberto = false; continue; }
    if (semValor) continue;
    // percentual: coluna "Descrição p/ Cálculo" numérica (0..1) e os valores calculados a partir dela (ex.: =F53*C56)
    const refC = new RegExp("\\$?C\\$?" + (i + 1) + "\\b");
    const semFormulas = cols.every((c) => !formula(i, c));
    // com fórmula: referência à célula C da linha; sem fórmulas (PDF/valores): coluna C numérica (0..1) = percentual, modo inferido depois
    const usaPct = typeof calc === "number" && calc > 0 && calc < 1 && (cols.some((c) => refC.test(formula(i, c))) || semFormulas);
    if (usaPct) {
      let faixa = null;
      const modos = [], fixos = [];
      cols.forEach((c) => {
        const f = formula(i, c);
        const mFaixa = f.match(/SUM\(\s*\$?[A-Z]+\$?(\d+)\s*:\s*\$?[A-Z]+\$?(\d+)\s*\)\s*\*/i);
        if (!f || !refC.test(f)) { modos.push("fixo"); fixos.push(typeof (c >= 0 ? r[c] : null) === "number" ? fmtNumPO(r[c]) : ""); }
        else if (mFaixa) { modos.push("sobre"); fixos.push(""); if (!faixa) faixa = [+mFaixa[1], +mFaixa[2]]; }
        else if (/\*\s*100\s*\/\s*\d/.test(f)) { modos.push("dentro"); fixos.push(""); }
        else { modos.push("sobre"); fixos.push(""); }
      });
      if (cols[2] < 0) { modos[2] = modos[1]; fixos[2] = fixos[1]; } // sem coluna Futuro na planilha: segue a regra do Realizado
      linhas.push({ tipo: "percentual", nome: label.replace(/\s+/g, " "), pct: String(Math.round(calc * 10000) / 100).replace(".", ","), modos, fixos, faixaSheet: faixa, vals: [], srow: i + 1, sem_formula: semFormulas });
      continue;
    }
    const v = [txt(vals[0]), txt(vals[1]), txt(vals[2])];
    const futVazio = vazio(vals[2]);
    if (!futVazio) temFuturoPreenchido = true;
    linhas.push({ tipo: "item", nome: label.replace(/\s+/g, " "), vals: v, futAuto: futVazio, srow: i + 1 });
  }
  PO_MANUAL.colunas = ["Previsto", "Realizado", "Futuro"];
  PO_MANUAL.colMeta = [{ papel: "previsto" }, { papel: "realizado" }, { papel: "futuro" }];
  PO_MANUAL.linhas = linhas;
  PO_MANUAL.usado = true; PO_MANUAL.importado = true;
  poGarantirIds();
  poInferirModos();
  // percentual sobre um intervalo de itens (ex.: SUM(F30:F32)*C33): traduz as linhas da planilha para os ids das linhas importadas
  linhas.forEach((l) => {
    if (!l.faixaSheet) return;
    l.baseIds = linhas.filter((x) => x.tipo === "item" && x.srow >= l.faixaSheet[0] && x.srow <= l.faixaSheet[1]).map((x) => x.id);
  });
  // Rateio vinculado: Cota = último total antes do fundo de reserva; Fundo = sub-total do grupo do fundo; ÷ unidades
  const iFundo = linhas.findIndex((l) => (l.tipo === "grupo" || l.tipo === "percentual") && /fundo/i.test(l.nome));
  let iCota = -1, iFundoSub = -1;
  if (iFundo >= 0) {
    for (let k = iFundo - 1; k >= 0; k--) if (linhas[k].tipo === "total") { iCota = k; break; }
    iFundoSub = linhas.findIndex((l, k) => k >= iFundo && l.tipo === "subtotal");
    if (iFundoSub < 0) iFundoSub = linhas.findIndex((l, k) => k >= iFundo && l.tipo === "total");
  }
  const blocoR = (titulo, papel) => {
    const b = poRateioModelo(titulo);
    b.unidades = unidades;
    const iTotalDesp = linhas.map((l) => l.tipo).lastIndexOf("total");
    if (iTotalDesp >= 0) {
      b.itens = [{ nome: "Recibo completo (Cota + Fundo + taxas)", recibo: "", origem: { linhaId: linhas[iTotalDesp].id, papel } }, { nome: "Água / Esgoto", recibo: "Conforme consumo" }];
    } else if (iCota >= 0 && iFundoSub >= 0) {
      b.itens = [{ nome: "Cota Condominial", recibo: "", origem: { linhaId: linhas[iCota].id, papel } }, { nome: "Fundo de Reserva", recibo: "", origem: { linhaId: linhas[iFundoSub].id, papel } }, { nome: "Água / Esgoto", recibo: "Conforme consumo" }];
    }
    return b;
  };
  PO_MANUAL.rateio = { atual: blocoR("Rateio atual", "realizado"), previsto: blocoR("Previsão (novo rateio)", "futuro") };
  return { itens: linhas.filter((l) => l.tipo === "item").length, temFuturoPreenchido, rateioVinculado: iCota >= 0 && iFundoSub >= 0, unidades };
}

// Abas da planilha que têm tabela de previsão (cabeçalho Atual/Previsto), com o condomínio do título (célula A1).
function listarAbasPrevisao(workbook) {
  return workbook.SheetNames.map((n) => {
    const rows = XLSX.utils.sheet_to_json(workbook.Sheets[n], { header: 1, raw: true, defval: null });
    const h = prevHeaderInfo(rows) || cabecalho3Col(rows);
    const a1 = rows[0] && rows[0][0] ? String(rows[0][0]) : "";
    const m = a1.split(/[-–]/);
    const cond = /previs[ãa]o or[çc]ament/i.test(a1) && m.length > 1 ? m.slice(1).join("-").trim() : "";
    return { nome: n, valida: !!h, condominio: cond, tres: !!cabecalho3Col(rows) };
  });
}
// Aba usada: 1) a aba ATIVA do arquivo (a que estava selecionada ao salvar), se tiver previsão;
// 2) senão a regra por nome/colunas (período no nome, Atual+Previsto, não ser "ATUAL", a mais à direita).
function pickPrevisaoSheet(workbook) {
  const names = workbook.SheetNames;
  if (names.length === 1) return names[0];
  const abas = listarAbasPrevisao(workbook);
  const wbv = workbook.Workbook && workbook.Workbook.WBView && workbook.Workbook.WBView[0]; // SheetJS guarda a aba ativa aqui
  const idx = wbv ? Number(wbv.activeTab) : undefined;
  if (typeof idx === "number" && !isNaN(idx) && abas[idx] && abas[idx].valida) return abas[idx].nome;
  const info = names.map((n, i) => {
    const rows = XLSX.utils.sheet_to_json(workbook.Sheets[n], { header: 1, raw: true, defval: null });
    const h = prevHeaderInfo(rows);
    return { n, idx: i, h, completa: !!(h && h.atual >= 0), periodo: /\d{4}\s*[-\/]\s*\d{4}/.test(n), atualAba: n.trim().toUpperCase() === "ATUAL" };
  });
  const validas = info.filter((x) => x.h);
  const baseL = validas.length ? validas : info;
  const score = (x) => (x.completa ? 4 : 0) + (x.periodo ? 2 : 0) + (x.atualAba ? 0 : 1);
  let best = baseL[0];
  baseL.forEach((x) => { if (score(x) >= score(best)) best = x; });
  return best.n;
}

function parsePrevisao(workbook, abaEscolhida) {
  const sheetName = abaEscolhida && workbook.SheetNames.indexOf(abaEscolhida) >= 0 ? abaEscolhida : pickPrevisaoSheet(workbook);
  const sheet = workbook.Sheets[sheetName];
  const rows = XLSX.utils.sheet_to_json(sheet, { header: 1, raw: true, defval: null });
  const hdr = prevHeaderInfo(rows) || { row: 1, atual: 4, previsto: 5, desc: 1 };
  const cDesc = hdr.desc, cAtual = hdr.atual, cPrev = hdr.previsto;
  const cell = (r, c) => (c >= 0 && r && r[c] !== undefined ? r[c] : null);
  const vazio = (v) => v === null || v === undefined || (typeof v === "string" && v.trim() === "");

  const categories = [];
  const itensCompletos = [];
  let totalDespesasAtual = null, totalDespesasPrevisto = null;
  let condominioNome = null;

  for (let i = 0; i < Math.min(rows.length, 3); i++) {
    for (const c0 of rows[i] || []) {
      if (typeof c0 === "string" && /previs[ãa]o or[çc]ament/i.test(c0)) {
        const m = c0.split(/[-–]/);
        if (m.length > 1) condominioNome = m.slice(1).join("-").trim();
      }
    }
  }

  const isRateioTitulo = (t) => /^rateio\b/i.test(t) || /^previs[ãa]o\b.*(-|rateio)/i.test(t);
  const codigoGrupo = (v) => (typeof v === "number" && Number.isInteger(v)) || /^\d+$/.test(prevNorm(v));
  const codigoItem = (v) => /^\d+\.\d+/.test(prevNorm(v));

  let currentGroupLabel = null;
  let fimItens = false;

  for (let i = hdr.row + 1; i < rows.length && !fimItens; i++) {
    const r = rows[i] || [];
    const labelRaw = cell(r, cDesc);
    if (vazio(labelRaw)) continue;
    const label = prevNorm(labelRaw);
    const colA = r[0];
    const colE = cell(r, cAtual), colF = cell(r, cPrev);
    const semValor = vazio(colE) && vazio(colF);

    if (isRateioTitulo(label)) break;

    // próxima linha preenchida: se for item numerado (1.1, 2.3...), esta linha é um cabeçalho de grupo
    let proxCodigo = false;
    for (let j = i + 1; j < Math.min(i + 3, rows.length); j++) {
      if (!vazio(cell(rows[j], cDesc))) { proxCodigo = codigoItem((rows[j] || [])[0]); break; }
    }
    const isGroupHeader = semValor && label.length > 2 && !/^sub-?total$|^total/i.test(label) &&
      (codigoGrupo(colA) || (vazio(colA) && proxCodigo));
    if (isGroupHeader) { currentGroupLabel = label; itensCompletos.push({ tipo: "grupo", nome: label }); continue; }

    function sufixoCalculo(rowArr, colD) {
      const calc = rowArr[colD + 1];
      if (typeof calc === "number" && calc > 0 && calc < 1) {
        return " (" + String(Math.round(calc * 10000) / 100).replace(".", ",") + "%)";
      }
      if (typeof calc === "string" && calc.trim()) {
        const txt = calc.trim().replace(/\s+/g, " ");
        return /^\(.*\)$/.test(txt) ? " " + txt : " (" + txt + ")";
      }
      return "";
    }

    if (/^sub-?total$/i.test(label) && currentGroupLabel) {
      categories.push({ nome: currentGroupLabel, atual: parseBRNumber(colE), previsto: parseBRNumber(colF) });
      itensCompletos.push({ tipo: "subtotal", nome: "Sub-Total", atualRaw: colE, previstoRaw: colF });
      currentGroupLabel = null;
      continue;
    }

    // Linhas de fechamento fecham qualquer grupo aberto antes (evita duplicar Fundo de Reserva e Total despesas)
    if (/^total de pagamentos$/i.test(label)) {
      currentGroupLabel = null;
      categories.push({ nome: "Total de pagamentos", atual: parseBRNumber(colE), previsto: parseBRNumber(colF) });
      itensCompletos.push({ tipo: "subtotal", nome: "Total de pagamentos", atualRaw: colE, previstoRaw: colF });
    } else if (/^(garantidora|inadimpl[êe]ncia)/i.test(label)) {
      currentGroupLabel = null;
      categories.push({ nome: /^garantidora/i.test(label) ? "Garantidora" : "Inadimplência", atual: parseBRNumber(colE), previsto: parseBRNumber(colF) });
      itensCompletos.push({ tipo: "item", nome: label + sufixoCalculo(r, cDesc), atualRaw: colE, previstoRaw: colF });
    } else if (/^fundo de reserva/i.test(label) && !semValor) {
      currentGroupLabel = null;
      categories.push({ nome: "Fundo de Reserva", atual: parseBRNumber(colE), previsto: parseBRNumber(colF) });
      itensCompletos.push({ tipo: "item", nome: label + sufixoCalculo(r, cDesc), atualRaw: colE, previstoRaw: colF });
    } else if (/^sub-?total$/i.test(label) && !currentGroupLabel) {
      itensCompletos.push({ tipo: "subtotal", nome: "Sub-Total", atualRaw: colE, previstoRaw: colF });
    } else if (/^total despesas$/i.test(label)) {
      currentGroupLabel = null;
      totalDespesasAtual = parseBRNumber(colE);
      totalDespesasPrevisto = parseBRNumber(colF);
      itensCompletos.push({ tipo: "subtotal", nome: "Total despesas", atualRaw: colE, previstoRaw: colF });
    } else if (currentGroupLabel && !semValor) {
      itensCompletos.push({ tipo: "item", nome: label + sufixoCalculo(r, cDesc), atualRaw: colE, previstoRaw: colF });
    }
  }

  // ---- Rateio: blocos = título + linha "Fração" + linhas + "Total" (aceita blocos empilhados ou lado a lado) ----
  function lerBloco(startIdx, colTitulo) {
    const colBase = colTitulo !== undefined ? colTitulo : cDesc;
    let headerIdx = -1;
    let cFrac = colBase;
    for (let i = startIdx + 1; i < Math.min(startIdx + 6, rows.length); i++) {
      const r = rows[i] || [];
      for (let c = Math.max(0, colBase - 1); c < Math.min(r.length, colBase + 6); c++) {
        if (/^fra[çc][ãa]o$/i.test(prevNorm(r[c]))) { headerIdx = i; cFrac = c; break; }
      }
      if (headerIdx !== -1) break;
    }
    if (headerIdx === -1) return null;
    const heads = (rows[headerIdx] || []).map(prevNorm);
    let cTotal = -1, cRecibo = -1, cVTotal = -1, cUnid = -1;
    const colsExtrasRecibo = [];
    for (let c = cFrac + 1; c < heads.length; c++) {
      const t = heads[c];
      if (!t && c > cFrac + 1 && colsExtrasRecibo.length + (cTotal >= 0 ? 1 : 0) + (cRecibo >= 0 ? 1 : 0) > 0) break;
      if (/^fra[çc][ãa]o$/i.test(t)) break;
      if (/total.*rateio|rateio.*total/i.test(t)) cTotal = c;
      else if (/^(v\.?\s*total|total\s*unit|valor\s*total|recibo\s*total|valor\s*final)/i.test(t) && cVTotal < 0) cVTotal = c;
      else if (/^(recibo|cota)/i.test(t) && cRecibo < 0) cRecibo = c;
      else if (/^(unid|qtd|quant)/i.test(t) && cUnid < 0) cUnid = c;
      else if (/fundo|garant|taxa|extra/i.test(t)) colsExtrasRecibo.push(c);
    }
    if (cTotal < 0) {
      for (let c = cFrac + 1; c < heads.length; c++) {
        if (/^fra[çc][ãa]o$/i.test(heads[c])) break;
        if (heads[c]) cTotal = c;
      }
    }
    const colUnitPrincipal = cVTotal >= 0 ? cVTotal : (cRecibo >= 0 ? cRecibo : cFrac + 1);
    const somarExtras = cVTotal < 0 && colsExtrasRecibo.length > 0;

    const itens = [];
    let total = null, reajuste = null, reajusteLabel = null, reajusteIsPercent = false, somaUnidades = 0;
    let fim = headerIdx;
    for (let i = headerIdx + 1; i < Math.min(headerIdx + 40, rows.length); i++) {
      const nome = cell(rows[i], cFrac);
      if (vazio(nome)) break;
      const nomeTrim = prevNorm(nome);
      fim = i;
      if (/^total$/i.test(nomeTrim)) { total = parseBRNumber(cell(rows[i], cTotal)); break; }
      let valRecibo = parseBRNumber(cell(rows[i], colUnitPrincipal));
      if (somarExtras && typeof valRecibo === "number") {
        colsExtrasRecibo.forEach((ce) => {
          const ve = parseBRNumber(cell(rows[i], ce));
          if (typeof ve === "number") valRecibo += ve;
        });
      }
      const valTotal = parseBRNumber(cell(rows[i], cTotal));
      if (cUnid >= 0) {
        const u = parseBRNumber(cell(rows[i], cUnid));
        if (typeof u === "number" && u > 0) {
          somaUnidades += u;
          if (cVTotal < 0 && !somarExtras && typeof valTotal === "number" && valTotal > 0) {
            valRecibo = Math.round((valTotal / u) * 100) / 100;
          }
        }
      }
      itens.push({ nome: nomeTrim, recibo: valRecibo, total: valTotal });
    }
    for (let i = startIdx; i <= Math.min(fim + 1, rows.length - 1) && reajuste === null; i++) {
      const r = rows[i] || [];
      for (let c = 0; c < r.length; c++) {
        if (c === cFrac || typeof r[c] !== "string" || !/^\s*(reajuste|diferen[çc]a)/i.test(r[c])) continue;
        for (let j = i; j <= Math.min(fim + 2, rows.length - 1); j++) {
          const v = (rows[j] || [])[c];
          if (typeof v === "number") { reajuste = v; reajusteLabel = r[c].trim(); reajusteIsPercent = /reajuste/i.test(r[c]); break; }
        }
        if (reajuste !== null) break;
      }
    }
    return { itens, total, reajuste, reajusteLabel, reajusteIsPercent, unidades: somaUnidades > 0 ? Math.round(somaUnidades) : null };
  }

  const ehTituloBlocoRateio = (t) => /^rateio\b|^novo\s+rateio\b|^simula[çc][ãa]o.*rateio|^previs[ãa]o\b.*(-|rateio|cota)/i.test(t);
  const blocos = [];
  for (let i = hdr.row + 1; i < rows.length; i++) {
    const r = rows[i] || [];
    for (let c = 0; c < r.length; c++) {
      const t = prevNorm(r[c]);
      if (t && ehTituloBlocoRateio(t)) {
        const b = lerBloco(i, c);
        if (b && b.itens.length) blocos.push({ titulo: t, idx: i, col: c, bloco: b });
      }
    }
  }
  let rateioPrevisto = null;
  for (let k = blocos.length - 1; k >= 0; k--) {
    if (/^previs[ãa]o|previsto|novo\s+rateio|simula/i.test(blocos[k].titulo)) { rateioPrevisto = blocos[k].bloco; break; }
  }
  const anteriores = blocos.filter((b) => !/^previs[ãa]o|previsto|novo\s+rateio|simula/i.test(b.titulo));
  const comAtual = anteriores.filter((b) => /atual/i.test(b.titulo));
  const escolhido = comAtual.length ? comAtual[comAtual.length - 1] : anteriores[anteriores.length - 1];
  const rateioAtual = escolhido ? escolhido.bloco : null;

  // Se a planilha trouxer só a tabela do Rateio atual por fração, projeta o Rateio previsto pelo fator de reajuste do orçamento
  if (rateioAtual && rateioAtual.itens && rateioAtual.itens.length && (!rateioPrevisto || !rateioPrevisto.itens || !rateioPrevisto.itens.length)) {
    if (typeof totalDespesasAtual === "number" && totalDespesasAtual > 0 && typeof totalDespesasPrevisto === "number" && totalDespesasPrevisto > 0) {
      const fator = totalDespesasPrevisto / totalDespesasAtual;
      rateioPrevisto = {
        itens: rateioAtual.itens.map((it) => ({
          nome: it.nome,
          recibo: typeof it.recibo === "number" ? Math.round(it.recibo * fator * 100) / 100 : it.recibo,
          total: typeof it.total === "number" ? Math.round(it.total * fator * 100) / 100 : it.total,
        })),
        total: typeof rateioAtual.total === "number" ? Math.round(rateioAtual.total * fator * 100) / 100 : totalDespesasPrevisto,
        reajuste: fator - 1,
        reajusteLabel: "Reajuste de",
        reajusteIsPercent: true,
        unidades: rateioAtual.unidades,
        projetado: true,
      };
    }
  }

  if (!categories.length) {
    throw new Error("Não encontrei as categorias (Sub-Total) na planilha de previsão orçamentária.");
  }

  return {
    sheetUsed: sheetName,
    condominioNome,
    categories,
    itensCompletos,
    totalDespesasAtual,
    totalDespesasPrevisto,
    rateioAtual,
    rateioPrevisto,
  };
}

// Formata um valor de Atual/Previsto da previsão orçamentária: preserva textos como "Conforme consumo" e "ISENTO COTA".
function formatPOValue(raw) {
  if (raw === null || raw === undefined) return "—";
  if (typeof raw === "number") return fmtBRL(raw);
  const s = String(raw).trim();
  if (!s || /^[-—]+$/.test(s)) return "—";
  if (/conf(orme|\.?)\s*(o\s*)?consumo/i.test(s)) return "Conforme consumo";
  const n = parseBRNumber(s);
  if (n !== null && !isNaN(n)) return fmtBRL(n);
  return s;
}

/* ===================== PARSER: 16B (texto colado) ===================== */
function parse16B(raw) {
  if (!raw || !raw.trim()) return null;
  const lines = raw.split("\n").map((l) => l.trim()).filter(Boolean);
  const contas = [];
  const grupos = [];
  let mode = null; // 'contas' | 'grupos'
  let saldoFinalGeral = null;
  let periodo = ""; // "De 01/11/2025 até 31/10/2026", como vem no cabeçalho do relatório

  const numRe = /-?\(?[\d.]+,\d{2}\)?/g;

  for (const line of lines) {
    const mPer = /^de\s+(\d{1,2}\/\d{1,2}\/\d{4})\s+at\S?\s*(\d{1,2}\/\d{1,2}\/\d{4})/i.exec(line);
    if (mPer && !periodo) { periodo = "De " + mPer[1] + " até " + mPer[2]; continue; }
    if (/^conta\b/i.test(line) && /saldo/i.test(line)) { mode = "contas"; continue; }
    if (/^grupos de saldo/i.test(line)) { mode = "grupos"; continue; }
    if (/^\(\*\)/.test(line)) continue;

    const nums = line.match(numRe);
    // linha de fechamento com um valor só: "Saldo final 125.510,60"
    if (nums && nums.length === 1 && /^saldo final\b/i.test(line)) { saldoFinalGeral = parseBRNumber(nums[0]); continue; }
    if (!nums || nums.length < 2) continue;

    // nome = tudo antes do primeiro número
    const firstNumIdx = line.search(numRe);
    let nome = line.slice(0, firstNumIdx).trim();
    if (!nome) continue;

    const vals = nums.map(parseBRNumber);
    const entry = { nome, saldoAnt: vals[0], creditos: vals[1], debitos: vals[2], saldoFinal: vals[vals.length - 1] };

    if (/^saldo final$/i.test(nome)) {
      saldoFinalGeral = entry.saldoFinal;
      continue;
    }

    if (mode === "grupos") grupos.push(entry);
    else contas.push(entry);
  }

  if (!contas.length && !grupos.length) {
    throw new Error("Não consegui reconhecer nenhuma linha do 16B no texto colado. Confira se copiou a tabela inteira.");
  }

  return { contas, grupos, saldoFinalGeral, periodo };
}

/* ===================== UI: upload handlers, pauta, previews ===================== */
let PARSED = { a11: null, prev: null, b16: null };

function setupDropzone(dzId, inputId, onFile, onClear) {
  const dz = document.getElementById(dzId);
  const input = document.getElementById(inputId);
  // Mostra o nome do arquivo sem destruir o <input> (assim dá pra escolher outro arquivo depois) e
  // oferece "Remover arquivo" logo abaixo: limpa o upload e volta o dropzone ao estado inicial.
  const textoInicial = dz.textContent.trim();
  dz._mostrar = (nome) => {
    dz.innerHTML = "";
    dz.appendChild(input);
    const sp = document.createElement("span");
    sp.className = "fname";
    sp.textContent = nome;
    dz.appendChild(sp);
  };
  const remover = document.createElement("button");
  remover.type = "button";
  remover.className = "dz-remove";
  remover.textContent = "\u2715 Remover arquivo enviado";
  remover.hidden = true;
  dz.insertAdjacentElement("afterend", remover);
  new MutationObserver(() => { remover.hidden = !dz.classList.contains("ok"); }).observe(dz, { attributes: true, attributeFilter: ["class"] });
  remover.addEventListener("click", () => {
    input.value = "";
    dz.innerHTML = "";
    dz.appendChild(input);
    dz.appendChild(document.createTextNode(textoInicial));
    dz.classList.remove("ok");
    if (onClear) onClear();
    if (typeof refreshPreview === "function") refreshPreview();
  });
  dz.addEventListener("click", () => input.click());
  dz.addEventListener("dragover", (e) => { e.preventDefault(); dz.classList.add("drag"); });
  dz.addEventListener("dragleave", () => dz.classList.remove("drag"));
    dz.addEventListener("drop", (e) => {
    e.preventDefault(); dz.classList.remove("drag");
    if (e.dataTransfer.files[0]) { input.files = e.dataTransfer.files; onFile(e.dataTransfer.files[0]); }
  });
  input.addEventListener("change", () => { if (input.files[0]) onFile(input.files[0]); });
}

// Alguns geradores gravam a dimensão da aba errada (ex.: "A1:A17" com dados até a coluna E) e o SheetJS ignora o resto.
// Recalcula o intervalo a partir das células que realmente existem.
function corrigirRefPlanilhas(wb) {
  (wb.SheetNames || []).forEach((n) => {
    const ws = wb.Sheets[n];
    if (!ws) return;
    let r0 = Infinity, c0 = Infinity, r1 = -1, c1 = -1;
    Object.keys(ws).forEach((k) => {
      if (k[0] === "!") return;
      const { r, c } = XLSX.utils.decode_cell(k);
      if (r < r0) r0 = r; if (c < c0) c0 = c; if (r > r1) r1 = r; if (c > c1) c1 = c;
    });
    if (r1 < 0) return;
    const atual = ws["!ref"] ? XLSX.utils.decode_range(ws["!ref"]) : null;
    if (!atual || atual.e.r < r1 || atual.e.c < c1 || atual.s.r > r0 || atual.s.c > c0) {
      ws["!ref"] = XLSX.utils.encode_range({ s: { r: Math.min(r0, atual ? atual.s.r : r0), c: Math.min(c0, atual ? atual.s.c : c0) }, e: { r: Math.max(r1, atual ? atual.e.r : r1), c: Math.max(c1, atual ? atual.e.c : c1) } });
    }
  });
}

function readWorkbook(file) {
  return new Promise((resolve, reject) => {
    const reader = new FileReader();
    reader.onload = (e) => {
      try {
        const wb = XLSX.read(e.target.result, { type: "array" });
        corrigirRefPlanilhas(wb);
        resolve(wb);
      } catch (err) { reject(err); }
    };
    reader.onerror = () => reject(new Error("Falha ao ler o arquivo."));
    reader.readAsArrayBuffer(file);
  });
}

function setStatus(elId, ok, msg) {
  const el = document.getElementById(elId);
  el.className = "status-line " + (ok ? "ok" : "err");
  el.innerHTML = '<span class="dot"></span>' + msg;
}

/* ---------- Aviso de arquivo de outro condomínio (não bloqueia o upload) ---------- */
// Lê o nome do condomínio no cabeçalho do documento: "W011A CONDOMINIO DO EDIFICIO X (144)" (11A/16B) ou "Previsão Orçamentária - X".
// Sem nome identificável devolve null (nenhum alerta).
function nomeCondominioDoDocumento(linhas) {
  for (const l0 of (linhas || []).slice(0, 8)) {
    const l = String(l0 || "").replace(/\s+/g, " ").trim();
    if (!l) continue;
    let m = /previs[ãa]o\s+or[çc]ament\S*\s*[-–]\s*(.+)$/i.exec(l);
    if (m && m[1].trim().length >= 3) return m[1].trim();
    m = /^[A-Za-z0-9]{3,6}\s+(.{4,}?)\s*\(\d+\)\s*$/.exec(l);
    if (m) return m[1].trim();
  }
  return null;
}
function linhasIniciaisDaAba(wb, aba) {
  const ws = wb && wb.Sheets && wb.Sheets[aba || wb.SheetNames[0]];
  if (!ws) return [];
  const rows = XLSX.utils.sheet_to_json(ws, { header: 1, raw: true, defval: null }).slice(0, 8);
  const out = [];
  rows.forEach((r) => (r || []).forEach((c) => { if (typeof c === "string" && c.trim()) out.push(c); }));
  return out;
}
const CONDO_PALAVRAS_GENERICAS = new Set(["condominio", "cond", "edificio", "ed", "edif", "residencial", "res", "conjunto", "conj", "do", "da", "de", "dos", "das", "e"]);
function tokensCondominio(nome) {
  return String(nome || "").normalize("NFD").replace(/[\u0300-\u036f]/g, "").toLowerCase().split(/[^a-z0-9]+/).filter((t) => t && !CONDO_PALAVRAS_GENERICAS.has(t));
}
// Mesmo condomínio quando pelo menos metade das palavras do nome menor aparece no outro (tolera abreviações como "Nucleo SP").
function mesmoCondominio(a, b) {
  const ta = tokensCondominio(a), tb = tokensCondominio(b);
  if (!ta.length || !tb.length) return true; // sem como comparar: não alerta
  const [menor, maior] = ta.length <= tb.length ? [ta, tb] : [tb, ta];
  return menor.filter((t) => maior.indexOf(t) >= 0).length / menor.length >= 0.5;
}
function avisarCondominio(statusId, nomeDoc) {
  const st = document.getElementById(statusId);
  if (!st) return;
  let av = document.getElementById("aviso_" + statusId);
  if (!av) { av = document.createElement("div"); av.id = "aviso_" + statusId; av.className = "aviso-cond"; av.hidden = true; st.insertAdjacentElement("afterend", av); }
  const sel = (document.getElementById("f_condominio").value || "").trim();
  if (!nomeDoc || !sel || mesmoCondominio(nomeDoc, sel)) { av.hidden = true; av.textContent = ""; return; }
  av.textContent = "⚠ Atenção: este arquivo parece ser do condomínio '" + nomeDoc + "', mas o condomínio selecionado é '" + sel + "'. Confira antes de continuar.";
  av.hidden = false;
}

/* ---------- Modal de erro (pra leigos entenderem o que houve e o que fazer) ---------- */
function showErrorModal(title, msg, suggestion) {
  document.getElementById("modalTitle").textContent = "⚠ " + title;
  document.getElementById("modalMsg").textContent = msg;
  document.getElementById("modalSuggestion").textContent = suggestion || "";
  document.getElementById("modalSuggestion").style.display = suggestion ? "block" : "none";
  document.getElementById("errorModalOverlay").classList.remove("hidden");
}
document.getElementById("modalCloseBtn").addEventListener("click", () => {
  document.getElementById("errorModalOverlay").classList.add("hidden");
});
document.getElementById("errorModalOverlay").addEventListener("click", (e) => {
  if (e.target.id === "errorModalOverlay") document.getElementById("errorModalOverlay").classList.add("hidden");
});

// Reporta um erro no texto discreto (status-line) E num modal com explicação simples + sugestão.
function reportError(elId, title, msg, suggestion) {
  setStatus(elId, false, msg);
  showErrorModal(title, msg, suggestion);
}

/* ===================== Busca automática do 011A (API Superlógica) ===================== */
const BAL_API_BASE = ["localhost", "127.0.0.1"].includes(location.hostname) ? "" : "https://jetro-superlogica-api-agencia-lumno.vercel.app";
const MESES_ABBR = ["Jan", "Fev", "Mar", "Abr", "Mai", "Jun", "Jul", "Ago", "Set", "Out", "Nov", "Dez"];

function pad2(n) { return String(n).padStart(2, "0"); }
function lastDayOfMonth(mes, ano) { return new Date(ano, mes, 0).getDate(); }

async function fetchBalanceteMes(idCondominio, mes, ano) {
  const dtInicio = pad2(mes) + "/01/" + ano;
  const dtFim = pad2(mes) + "/" + pad2(lastDayOfMonth(mes, ano)) + "/" + ano;
  const url = BAL_API_BASE + "/api/superlogica/balancetes/index?idCondominio=" + idCondominio + "&dtInicio=" + dtInicio + "&dtFim=" + dtFim + "&agrupadoPorMes=1";
  const res = await fetch(url);
  if (!res.ok) throw new Error("HTTP " + res.status + " ao buscar " + dtInicio + " - " + dtFim);
  return res.json();
}

function parseBalanceteMes(apiResponse) {
  if (!Array.isArray(apiResponse) || !apiResponse[0] || !Array.isArray(apiResponse[0].itens)) return null;
  const itens = apiResponse[0].itens.map((arr) => arr[0]).filter(Boolean);
  function num(v) { const n = parseFloat(v); return isNaN(n) ? 0 : n; }
  function semAcento(s) { return (s || "").toString().normalize("NFD").replace(/[\u0300-\u036f]/g, "").toLowerCase(); }
  let totalReceitas = null, totalDespesas = null, somaRecFolha = 0, somaDespFolha = 0;
  let energia = 0, agua = 0, foundEnergia = false, foundAgua = false;

  itens.forEach((it) => {
    const descOriginal = it.descricaocomcontacategoria || it.descricao || "";
    const desc = semAcento(descOriginal).replace(/^\d[\d.]*\s*/, "").trim();
    const nat = String(it.natureza || "");
    const conta = String(it.conta || "").trim();

    if (nat === "0" && !conta.includes(".")) {
      if (/^receitas?$/i.test(desc) || conta === "1") totalReceitas = num(it.valor);
      if (/^despesas?$/i.test(desc) || conta === "2") totalDespesas = num(it.valor);
    }
    if (nat === "1") somaRecFolha += num(it.valor);
    if (nat === "2") {
      somaDespFolha += num(it.valor);
      if (/energia/.test(desc)) { energia += num(it.valor); foundEnergia = true; }
      if (/\bagua\b/.test(desc) && !/caixa|limpeza/.test(desc)) { agua += num(it.valor); foundAgua = true; }
    }
  });

  if (totalReceitas === null && somaRecFolha > 0) totalReceitas = somaRecFolha;
  if (totalDespesas === null && somaDespFolha > 0) totalDespesas = somaDespFolha;

  const categoriasFolha = itens
    .filter((it) => it.natureza === "1" || it.natureza === "2")
    .map((it) => ({
      secao: it.natureza === "1" ? "receitas" : "despesas",
      label: (it.descricaocomcontacategoria || it.descricao || "").replace(/^\d[\d.]*\s*/, "").trim(),
      valor: num(it.valor),
    }));
  return {
    totalReceitas,
    totalDespesas,
    energia: foundEnergia ? energia : null,
    agua: foundAgua ? agua : null,
    categoriasFolha,
    itensBrutos: itens.map((it) => ({
      conta: String(it.conta || "").trim(),
      natureza: String(it.natureza || ""),
      nome: (it.descricao || "").replace(/^\d[\d.]*\s*/, "").trim(),
      valor: num(it.valor),
    })),
  };
}

/* ===== Plano de contas real do condomínio (agnóstico a código de conta) ===== */
const PLANO_CONTAS_CACHE = {};

async function fetchPlanoContas(idCondominio) {
  if (PLANO_CONTAS_CACHE[idCondominio]) return PLANO_CONTAS_CACHE[idCondominio];
  const res = await fetch(BAL_API_BASE + "/api/superlogica/planocontas/index?idCondominio=" + idCondominio);
  if (!res.ok) throw new Error("HTTP " + res.status + " ao buscar o plano de contas");
  const data = await res.json();
  if (!Array.isArray(data)) throw new Error((data && data.msg) || "resposta inesperada ao buscar o plano de contas");
  PLANO_CONTAS_CACHE[idCondominio] = data;
  return data;
}

function montarPlanoContas(planoApi, mesesItens) {
  const nomesPlano = {};
  (planoApi || []).forEach((p) => {
    const cod = String(p.st_conta_cont || "").trim();
    if (cod && !nomesPlano[cod]) nomesPlano[cod] = (p.st_descricao_cont || "").trim();
  });
  const n = mesesItens.length;
  const mapa = {};
  const secaoPorPrefixo = {};

  // Descobre qual prefixo raiz pertence a receitas (natureza "1") e qual pertence a despesas (natureza "2")
  mesesItens.forEach((itens) => {
    itens.forEach((it) => {
      if ((it.natureza === "1" || it.natureza === "2") && it.conta) {
        const raiz = it.conta.split(".")[0];
        secaoPorPrefixo[raiz] = it.natureza === "1" ? "receitas" : "despesas";
      }
    });
  });

  mesesItens.forEach((itens, i) => {
    itens.forEach((it) => {
      if (!/^\d+(\.\d+)*$/.test(it.conta)) return;
      const raiz = it.conta.split(".")[0];
      const secao = it.natureza === "1" ? "receitas" : (it.natureza === "2" ? "despesas" : secaoPorPrefixo[raiz]);
      if (!secao) return;
      if (!mapa[it.conta]) {
        mapa[it.conta] = {
          conta: it.conta,
          nome: it.nome || nomesPlano[it.conta] || it.conta,
          natureza: it.natureza,
          secao,
          valores: new Array(n).fill(null),
        };
      } else if (it.nome && mapa[it.conta].nome === it.conta) {
        mapa[it.conta].nome = it.nome;
      }
      mapa[it.conta].valores[i] = it.valor;
    });
  });

  const folhas = Object.values(mapa).filter((c) => c.natureza === "1" || c.natureza === "2");
  folhas.forEach((f) => {
    const partes = f.conta.split(".");
    for (let k = 1; k < partes.length; k++) {
      const pai = partes.slice(0, k).join(".");
      if (!mapa[pai]) {
        mapa[pai] = { conta: pai, nome: nomesPlano[pai] || pai, natureza: "0", secao: f.secao, valores: new Array(n).fill(null) };
      }
    }
  });

  Object.values(mapa).filter((c) => c.natureza === "0").forEach((g) => {
    for (let i = 0; i < n; i++) {
      if (g.valores[i] !== null) continue;
      g.valores[i] = folhas.filter((x) => x.conta.indexOf(g.conta + ".") === 0).reduce((t, x) => t + (x.valores[i] || 0), 0);
    }
  });

  // Remove a conta raiz única de cada seção (ex: "1 - Receitas" e "2 - Despesas") para não duplicar o total da seção
  ["receitas", "despesas"].forEach((sec) => {
    const daSec = Object.values(mapa).filter((c) => c.secao === sec);
    const raizes = daSec.filter((c) => !c.conta.includes("."));
    if (raizes.length === 1 && raizes[0].natureza === "0") {
      delete mapa[raizes[0].conta];
    }
  });

  const cmp = (a, b) => {
    const A = a.split(".").map(Number), B = b.split(".").map(Number);
    for (let i = 0; i < Math.max(A.length, B.length); i++) {
      if ((A[i] === undefined ? -1 : A[i]) !== (B[i] === undefined ? -1 : B[i])) return (A[i] === undefined ? -1 : A[i]) - (B[i] === undefined ? -1 : B[i]);
    }
    return 0;
  };
  return Object.values(mapa)
    .filter((c) => c.valores.some((v) => v))
    .sort((a, b) => cmp(a.conta, b.conta))
    .map((c) => ({
      conta: c.conta,
      nome: c.nome,
      nivel: Math.max(0, c.conta.split(".").length - 2),
      natureza: c.natureza,
      secao: c.secao,
      valores: c.valores,
    }));
}

/* ===== Contas bancárias e saldo (API Superlógica) — compartilhado entre 011A e 16B ===== */
const CONTAS_BANCARIAS_CACHE = {};

async function fetchContasBancarias(idCondominio) {
  if (CONTAS_BANCARIAS_CACHE[idCondominio]) return CONTAS_BANCARIAS_CACHE[idCondominio];
  const url = BAL_API_BASE + "/api/superlogica/contabancos/index?exibirDadosAgencia=1&exibirContasFechadas=0&exibirDadosBanco=1&exibirContasAtivas=1&idCondominio=" + idCondominio;
  const res = await fetch(url);
  if (!res.ok) throw new Error("HTTP " + res.status + " ao listar contas bancárias");
  const data = await res.json();
  if (!Array.isArray(data)) throw new Error((data && data.msg) || "resposta inesperada ao listar contas bancárias");
  CONTAS_BANCARIAS_CACHE[idCondominio] = data;
  return data;
}

function renderContasSeletor(contas) {
  const wrap = document.getElementById("contasSeletorWrap");
  if (!wrap) return;
  if (!contas.length) { wrap.innerHTML = ""; return; }
  wrap.innerHTML = '<label style="display:block;margin-bottom:4px;">Contas bancárias a incluir no 16B</label>' +
    contas.map((c) => {
      const nome = c.st_descricao_cb || c.st_conta_cb || ("Conta " + c.id_contabanco_cb);
      return '<label class="conta-check" style="display:block;font-weight:400;margin-bottom:4px;">' +
        '<input type="checkbox" class="conta-chk" value="' + c.id_contabanco_cb + '" checked style="margin-right:6px;">' + nome + '</label>';
    }).join("");
}

async function carregarContasParaSeletor() {
  const idCondominio = document.getElementById("f_condominio_id").value;
  if (!idCondominio) return;
  try {
    const contas = await fetchContasBancarias(idCondominio);
    renderContasSeletor(contas);
  } catch (e) { /* seletor fica vazio — a busca automática ainda tenta buscar todas as contas */ }
}

function getContasSelecionadasIds() {
  const checks = document.querySelectorAll(".conta-chk");
  if (!checks.length) return null; // seletor nunca carregou — não filtra, busca todas
  return Array.from(checks).filter((c) => c.checked).map((c) => c.value);
}

// A resposta exata do endpoint de saldo não está documentada — essa função tenta
// reconhecer o valor em formatos plausíveis (objeto com campo contendo "saldo", número ou string).
function parseSaldoResponse(apiResponse) {
  let obj = apiResponse;
  if (Array.isArray(obj)) obj = obj[0];
  if (obj === null || obj === undefined) return null;
  if (typeof obj === "number") return obj;
  if (typeof obj === "string") { const n = parseFloat(obj.replace(",", ".")); return isNaN(n) ? null : n; }
  if (typeof obj === "object") {
    const keys = Object.keys(obj);
    const saldoKey = keys.find((k) => /saldo/i.test(k));
    if (saldoKey) { const n = parseFloat(String(obj[saldoKey]).replace(",", ".")); if (!isNaN(n)) return n; }
  }
  return null;
}

async function fetchSaldoConta(idConta, idCondominio, dataRefMMDDAAAA, ateAData) {
  const url = BAL_API_BASE + "/api/superlogica/caixa/saldo?idConta=" + idConta + "&idEmpresa=" + idCondominio + "&dtInicio=" + dataRefMMDDAAAA + "&ateAData=" + ateAData;
  const res = await fetch(url);
  if (!res.ok) throw new Error("HTTP " + res.status + " ao buscar saldo da conta " + idConta);
  const data = await res.json();
  return parseSaldoResponse(data);
}

// Soma o saldo de todas as contas ativas numa data de referência. Contas com erro são ignoradas
// (não derrubam o total) — se nenhuma conta responder, retorna null.
async function fetchSaldoConsolidado(contas, idCondominio, dataRefMMDDAAAA, ateAData) {
  let total = 0, algumEncontrado = false;
  for (const c of contas) {
    try {
      const v = await fetchSaldoConta(c.id_contabanco_cb, idCondominio, dataRefMMDDAAAA, ateAData);
      if (v !== null) { total += v; algumEncontrado = true; }
    } catch (e) { /* conta individual falhou — segue somando as demais */ }
  }
  return algumEncontrado ? total : null;
}

(function initPeriodo11A() {
  const elIni = document.getElementById("f_11a_mes_ini");
  const elFim = document.getElementById("f_11a_mes_fim");
  if (!elIni || !elFim) return;
  const hoje = new Date();
  let mFim = hoje.getMonth();
  let yFim = hoje.getFullYear();
  if (mFim === 0) { mFim = 12; yFim -= 1; }
  let mIni = mFim - 11;
  let yIni = yFim;
  if (mIni <= 0) { mIni += 12; yIni -= 1; }
  elIni.value = yIni + "-" + pad2(mIni);
  elFim.value = yFim + "-" + pad2(mFim);
})();

async function buscar11aAutomatico() {
  const idCondominio = document.getElementById("f_condominio_id").value;
  if (!idCondominio) { reportError("status11a", "Falta selecionar o condomínio", "Selecione um condomínio da lista de busca (aba 'Dados da assembleia') antes de buscar.", "Vá até a primeira aba, digite o nome do condomínio no campo de busca e escolha um item da lista que aparecer."); return; }
  const btn = document.getElementById("btnBuscar11a");
  const originalText = btn.textContent;
  btn.disabled = true;
  btn.innerHTML = '<span class="spinner" style="border-color:var(--button);border-top-color:transparent;"></span> Buscando...';
  setStatus("status11a", true, "Buscando os últimos 12 meses na Superlógica...");

  const valIni = document.getElementById("f_11a_mes_ini")?.value || "";
  const valFim = document.getElementById("f_11a_mes_fim")?.value || "";
  if (!valIni || !valFim) {
    btn.disabled = false;
    btn.textContent = originalText;
    reportError("status11a", "Falta o período", "Informe o mês/ano de início e fim antes de buscar.", "Preencha os dois campos de período no topo desta aba.");
    return;
  }
  const [yIni, mIni] = valIni.split("-").map(Number);
  const [yFim, mFim] = valFim.split("-").map(Number);
  if (yIni * 12 + mIni > yFim * 12 + mFim) {
    btn.disabled = false;
    btn.textContent = originalText;
    reportError("status11a", "Período inválido", "O mês/ano inicial não pode ser maior que o mês/ano final.", "Ajuste as datas do período.");
    return;
  }

  const periodos = [];
  let m = mIni, y = yIni;
  while (y * 12 + m <= yFim * 12 + mFim && periodos.length < 36) {
    periodos.push({ mes: m, ano: y });
    m += 1;
    if (m === 13) { m = 1; y += 1; }
  }

  const months = [], totalReceitas = [], totalDespesas = [], energiaEletrica = [], aguaEsgotoDespesa = [], saldoAnterior = [], saldoFinal = [];
  const categoriasMap = {}; // "secao|label" -> { secao, label, valores: [] } alinhado a months[]
  const mesesItens = []; // itens brutos do 11A por mês, usados pra montar o plano de contas

  let contas = [];
  let saldoDisponivel = true;
  try {
    contas = await fetchContasBancarias(idCondominio);
    if (!contas.length) saldoDisponivel = false;
  } catch (e) {
    saldoDisponivel = false;
  }

  let saldoFinalMesAnterior = null; // reaproveitado: saldo final do mês M-1 = saldo anterior do mês M

  try {
    for (let i = 0; i < periodos.length; i++) {
      const p = periodos[i];
      const resp = await fetchBalanceteMes(idCondominio, p.mes, p.ano);
      const parsed = parseBalanceteMes(resp);
      months.push(MESES_ABBR[p.mes - 1] + "/" + p.ano);
      mesesItens.push(parsed ? parsed.itensBrutos : []);
      totalReceitas.push(parsed ? parsed.totalReceitas : null);
      totalDespesas.push(parsed ? parsed.totalDespesas : null);
      energiaEletrica.push(parsed ? parsed.energia : null);
      aguaEsgotoDespesa.push(parsed ? parsed.agua : null);

      const chavesTocadas = new Set();
      (parsed ? parsed.categoriasFolha : []).forEach((cat) => {
        const chave = cat.secao + "|" + cat.label;
        if (!categoriasMap[chave]) {
          categoriasMap[chave] = { secao: cat.secao, label: cat.label, valores: new Array(months.length - 1).fill(null) };
        }
        categoriasMap[chave].valores.push(cat.valor);
        chavesTocadas.add(chave);
      });
      Object.keys(categoriasMap).forEach((chave) => {
        if (!chavesTocadas.has(chave)) categoriasMap[chave].valores.push(null);
      });

      if (saldoDisponivel) {
        let saldoAntMes = saldoFinalMesAnterior;
        if (saldoAntMes === null) {
          const primeiroDia = pad2(p.mes) + "/01/" + p.ano;
          saldoAntMes = await fetchSaldoConsolidado(contas, idCondominio, primeiroDia, 0);
        }
        const ultimoDia = pad2(p.mes) + "/" + pad2(lastDayOfMonth(p.mes, p.ano)) + "/" + p.ano;
        const saldoFimMes = await fetchSaldoConsolidado(contas, idCondominio, ultimoDia, 1);
        saldoAnterior.push(saldoAntMes);
        saldoFinal.push(saldoFimMes);
        saldoFinalMesAnterior = saldoFimMes;
      } else {
        saldoAnterior.push(null);
        saldoFinal.push(null);
      }
      setStatus("status11a", true, "Buscando... " + months.length + "/" + periodos.length + " meses (" + months[months.length - 1] + ")");
    }

    let contasPlano = null;
    try {
      contasPlano = montarPlanoContas(await fetchPlanoContas(idCondominio), mesesItens);
      if (!contasPlano.length) contasPlano = null;
    } catch (e) { /* sem plano de contas: o slide de categorias simplesmente não é gerado */ }

    const data = {
      months,
      contasPlano,
      saldoAnterior,
      saldoFinal,
      totalReceitas,
      totalDespesas,
      energiaEletrica: energiaEletrica.some((v) => v !== null) ? energiaEletrica : null,
      aguaEsgotoDespesa: aguaEsgotoDespesa.some((v) => v !== null) ? aguaEsgotoDespesa : null,
      categoriasDetalhadas: Object.values(categoriasMap).sort((a, b) => a.label.localeCompare(b.label, "pt-BR")),
    };
    PARSED.a11 = data;
    document.getElementById("dz11a").classList.add("ok");
    const msgSaldo = saldoDisponivel
      ? "Saldo anterior/final também buscados das contas bancárias."
      : "Não achei contas bancárias ativas para esse condomínio — saldo anterior/final seguem manuais.";
    setStatus("status11a", true, months.length + " meses buscados automaticamente (" + months[0] + " a " + months[months.length - 1] + "). " + msgSaldo);
    renderPreview11A(data);
    if (typeof refreshPreview === "function") refreshPreview();
  } catch (err) {
    PARSED.a11 = null;
    reportError("status11a", "Não consegui buscar o 11A", "Erro ao buscar: " + err.message, "Confira sua conexão com a internet e se o condomínio selecionado está correto. Se o erro continuar, tente novamente em alguns minutos ou envie o arquivo .xlsx manualmente.");
  } finally {
    btn.disabled = false;
    btn.textContent = originalText;
  }
}

window.onCondominioSelected = function (id) {
  const btn11a = document.getElementById("btnBuscar11a");
  if (btn11a) btn11a.disabled = !id;
  const btnInad = document.getElementById("btnBuscarInad");
  if (btnInad) btnInad.disabled = !id;
  const btn16b = document.getElementById("btnBuscar16b");
  if (btn16b) btn16b.disabled = !id;
  if (id) carregarContasParaSeletor();
};



document.getElementById("btnBuscar11a").addEventListener("click", buscar11aAutomatico);

/* ===== Inadimplência automática ===== */
// Busca o total de unidades do condomínio (inadimplentes ou não), paginando até a API
// devolver menos que o tamanho de página pedido. Usado só para calcular o % de inadimplência.
async function fetchTotalUnidades(idCondominio) {
  const ITENS_POR_PAGINA = 50;
  let total = 0;
  let pagina = 1;
  while (true) {
    const url = BAL_API_BASE + "/api/superlogica/unidades/index?idCondominio=" + idCondominio +
      "&exibirGruposDasUnidades=0&itensPorPagina=" + ITENS_POR_PAGINA + "&pagina=" + pagina + "&exibirDadosDosContatos=0";
    const res = await fetch(url);
    if (!res.ok) throw new Error("HTTP " + res.status + " ao listar unidades (página " + pagina + ")");
    const data = await res.json();
    if (!Array.isArray(data)) throw new Error((data && data.msg) || "resposta inesperada ao listar unidades");
    total += data.length;
    if (data.length < ITENS_POR_PAGINA) break;
    pagina += 1;
    if (pagina > 40) break; // trava de segurança
  }
  return total;
}

// % de inadimplência = unidades inadimplentes ÷ total de unidades × 100 (mesma conta do backend).
function pctInadimplencia(unidades, totalUnidades) {
  const u = parseInt(String(unidades).replace(/\D/g, ""), 10);
  const t = parseInt(String(totalUnidades).replace(/\D/g, ""), 10);
  if (isNaN(u) || isNaN(t) || t <= 0) return null;
  return ((u / t) * 100).toFixed(2).replace(".", ",");
}
function recalcInadPct() {
  const pct = pctInadimplencia(document.getElementById("f_inad_unidades").value, document.getElementById("f_inad_unidades_total").value);
  if (pct !== null) document.getElementById("f_inad_total").value = pct;
}
["f_inad_unidades", "f_inad_unidades_total"].forEach((id) => document.getElementById(id).addEventListener("input", recalcInadPct));

// Cobrança paga pela garantidora: o campo fl_statusgarantidora_recb vem vazio quando não há garantidora.
function cobrancaGarantida(c) {
  const v = String(c && c.fl_statusgarantidora_recb !== undefined && c.fl_statusgarantidora_recb !== null ? c.fl_statusgarantidora_recb : "").trim();
  return v !== "" && v !== "0";
}
async function fetchCondominioGarantido(idCondominio) {
  const res = await fetch(BAL_API_BASE + "/api/superlogica/condominios/get?id=" + idCondominio);
  if (!res.ok) throw new Error("HTTP " + res.status);
  const data = await res.json();
  const c = Array.isArray(data) ? data[0] : data;
  const v = String(c && c.fl_garantido_cond !== undefined && c.fl_garantido_cond !== null ? c.fl_garantido_cond : "").trim();
  return v !== "" && v !== "0";
}

async function buscarInadimplenciaAutomatico() {
  const idCondominio = document.getElementById("f_condominio_id").value;
  if (!idCondominio) { reportError("statusInad", "Falta selecionar o condomínio", "Selecione um condomínio da lista de busca (aba 'Dados da assembleia') antes de buscar.", "Vá até a primeira aba, digite o nome do condomínio no campo de busca e escolha um item da lista que aparecer."); return; }
  const btn = document.getElementById("btnBuscarInad");
  const originalText = btn.textContent;
  btn.disabled = true;
  btn.innerHTML = '<span class="spinner" style="border-color:var(--button);border-top-color:transparent;"></span> Buscando...';
  setStatus("statusInad", true, "Buscando inadimplência na Superlógica...");

  try {
    const hoje = new Date();
    const posicaoEm = pad2(hoje.getMonth() + 1) + "/" + pad2(hoje.getDate()) + "/" + hoje.getFullYear();
    const ITENS_POR_PAGINA = 50;
    let todasUnidades = [];
    let pagina = 1;
    while (true) {
      const url = BAL_API_BASE + "/api/superlogica/inadimplencia/index?posicaoEm=" + posicaoEm +
        "&comValoresAtualizados=0&comValoresAtualizadosPorComposicao=0&idCondominio=" + idCondominio +
        "&comDadosDaReceita=1&itensPorPagina=" + ITENS_POR_PAGINA + "&pagina=" + pagina + "&semAcordo=0&semProcesso=0";
      // Sem cobrancaDoTipo: esse filtro devolvia só cobranças rotuladas INADIMPLENTE e escondia unidades/valores que o painel conta.
      const res = await fetch(url);
      if (!res.ok) throw new Error("HTTP " + res.status + " (página " + pagina + ")");
      const data = await res.json();
      if (!Array.isArray(data)) throw new Error((data && data.msg) || "resposta inesperada");
      todasUnidades = todasUnidades.concat(data);
      setStatus("statusInad", true, "Buscando inadimplência... " + todasUnidades.length + " unidades encontradas até agora");
      if (data.length < ITENS_POR_PAGINA) break; // última página
      pagina += 1;
      if (pagina > 40) break; // trava de segurança
    }

    // Cada item do array é uma unidade; "recebimento" é a lista de cobranças em aberto dela.
    const unidadesInad = todasUnidades.length;
    let valorTotal = 0, cobrancasTotal = 0;
    todasUnidades.forEach((unidade) => {
      const cobrancas = Array.isArray(unidade.recebimento) ? unidade.recebimento : [];
      cobrancasTotal += cobrancas.length;
      cobrancas.forEach((c) => {
        valorTotal += parseFloat(c.vl_total_recb) || 0;
      });
    });

    document.getElementById("f_inad_valor").value = valorTotal.toLocaleString("pt-BR", { minimumFractionDigits: 2, maximumFractionDigits: 2 });
    document.getElementById("f_inad_cobrancas").value = String(cobrancasTotal);
    document.getElementById("f_inad_unidades").value = String(unidadesInad);

    let totalUnidades = null;
    try {
      totalUnidades = await fetchTotalUnidades(idCondominio);
    } catch (e) { /* endpoint de unidades falhou — degrada sem quebrar o resto */ }

    // Garantidora: só existe se o condomínio tem fl_garantido_cond (1/-1). Sem isso (vazio) => 0%.
    // Uma mesma unidade pode ter cobranças orgânicas e garantidas simultaneamente, então cada carteira é contada de forma independente.
    let temGarantidora = false;
    try { temGarantidora = await fetchCondominioGarantido(idCondominio); } catch (e) { /* assume sem garantidora */ }
    const garantidas = temGarantidora
      ? todasUnidades.filter((u) => (Array.isArray(u.recebimento) ? u.recebimento : []).some((c) => cobrancaGarantida(c))).length
      : 0;
    const organicas = temGarantidora
      ? todasUnidades.filter((u) => (Array.isArray(u.recebimento) ? u.recebimento : []).some((c) => !cobrancaGarantida(c))).length
      : unidadesInad;

    if (totalUnidades) {
      document.getElementById("f_inad_unidades_total").value = String(totalUnidades);
      recalcInadPct();
      const pctDe = (n) => ((n / totalUnidades) * 100).toFixed(2).replace(".", ",");
      document.getElementById("f_inad_garantidora").value = temGarantidora ? pctDe(garantidas) : "0,00";
      document.getElementById("f_inad_organica").value = pctDe(organicas);
      setStatus("statusInad", true, unidadesInad + " de " + totalUnidades + " unidades inadimplentes (" + cobrancasTotal + " cobranças) — valor, cobranças e as 3 porcentagens preenchidos" + (temGarantidora ? " (orgânica: " + organicas + " unid. | garantidora: " + garantidas + " unid.)." : " (condomínio sem garantidora: garantidora = 0%)."));
    } else {
      setStatus("statusInad", true, unidadesInad + " unidades inadimplentes (" + cobrancasTotal + " cobranças) — valor e cobranças preenchidos. Não consegui o total de unidades; as 3 % seguem manuais.");
    }
    if (typeof refreshPreview === "function") refreshPreview();
  } catch (err) {
    reportError("statusInad", "Não consegui buscar a inadimplência", "Erro ao buscar: " + err.message, "Confira sua conexão e se o condomínio está selecionado corretamente. Se continuar falhando, tente de novo em alguns minutos.");
  } finally {
    btn.disabled = false;
    btn.textContent = originalText;
  }
}

document.getElementById("btnBuscarInad").addEventListener("click", buscarInadimplenciaAutomatico);

/* ===== 16B automático (saldo anterior/final por conta bancária) ===== */
// Busca as movimentações de uma conta no período e soma créditos/débitos.
// fl_dc_mov: "C" (crédito, vl_valor_mov positivo) ou "D" (débito, vl_valor_mov negativo).
async function fetchMovimentacoesConta(idConta, idCondominio, dtIniFmt, dtFimFmt) {
  const url = BAL_API_BASE + "/api/superlogica/caixa?idCondominio=" + idCondominio +
    "&dtInicio=" + dtIniFmt + "&dtFim=" + dtFimFmt + "&tipoFiltroData=periodo&idContaBancaria=" + idConta + "&semSaldoInicial=1";
  const res = await fetch(url);
  if (!res.ok) throw new Error("HTTP " + res.status + " ao buscar movimentações da conta " + idConta);
  const data = await res.json();
  if (!Array.isArray(data)) throw new Error((data && data.msg) || "resposta inesperada");
  let creditos = 0, debitos = 0;
  const movimentos = [];
  data.forEach((mov) => {
    const v = parseFloat(mov.vl_valor_mov) || 0;
    if (mov.fl_dc_mov === "C") creditos += v;
    else if (mov.fl_dc_mov === "D") debitos += Math.abs(v);
    movimentos.push({
      data: (mov.dt_data_mov || "").slice(0, 10),
      descricao: mov.st_descricao_mov || "",
      tipo: mov.fl_dc_mov,
      valor: Math.abs(v),
    });
  });
  movimentos.sort((a, b) => a.data.localeCompare(b.data));
  return { creditos, debitos, movimentos };
}

async function buscar16bAutomatico() {
  const idCondominio = document.getElementById("f_condominio_id").value;
  const dtIni = document.getElementById("f_16b_data_ini").value; // AAAA-MM-DD
  const dtFim = document.getElementById("f_16b_data_fim").value;
  if (!idCondominio) { reportError("status16b", "Falta selecionar o condomínio", "Selecione um condomínio da lista de busca (aba 'Dados da assembleia') antes de buscar.", "Vá até a primeira aba, digite o nome do condomínio no campo de busca e escolha um item da lista que aparecer."); return; }
  if (!dtIni || !dtFim) { reportError("status16b", "Faltam as datas do período", "Preencha a data de início e fim do período.", "Preencha os dois campos de data no topo desta aba antes de clicar em buscar."); return; }

  const btn = document.getElementById("btnBuscar16b");
  const originalText = btn.textContent;
  btn.disabled = true;
  btn.innerHTML = '<span class="spinner" style="border-color:var(--button);border-top-color:transparent;"></span> Buscando...';
  setStatus("status16b", true, "Buscando contas bancárias, saldos e movimentações...");

  try {
    let contas = await fetchContasBancarias(idCondominio);
    const selecionadas = getContasSelecionadasIds();
    if (selecionadas) contas = contas.filter((c) => selecionadas.includes(String(c.id_contabanco_cb)));
    if (!contas.length) throw new Error("nenhuma conta bancária selecionada (confira o seletor de contas no topo desta aba).");

    const [anoI, mesI, diaI] = dtIni.split("-");
    const [anoF, mesF, diaF] = dtFim.split("-");
    const dtIniFmt = mesI + "/" + diaI + "/" + anoI;
    const dtFimFmt = mesF + "/" + diaF + "/" + anoF;

    const linhas = [];
    let algumaMovimentacaoFalhou = false;
    for (const c of contas) {
      let saldoAnt = null, saldoFim = null, creditos = null, debitos = null, movimentos = [];
      try { saldoAnt = await fetchSaldoConta(c.id_contabanco_cb, idCondominio, dtIniFmt, 0); } catch (e) { /* ignora, fica — */ }
      try { saldoFim = await fetchSaldoConta(c.id_contabanco_cb, idCondominio, dtFimFmt, 1); } catch (e) { /* ignora, fica — */ }
      try {
        const mov = await fetchMovimentacoesConta(c.id_contabanco_cb, idCondominio, dtIniFmt, dtFimFmt);
        creditos = mov.creditos;
        debitos = mov.debitos;
        movimentos = mov.movimentos;
      } catch (e) { algumaMovimentacaoFalhou = true; }
      linhas.push({
        nome: c.st_descricao_cb || c.st_conta_cb || ("Conta " + c.id_contabanco_cb),
        saldoAnt, creditos, debitos, saldoFinal: saldoFim, movimentos,
      });
    }

    PARSED.b16 = { contas: linhas, grupos: [], saldoFinalGeral: null };
    document.getElementById("f_16b_periodo").value = "De " + dtIniFmt + " até " + dtFimFmt;
    const msgFinal = algumaMovimentacaoFalhou
      ? linhas.length + " conta(s) encontradas — algumas movimentações não puderam ser buscadas (ficaram como —); confira ou cole o PDF manualmente."
      : linhas.length + " conta(s) encontradas — saldo anterior/final e créditos/débitos preenchidos automaticamente.";
    setStatus("status16b", true, msgFinal);
    renderPreview16B(PARSED.b16);
    if (typeof refreshPreview === "function") refreshPreview();
  } catch (err) {
    PARSED.b16 = null;
    reportError("status16b", "Não consegui buscar o 16B", "Erro: " + err.message, "Confira sua conexão, o condomínio selecionado e o período informado. Se preferir, envie o relatório 16B (PDF ou planilha) logo abaixo.");
  } finally {
    btn.disabled = false;
    btn.textContent = originalText;
  }
}

document.getElementById("btnBuscar16b").addEventListener("click", buscar16bAutomatico);

/* Previsão Orçamentária: testamos os endpoints W025A e W046A (família /relatorios/id/XXXA) e
   eles retornam vazio ({"id_impressao":"<div id=\"relatorio\"></div>"}) — essa família de endpoints
   parece depender de o relatório já ter sido gerado manualmente no painel da Superlógica antes,
   diferente do /balancetes/index (usado no 011A), que calcula os dados na hora. Não há automação
   possível com o que está mapeado até agora — segue só o upload manual do .xlsx (abaixo). */

setupDropzone("dz11a", "file11a", async (file) => {
  const dz = document.getElementById("dz11a");
  dz._mostrar(file.name);
  try {
    const wb = await lerArquivoTabela(file, "11a");
    const data = parse11A(wb);
    PARSED.a11 = data;
    dz.classList.add("ok");
    avisarCondominio("status11a", nomeCondominioDoDocumento(linhasIniciaisDaAba(wb)));
    setStatus("status11a", true, data.months.length + " meses encontrados (" + data.months[0] + " a " + data.months[data.months.length - 1] + ")");
    renderPreview11A(data);
    if (typeof refreshPreview === "function") refreshPreview();
  } catch (err) {
    PARSED.a11 = null;
    dz.classList.remove("ok");
    avisarCondominio("status11a", null);
    reportError("status11a", "Não consegui ler esse arquivo", "Erro: " + err.message, "Confira se esse é mesmo o arquivo .xlsx exportado do relatório 11A da Superlógica, sem edições manuais na estrutura das colunas.");
    document.getElementById("preview11a").innerHTML = "";
  }
}, () => {
  PARSED.a11 = null;
  document.getElementById("preview11a").innerHTML = "";
  document.getElementById("status11a").textContent = "";
  document.getElementById("status11a").className = "status-line";
  avisarCondominio("status11a", null);
});

let PO_WB = null; // planilha de previsão enviada (para trocar de aba sem reenviar)
// Importa uma aba de 3 colunas para a tabela manual e mostra o resultado (a tabela manual fica visível: é onde se edita o Futuro).
function poAplicarImport3Col(wb, aba) {
  const r = poImportar3Col(wb, aba);
  document.getElementById("poManual").hidden = false;
  poRenderTabela();
  poAtualizarDados();
  setStatus("statuspo", true, 'aba "' + aba + '" (modelo de 3 colunas): ' + r.itens + " itens importados com Previsto e Realizado. " +
    (r.temFuturoPreenchido ? "Os valores de Futuro que já estavam na planilha foram mantidos; os vazios usam Futuro = Realizado (editável)." : "Futuro foi preenchido com o valor do Realizado (sugestão; edite linha a linha ou use o ajuste em massa)."));
}
// Seletor visível de aba: só aparece quando mais de uma aba tem previsão; mostra o condomínio de cada uma.
function poMontarSeletorAba(wb, usada) {
  const wrap = document.getElementById("poAbaWrap");
  const abas = listarAbasPrevisao(wb).filter((a) => a.valida);
  if (abas.length < 2) { wrap.hidden = true; return; }
  const sel = document.getElementById("poAbaSel");
  sel.innerHTML = abas.map((a) => '<option value="' + poAttr(a.nome) + '"' + (a.nome === usada ? " selected" : "") + ">" + poAttr(a.nome + (a.condominio ? " — " + a.condominio : "")) + "</option>").join("");
  wrap.hidden = false;
}
setupDropzone("dzpo", "filepo", async (file) => {
  const dz = document.getElementById("dzpo");
  dz._mostrar(file.name);
  try {
    const wb = await lerArquivoTabela(file, "previsao");
    PO_WB = wb;
    const abaInicial = pickPrevisaoSheet(wb);
    const avisarPO = (aba) => avisarCondominio("statuspo", nomeCondominioDoDocumento(linhasIniciaisDaAba(wb, aba)));
    if (abaEh3Col(wb, abaInicial)) { // modelo de 3 colunas: preenche a tabela manual (Futuro = Realizado)
      poAplicarImport3Col(wb, abaInicial);
      avisarPO(abaInicial);
      dz.classList.add("ok");
      poMontarSeletorAba(wb, abaInicial);
      if (typeof refreshPreview === "function") refreshPreview();
      return;
    }
    document.getElementById("poManual").hidden = true;
    const data = parsePrevisao(wb);
    PARSED.prev = data;
    const unPlanilha = (data.rateioPrevisto && data.rateioPrevisto.unidades) || (data.rateioAtual && data.rateioAtual.unidades);
    const elUnPo = document.getElementById("f_po_unidades_total");
    if (unPlanilha && elUnPo && !elUnPo.value) elUnPo.value = String(unPlanilha);
    poMontarSeletorAba(wb, data.sheetUsed);
    dz.classList.add("ok");
    avisarPO(data.sheetUsed);
    // modelo antigo (Atual/Previsto): vale o que foi feito por último, a tabela manual some enquanto há planilha enviada
    setStatus("statuspo", true, 'lendo aba "' + data.sheetUsed + '" — ' + data.categories.length + " categorias, total previsto " + fmtBRL(data.totalDespesasPrevisto, true));
    renderPreviewPO(data);
    if (typeof refreshPreview === "function") refreshPreview();
  } catch (err) {
    PARSED.prev = null;
    dz.classList.remove("ok");
    avisarCondominio("statuspo", null);
    reportError("statuspo", "Não consegui ler essa planilha", "Erro: " + err.message, "Confira se esse é o arquivo .xlsx da previsão orçamentária, com as colunas Atual e Previsto, sem edições na estrutura das abas.");
    document.getElementById("previewpo").innerHTML = "";
  }
}, () => {
  if (PARSED.prev && PARSED.prev.sheetUsed !== "preenchimento manual") PARSED.prev = null;
  document.getElementById("poManual").hidden = false; // ao remover a planilha, a tabela manual volta
  if (PO_MANUAL.importado) poResetarModelo(); // os dados importados saem junto com o arquivo: volta ao modelo em branco
  PO_WB = null;
  document.getElementById("poAbaWrap").hidden = true;
  document.getElementById("previewpo").innerHTML = "";
  document.getElementById("statuspo").textContent = "";
  document.getElementById("statuspo").className = "status-line";
  avisarCondominio("statuspo", null);
  if (typeof poAtualizarDados === "function") poAtualizarDados(); // se há tabela manual, ela volta a valer
});

/* ---------- Previsão orçamentária: preenchimento manual ---------- */
// Linhas: { tipo, nome, vals: [texto por coluna], pct } com tipo:
//   grupo      cabeçalho de grupo (vira título do slide)
//   item       linha de valores digitados
//   subtotal   soma dos itens do grupo acima (automático)
//   total      valor do total anterior + subtotais e percentuais desde ele (automático): Total de pagamentos, Sub-Total, Total despesas
//   percentual % digitado x valor do último total (automático): Garantidora, Fundo de reserva
// Colunas de valor são livres (nome, ordem, quantidade). Começa com o modelo da planilha MINUTA PO.
const PO_MODELO_GRUPOS = [
  ["CONTRATOS", ["Manutenção Mat. Eletrico/Hidraulico", "Material de Construção", "Manutenção Extintores de Incêndio", "Material de Limpeza", "Manutenção de Interforne / Portão", "Limpeza de caixa d`agua", "Controle de Pragas", "Manutenção de Jardinagem", "Material de Jardinagem", "Manutenção de Bombas", "Limpeza caixa de gordura", "Aquisições", "Eventual necessidade"]],
  ["PESSOAL", ["Folha salarial Porteiro (Diurno)", "Folha salarial Zelador"]],
  ["TARIFAS PUBLICAS", ["Água e esgoto|Conf. Consumo", "Energia Elétrica", "Telefone/Internet"]],
  ["ADMINISTRATIVAS", ["Gestão Administrativa", "Material de Escritório / Expediente", "Certificado digital", "ISS", "Registro de ATAS", "DIRF/RAIS/SST - Esocial", "Despesas Bancárias/repasse boletos", "Pró labore síndico", "Pró labore Subsíndico"]],
  ["DIVERSOS", ["Seguro Predial Obrigatório"]],
];
function poLinhasModelo() {
  const L = [];
  PO_MODELO_GRUPOS.forEach(([g, itens]) => {
    L.push({ tipo: "grupo", nome: g, vals: [] });
    itens.forEach((t) => { const [nome, fixo] = t.split("|"); L.push({ tipo: "item", nome, vals: fixo ? [fixo, fixo] : [] }); });
    L.push({ tipo: "subtotal", nome: "Sub-Total", vals: [] });
  });
  L.push({ tipo: "total", nome: "Total de pagamentos", vals: [] });
  L.push({ tipo: "percentual", nome: "Garantidora", pct: "10", vals: [] });
  L.push({ tipo: "total", nome: "Sub-Total", vals: [] });
  L.push({ tipo: "grupo", nome: "FUNDO DE RESERVA", vals: [] });
  L.push({ tipo: "percentual", nome: "Fundo de reserva", pct: "5", vals: [] });
  L.push({ tipo: "subtotal", nome: "Sub-Total", vals: [] });
  L.push({ tipo: "total", nome: "Total despesas", vals: [] });
  return L;
}
function poRateioModelo(titulo) {
  return { titulo, unidades: "", itens: [{ nome: "Cota Condominial", recibo: "" }, { nome: "Fundo de Reserva", recibo: "" }, { nome: "Agua/Esgoto", recibo: "Conf. Consumo" }] };
}
const PO_MANUAL = { colunas: ["Atual", "Previsto"], colMeta: [{}, {}], linhas: poLinhasModelo(), rateio: { atual: poRateioModelo("Rateio atual"), previsto: poRateioModelo("Previsão (novo rateio)") }, usado: false, importado: false };
// Papéis de coluna (modelo de 3 colunas): "previsto" (última previsão aprovada), "realizado" e "futuro" (nova previsão).
// colMeta[c] acompanha colunas[c] em qualquer inserção/remoção/reordenação. Linha de item com futAuto=true usa
// Futuro = Realizado (sugestão) até o usuário editar aquela célula. Linhas têm id estável (rateio aponta para elas).
let PO_ID = 0;
function poGarantirIds() {
  PO_MANUAL.linhas.forEach((l) => { if (!l.id) l.id = ++PO_ID; });
  while (PO_MANUAL.colMeta.length < PO_MANUAL.colunas.length) PO_MANUAL.colMeta.push({});
  PO_MANUAL.colMeta.length = PO_MANUAL.colunas.length;
}
function poIdxPapel(papel) { return PO_MANUAL.colMeta.findIndex((m) => m && m.papel === papel); }

function poNumero(txt) {
  if (txt === null || txt === undefined || String(txt).trim() === "") return null;
  if (/conf(orme|\.?)\s*(o\s*)?consumo/i.test(String(txt))) return null;
  const n = parseBRNumber(String(txt));
  return n === null || isNaN(n) ? null : n;
}
function poValorBruto(txt) {
  const t = (txt === null || txt === undefined) ? "" : String(txt).trim();
  if (t === "") return null;
  if (/conf(orme|\.?)\s*(o\s*)?consumo/i.test(t)) return "Conforme consumo";
  const n = parseBRNumber(t);
  return n === null || isNaN(n) ? t : n;
}
const poSoma = (a, b) => (a === null && b === null ? null : (a || 0) + (b || 0));

// Calcula, para cada linha, o valor por coluna (itens: digitado; subtotal/total/percentual: automático).
function poCalcular() {
  poGarantirIds();
  const nC = PO_MANUAL.colunas.length;
  const iReal = poIdxPapel("realizado"), iFut = poIdxPapel("futuro");
  const vazio = () => new Array(nC).fill(null);
  let grupo = vazio(), desdeTotal = vazio(), ultimoTotal = null, emGrupo = false;
  const porId = {}; // valores já calculados de cada linha (id -> vetor), para percentuais com base em um intervalo de itens
  return PO_MANUAL.linhas.map((l) => { const v = poCalcularLinha(l); porId[l.id] = v; return v; });
  function poCalcularLinha(l) {
    if (l.tipo === "grupo") { grupo = vazio(); emGrupo = true; return null; }
    if (l.tipo === "item") {
      const v = [];
      for (let c = 0; c < nC; c++) {
        const fonte = (c === iFut && l.futAuto && iReal >= 0) ? l.vals[iReal] : l.vals[c]; // Futuro = Realizado até editar
        v.push(poNumero(fonte)); grupo[c] = poSoma(grupo[c], v[c]);
      }
      return v;
    }
    if (l.tipo === "subtotal") {
      const v = grupo.slice();
      for (let c = 0; c < nC; c++) desdeTotal[c] = poSoma(desdeTotal[c], v[c]);
      grupo = vazio(); emGrupo = false;
      return v;
    }
    if (l.tipo === "percentual") {
      const f = (poNumero(l.pct) || 0) / 100;
      // modos por coluna: "sobre" = % x base; "dentro" = % por dentro (base x %/(1-%)), como em "(total*100/94)*6%";
      // "fixo" = valor digitado na planilha (sem fórmula). Base = total anterior, ou a soma de um intervalo de itens (baseIds).
      const v = vazio().map((_, c) => {
        const modo = (l.modos && l.modos[c]) || "sobre";
        if (modo === "fixo") return poNumero(l.fixos && l.fixos[c]);
        let base = null;
        if (l.baseIds && l.baseIds.length) l.baseIds.forEach((id) => { if (porId[id] && porId[id][c] !== null) base = poSoma(base, porId[id][c]); });
        else base = ultimoTotal ? ultimoTotal[c] : null;
        if (base === null) return null;
        return modo === "dentro" && f < 1 ? base * f / (1 - f) : base * f;
      });
      // dentro de um grupo (ex.: FUNDO DE RESERVA) entra no sub-total do grupo; fora, soma direto no total seguinte
      const alvo = emGrupo ? grupo : desdeTotal;
      for (let c = 0; c < nC; c++) alvo[c] = poSoma(alvo[c], v[c]);
      return v;
    }
    // total
    const v = vazio().map((_, c) => poSoma(ultimoTotal ? ultimoTotal[c] : null, desdeTotal[c]));
    ultimoTotal = v.slice();
    desdeTotal = vazio(); emGrupo = false;
    return v;
  }
}

// Recibo de um item do rateio: digitado, ou (se vinculado) valor da linha da tabela ÷ unidades, recalculado sozinho.
function poReciboItem(it, un, calc) {
  if (it.origem) {
    const ix = PO_MANUAL.linhas.findIndex((l) => l.id === it.origem.linhaId), c = poIdxPapel(it.origem.papel);
    return ix >= 0 && c >= 0 && calc[ix] && calc[ix][c] !== null && un ? calc[ix][c] / un : null;
  }
  return poNumero(it.recibo);
}
function poItemPreenchido(it, un, calc) {
  if (it.origem) return poReciboItem(it, un, calc) !== null;
  return String(it.recibo === null || it.recibo === undefined ? "" : it.recibo).trim() !== "";
}

function poRateioBloco(b, calc) {
  const un = poNumero(b.unidades);
  const itens = b.itens.filter((it) => poItemPreenchido(it, un, calc)).map((it) => {
    const rec = poReciboItem(it, un, calc);
    return { nome: (it.nome || "").trim() || "Item", recibo: rec, total: rec !== null && un ? rec * un : null };
  });
  const totRec = itens.reduce((a, i) => poSoma(a, i.recibo), null);
  const totTot = itens.reduce((a, i) => poSoma(a, i.total), null);
  return { itens, total: totTot, totalRecibo: totRec, reajuste: null, reajusteLabel: null, reajusteIsPercent: false };
}

function poConstruirDados() {
  const L = PO_MANUAL.linhas, calc = poCalcular(), nC = PO_MANUAL.colunas.length;
  const temValor = L.some((l, i) => l.tipo === "item" && calc[i] && calc[i].some((v) => v !== null));
  const iReal = poIdxPapel("realizado"), iFut = poIdxPapel("futuro");
  const unA = poNumero(PO_MANUAL.rateio.atual.unidades), unP = poNumero(PO_MANUAL.rateio.previsto.unidades);
  const rateioAtualOk = PO_MANUAL.rateio.atual.itens.some((i) => poReciboItem(i, unA, calc) !== null);
  const rateioPrevOk = PO_MANUAL.rateio.previsto.itens.some((i) => poReciboItem(i, unP, calc) !== null);
  if (!temValor && !rateioAtualOk && !rateioPrevOk) return null;
  const categories = [], itensCompletos = [];
  let grupoNome = null, totalFinal = null;
  const bruto = (i) => (calc[i] || new Array(nC).fill(null));
  L.forEach((l, i) => {
    const nome = (l.nome || "").trim();
    if (l.tipo === "grupo") { grupoNome = nome || "Grupo"; itensCompletos.push({ tipo: "grupo", nome: grupoNome }); return; }
    if (l.tipo === "item") {
      const vs = l.vals.slice(0, nC).concat(new Array(Math.max(0, nC - l.vals.length)).fill(null));
      if (l.futAuto && iFut >= 0 && iReal >= 0) vs[iFut] = vs[iReal]; // Futuro = Realizado (sugestão) enquanto não editado
      itensCompletos.push({ tipo: "item", nome: nome || "Item", valores: vs.map(poValorBruto) });
      return;
    }
    const v = bruto(i);
    if (l.tipo === "percentual") {
      itensCompletos.push({ tipo: "item", nome: (nome || "Percentual") + " (" + String(poNumero(l.pct) || 0).replace(".", ",") + "%)", valores: v });
      return;
    }
    itensCompletos.push({ tipo: "subtotal", nome: nome || "Sub-Total", valores: v });
    if (l.tipo === "subtotal") categories.push({ nome: grupoNome || nome || "Sub-Total", atual: v[iReal >= 0 ? iReal : 0], previsto: v[iFut >= 0 ? iFut : 1] === undefined ? null : v[iFut >= 0 ? iFut : 1] });
    else totalFinal = v;
  });
  if (!categories.length) {
    itensCompletos.filter((i) => i.tipo === "item").forEach((i) => categories.push({ nome: i.nome, atual: poNumero(i.valores[0]), previsto: poNumero(i.valores[1]) }));
  }
  // linhas de valor: 1ª coluna = Atual, 2ª = Previsto no resumo/cartões (quando existirem)
  itensCompletos.forEach((it) => { if (it.valores) { it.atualRaw = it.valores[iReal >= 0 ? iReal : 0]; it.previstoRaw = it.valores[iFut >= 0 ? iFut : 1]; } });
  let rateioAtual = null, rateioPrevisto = null;
  if (rateioAtualOk) rateioAtual = poRateioBloco(PO_MANUAL.rateio.atual, calc);
  if (rateioPrevOk) rateioPrevisto = poRateioBloco(PO_MANUAL.rateio.previsto, calc);
  if (rateioAtual && rateioPrevisto && rateioAtual.totalRecibo && rateioPrevisto.totalRecibo !== null) {
    rateioPrevisto.reajuste = rateioPrevisto.totalRecibo / rateioAtual.totalRecibo - 1;
    rateioPrevisto.reajusteLabel = "Reajuste de";
    rateioPrevisto.reajusteIsPercent = true;
  }
  const unManual = unP || unA || parseInt(document.getElementById("f_po_unidades_total")?.value || "", 10) || null;
  return {
    sheetUsed: "preenchimento manual", condominioNome: null, categories, itensCompletos,
    colunasValores: PO_MANUAL.colunas.map((n) => (n || "").trim() || "Coluna"),
    totalDespesasAtual: totalFinal ? totalFinal[iReal >= 0 ? iReal : 0] : null, totalDespesasPrevisto: totalFinal && totalFinal[iFut >= 0 ? iFut : 1] !== undefined ? totalFinal[iFut >= 0 ? iFut : 1] : null,
    rateioAtual, rateioPrevisto, unidadesTotal: unManual,
  };
}
function poAtualizarDados() {
  const dados = poConstruirDados();
  if (!dados) {
    if (PARSED.prev && PARSED.prev.sheetUsed === "preenchimento manual") PARSED.prev = null;
    if (PO_MANUAL.usado) setStatus("statuspo", true, "Tabela manual sem valores — digite ao menos um valor.");
  } else {
    PARSED.prev = dados;
    setStatus("statuspo", true, "Previsão preenchida manualmente — " + dados.itensCompletos.filter((i) => i.tipo === "item").length + " itens" + (dados.totalDespesasPrevisto !== null && dados.totalDespesasPrevisto !== undefined ? ", total " + fmtBRL(dados.totalDespesasPrevisto, true) : ""));
  }
  if (typeof refreshPreview === "function") refreshPreview();
}

function poAttr(v) { return String(v === null || v === undefined ? "" : v).replace(/&/g, "&amp;").replace(/"/g, "&quot;").replace(/</g, "&lt;"); }
function poRenderTabela() {
  const cols = PO_MANUAL.colunas, nC = cols.length, calc = poCalcular();
  const iReal = poIdxPapel("realizado"), iFut = poIdxPapel("futuro");
  const massa = document.getElementById("poMassa");
  if (massa) massa.hidden = iFut < 0; // o ajuste em massa só aparece quando há uma coluna "Futuro"
  const ico = { drag: "&#10303;", dup: "&#10697;", del: "&#128465;", add: "+" };
  // Colunas: ao passar o mouse no cabeçalho aparecem arrastar / duplicar / excluir; o "+" na borda direita adiciona antes/depois.
  let h = '<table class="po-tabela"><tr><th class="po-th-acoes"></th><th>Descrição</th>' + cols.map((n, k) =>
    '<th data-po-th="' + k + '"' + (PO_MANUAL.colMeta[k] && PO_MANUAL.colMeta[k].papel ? ' class="po-papel-' + PO_MANUAL.colMeta[k].papel + '"' : "") + '><div class="po-hover po-hover-col">' +
    '<span class="po-ic po-grip" draggable="true" data-po-drag="' + k + '" title="Arrastar para reordenar a coluna">' + ico.drag + '</span>' +
    '<button type="button" class="po-ic" data-po-dupcol="' + k + '" title="Duplicar coluna">' + ico.dup + '</button>' +
    '<button type="button" class="po-ic po-ic-del" data-po-delcol="' + k + '" title="Excluir coluna">' + ico.del + '</button></div>' +
    '<input type="text" data-po-col="' + k + '" value="' + poAttr(n) + '" placeholder="Coluna">' +
    '<button type="button" class="po-plus po-plus-col" data-po-addcol="' + k + '" title="Adicionar coluna">' + ico.add + '</button></th>').join("") + "</tr>";
  const rot = { grupo: "Grupo", item: "Item", subtotal: "Sub-total", total: "Total", percentual: "%" };
  PO_MANUAL.linhas.forEach((l, i) => {
    // Linhas: arrastar / duplicar / excluir no hover da linha; "+" na divisória de baixo adiciona antes/depois.
    h += '<tr class="po-' + l.tipo + '" data-po-tr="' + i + '"><td class="po-acoes"><span class="po-tipo">' + rot[l.tipo] + '</span><div class="po-hover po-hover-row">' +
      '<span class="po-ic po-grip" draggable="true" data-po-rdrag="' + i + '" title="Arrastar para reordenar a linha">' + ico.drag + '</span>' +
      '<button type="button" class="po-ic" data-po-duprow="' + i + '" title="Duplicar linha">' + ico.dup + '</button>' +
      '<button type="button" class="po-ic po-ic-del" data-po-del="' + i + '" title="Excluir linha">' + ico.del + '</button></div>' +
      '<button type="button" class="po-plus po-plus-row" data-po-addrow="' + i + '" title="Adicionar linha">' + ico.add + '</button></td><td>' +
      '<input type="text" data-po-nome="' + i + '" value="' + poAttr(l.nome) + '" placeholder="' + (l.tipo === "grupo" ? "Nome do grupo" : "Descrição") + '">' +
      (l.tipo === "percentual" ? '<input class="po-pct" type="text" data-po-pct="' + i + '" value="' + poAttr(l.pct) + '" placeholder="%" title="Percentual sobre o total acima">' : "") + "</td>";
    if (l.tipo === "grupo") h += '<td colspan="' + nC + '"></td>';
    else for (let c = 0; c < nC; c++) {
      if (l.tipo === "item") {
        const auto = c === iFut && l.futAuto && iReal >= 0;
        h += '<td><input type="text" data-po-val="' + i + ":" + c + '"' + (auto ? ' data-po-fauto="' + i + '" class="po-auto" title="Sugestão: igual ao Realizado. Edite para ajustar."' : "") + ' value="' + poAttr(auto ? l.vals[iReal] : l.vals[c]) + '" placeholder="0,00"></td>';
      }
      else h += '<td><input type="text" disabled data-po-calc="' + i + ":" + c + '" value="' + poAttr(calc[i] && calc[i][c] !== null ? fmtBRL(calc[i][c]) : "") + '"></td>';
    }
    h += "</tr>";
  });
  h += "</table>";
  const blocoR = (chave, b) => {
    let t = '<div class="po-rateio-bloco"><h4>' + poAttr(b.titulo) + '</h4><label>Unidades <input type="text" data-po-un="' + chave + '" value="' + poAttr(b.unidades) + '" placeholder="Ex: 98"></label>' +
      '<table class="po-tabela"><tr><th>Fração</th><th>Recibo (R$)</th><th></th></tr>';
    b.itens.forEach((it, k) => {
      t += '<tr><td><input type="text" data-po-rn="' + chave + ":" + k + '" value="' + poAttr(it.nome) + '"></td><td><input type="text" data-po-rr="' + chave + ":" + k + '"' + (it.origem ? ' class="po-auto" data-po-rauto="' + chave + ":" + k + '" title="Calculado: valor da tabela ÷ unidades. Edite para digitar outro valor."' : "") + ' value="' + poAttr(it.origem ? (poReciboItem(it, poNumero(b.unidades), calc) === null ? "" : fmtBRL(poReciboItem(it, poNumero(b.unidades), calc))) : it.recibo) + '" placeholder="0,00 ou Conf. Consumo"></td>' +
        '<td><button type="button" class="po-x" data-po-rdel="' + chave + ":" + k + '" title="Excluir">×</button></td></tr>';
    });
    return t + '</table><button type="button" class="btn-sec" data-po-radd="' + chave + '">+ Fração</button></div>';
  };
  h += '<h3 class="po-sub-h">Rateio</h3><p class="hint" style="margin:0 0 6px;">Total do rateio = recibo × unidades. O reajuste é calculado comparando o total do recibo previsto com o atual.</p><div class="po-rateio">' + blocoR("atual", PO_MANUAL.rateio.atual) + blocoR("previsto", PO_MANUAL.rateio.previsto) + "</div>";
  document.getElementById("poManualTabela").innerHTML = h;
}
// Ajuste em massa (opcional): multiplica TODOS os "Futuro" numéricos por (1 + p/100). Cada valor continua editável depois.
function poAjusteMassa(p) {
  const iFut = poIdxPapel("futuro"), iReal = poIdxPapel("realizado");
  if (iFut < 0) return 0;
  let n = 0;
  PO_MANUAL.linhas.forEach((l) => {
    if (l.tipo !== "item") return;
    const atual = poNumero((l.futAuto && iReal >= 0) ? l.vals[iReal] : l.vals[iFut]);
    if (atual === null) return;
    const novo = Math.round(atual * (1 + p / 100) * 100) / 100;
    while (l.vals.length <= iFut) l.vals.push("");
    l.vals[iFut] = novo.toLocaleString("pt-BR", { minimumFractionDigits: 2, maximumFractionDigits: 2 });
    l.futAuto = false;
    n++;
  });
  PO_MANUAL.usado = true;
  poRenderTabela(); poAtualizarDados();
  return n;
}
// atualiza só os campos calculados (sem recriar a tabela e perder o foco enquanto digita)
function poAtualizarCalculados() {
  const calc = poCalcular();
  document.querySelectorAll("#poManualTabela [data-po-calc]").forEach((inp) => {
    const [i, c] = inp.dataset.poCalc.split(":").map(Number);
    inp.value = calc[i] && calc[i][c] !== null ? fmtBRL(calc[i][c]) : "";
  });
  const iReal = poIdxPapel("realizado");
  document.querySelectorAll("#poManualTabela [data-po-fauto]").forEach((inp) => { // Futuro sugerido acompanha o Realizado
    const l = PO_MANUAL.linhas[+inp.dataset.poFauto];
    if (l && l.futAuto && iReal >= 0) inp.value = l.vals[iReal] || "";
  });
  document.querySelectorAll("#poManualTabela [data-po-rauto]").forEach((inp) => { // recibos vinculados do rateio
    const [bl, k] = inp.dataset.poRauto.split(":");
    const b = PO_MANUAL.rateio[bl], it = b && b.itens[+k];
    if (it && it.origem) { const r = poReciboItem(it, poNumero(b.unidades), calc); inp.value = r === null ? "" : fmtBRL(r); }
  });
}
// Popover "adicionar antes ou depois" (linha: também escolhe o tipo; coluna: só a posição).
function poAbrirPopover(x, y, titulo, comTipo, aoEscolher) {
  const velho = document.getElementById("poPop");
  if (velho) velho.remove();
  const pop = document.createElement("div");
  pop.id = "poPop";
  pop.className = "po-pop";
  pop.innerHTML = '<div class="po-pop-t">' + poAttr(titulo) + "</div>" +
    (comTipo ? '<label>Tipo da linha <select id="poPopTipo"><option value="item">Item (valores digitados)</option><option value="grupo">Grupo (título do slide)</option><option value="subtotal">Sub-total (soma o grupo)</option><option value="total">Total (soma subtotais)</option><option value="percentual">Percentual (% do total acima)</option></select></label>' : "") +
    '<div class="po-pop-b"><button type="button" data-pos="antes">Adicionar antes</button><button type="button" data-pos="depois">Adicionar depois</button><button type="button" data-pos="x" class="po-pop-x">Cancelar</button></div>';
  document.body.appendChild(pop);
  const w = pop.offsetWidth;
  pop.style.left = Math.max(8, Math.min(x, window.innerWidth - w - 8)) + "px";
  pop.style.top = Math.max(8, Math.min(y, window.innerHeight - pop.offsetHeight - 8)) + "px";
  pop.addEventListener("click", (e) => {
    const bt = e.target.closest("button");
    if (!bt) return;
    const tipo = comTipo ? document.getElementById("poPopTipo").value : null;
    pop.remove();
    if (bt.dataset.pos !== "x") aoEscolher(bt.dataset.pos, tipo);
  });
  setTimeout(() => document.addEventListener("mousedown", function fora(ev) { if (!pop.contains(ev.target)) { pop.remove(); document.removeEventListener("mousedown", fora); } }), 0);
}
function poResetarModelo() {
  PO_MANUAL.colunas = ["Atual", "Previsto"]; PO_MANUAL.colMeta = [{}, {}]; PO_MANUAL.linhas = poLinhasModelo();
  PO_MANUAL.rateio = { atual: poRateioModelo("Rateio atual"), previsto: poRateioModelo("Previsão (novo rateio)") };
  PO_MANUAL.usado = false; PO_MANUAL.importado = false; poRenderTabela(); poAtualizarDados();
}
function poNovaLinha(tipo) {
  return { tipo, nome: tipo === "subtotal" ? "Sub-Total" : tipo === "total" ? "Total" : "", vals: [], pct: tipo === "percentual" ? "" : undefined };
}
function poMover(arr, de, para) { const [x] = arr.splice(de, 1); arr.splice(para, 0, x); }

(function setupPOManual() {
  const wrap = document.getElementById("poManualTabela");
  if (!wrap) return;
  const bind = (id, fn) => document.getElementById(id).addEventListener("click", fn);
  bind("poModelo", () => {
    if (!confirm("Voltar à estrutura do modelo? O que foi digitado na tabela manual será perdido.")) return;
    poResetarModelo();
  });
  document.getElementById("poMassaBtn").addEventListener("click", () => {
    const bruto = document.getElementById("poMassaPct").value.trim().replace("%", "").replace(",", ".");
    const p = parseFloat(bruto);
    const res = document.getElementById("poMassaRes");
    if (isNaN(p)) { res.textContent = "Digite um percentual, por exemplo +10 ou -5."; return; }
    const n = poAjusteMassa(p);
    res.textContent = n ? "Aplicado em " + n + " valores de Futuro (" + (p > 0 ? "+" : "") + String(p).replace(".", ",") + "%). Você ainda pode editar cada linha." : "Não há valores numéricos de Futuro para ajustar.";
  });
  wrap.addEventListener("input", (e) => {
    const t = e.target, d = t.dataset;
    PO_MANUAL.usado = true;
    if (d.poNome !== undefined) PO_MANUAL.linhas[+d.poNome].nome = t.value;
    else if (d.poPct !== undefined) { PO_MANUAL.linhas[+d.poPct].pct = t.value; poAtualizarCalculados(); }
    else if (d.poVal !== undefined) { const [i, c] = d.poVal.split(":").map(Number); PO_MANUAL.linhas[i].vals[c] = t.value; if (c === poIdxPapel("futuro") && PO_MANUAL.linhas[i].futAuto) { PO_MANUAL.linhas[i].futAuto = false; t.classList.remove("po-auto"); } poAtualizarCalculados(); }
    else if (d.poCol !== undefined) PO_MANUAL.colunas[+d.poCol] = t.value;
    else if (d.poUn !== undefined) { PO_MANUAL.rateio[d.poUn].unidades = t.value; poAtualizarCalculados(); }
    else if (d.poRn !== undefined) { const [b, k] = d.poRn.split(":"); PO_MANUAL.rateio[b].itens[+k].nome = t.value; }
    else if (d.poRr !== undefined) { const [b, k] = d.poRr.split(":"); const it = PO_MANUAL.rateio[b].itens[+k]; it.recibo = t.value; if (it.origem) { delete it.origem; t.classList.remove("po-auto"); } poAtualizarCalculados(); }
    else return;
    poAtualizarDados();
  });
  wrap.addEventListener("click", (e) => {
    const b = e.target.closest("button");
    if (!b) return;
    const d = b.dataset;
    const r = b.getBoundingClientRect();
    if (d.poAddrow !== undefined) {
      const i = +d.poAddrow;
      poAbrirPopover(r.left, r.bottom + 6, "Adicionar linha em relação a \"" + (PO_MANUAL.linhas[i].nome || "esta linha") + "\"", true, (pos, tipo) => {
        PO_MANUAL.linhas.splice(pos === "antes" ? i : i + 1, 0, poNovaLinha(tipo));
        PO_MANUAL.usado = true; poRenderTabela(); poAtualizarDados();
      });
      return;
    }
    if (d.poAddcol !== undefined) {
      const k = +d.poAddcol;
      poAbrirPopover(r.left, r.bottom + 6, "Adicionar coluna em relação a \"" + (PO_MANUAL.colunas[k] || "esta coluna") + "\"", false, (pos) => {
        const at = pos === "antes" ? k : k + 1;
        PO_MANUAL.colunas.splice(at, 0, "Coluna " + (PO_MANUAL.colunas.length + 1)); PO_MANUAL.colMeta.splice(at, 0, {});
        PO_MANUAL.linhas.forEach((l) => { while (l.vals.length < at) l.vals.push(""); l.vals.splice(at, 0, ""); });
        PO_MANUAL.usado = true; poRenderTabela(); poAtualizarDados();
      });
      return;
    }
    if (d.poDel !== undefined) PO_MANUAL.linhas.splice(+d.poDel, 1);
    else if (d.poDuprow !== undefined) {
      const i = +d.poDuprow, o = PO_MANUAL.linhas[i];
      PO_MANUAL.linhas.splice(i + 1, 0, { tipo: o.tipo, nome: o.nome, vals: o.vals.slice(), pct: o.pct, futAuto: o.futAuto, modos: o.modos ? o.modos.slice() : undefined, fixos: o.fixos ? o.fixos.slice() : undefined, baseIds: o.baseIds ? o.baseIds.slice() : undefined });
    } else if (d.poDupcol !== undefined) {
      const k = +d.poDupcol;
      PO_MANUAL.colunas.splice(k + 1, 0, (PO_MANUAL.colunas[k] || "Coluna") + " (cópia)"); PO_MANUAL.colMeta.splice(k + 1, 0, {});
      PO_MANUAL.linhas.forEach((l) => { while (l.vals.length < k + 1) l.vals.push(""); l.vals.splice(k + 1, 0, l.vals[k]); });
    } else if (d.poDelcol !== undefined) {
      if (PO_MANUAL.colunas.length <= 1) { alert("A tabela precisa de ao menos uma coluna de valores."); return; }
      const k = +d.poDelcol;
      PO_MANUAL.colunas.splice(k, 1); PO_MANUAL.colMeta.splice(k, 1);
      PO_MANUAL.linhas.forEach((l) => l.vals.splice(k, 1));
    } else if (d.poRdel !== undefined) { const [bl, k] = d.poRdel.split(":"); PO_MANUAL.rateio[bl].itens.splice(+k, 1); }
    else if (d.poRadd !== undefined) PO_MANUAL.rateio[d.poRadd].itens.push({ nome: "", recibo: "" });
    else return;
    PO_MANUAL.usado = true;
    poRenderTabela(); poAtualizarDados();
  });
  // reordenar LINHAS arrastando a alça (⠿) de uma linha até outra
  let linhaOrigem = null;
  wrap.addEventListener("dragstart", (e) => {
    const g = e.target.closest("[data-po-rdrag]");
    if (!g) return;
    linhaOrigem = +g.dataset.poRdrag;
    e.dataTransfer.effectAllowed = "move";
    try { e.dataTransfer.setData("text/plain", "linha" + linhaOrigem); } catch (err) { /* ok */ }
  });
  wrap.addEventListener("dragover", (e) => { if (linhaOrigem !== null && e.target.closest("[data-po-tr]")) e.preventDefault(); });
  wrap.addEventListener("drop", (e) => {
    const tr = e.target.closest("[data-po-tr]");
    if (linhaOrigem === null || !tr) return;
    e.preventDefault();
    const destino = +tr.dataset.poTr;
    if (destino !== linhaOrigem) { poMover(PO_MANUAL.linhas, linhaOrigem, destino); PO_MANUAL.usado = true; poRenderTabela(); poAtualizarDados(); }
    linhaOrigem = null;
  });
  // reordenar colunas arrastando a alça (⠿) de um cabeçalho até outro
  let colOrigem = null;
  wrap.addEventListener("dragstart", (e) => {
    const g = e.target.closest("[data-po-drag]");
    if (!g) return;
    colOrigem = +g.dataset.poDrag;
    e.dataTransfer.effectAllowed = "move";
    try { e.dataTransfer.setData("text/plain", String(colOrigem)); } catch (err) { /* ok */ }
  });
  wrap.addEventListener("dragover", (e) => { if (colOrigem !== null && e.target.closest("[data-po-th]")) e.preventDefault(); });
  wrap.addEventListener("drop", (e) => {
    const th = e.target.closest("[data-po-th]");
    if (colOrigem === null || !th) return;
    e.preventDefault();
    const destino = +th.dataset.poTh;
    if (destino !== colOrigem) {
      const mover = (arr) => { const [x] = arr.splice(colOrigem, 1); arr.splice(destino, 0, x); };
      mover(PO_MANUAL.colunas); mover(PO_MANUAL.colMeta);
      PO_MANUAL.linhas.forEach((l) => { while (l.vals.length < PO_MANUAL.colunas.length) l.vals.push(""); mover(l.vals); });
      PO_MANUAL.usado = true;
      poRenderTabela(); poAtualizarDados();
    }
    colOrigem = null;
  });
  wrap.addEventListener("dragend", () => { colOrigem = null; });
  poRenderTabela();
})();

// Em vez de uma tabela larga (1 categoria x 12 meses lado a lado, ilegível no celular),
// mostra um seletor de mês + lista vertical de 2 colunas (Categoria / Valor) pro mês escolhido.
function renderPreview11A(data) {
  const months = data.months;
  const categorias = data.categoriasDetalhadas || [];
  const receitas = categorias.filter((c) => c.secao === "receitas");
  const despesas = categorias.filter((c) => c.secao === "despesas");
  const defaultIdx = months.length - 1;

  let html = '<label for="mesSelect11a" style="margin-top:6px;">Ver mês</label>';
  html += '<select id="mesSelect11a">';
  months.forEach((m, i) => (html += '<option value="' + i + '"' + (i === defaultIdx ? " selected" : "") + ">" + m + "</option>"));
  html += "</select>";
  html += '<div id="mesTabelaWrap11a"></div>';
  document.getElementById("preview11a").innerHTML = html;

  function renderMes(idx) {
    let t = '<table class="preview-table" style="min-width:auto;"><tr><th>—</th><th>Valor</th></tr>';
    function row(label, arr, bold) {
      t += "<tr><td" + (bold ? ' style="font-weight:700;"' : "") + ">" + label + "</td><td" + (bold ? ' style="font-weight:700;"' : "") + ">" + fmtBRL(arr[idx]) + "</td></tr>";
    }
    row("Saldo Anterior", data.saldoAnterior);
    row("Total Receitas", data.totalReceitas, true);
    row("Total Despesas", data.totalDespesas, true);
    row("Saldo Final", data.saldoFinal, true);
    if (receitas.length) {
      t += '<tr><td colspan="2" style="font-weight:700;padding-top:12px;color:var(--button);">Receitas</td></tr>';
      receitas.forEach((c) => (t += "<tr><td>" + c.label + "</td><td>" + fmtBRL(c.valores[idx]) + "</td></tr>"));
    }
    if (despesas.length) {
      t += '<tr><td colspan="2" style="font-weight:700;padding-top:12px;color:var(--button);">Despesas</td></tr>';
      despesas.forEach((c) => (t += "<tr><td>" + c.label + "</td><td>" + fmtBRL(c.valores[idx]) + "</td></tr>"));
    }
    t += "</table>";
    document.getElementById("mesTabelaWrap11a").innerHTML = t;
  }
  renderMes(defaultIdx);
  document.getElementById("mesSelect11a").addEventListener("change", (e) => renderMes(parseInt(e.target.value, 10)));
}

function renderPreviewPO(data) {
  let html = '<table class="preview-table"><tr><th>Categoria</th><th>Atual</th><th>Previsto</th></tr>';
  data.categories.forEach((c) => {
    html += "<tr><td>" + c.nome + "</td><td>" + fmtBRL(c.atual) + "</td><td>" + fmtBRL(c.previsto) + "</td></tr>";
  });
  html += "</table>";
  if (data.rateioAtual && data.rateioPrevisto) {
    html += '<p class="hint">Reajuste do rateio: ' + (data.rateioAtual.reajuste !== null && data.rateioAtual.reajuste !== undefined ? (data.rateioAtual.reajuste * 100).toFixed(2) + "%" : "—") + "</p>";
  }

  if (data.itensCompletos && data.itensCompletos.length) {
    html += '<details style="margin-top:10px;"><summary style="cursor:pointer;font-size:12.5px;color:var(--button);font-weight:600;">Ver todos os itens (' + data.itensCompletos.filter((it) => it.tipo === "item").length + ')</summary>';
    html += '<table class="preview-table" style="margin-top:8px;"><tr><th>Descrição</th><th>Atual</th><th>Previsto</th></tr>';
    data.itensCompletos.forEach((it) => {
      if (it.tipo === "grupo") {
        html += '<tr><td colspan="3" style="font-weight:700;background:#f5f3fb;color:var(--button);">' + it.nome + "</td></tr>";
      } else if (it.tipo === "subtotal") {
        html += "<tr><td style=\"font-weight:700;\">" + it.nome + "</td><td style=\"font-weight:700;\">" + formatPOValue(it.atualRaw) + "</td><td style=\"font-weight:700;\">" + formatPOValue(it.previstoRaw) + "</td></tr>";
      } else {
        html += "<tr><td>" + it.nome + "</td><td>" + formatPOValue(it.atualRaw) + "</td><td>" + formatPOValue(it.previstoRaw) + "</td></tr>";
      }
    });
    html += "</table></details>";
  }

  document.getElementById("previewpo").innerHTML = html;
}

function renderPreview16B(data) {
  let html = "";
  if (data.contas.length) {
    html += '<div class="preview-scroll"><table class="preview-table"><tr><th>Conta</th><th>Saldo ant.</th><th>Créditos</th><th>Débitos</th><th>Saldo final</th></tr>';
    data.contas.forEach((c) => {
      html += "<tr><td>" + c.nome + "</td><td>" + fmtBRL(c.saldoAnt) + "</td><td>" + fmtBRL(c.creditos) + "</td><td>" + fmtBRL(c.debitos) + "</td><td>" + fmtBRL(c.saldoFinal) + "</td></tr>";
    });
    html += "</table></div>";

    data.contas.forEach((c) => {
      if (!c.movimentos || !c.movimentos.length) return;
      html += '<details style="margin-top:8px;"><summary style="cursor:pointer;font-size:12.5px;color:var(--button);font-weight:600;">Ver lançamentos de "' + c.nome + '" (' + c.movimentos.length + ')</summary>';
      html += '<div class="preview-scroll"><table class="preview-table"><tr><th>Data</th><th>Descrição</th><th>Tipo</th><th>Valor</th></tr>';
      c.movimentos.forEach((m) => {
        html += "<tr><td>" + m.data + "</td><td style=\"text-align:left;white-space:normal;\">" + m.descricao + "</td><td>" + (m.tipo === "C" ? "Crédito" : "Débito") + "</td><td>" + fmtBRL(m.valor) + "</td></tr>";
      });
      html += "</table></div></details>";
    });
  }
  if (data.grupos.length) {
    html += '<div class="preview-scroll"><table class="preview-table"><tr><th>Grupo de saldo</th><th>Saldo ant.</th><th>Créditos</th><th>Débitos</th><th>Saldo final</th></tr>';
    data.grupos.forEach((c) => {
      html += "<tr><td>" + c.nome + "</td><td>" + fmtBRL(c.saldoAnt) + "</td><td>" + fmtBRL(c.creditos) + "</td><td>" + fmtBRL(c.debitos) + "</td><td>" + fmtBRL(c.saldoFinal) + "</td></tr>";
    });
    html += "</table></div>";
  }
  document.getElementById("preview16b").innerHTML = html;
}

/* ---------- Pauta: lista dinâmica ---------- */
// role: "prestacao" | "previsao" | "assuntos" | "topico".
// "topico" (também o que o usuário adiciona manualmente) ganha uma seção própria na sidebar,
// ligada ao item da pauta por um topicId único.
const DEFAULT_PAUTA = [
  { text: "Prestação de contas da administração", role: "prestacao" },
  { text: "Eleição de síndico, subsíndico e conselho fiscal", role: "topico" },
  { text: "Previsão orçamentária", role: "previsao" },
  { text: "Assuntos gerais (exclusivo para discussão, sem caráter deliberativo)", role: "assuntos" },
];

// Abas fixas que dependem de um item da pauta. Se o item sai da pauta, as abas ficam indisponíveis
// (e o backend não gera os slides, porque só itera sobre a pauta recebida).
const STRUCTURAL_TABS = {
  prestacao: { nome: "Prestação de contas", panels: ["2", "4", "5"] },
  previsao: { nome: "Previsão orçamentária", panels: ["3"] },
};

let pautaDragSrc = null;
let topicSeq = 0;

function novoTopicId() {
  topicSeq += 1;
  return "t" + topicSeq;
}

/* ---------- Formulário do tópico "Eleição" ---------- */
// Os campos montam o texto do slide (uma linha por informação preenchida) dentro do textarea oculto
// f_topic_<id>, que é o que o resto do app já lê; as observações livres entram no fim.
function tituloEhEleicao(label) { return /elei[çc][ãa]o/i.test(label || ""); }
const ELEICAO_VOTACAO = ["A definir na assembleia", "Aclamação", "Votação aberta (voto declarado)", "Cédula (voto secreto)"];
function montarFormEleicao(section, topicId) {
  const ta = section.querySelector("#f_topic_" + topicId);
  if (!ta || section.querySelector(".el-form")) return;
  ta.style.display = "none";
  const campo = (id, rotulo, tipo, ph) => '<div class="el-campo"><label>' + rotulo + "</label>" +
    (tipo === "area" ? '<textarea rows="3" data-el="' + id + '" placeholder="' + ph + '"></textarea>' : '<input type="text" data-el="' + id + '" placeholder="' + ph + '">') + "</div>";
  const form = document.createElement("div");
  form.className = "el-form";
  form.innerHTML =
    '<h3 class="po-sub-h">Candidatos inscritos</h3>' +
    campo("candSindico", "Síndico", "area", "Um nome por linha") +
    campo("candSubsindico", "Subsíndico", "area", "Um nome por linha") +
    campo("candConselho", "Conselho fiscal", "area", "Um nome por linha") +
    '<h3 class="po-sub-h">Votação</h3>' +
    '<div class="el-campo"><label>Forma de votação</label><select data-el="votacao">' + ELEICAO_VOTACAO.map((o) => "<option>" + o + "</option>").join("") + "</select></div>" +
    '<h3 class="po-sub-h">Resultado</h3>' +
    campo("eleitoSindico", "Síndico eleito", "texto", "Nome") +
    campo("eleitoSubsindico", "Subsíndico eleito", "texto", "Nome") +
    campo("eleitosConselho", "Conselho fiscal eleito", "area", "Um nome por linha") +
    '<h3 class="po-sub-h">Mandato</h3>' +
    campo("mandato", "Mandato", "texto", "Ex: 2 anos, de 11/2026 a 11/2028") +
    '<h3 class="po-sub-h">Observações adicionais</h3>' +
    campo("obs", "Texto livre (uma linha por item do slide)", "area", "Qualquer informação extra");
  ta.insertAdjacentElement("afterend", form);
  form.querySelector('[data-el="mandato"]').value = "Até 2 anos, permitida a reeleição";
  const v = (id) => { const e = form.querySelector('[data-el="' + id + '"]'); return e ? e.value.trim() : ""; };
  const lista = (id) => v(id).split("\n").map((x) => x.trim()).filter(Boolean).join("; ");
  const compor = () => {
    const L = [];
    const add = (rot, val) => { if (val) L.push(rot + ": " + val); };
    add("Candidatos a síndico", lista("candSindico"));
    add("Candidatos a subsíndico", lista("candSubsindico"));
    add("Candidatos ao conselho fiscal", lista("candConselho"));
    if (v("votacao") && v("votacao") !== ELEICAO_VOTACAO[0]) add("Forma de votação", v("votacao"));
    else if (L.length) add("Forma de votação", ELEICAO_VOTACAO[0]);
    add("Síndico eleito", v("eleitoSindico"));
    add("Subsíndico eleito", v("eleitoSubsindico"));
    add("Conselho fiscal eleito", lista("eleitosConselho"));
    if (L.length || v("obs")) add("Mandato", v("mandato"));
    v("obs").split("\n").map((x) => x.trim()).filter(Boolean).forEach((x) => L.push(x));
    ta.value = L.join("\n");
  };
  form.addEventListener("input", compor);
  form.addEventListener("change", compor);
  compor();
}

/* ---------- Seções dinâmicas dos tópicos (aba na sidebar + painel com textarea e anexos) ---------- */
const TOPIC_ATTACHMENTS = {};
let topicAnexoSeq = 0;

function renderListaAnexosTopico(topicId) {
  const listEl = document.querySelector("#panel-topic-" + topicId + " .topic-anexos-list");
  if (!listEl) return;
  const itens = TOPIC_ATTACHMENTS[topicId] || [];
  if (!itens.length) { listEl.innerHTML = ""; return; }
  listEl.innerHTML = itens.map((anx, idx) =>
    '<div class="topic-anexo-item">' +
    '<img class="topic-anexo-thumb" src="' + anx.dataUrl + '" alt="">' +
    '<span class="topic-anexo-nome" title="' + poAttr(anx.name) + '">' + (idx + 1) + ". " + pvEsc(anx.name) + "</span>" +
    '<div class="topic-anexo-acoes">' +
    '<button type="button" data-anx-up="' + idx + '"' + (idx === 0 ? " disabled" : "") + ' title="Mover para cima">↑</button>' +
    '<button type="button" data-anx-dn="' + idx + '"' + (idx === itens.length - 1 ? " disabled" : "") + ' title="Mover para baixo">↓</button>' +
    '<button type="button" class="del" data-anx-rm="' + idx + '" title="Remover anexo">×</button>' +
    "</div></div>"
  ).join("");
}

function redimensionarImagemDataUrl(file) {
  return new Promise((resolve, reject) => {
    const reader = new FileReader();
    reader.onload = (ev) => {
      const img = new Image();
      img.onload = () => {
        const maxW = 1600;
        const escala = img.width > maxW ? maxW / img.width : 1;
        const w = Math.round(img.width * escala);
        const h = Math.round(img.height * escala);
        const canvas = document.createElement("canvas");
        canvas.width = w;
        canvas.height = h;
        canvas.getContext("2d").drawImage(img, 0, 0, w, h);
        resolve({ dataUrl: canvas.toDataURL("image/jpeg", 0.86), width: w, height: h });
      };
      img.onerror = () => reject(new Error("Imagem inválida: " + file.name));
      img.src = ev.target.result;
    };
    reader.onerror = () => reject(new Error("Falha ao ler " + file.name));
    reader.readAsDataURL(file);
  });
}

async function processarArquivosTopico(topicId, fileList, statusEl) {
  const files = Array.from(fileList || []);
  if (!files.length) return;
  TOPIC_ATTACHMENTS[topicId] = TOPIC_ATTACHMENTS[topicId] || [];
  statusEl.className = "status-line ok";
  statusEl.innerHTML = '<span class="dot"></span>Processando arquivo(s)...';
  try {
    for (const file of files) {
      if (/\.pdf$/i.test(file.name)) {
        await carregarPdfJs();
        const buf = await file.arrayBuffer();
        const pdf = await window.pdfjsLib.getDocument({ data: buf }).promise;
        const totalPags = Math.min(pdf.numPages, 20);
        for (let p = 1; p <= totalPags; p++) {
          const page = await pdf.getPage(p);
          const vp1 = page.getViewport({ scale: 1 });
          const escala = Math.min(2, 1600 / vp1.width);
          const vp = page.getViewport({ scale: escala });
          const canvas = document.createElement("canvas");
          canvas.width = Math.round(vp.width);
          canvas.height = Math.round(vp.height);
          await page.render({ canvasContext: canvas.getContext("2d"), viewport: vp }).promise;
          topicAnexoSeq += 1;
          TOPIC_ATTACHMENTS[topicId].push({
            id: "anx" + topicAnexoSeq,
            name: file.name + (pdf.numPages > 1 ? " — Pág. " + p + "/" + pdf.numPages : ""),
            dataUrl: canvas.toDataURL("image/jpeg", 0.86),
            width: canvas.width,
            height: canvas.height,
          });
        }
      } else if (file.type.startsWith("image/") || /\.(png|jpe?g|webp|gif)$/i.test(file.name)) {
        const res = await redimensionarImagemDataUrl(file);
        topicAnexoSeq += 1;
        TOPIC_ATTACHMENTS[topicId].push({
          id: "anx" + topicAnexoSeq,
          name: file.name,
          dataUrl: res.dataUrl,
          width: res.width,
          height: res.height,
        });
      } else {
        throw new Error("Formato não suportado: " + file.name + ". Use imagens (.jpg, .png) ou .pdf.");
      }
    }
    const qtd = TOPIC_ATTACHMENTS[topicId].length;
    statusEl.className = "status-line ok";
    statusEl.innerHTML = '<span class="dot"></span>' + qtd + " slide(s) anexado(s) nesta pauta.";
    renderListaAnexosTopico(topicId);
    if (typeof refreshPreview === "function") refreshPreview();
  } catch (err) {
    statusEl.className = "status-line err";
    statusEl.innerHTML = '<span class="dot"></span>' + err.message;
  }
}

function criarSecaoTopico(topicId, label) {
  const nome = (label || "").trim() || "Novo tópico";

  const btn = document.createElement("button");
  btn.type = "button";
  btn.className = "nav-btn";
  btn.dataset.panel = "topic-" + topicId;
  btn.dataset.label = nome;
  btn.innerHTML = '<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="1.8"><path d="M4 4h16v13H7l-3 3V4Z"></path><path d="M8 9h8M8 13h5"></path></svg>';
  document.getElementById("dynamicTopicsNav").appendChild(btn);

  const section = document.createElement("section");
  section.className = "panel";
  section.id = "panel-topic-" + topicId;
  section.dataset.panel = "topic-" + topicId;
  section.innerHTML =
    '<div class="card">' +
    '<h2 class="topic-title" style="margin-left:0;"></h2>' +
    '<p class="sub" style="margin-left:0;">Opcional — um item por linha. Deixe em branco para gerar só o slide divisor.</p>' +
    '<div class="body" style="margin-left:0;">' +
    '<textarea id="f_topic_' + topicId + '" rows="4" placeholder="Ex: um item por linha"></textarea>' +
    '<div class="topic-anexos-box" style="margin-top:16px;">' +
    '<label>Imagens ou PDF para espelhar nos slides (opcional)</label>' +
    '<p class="hint" style="margin:0 0 8px;">Suba fotos (ex: melhorias feitas no condomínio) ou um arquivo PDF (ex: orçamento). Cada imagem ou página do PDF vira um slide nesta pauta.</p>' +
    '<div class="dropzone topic-dz"><input type="file" accept="image/*,.pdf" multiple>Arraste imagens (.jpg, .png) ou PDF aqui, ou clique para selecionar</div>' +
    '<div class="status-line topic-anexo-status"></div>' +
    '<div class="topic-anexos-list"></div>' +
    '</div>' +
    '</div></div>';
  section.querySelector(".topic-title").textContent = nome;

  const dz = section.querySelector(".topic-dz");
  const inp = dz.querySelector("input");
  const stEl = section.querySelector(".topic-anexo-status");
  const listEl = section.querySelector(".topic-anexos-list");
  dz.addEventListener("click", () => inp.click());
  dz.addEventListener("dragover", (e) => { e.preventDefault(); dz.classList.add("drag"); });
  dz.addEventListener("dragleave", () => dz.classList.remove("drag"));
  dz.addEventListener("drop", (e) => {
    e.preventDefault();
    dz.classList.remove("drag");
    if (e.dataTransfer.files.length) processarArquivosTopico(topicId, e.dataTransfer.files, stEl);
  });
  inp.addEventListener("change", () => {
    if (inp.files.length) {
      processarArquivosTopico(topicId, inp.files, stEl);
      inp.value = "";
    }
  });
  listEl.addEventListener("click", (e) => {
    const b = e.target.closest("button");
    if (!b) return;
    const arr = TOPIC_ATTACHMENTS[topicId] || [];
    if (b.dataset.anxRm !== undefined) arr.splice(+b.dataset.anxRm, 1);
    else if (b.dataset.anxUp !== undefined) { const i = +b.dataset.anxUp; if (i > 0) [arr[i - 1], arr[i]] = [arr[i], arr[i - 1]]; }
    else if (b.dataset.anxDn !== undefined) { const i = +b.dataset.anxDn; if (i < arr.length - 1) [arr[i + 1], arr[i]] = [arr[i], arr[i + 1]]; }
    else return;
    stEl.className = arr.length ? "status-line ok" : "status-line";
    stEl.innerHTML = arr.length ? '<span class="dot"></span>' + arr.length + " slide(s) anexado(s) nesta pauta." : "";
    renderListaAnexosTopico(topicId);
    if (typeof refreshPreview === "function") refreshPreview();
  });

  document.getElementById("dynamicTopicsPanels").appendChild(section);
  if (tituloEhEleicao(nome)) montarFormEleicao(section, topicId);
}

function atualizarRotuloTopico(topicId, label) {
  const nome = (label || "").trim() || "Novo tópico";
  const btn = document.querySelector('.sidebar nav button[data-panel="topic-' + topicId + '"]');
  if (btn) btn.dataset.label = nome;
  const titulo = document.querySelector("#panel-topic-" + topicId + " .topic-title");
  if (titulo) titulo.textContent = nome;
  const ta = document.getElementById("f_topic_" + topicId);
  if (ta && !ta.value.trim() && tituloEhEleicao(nome)) montarFormEleicao(document.getElementById("panel-topic-" + topicId), topicId);
}

function removerSecaoTopico(topicId) {
  delete TOPIC_ATTACHMENTS[topicId];
  const btn = document.querySelector('.sidebar nav button[data-panel="topic-' + topicId + '"]');
  const panel = document.getElementById("panel-topic-" + topicId);
  const eraAtiva = (btn && btn.classList.contains("active")) || (panel && panel.classList.contains("active"));
  if (btn) btn.remove();
  if (panel) panel.remove();
  if (eraAtiva) switchPanel("1");
}

/* ---------- Abas estruturais: habilitadas só enquanto o item correspondente está na pauta ---------- */
/* ---------- Cópias independentes de Prestação de Contas e Previsão Orçamentária ---------- */
const EXTRA_SECTIONS = {};
let extraSeq = 0;

function removerSecaoExtra(extraId) {
  delete EXTRA_SECTIONS[extraId];
  const btn = document.querySelector('.sidebar nav button[data-panel="extra-' + extraId + '"]');
  const panel = document.getElementById("panel-extra-" + extraId);
  const eraAtiva = (btn && btn.classList.contains("active")) || (panel && panel.classList.contains("active"));
  if (btn) btn.remove();
  if (panel) panel.remove();
  const pautaDiv = document.querySelector('#pautaList .repeat-block[data-extra-id="' + extraId + '"]');
  if (pautaDiv) pautaDiv.remove();
  updatePautaHint();
  atualizarAbasEstruturais();
  if (eraAtiva) switchPanel("1");
  if (typeof refreshPreview === "function") refreshPreview();
}

function criarSecaoExtra(role, tituloInicial) {
  extraSeq += 1;
  const extraId = "ext" + extraSeq;
  const nCopia = Object.values(EXTRA_SECTIONS).filter((s) => s.role === role).length + 2;
  const titulo = (tituloInicial || "").trim() || (role === "prestacao" ? "Prestação de contas (" + nCopia + ")" : "Previsão orçamentária (" + nCopia + ")");
  EXTRA_SECTIONS[extraId] = { id: extraId, role, a11: null, b16: null, periodo16b: "", inadimplencia: null, prev: null, unidadesTotal: "" };

  const navWrap = document.getElementById("dynamicExtraNav");
  const panelsWrap = document.getElementById("dynamicExtraPanels");
  if (navWrap && panelsWrap) {
    const btn = document.createElement("button");
    btn.type = "button";
    btn.className = "nav-btn";
    btn.dataset.panel = "extra-" + extraId;
    btn.dataset.label = titulo;
    btn.innerHTML = role === "prestacao"
      ? '<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="1.8"><path d="M3 21V9M9 21V3M15 21v-7M21 21V11"></path></svg>'
      : '<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="1.8"><circle cx="12" cy="12" r="9"></circle><path d="M12 7v5l3 3"></path></svg>';
    navWrap.appendChild(btn);

    const section = document.createElement("section");
    section.className = "panel";
    section.id = "panel-extra-" + extraId;
    section.dataset.panel = "extra-" + extraId;

    if (role === "prestacao") {
      const defIni = document.getElementById("f_11a_mes_ini")?.value || "";
      const defFim = document.getElementById("f_11a_mes_fim")?.value || "";
      const temCond = !!document.getElementById("f_condominio_id")?.value;
      const disAttr = temCond ? "" : " disabled";
      section.innerHTML =
        '<div class="card">' +
        '<div class="card-head-row"><div>' +
        '<h2 class="extra-title" style="margin-left:0;"></h2>' +
        '<p class="sub" style="margin-left:0;">Cópia independente — preencha o 11A, 16B e Inadimplência específicos deste período</p>' +
        '</div><button type="button" class="secondary btn-rm-extra">✕ Remover esta cópia</button></div>' +
        '<div class="body" style="margin-left:0;">' +
        '<h3 class="po-sub-h">1. Relatório 11A — Demonstrativo de Receitas e Despesas</h3>' +
        '<div class="row" style="margin-bottom:10px;">' +
        '<div><label>Mês/ano início</label><input type="month" class="ext-11a-ini" value="' + poAttr(defIni) + '"></div>' +
        '<div><label>Mês/ano fim</label><input type="month" class="ext-11a-fim" value="' + poAttr(defFim) + '"></div>' +
        '</div>' +
        '<button type="button" class="primary ext-11a-buscar" style="width:auto;padding:9px 16px;font-size:13px;"' + disAttr + '>Buscar 11A automaticamente</button>' +
        '<div class="status-line ext-status"></div>' +
        '<div class="ext-preview"></div>' +
        '<details style="margin-top:10px;"><summary style="cursor:pointer;font-size:12.5px;color:var(--muted);">Ou enviar arquivo .xlsx do 11A manualmente</summary>' +
        '<div class="dropzone ext-dz" style="margin-top:8px;"><input type="file" accept=".xlsx,.xls">Arraste o arquivo .xlsx do relatório 11A aqui, ou clique para selecionar</div>' +
        '</details>' +
        '</div></div>' +

        '<div class="card">' +
        '<h2 style="margin-left:0;">2. Relatório 16B — Resumo Financeiro (desta cópia)</h2>' +
        '<p class="sub" style="margin-left:0;">Opcional — se ficar vazio, não será repetido nesta prestação de contas</p>' +
        '<div class="body" style="margin-left:0;">' +
        '<div class="row">' +
        '<div><label>Data início</label><input type="date" class="ext-16b-ini"></div>' +
        '<div><label>Data fim</label><input type="date" class="ext-16b-fim"></div>' +
        '</div>' +
        '<button type="button" class="primary ext-16b-buscar" style="width:auto;padding:9px 16px;font-size:13px;margin-top:10px;"' + disAttr + '>Buscar 16B automaticamente</button>' +
        '<div class="status-line ext-16b-status"></div>' +
        '<p class="sub" style="margin:12px 0 6px;">Ou envie o relatório 16B (PDF ou planilha)</p>' +
        '<div class="dropzone ext-16b-dz"><input type="file" accept=".pdf,.xlsx,.xls,.csv">Arraste o relatório 16B (PDF ou planilha) aqui, ou clique para selecionar</div>' +
        '</div></div>' +

        '<div class="card">' +
        '<h2 style="margin-left:0;">3. Inadimplência (desta cópia)</h2>' +
        '<p class="sub" style="margin-left:0;">Opcional — preencha apenas se quiser exibir o slide de inadimplência nesta prestação de contas</p>' +
        '<div class="body" style="margin-left:0;">' +
        '<button type="button" class="primary ext-inad-buscar" style="width:auto;padding:9px 16px;font-size:13px;"' + disAttr + '>Buscar inadimplência automaticamente</button>' +
        '<div class="status-line ext-inad-status"></div>' +
        '<div class="row3" style="margin-top:12px;">' +
        '<div><label>% Inadimplência orgânica</label><input type="text" class="ext-inad-org" placeholder="Ex: 22,00"></div>' +
        '<div><label>% Inadimplência garantidora</label><input type="text" class="ext-inad-gar" placeholder="Ex: 39,00"></div>' +
        '<div><label>% Inadimplência total</label><input type="text" class="ext-inad-tot" placeholder="Ex: 44,00"></div>' +
        '</div>' +
        '<div class="row3">' +
        '<div><label>Valor acumulado (R$)</label><input type="text" class="ext-inad-val" placeholder="Ex: 85.253,35"></div>' +
        '<div><label>Nº de cobranças</label><input type="text" class="ext-inad-cob" placeholder="Ex: 229"></div>' +
        '<div><label>Nº de unidades inadimplentes</label><input type="text" class="ext-inad-un" placeholder="Ex: 12"></div>' +
        '</div>' +
        '<div class="row">' +
        '<div><label>Nº total de unidades do condomínio</label><input type="text" class="ext-inad-untot" placeholder="Ex: 44"></div>' +
        '<div></div>' +
        '</div>' +
        '</div></div>';

      const setStEl = (el, ok, msg) => { el.className = "status-line " + (ok ? "ok" : "err"); el.innerHTML = '<span class="dot"></span>' + msg; };

      // --- 11A da cópia ---
      const st11 = section.querySelector(".ext-status");
      const pv11 = section.querySelector(".ext-preview");
      const renderMini11A = (d) => {
        pv11.innerHTML = '<p class="hint">Período carregado: <strong>' + pvEsc(d.months[0]) + " a " + pvEsc(d.months[d.months.length - 1]) + "</strong> (" + d.months.length + " meses).</p>";
      };

      section.querySelector(".ext-11a-buscar").addEventListener("click", async () => {
        const idCond = document.getElementById("f_condominio_id")?.value;
        if (!idCond) { showErrorModal("Falta selecionar o condomínio", "Selecione um condomínio na aba 'Dados da assembleia' antes de buscar."); return; }
        const vIni = section.querySelector(".ext-11a-ini").value;
        const vFim = section.querySelector(".ext-11a-fim").value;
        if (!vIni || !vFim) { showErrorModal("Falta o período", "Preencha mês/ano inicial e final."); return; }
        const [yI, mI] = vIni.split("-").map(Number);
        const [yF, mF] = vFim.split("-").map(Number);
        if (yI * 12 + mI > yF * 12 + mF) { showErrorModal("Período inválido", "O mês inicial não pode ser maior que o final."); return; }
        const btnB = section.querySelector(".ext-11a-buscar");
        const txtOrig = btnB.textContent;
        btnB.disabled = true;
        btnB.innerHTML = '<span class="spinner" style="border-color:var(--button);border-top-color:transparent;"></span> Buscando...';
        try {
          const periodos = [];
          let m = mI, y = yI;
          while (y * 12 + m <= yF * 12 + mF && periodos.length < 36) { periodos.push({ mes: m, ano: y }); m += 1; if (m === 13) { m = 1; y += 1; } }
          const months = [], totalReceitas = [], totalDespesas = [], energiaEletrica = [], aguaEsgotoDespesa = [], saldoAnterior = [], saldoFinal = [], mesesItens = [];
          let contas = [], saldoOk = true;
          try { contas = await fetchContasBancarias(idCond); if (!contas.length) saldoOk = false; } catch (e) { saldoOk = false; }
          let saldoAntUlt = null;
          for (let i = 0; i < periodos.length; i++) {
            const p = periodos[i];
            const parsed = parseBalanceteMes(await fetchBalanceteMes(idCond, p.mes, p.ano));
            months.push(MESES_ABBR[p.mes - 1] + "/" + p.ano);
            mesesItens.push(parsed ? parsed.itensBrutos : []);
            totalReceitas.push(parsed ? parsed.totalReceitas : null);
            totalDespesas.push(parsed ? parsed.totalDespesas : null);
            energiaEletrica.push(parsed ? parsed.energia : null);
            aguaEsgotoDespesa.push(parsed ? parsed.agua : null);
            if (saldoOk) {
              const sAnt = saldoAntUlt !== null ? saldoAntUlt : await fetchSaldoConsolidado(contas, idCond, pad2(p.mes) + "/01/" + p.ano, 0);
              const sFim = await fetchSaldoConsolidado(contas, idCond, pad2(p.mes) + "/" + pad2(lastDayOfMonth(p.mes, p.ano)) + "/" + p.ano, 1);
              saldoAnterior.push(sAnt); saldoFinal.push(sFim); saldoAntUlt = sFim;
            } else { saldoAnterior.push(null); saldoFinal.push(null); }
            setStEl(st11, true, "Buscando... " + months.length + "/" + periodos.length + " meses");
          }
          let contasPlano = null;
          try { contasPlano = montarPlanoContas(await fetchPlanoContas(idCond), mesesItens); if (!contasPlano.length) contasPlano = null; } catch (e) {}
          const d = {
            months, contasPlano, saldoAnterior, saldoFinal, totalReceitas, totalDespesas,
            energiaEletrica: energiaEletrica.some((v) => v !== null) ? energiaEletrica : null,
            aguaEsgotoDespesa: aguaEsgotoDespesa.some((v) => v !== null) ? aguaEsgotoDespesa : null,
            categoriasDetalhadas: [],
          };
          EXTRA_SECTIONS[extraId].a11 = d;
          setStEl(st11, true, months.length + " meses buscados (" + months[0] + " a " + months[months.length - 1] + ").");
          renderMini11A(d);
          if (typeof refreshPreview === "function") refreshPreview();
        } catch (err) {
          setStEl(st11, false, "Erro ao buscar: " + err.message);
        } finally { btnB.disabled = false; btnB.textContent = txtOrig; }
      });

      const dz11 = section.querySelector(".ext-dz");
      const inp11 = dz11.querySelector("input");
      dz11.addEventListener("click", () => inp11.click());
      inp11.addEventListener("change", async () => {
        if (!inp11.files[0]) return;
        try {
          const wb = await readWorkbook(inp11.files[0]);
          const d = parse11A(wb);
          EXTRA_SECTIONS[extraId].a11 = d;
          dz11.classList.add("ok");
          dz11.textContent = inp11.files[0].name;
          setStEl(st11, true, d.months.length + " meses carregados (" + d.months[0] + " a " + d.months[d.months.length - 1] + ")");
          renderMini11A(d);
          if (typeof refreshPreview === "function") refreshPreview();
        } catch (err) { setStEl(st11, false, "Erro: " + err.message); }
      });

      // --- 16B da cópia ---
      const st16 = section.querySelector(".ext-16b-status");
      section.querySelector(".ext-16b-buscar").addEventListener("click", async () => {
        const idCond = document.getElementById("f_condominio_id")?.value;
        const dtIni = section.querySelector(".ext-16b-ini").value;
        const dtFim = section.querySelector(".ext-16b-fim").value;
        if (!idCond) { showErrorModal("Falta selecionar o condomínio", "Selecione um condomínio na aba 'Dados da assembleia' antes de buscar."); return; }
        if (!dtIni || !dtFim) { showErrorModal("Faltam as datas", "Informe a data de início e fim do 16B."); return; }
        const btnB = section.querySelector(".ext-16b-buscar");
        const txtOrig = btnB.textContent;
        btnB.disabled = true;
        btnB.innerHTML = '<span class="spinner" style="border-color:var(--button);border-top-color:transparent;"></span> Buscando...';
        try {
          let contas = await fetchContasBancarias(idCond);
          const selecionadas = getContasSelecionadasIds();
          if (selecionadas) contas = contas.filter((c) => selecionadas.includes(String(c.id_contabanco_cb)));
          const [anoI, mesI, diaI] = dtIni.split("-");
          const [anoF, mesF, diaF] = dtFim.split("-");
          const dtIniFmt = mesI + "/" + diaI + "/" + anoI;
          const dtFimFmt = mesF + "/" + diaF + "/" + anoF;
          const linhas = [];
          for (const c of contas) {
            let saldoAnt = null, saldoFim = null, creditos = null, debitos = null, movimentos = [];
            try { saldoAnt = await fetchSaldoConta(c.id_contabanco_cb, idCond, dtIniFmt, 0); } catch (e) {}
            try { saldoFim = await fetchSaldoConta(c.id_contabanco_cb, idCond, dtFimFmt, 1); } catch (e) {}
            try {
              const mov = await fetchMovimentacoesConta(c.id_contabanco_cb, idCond, dtIniFmt, dtFimFmt);
              creditos = mov.creditos; debitos = mov.debitos; movimentos = mov.movimentos;
            } catch (e) {}
            linhas.push({ nome: c.st_descricao_cb || c.st_conta_cb || ("Conta " + c.id_contabanco_cb), saldoAnt, creditos, debitos, saldoFinal: saldoFim, movimentos });
          }
          EXTRA_SECTIONS[extraId].b16 = { contas: linhas, grupos: [], saldoFinalGeral: null };
          EXTRA_SECTIONS[extraId].periodo16b = "De " + dtIniFmt + " até " + dtFimFmt;
          setStEl(st16, true, linhas.length + " conta(s) carregadas (" + dtIniFmt + " a " + dtFimFmt + ").");
          if (typeof refreshPreview === "function") refreshPreview();
        } catch (err) {
          setStEl(st16, false, "Erro ao buscar 16B: " + err.message);
        } finally { btnB.disabled = false; btnB.textContent = txtOrig; }
      });

      const dz16 = section.querySelector(".ext-16b-dz");
      const inp16 = dz16.querySelector("input");
      dz16.addEventListener("click", () => inp16.click());
      inp16.addEventListener("change", async () => {
        if (!inp16.files[0]) return;
        try {
          const texto = await arquivoParaTexto16B(inp16.files[0]);
          const data = parse16B(texto);
          EXTRA_SECTIONS[extraId].b16 = data;
          EXTRA_SECTIONS[extraId].periodo16b = data.periodo || "";
          dz16.classList.add("ok");
          dz16.textContent = inp16.files[0].name;
          setStEl(st16, true, (data.contas.length + data.grupos.length) + " linhas reconhecidas no 16B.");
          if (typeof refreshPreview === "function") refreshPreview();
        } catch (err) { setStEl(st16, false, "Erro no 16B: " + err.message); }
      });

      // --- Inadimplência da cópia ---
      const stInad = section.querySelector(".ext-inad-status");
      const elOrg = section.querySelector(".ext-inad-org");
      const elGar = section.querySelector(".ext-inad-gar");
      const elTot = section.querySelector(".ext-inad-tot");
      const elVal = section.querySelector(".ext-inad-val");
      const elCob = section.querySelector(".ext-inad-cob");
      const elUn = section.querySelector(".ext-inad-un");
      const elUnTot = section.querySelector(".ext-inad-untot");

      const syncInadCopia = () => {
        const pct = pctInadimplencia(elUn.value, elUnTot.value);
        if (pct !== null) elTot.value = pct;
        const org = elOrg.value.trim(), gar = elGar.value.trim(), tot = elTot.value.trim(), val = elVal.value.trim();
        EXTRA_SECTIONS[extraId].inadimplencia = (org || gar || tot || val) ? {
          organica: org, garantidora: gar, total: tot, valor: val,
          cobrancas: elCob.value.trim(), unidades: elUn.value.trim(), unidades_total: elUnTot.value.trim(),
          posicao: dataHojeBR(),
        } : null;
        if (typeof refreshPreview === "function") refreshPreview();
      };
      [elOrg, elGar, elTot, elVal, elCob, elUn, elUnTot].forEach((inp) => inp.addEventListener("input", syncInadCopia));

      section.querySelector(".ext-inad-buscar").addEventListener("click", async () => {
        const idCond = document.getElementById("f_condominio_id")?.value;
        if (!idCond) { showErrorModal("Falta selecionar o condomínio", "Selecione um condomínio na aba 'Dados da assembleia' antes de buscar."); return; }
        const btnB = section.querySelector(".ext-inad-buscar");
        const txtOrig = btnB.textContent;
        btnB.disabled = true;
        btnB.innerHTML = '<span class="spinner" style="border-color:var(--button);border-top-color:transparent;"></span> Buscando...';
        try {
          const hoje = new Date();
          const posicaoEm = pad2(hoje.getMonth() + 1) + "/" + pad2(hoje.getDate()) + "/" + hoje.getFullYear();
          let todasUnidades = [], pagina = 1;
          while (true) {
            const url = BAL_API_BASE + "/api/superlogica/inadimplencia/index?posicaoEm=" + posicaoEm +
              "&comValoresAtualizados=0&comValoresAtualizadosPorComposicao=0&idCondominio=" + idCond +
              "&comDadosDaReceita=1&itensPorPagina=50&pagina=" + pagina + "&semAcordo=0&semProcesso=0";
            const res = await fetch(url);
            if (!res.ok) throw new Error("HTTP " + res.status);
            const data = await res.json();
            if (!Array.isArray(data)) throw new Error((data && data.msg) || "resposta inesperada");
            todasUnidades = todasUnidades.concat(data);
            if (data.length < 50 || ++pagina > 40) break;
          }
          let valorTotal = 0, cobrancasTotal = 0;
          todasUnidades.forEach((u) => {
            const cobs = Array.isArray(u.recebimento) ? u.recebimento : [];
            cobrancasTotal += cobs.length;
            cobs.forEach((c) => { valorTotal += parseFloat(c.vl_total_recb) || 0; });
          });
          elVal.value = valorTotal.toLocaleString("pt-BR", { minimumFractionDigits: 2, maximumFractionDigits: 2 });
          elCob.value = String(cobrancasTotal);
          elUn.value = String(todasUnidades.length);
          let totalUnidades = null;
          try { totalUnidades = await fetchTotalUnidades(idCond); } catch (e) {}
          let temGar = false;
          try { temGar = await fetchCondominioGarantido(idCond); } catch (e) {}
          const garantidas = temGar ? todasUnidades.filter((u) => (Array.isArray(u.recebimento) ? u.recebimento : []).some((c) => cobrancaGarantida(c))).length : 0;
          const organicas = temGar ? todasUnidades.filter((u) => (Array.isArray(u.recebimento) ? u.recebimento : []).some((c) => !cobrancaGarantida(c))).length : todasUnidades.length;
          if (totalUnidades) {
            elUnTot.value = String(totalUnidades);
            const pctDe = (n) => ((n / totalUnidades) * 100).toFixed(2).replace(".", ",");
            elGar.value = temGar ? pctDe(garantidas) : "0,00";
            elOrg.value = pctDe(organicas);
          }
          syncInadCopia();
          setStEl(stInad, true, todasUnidades.length + " unidades inadimplentes carregadas.");
        } catch (err) {
          setStEl(stInad, false, "Erro ao buscar inadimplência: " + err.message);
        } finally { btnB.disabled = false; btnB.textContent = txtOrig; }
      });
    } else {
      section.innerHTML =
        '<div class="card">' +
        '<div class="card-head-row"><div>' +
        '<h2 class="extra-title" style="margin-left:0;"></h2>' +
        '<p class="sub" style="margin-left:0;">Cópia independente da Previsão Orçamentária — envie a planilha .xlsx deste cenário</p>' +
        '</div><button type="button" class="secondary btn-rm-extra">✕ Remover esta cópia</button></div>' +
        '<div class="body" style="margin-left:0;">' +
        '<div class="dropzone ext-dz"><input type="file" accept=".xlsx,.xls">Arraste o arquivo .xlsx desta previsão orçamentária aqui, ou clique para selecionar</div>' +
        '<div class="row" style="margin-top:12px;"><div><label>Total de unidades do condomínio (para média do rateio)</label><input type="number" class="ext-po-un" min="1" step="1" placeholder="Ex: 73"></div><div></div></div>' +
        '<div class="status-line ext-status"></div>' +
        '<div class="ext-preview"></div>' +
        '</div></div>';

      const stEl = section.querySelector(".ext-status");
      const pvEl = section.querySelector(".ext-preview");
      const unEl = section.querySelector(".ext-po-un");
      const setSt = (ok, msg) => { stEl.className = "status-line " + (ok ? "ok" : "err"); stEl.innerHTML = '<span class="dot"></span>' + msg; };
      unEl.addEventListener("input", () => {
        EXTRA_SECTIONS[extraId].unidadesTotal = unEl.value;
        if (typeof refreshPreview === "function") refreshPreview();
      });

      const dz = section.querySelector(".ext-dz");
      const inp = dz.querySelector("input");
      dz.addEventListener("click", () => inp.click());
      inp.addEventListener("change", async () => {
        if (!inp.files[0]) return;
        try {
          const wb = await readWorkbook(inp.files[0]);
          const d = parsePrevisao(wb);
          EXTRA_SECTIONS[extraId].prev = d;
          const unPlan = (d.rateioPrevisto && d.rateioPrevisto.unidades) || (d.rateioAtual && d.rateioAtual.unidades);
          if (unPlan && !unEl.value) { unEl.value = String(unPlan); EXTRA_SECTIONS[extraId].unidadesTotal = String(unPlan); }
          dz.classList.add("ok");
          dz.textContent = inp.files[0].name;
          setSt(true, 'Aba "' + d.sheetUsed + '" — ' + d.categories.length + " categorias, total previsto " + fmtBRL(d.totalDespesasPrevisto, true));
          pvEl.innerHTML = '<p class="hint">Planilha carregada com sucesso (' + d.categories.length + ' categorias).</p>';
          if (typeof refreshPreview === "function") refreshPreview();
        } catch (err) { setSt(false, "Erro: " + err.message); }
      });
    }

    section.querySelector(".extra-title").textContent = titulo;
    section.querySelector(".btn-rm-extra").addEventListener("click", () => removerSecaoExtra(extraId));
    panelsWrap.appendChild(section);
  }

  return { extraId, titulo };
}

function atualizarRotuloExtra(extraId, label) {
  const nome = (label || "").trim() || "Cópia";
  const btn = document.querySelector('.sidebar nav button[data-panel="extra-' + extraId + '"]');
  if (btn) btn.dataset.label = nome;
  const tit = document.querySelector("#panel-extra-" + extraId + " .extra-title");
  if (tit) tit.textContent = nome;
}

function duplicarSecaoEstrutural(role, tituloBase) {
  const { extraId, titulo } = criarSecaoExtra(role, tituloBase);
  addPautaItem(titulo, role, extraId);
  switchPanel("extra-" + extraId);
  if (typeof refreshPreview === "function") refreshPreview();
}

function atualizarAbasEstruturais() {
  const itens = getPautaItemsComRole();
  const roles = new Set(itens.filter((it) => !it.extraId).map((it) => it.role));
  Object.keys(STRUCTURAL_TABS).forEach((role) => {
    const cfg = STRUCTURAL_TABS[role];
    const disponivel = roles.has(role);
    cfg.panels.forEach((n) => {
      const btn = document.querySelector('.sidebar nav button[data-panel="' + n + '"]');
      const panel = document.getElementById("panel-" + n);
      if (btn) {
        if (!btn.dataset.labelOriginal) btn.dataset.labelOriginal = btn.dataset.label;
        btn.classList.toggle("nav-disabled", !disponivel);
        btn.setAttribute("aria-disabled", String(!disponivel));
        btn.dataset.label = disponivel
          ? btn.dataset.labelOriginal
          : btn.dataset.labelOriginal + " — indisponível porque o item '" + cfg.nome + "' foi removido da pauta";
        if (!disponivel && btn.classList.contains("active")) switchPanel("1");
      }
      if (panel) panel.classList.toggle("panel-disabled", !disponivel);
    });
  });
}

function updatePautaHint() {
  const total = document.querySelectorAll(".pauta-input").length;
  document.getElementById("pautaHint").textContent =
    total + " item(ns) na pauta. Arraste para reordenar; clique em ⧉ para duplicar Prestação de contas ou Previsão orçamentária.";
}

function addPautaItem(value, role, extraId) {
  role = role || "topico";
  const list = document.getElementById("pautaList");
  const div = document.createElement("div");
  div.className = "repeat-block draggable";
  div.draggable = true;
  div.dataset.role = role;
  if (extraId) div.dataset.extraId = extraId;

  const podeDuplicar = role === "prestacao" || role === "previsao";
  div.innerHTML = '<span class="drag-handle" title="Arraste para reordenar">⋮⋮</span>' +
    '<input type="text" class="pauta-input" style="flex:1;">' +
    (podeDuplicar ? '<button type="button" class="secondary dup-pauta" title="Duplicar esta seção (cria uma nova aba independente)">⧉</button>' : "") +
    '<button type="button" class="secondary rm" style="padding:2px 8px;">×</button>';
  const input = div.querySelector(".pauta-input");
  input.value = value || "";

  if (role === "topico") {
    const topicId = novoTopicId();
    div.dataset.topicId = topicId;
    criarSecaoTopico(topicId, input.value);
  }

  input.addEventListener("input", () => {
    if (div.dataset.topicId) atualizarRotuloTopico(div.dataset.topicId, input.value);
    if (div.dataset.extraId) atualizarRotuloExtra(div.dataset.extraId, input.value);
    atualizarAbasEstruturais();
  });
  if (podeDuplicar) {
    div.querySelector(".dup-pauta").addEventListener("click", () => {
      duplicarSecaoEstrutural(role, (input.value.trim() || "Item") + " (2)");
    });
  }
  div.querySelector(".rm").addEventListener("click", () => {
    if (div.dataset.topicId) removerSecaoTopico(div.dataset.topicId);
    if (div.dataset.extraId) removerSecaoExtra(div.dataset.extraId);
    div.remove();
    updatePautaHint();
    atualizarAbasEstruturais();
    if (typeof refreshPreview === "function") refreshPreview();
  });
  div.addEventListener("dragstart", () => { pautaDragSrc = div; div.classList.add("dragging"); });
  div.addEventListener("dragend", () => { div.classList.remove("dragging"); pautaDragSrc = null; });
  div.addEventListener("dragover", (e) => { e.preventDefault(); });
  div.addEventListener("drop", (e) => {
    e.preventDefault();
    if (!pautaDragSrc || pautaDragSrc === div) return;
    const items = Array.from(list.children);
    const srcIdx = items.indexOf(pautaDragSrc);
    const tgtIdx = items.indexOf(div);
    if (srcIdx < tgtIdx) list.insertBefore(pautaDragSrc, div.nextSibling);
    else list.insertBefore(pautaDragSrc, div);
    if (typeof refreshPreview === "function") refreshPreview();
  });
  list.appendChild(div);
  updatePautaHint();
  atualizarAbasEstruturais();
}

function resetPautaList() {
  document.querySelectorAll("#pautaList .repeat-block[data-topic-id]").forEach((div) => removerSecaoTopico(div.dataset.topicId));
  Object.keys(EXTRA_SECTIONS).forEach((id) => removerSecaoExtra(id));
  document.getElementById("pautaList").innerHTML = "";
  DEFAULT_PAUTA.forEach((item) => addPautaItem(item.text, item.role));
}
resetPautaList();
document.getElementById("addPauta").addEventListener("click", () => { addPautaItem("", "topico"); if (typeof refreshPreview === "function") refreshPreview(); });
document.getElementById("resetPauta").addEventListener("click", () => {
  if (confirm("Isso vai descartar as alterações na pauta (itens adicionados, removidos ou editados) e o conteúdo escrito nas abas dos tópicos, voltando ao padrão original. Confirma?")) {
    resetPautaList();
    if (typeof refreshPreview === "function") refreshPreview();
  }
});

function getPautaItems() {
  return Array.from(document.querySelectorAll(".pauta-input")).map((i) => i.value.trim()).filter(Boolean);
}

/* ===================== Geração via modelo .pptx oficial (backend Python) ===================== */

// Lê o formulário de eleição de um tópico (campos data-el): listas por cargo, votação, resultado, mandato e observações.
function lerEleicao(topicId) {
  const form = document.querySelector("#panel-topic-" + topicId + " .el-form");
  if (!form) return null;
  const v = (id) => { const e = form.querySelector('[data-el="' + id + '"]'); return e ? e.value.trim() : ""; };
  const lista = (id) => v(id).split("\n").map((x) => x.trim()).filter(Boolean);
  return { candSindico: lista("candSindico"), candSubsindico: lista("candSubsindico"), candConselho: lista("candConselho"), votacao: v("votacao"),
    eleitoSindico: v("eleitoSindico"), eleitoSubsindico: v("eleitoSubsindico"), eleitosConselho: lista("eleitosConselho"), mandato: v("mandato"), obs: lista("obs") };
}
// Igual a _dados_eleicao no backend: null se não há candidatos nem resultado (cai no texto livre).
function eleicaoDados(el) {
  if (!el) return null;
  const L = (x) => (Array.isArray(x) ? x : String(x || "").split("\n")).map((n) => String(n).trim()).filter(Boolean);
  const d = {
    cand: [["Síndico", L(el.candSindico)], ["Subsíndico", L(el.candSubsindico)], ["Conselho fiscal", L(el.candConselho)]].filter((c) => c[1].length),
    res: [["Síndico eleito", L(el.eleitoSindico)], ["Subsíndico eleito", L(el.eleitoSubsindico)], ["Conselho fiscal eleito", L(el.eleitosConselho)]].filter((c) => c[1].length),
    votacao: String(el.votacao || "").trim(), mandato: String(el.mandato || "").trim(), obs: L(el.obs),
  };
  if (!d.cand.length && !d.res.length) return null;
  return d;
}
// Maior fonte (pt) em que a lista de nomes cabe no card (igual a _tam_nomes_card)
function tamNomesCard(nomes, larguraIn, alturaIn, tams) {
  tams = tams || [32, 28, 26, 24, 22, 20, 18, 16];
  for (const tam of tams) {
    const cpl = Math.max(8, Math.floor((larguraIn - 0.7) * 72 / (tam * 0.52)));
    const linhas = nomes.reduce((n, x) => n + Math.max(1, Math.ceil((x.length + 2) / cpl)), 0);
    if ((linhas * tam * 1.25 + nomes.length * 8) / 72 <= alturaIn) return tam;
  }
  return tams[tams.length - 1];
}
const ELEICAO_VOTACAO_PADRAO = "A definir na assembleia";
function especificarSlidesEleicao(d) {
  const votacao = d.votacao && d.votacao !== ELEICAO_VOTACAO_PADRAO ? d.votacao : "";
  const out = [];
  if (d.cand.length) {
    const extras = !d.res.length ? (d.mandato ? ["Mandato: " + d.mandato] : []).concat(d.obs) : [];
    out.push({ type: "eleicao", kind: "cand", label: "Eleição \u2014 Candidatos", titulo: "Eleição \u2014 Candidatos", cards: d.cand.map(([cargo, nomes]) => ({ cargo: cargo + " (" + nomes.length + ")", nomes })),
      alturaIn: (votacao || extras.length) ? 6.0 : 6.9, linhas: (votacao ? [{ txt: "Forma de votação: " + votacao, forte: true }] : []).concat(extras.length ? [{ txt: extras.join("   |   ").slice(0, 150) }] : []) });
  }
  if (d.res.length) {
    const linhas = [];
    if (votacao && !d.cand.length) linhas.push({ txt: "Forma de votação: " + votacao, forte: true });
    if (d.mandato) linhas.push({ txt: "Mandato: " + d.mandato, forte: true });
    d.obs.forEach((o) => linhas.push({ txt: o }));
    out.push({ type: "eleicao", kind: "res", label: "Eleição \u2014 Resultado", titulo: "Eleição \u2014 Resultado", cards: d.res.map(([cargo, nomes]) => ({ cargo, nomes })), alturaIn: 4.4, linhas });
  }
  return out;
}

// Itens da pauta na ordem atual, com o "papel" de cada um. Tópicos carregam o topicId e o
// conteúdo digitado na aba dinâmica deles (vazio = só o slide divisor).
function getPautaItemsComRole() {
  return Array.from(document.querySelectorAll("#pautaList .repeat-block")).map((div) => {
    const item = {
      texto: div.querySelector(".pauta-input").value.trim(),
      role: div.dataset.role || "topico",
    };
    if (div.dataset.topicId) {
      const ta = document.getElementById("f_topic_" + div.dataset.topicId);
      item.topicId = div.dataset.topicId;
      item.conteudo = ta ? ta.value.trim() : "";
      const el = lerEleicao(div.dataset.topicId);
      if (el) item.eleicao = el;
      const imgs = TOPIC_ATTACHMENTS[div.dataset.topicId];
      if (imgs && imgs.length) item.imagens = imgs.map((x) => ({ dataUrl: x.dataUrl, width: x.width, height: x.height, titulo: x.name }));
    }
    if (div.dataset.extraId) {
      item.extraId = div.dataset.extraId;
      const sec = EXTRA_SECTIONS[div.dataset.extraId];
      if (sec) {
        if (item.role === "prestacao") {
          if (sec.a11) item.a11 = sec.a11;
          if (sec.inadimplencia) item.inadimplencia = sec.inadimplencia;
          if (sec.b16) { item.b16 = sec.b16; item.periodo16b = sec.periodo16b || ""; }
        }
        if (item.role === "previsao" && sec.prev) {
          const un = parseInt(sec.unidadesTotal || "", 10)
            || parseInt(document.getElementById("f_po_unidades_total")?.value || "", 10)
            || parseInt(document.getElementById("f_inad_unidades_total")?.value || "", 10)
            || sec.prev.unidadesTotal
            || null;
          item.previsao = Object.assign({}, sec.prev, un ? { unidadesTotal: un } : {});
        }
      }
    }
    return item;
  }).filter((it) => it.texto);
}

// Monta o JSON exatamente no formato que o generate_pptx.py espera.
function montarDadosParaBackend() {
  const dados = {
    condominio: document.getElementById("f_condominio").value.trim(),
    tipo: document.getElementById("f_tipo").value,
    data: dateToExtensoPT(document.getElementById("f_data").value),
    local: document.getElementById("f_local").value.trim(),
    conv1: fmtHoraBR(document.getElementById("f_conv1").value),
    conv2: fmtHoraBR(document.getElementById("f_conv2").value),
    pauta: getPautaItemsComRole(),
    canais: {
      tel: document.getElementById("f_tel").value.trim(),
      whats: document.getElementById("f_whats").value.trim(),
      email: document.getElementById("f_email").value.trim(),
      site: document.getElementById("f_site").value.trim(),
      endereco: document.getElementById("f_endereco").value.trim(),
    },
  };

  if (PARSED.a11) dados.a11 = PARSED.a11;

  const organica = document.getElementById("f_inad_organica").value.trim();
  const garantidora = document.getElementById("f_inad_garantidora").value.trim();
  const total = document.getElementById("f_inad_total").value.trim();
  const valorAcum = document.getElementById("f_inad_valor").value.trim();
  const cobrancasAcum = document.getElementById("f_inad_cobrancas").value.trim();
  const unidadesAcum = document.getElementById("f_inad_unidades").value.trim();
  const unidadesTotal = document.getElementById("f_inad_unidades_total").value.trim();
  if (organica || garantidora || total || valorAcum) {
    dados.inadimplencia = {
      organica, garantidora, total,
      valor: valorAcum, cobrancas: cobrancasAcum, unidades: unidadesAcum, unidades_total: unidadesTotal,
      posicao: dataHojeBR(),
    };
  }

  ajustesCarregar();
  if (Object.keys(AJUSTES).length) dados.ajustes = AJUSTES; // linhas por slide escolhidas na prévia
  if (PARSED.b16) {
    dados.b16 = PARSED.b16;
    dados.periodo16b = document.getElementById("f_16b_periodo").value.trim();
  }

  if (PARSED.prev) {
    const unPo = parseInt(document.getElementById("f_po_unidades_total")?.value || "", 10)
      || parseInt(document.getElementById("f_inad_unidades_total")?.value || "", 10)
      || PARSED.prev.unidadesTotal
      || null;
    dados.previsao = Object.assign({}, PARSED.prev, unPo ? { unidadesTotal: unPo } : {});
  }

  return dados;
}

async function gerarApresentacaoModelo() {
  const msgEl = document.getElementById("genMsg");
  const btn = document.getElementById("btnGerar");
  msgEl.className = ""; msgEl.textContent = "";

  const condominio = document.getElementById("f_condominio").value.trim();
  if (!condominio) {
    reportError("genMsg", "Falta o nome do condomínio", "O nome do condomínio é obrigatório pra gerar a apresentação.", "Volte na aba \"Dados da assembleia\" e preencha ou selecione o condomínio.");
    return;
  }
  const pauta = getPautaItemsComRole();
  if (!pauta.length) {
    reportError("genMsg", "Pauta vazia", "A pauta não tem nenhum item com texto.", "Volte na aba \"Dados da assembleia\" e adicione ou preencha pelo menos um item da pauta.");
    return;
  }
  // Os relatórios só são exigidos se o item correspondente ainda está na pauta (incluindo cópias).
  const faltaPrest = pauta.find((it) => it.role === "prestacao" && !(it.extraId ? it.a11 : PARSED.a11));
  if (faltaPrest) {
    reportError("genMsg", "Falta o relatório 11A", 'Falta buscar ou enviar o relatório 11A de "' + faltaPrest.texto + '".', "Abra a aba correspondente e clique em buscar automaticamente ou envie o arquivo .xlsx.");
    return;
  }
  const faltaPrev = pauta.find((it) => it.role === "previsao" && !(it.extraId ? it.previsao : PARSED.prev));
  if (faltaPrev) {
    reportError("genMsg", "Falta a Previsão Orçamentária", 'Falta enviar a planilha de previsão orçamentária de "' + faltaPrev.texto + '".', "Abra a aba correspondente e envie o arquivo .xlsx.");
    return;
  }

  btn.disabled = true;
  btn.innerHTML = '<span class="spinner"></span> Gerando...';

  try {
    const dados = montarDadosParaBackend();
    const res = await fetch("/api/generate", {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify(dados),
    });

    if (!res.ok) {
      let detalhe = "HTTP " + res.status;
      try {
        const erroJson = await res.json();
        if (erroJson && erroJson.erro) detalhe = erroJson.erro;
      } catch (e) { /* resposta não era JSON — segue com o status HTTP */ }
      throw new Error(detalhe);
    }

    const blob = await res.blob();
    const nomeArquivo = "Apresentacao_" + condominio.replace(/[^a-zA-Z0-9]+/g, "_").slice(0, 40) + ".pptx";

    const url = URL.createObjectURL(blob);
    const a = document.createElement("a");
    a.href = url;
    a.download = nomeArquivo;
    document.body.appendChild(a);
    a.click();
    a.remove();
    URL.revokeObjectURL(url);

    msgEl.className = "ok";
    msgEl.textContent = "Apresentação gerada! O download deve começar automaticamente.";
  } catch (err) {
    console.error(err);
    msgEl.className = "err";
    msgEl.textContent = "Erro ao gerar a apresentação: " + err.message;
    showErrorModal("Não consegui gerar a apresentação", "Erro ao gerar a apresentação: " + err.message, "Confira se todos os dados obrigatórios estão preenchidos e tente novamente. Se o erro continuar, tire um print desta mensagem e envie pro responsável pelo sistema.");
  } finally {
    btn.disabled = false;
    btn.textContent = "Gerar apresentação (.pptx)";
  }
}

function dateToExtensoPT(isoDate) {
  if (!isoDate) return "";
  const MESES = ["janeiro", "fevereiro", "março", "abril", "maio", "junho", "julho", "agosto", "setembro", "outubro", "novembro", "dezembro"];
  const [y, m, d] = isoDate.split("-").map(Number);
  if (!y || !m || !d) return "";
  return parseInt(d, 10) + " de " + MESES[m - 1] + " de " + y;
}

document.getElementById("btnGerar").addEventListener("click", gerarApresentacaoModelo);

document.getElementById("btnConv2Auto").addEventListener("click", () => {
  const conv1El = document.getElementById("f_conv1");
  if (!conv1El.value) return;
  const [h, m] = conv1El.value.split(":").map(Number);
  const total = h * 60 + m + 30;
  const novaHora = Math.floor(total / 60) % 24;
  const novoMin = total % 60;
  document.getElementById("f_conv2").value = String(novaHora).padStart(2, "0") + ":" + String(novoMin).padStart(2, "0");
});

/* ===================== Prévia da apresentação (galeria, desktop) ===================== */
const PV_COLORS = {
  darkPurple: "#1D1339",
  gold: "#D08E02",
  chartBlue: "#4472C4",
  pptxPurple: "#452E86", // PURPLE do generate_pptx.py (linha dos gráficos de energia/água)
  chartRed: "#C0504D",
};

// Cabeçalho roxo: títulos longos (ex: "Resumo Financeiro — De 01/08/2026 até 31/08/2026") reduzem a fonte pra caber.
function pvHeaderBar(titulo) {
  // Igual ao .pptx: título em 24pt negrito no canto superior esquerdo sobre o fundo real do slide de conteúdo.
  return '<div class="pv-ct">' + pvEsc(titulo) + '</div>';
}

// Mesma regra de _cortar (api/generate_pptx.py): descrição longa não quebra linha na tabela de lançamentos.
function cortarTexto(texto, maximo) {
  const t = String(texto || "").replace(/\s+/g, " ").trim();
  return t.length <= maximo ? t : t.slice(0, maximo - 1).trimEnd() + "…";
}

function pvEsc(s) {
  const d = document.createElement("div");
  d.textContent = s == null ? "" : String(s);
  return d.innerHTML;
}

function pvMonthsSlice(arr, n) {
  if (!arr) return [];
  if (arr.length <= n) return arr;
  return arr.slice(arr.length - n);
}

// Mini gráfico de barras (1 ou 2 séries) em divs, sem lib nenhuma.
// Escala do eixo de valores: mesma regra de _eixo_valores em api/generate_pptx.py (folga de 25% no topo
// pros rótulos de valor; passo "redondo"), pra prévia e .pptx terem a mesma escala.
function pvEixoValores(maxv) {
  if (maxv <= 0) return [1, 1];
  const alvo = maxv * 1.25, bruto = alvo / 6;
  const mag = Math.pow(10, Math.floor(Math.log10(bruto)));
  let passo = mag;
  for (const m of [1, 2, 2.5, 5, 10]) { passo = m * mag; if (passo >= bruto) break; }
  return [Math.ceil(alvo / passo - 1e-9) * passo, passo];
}

// Gráfico nativo do .pptx reproduzido em SVG (1 unidade = 0,01"; área de 16,4" x 6,8" igual ao .pptx):
// barras agrupadas (Receitas x Despesas) ou linha com marcadores (Energia, Água), eixo com grade,
// rótulos de valor em reais ("R$ 6.141" nas barras, sem centavos; "R$ 6.141,12" na linha), sempre na horizontal.
function pvChartSvg(kind, labels, series, colors, legend, refMax) {
  const W = 1640, H = 680, esq = 120, dir = 20, topo = 30; // fontes em pt do .pptx x 1,389 (1 unidade = 0,01")
  const temLegenda = series.length > 1;
  const base = 52 + (temLegenda ? 56 : 0);
  const pw = W - esq - dir, ph = H - topo - base;
  const [maximo, passo] = pvEixoValores(refMax !== undefined && refMax !== null ? refMax : Math.max(0, ...series.flat().map((v) => v || 0)));
  const y = (v) => topo + ph * (1 - v / maximo);
  const fonte = "Calibri,Carlito,'Segoe UI',sans-serif";
  let s = '<svg viewBox="0 0 ' + W + " " + H + '" style="width:100%;height:100%;display:block;font-family:' + fonte + ';" preserveAspectRatio="xMidYMid meet">';
  for (let v = 0; v <= maximo + passo / 1000; v += passo) {
    s += '<line x1="' + esq + '" x2="' + (W - dir) + '" y1="' + y(v) + '" y2="' + y(v) + '" stroke="#D9D9D9" stroke-width="1.5"/>';
    s += '<text x="' + (esq - 10) + '" y="' + (y(v) + 10) + '" text-anchor="end" font-size="22.2" fill="#595959">' + Math.round(v).toLocaleString("pt-BR") + "</text>";
  }
  const n = labels.length, cw = pw / n;
  labels.forEach((lab, i) => {
    s += '<text x="' + (esq + cw * (i + 0.5)) + '" y="' + (topo + ph + 38) + '" text-anchor="middle" font-size="25" fill="#595959">' + pvEsc(lab) + "</text>";
  });
  if (kind === "line") {
    const pts = series[0].map((v, i) => [esq + cw * (i + 0.5), y(v || 0), v || 0]);
    s += '<polyline fill="none" stroke="' + colors[0] + '" stroke-width="3.5" stroke-linejoin="round" points="' + pts.map((p) => p[0] + "," + p[1]).join(" ") + '"/>';
    pts.forEach((p) => {
      s += '<circle cx="' + p[0] + '" cy="' + p[1] + '" r="4.5" fill="' + colors[0] + '"/>';
    });
  } else {
    const ns = series.length, bw = cw / (ns + 0.6); // espaço entre grupos de 60% (GAP_BARRAS_PCT do backend)
    series.forEach((serie, si) => {
      serie.forEach((v0, i) => {
        const v = v0 || 0;
        const x = esq + cw * i + (cw - ns * bw) / 2 + si * bw;
        s += '<rect x="' + x + '" y="' + y(v) + '" width="' + bw + '" height="' + (topo + ph - y(v)) + '" fill="' + colors[si] + '"/>';
      });
    });
  }
  s += '<line x1="' + esq + '" x2="' + (W - dir) + '" y1="' + (topo + ph) + '" y2="' + (topo + ph) + '" stroke="#BFBFBF" stroke-width="1.5"/>';
  if (temLegenda) {
    const larg = legend.reduce((t, l) => t + 60 + l.length * 16, 0);
    let x = (W - larg) / 2;
    legend.forEach((l, i) => {
      s += '<rect x="' + x + '" y="' + (H - 38) + '" width="22" height="22" fill="' + colors[i] + '"/><text x="' + (x + 32) + '" y="' + (H - 18) + '" font-size="25" fill="#595959">' + pvEsc(l) + "</text>";
      x += 60 + l.length * 16;
    });
  }
  return s + "</svg>";
}

/* ---------- Texto livre (bullets): fonte grande, igual a paginar_bullets em api/generate_pptx.py ---------- */
const BULLETS_FONTES = [54, 48, 44, 40, 36, 32, 28, 24];
const BULLETS_LARGURA_IN = 16.4, BULLETS_ALTURA_IN = 7.4;
function alturaBullets(itens, tam) {
  const cpl = Math.max(10, Math.floor(BULLETS_LARGURA_IN * 72 / (tam * 0.52)));
  return itens.reduce((a, t) => a + Math.max(1, Math.ceil((t.length + 2) / cpl)) * tam * 1.2 + 10, 0) / 72;
}
function paginarBullets(itens) {
  for (const tam of BULLETS_FONTES) if (alturaBullets(itens, tam) <= BULLETS_ALTURA_IN) return [{ tam, itens }];
  const tam = BULLETS_FONTES[BULLETS_FONTES.length - 1];
  const paginas = [];
  let atual = [];
  itens.forEach((it) => {
    if (atual.length && alturaBullets(atual.concat([it]), tam) > BULLETS_ALTURA_IN) { paginas.push({ tam, itens: atual }); atual = []; }
    atual.push(it);
  });
  if (atual.length) paginas.push({ tam, itens: atual });
  return paginas;
}

/* ---------- Paginação do slide de pauta (mesmas constantes de api/generate_pptx.py) ---------- */
const TEXTO_PAUTA_ASSUNTOS = "Assuntos gerais"; // exibição curta do item 'assuntos' na lista Ordem do dia (igual a api/generate_pptx.py)
const PAUTA_LINHAS_PRIMEIRA = 10;
const PAUTA_LINHAS_CONTINUACAO = 13;
const PAUTA_CHARS_POR_LINHA = 100; // estimativa pro pior caso: a 22pt cabem ~120 caracteres de Calibri na largura da lista

// Divide os textos em páginas [[{numero, texto}, ...], ...]; itens longos pesam mais de uma linha
// e a numeração é contínua. Cada página leva no mínimo 1 item e o corte é por índice (inicio:fim),
// então todo item cai em exatamente uma página, mesmo que a estimativa de peso esteja errada.
function paginarPauta(textos) {
  const pesos = textos.map((t) => Math.max(1, Math.ceil((t.length + 4) / PAUTA_CHARS_POR_LINHA)));
  const paginas = [];
  let inicio = 0, capacidade = PAUTA_LINHAS_PRIMEIRA;
  while (inicio < textos.length) {
    let fim = inicio + 1, usado = pesos[inicio];
    while (fim < textos.length && usado + pesos[fim] <= capacidade) { usado += pesos[fim]; fim++; }
    paginas.push(textos.slice(inicio, fim).map((texto, k) => ({ numero: inicio + k + 1, texto })));
    inicio = fim;
    capacidade = PAUTA_LINHAS_CONTINUACAO;
  }
  if (paginas.reduce((n, pg) => n + pg.length, 0) !== textos.length) throw new Error("paginarPauta perdeu itens da pauta");
  return paginas;
}

// Fonte dinâmica por página (manter igual a tamanho_fonte_pauta em api/generate_pptx.py).
// Maior fonte da escala em que os itens DESTA página ainda cabem na altura orçada (a mesma já medida
// no PowerPoint: linhas da capacidade x 0,506" a 22pt). Fonte maior => menos caracteres por linha.
const PAUTA_FONTES = [32, 30, 28, 26, 24, 22];
const PAUTA_FONT_BASE = 22;
const PAUTA_ALTURA_LINHA_IN = 0.506;

function tamanhoFontePauta(textos, primeira) {
  const orcamento = (primeira ? PAUTA_LINHAS_PRIMEIRA : PAUTA_LINHAS_CONTINUACAO) * PAUTA_ALTURA_LINHA_IN;
  for (const tam of PAUTA_FONTES) {
    const escala = tam / PAUTA_FONT_BASE;
    const chars = Math.floor(PAUTA_CHARS_POR_LINHA / escala);
    const linhas = textos.reduce((n, t) => n + Math.max(1, Math.ceil((t.length + 4) / chars)), 0);
    if (linhas * PAUTA_ALTURA_LINHA_IN * escala <= orcamento) return tam;
  }
  return PAUTA_FONT_BASE;
}

const LINHAS_POR_TABELA = 15; // igual a TABELA_LINHAS_POR_SLIDE em api/generate_pptx.py (fonte 18pt, linha de 0,44")
/* ---------- Ajuste manual de paginação por tabela (igual a _paginar_tabela / _layout_tabela no backend) ---------- */
// AJUSTES = { chave: { linhas: N } } — "linhas por slide" escolhido na prévia; vai ao backend em dados.ajustes.
// A fonte e a altura da linha saem do número de linhas (mais linhas => fonte menor, até o piso de 12pt).
const AJ_FONTE_PISO = 12, AJ_AREA_IN = 9.3 - 1.9;
const TIPOS_TABELA = {
  meses: { padrao: 17, fonte: 14, altura: 0.40, fonteG: 16, alturaG: 0.55, grande: 12 },
  previsao: { padrao: 15, fonte: 18, altura: 0.44, fonteG: 24, alturaG: 0.62, grande: 12 },
  lanc: { padrao: 15, fonte: 18, altura: 0.44, fonteG: 18, alturaG: 0.44, grande: 0 },
};
let AJUSTES = {}, AJ_CID = null;
function slugTabela(t) { return String(t || "").normalize("NFD").replace(/[\u0300-\u036f]/g, "").toLowerCase().replace(/[^a-z0-9]+/g, "-").replace(/^-+|-+$/g, ""); }
function capLinhasTabela(tipo) {
  const c = TIPOS_TABELA[tipo];
  return Math.max(c.padrao, Math.floor(AJ_AREA_IN / (c.altura * AJ_FONTE_PISO / c.fonte)) - 1);
}
// {faixas: [[ini, fim], ...], L: linhas por página usadas, manual: bool}
function paginarTabela(n, tipo, chave) {
  const c = TIPOS_TABELA[tipo], cap = capLinhasTabela(tipo);
  if (n <= 0) return { faixas: [], L: c.padrao, manual: false };
  const pedido = AJUSTES[chave] && AJUSTES[chave].linhas;
  let L = c.padrao, manual = false;
  if (pedido) { L = Math.max(1, Math.min(parseInt(pedido, 10) || c.padrao, cap)); manual = true; }
  let paginas = Math.ceil(n / L);
  if (!manual && paginas > 1) {
    const k = n - (paginas - 1) * L;
    if (k <= 2) { // última página órfã (1-2 linhas): as anteriores absorvem
      const L2 = L + Math.ceil(k / (paginas - 1));
      if (L2 <= cap) { L = L2; paginas = Math.ceil(n / L); }
    }
  }
  const faixas = [];
  for (let i = 0; i < paginas; i++) faixas.push([i * L, Math.min(n, (i + 1) * L)]);
  return { faixas, L, manual };
}
// ---- Fusão de blocos adjacentes (igual a _fundir_blocos no backend) ----
// AJUSTES[chave].juntarCom = chave do bloco imediatamente anterior: os dois viram um slide só ("A + B").
// Só vale se as linhas somadas couberem em um slide até o piso de legibilidade.
const SEPS_POR_BLOCO = { meses: 0, previsao: 1, lanc: 1 }; // linhas de título de bloco inseridas ao fundir
function tamUnidade(partes, tipo) {
  if (partes.length === 1) return partes[0].itens.length;
  return partes.reduce((n, p) => n + p.itens.length + SEPS_POR_BLOCO[tipo], 0);
}
function fundirBlocos(blocos, tipo) {
  const cap = capLinhasTabela(tipo), unidades = [];
  blocos.forEach((b, k) => {
    const aj = AJUSTES[b.chave], cur = unidades[unidades.length - 1];
    if (cur && k > 0 && aj && aj.juntarCom && aj.juntarCom === blocos[k - 1].chave && tamUnidade(cur.partes.concat([b]), tipo) <= cap) { cur.partes.push(b); return; }
    unidades.push({ partes: [b] });
  });
  return unidades.map((u) => {
    const ps = u.partes;
    let itens;
    if (ps.length === 1) itens = ps[0].itens.slice();
    else {
      itens = [];
      ps.forEach((pt) => {
        if (tipo === "previsao") itens.push({ tipo: "grupo", nome: pt.nome });
        else if (tipo === "lanc") itens.push({ sep: pt.nome });
        pt.itens.forEach((x) => itens.push(x));
      });
    }
    return { chave: ps[0].chave, ultimaChave: ps[ps.length - 1].chave, nomes: ps.map((pt) => pt.nome), itens, fundido: ps.length > 1, partes: ps };
  });
}
function juntarNomes(nomes, palavra) {
  const j = nomes.join(" + ");
  if (j.length <= 46) return j;
  const n = nomes.length - 1;
  return nomes[0] + " + " + n + " " + (n > 1 ? palavra : palavra.slice(0, -1));
}
function nomeCurtoPlano(grupo) {
  const m = grupo.match(/^\s*(receitas?|despesas?)\b[\s:\-\u2013\u2014]*/i);
  const resto = (m ? grupo.slice(m[0].length) : grupo).trim();
  return resto ? resto.charAt(0).toUpperCase() + resto.slice(1) : grupo.trim();
}
function layoutTabela(tipo, nDados) {
  const c = TIPOS_TABELA[tipo], n = nDados + 1;
  const grande = !!c.grande && n <= c.grande;
  const fb = grande ? c.fonteG : c.fonte, hb = grande ? c.alturaG : c.altura;
  const altura = Math.min(hb, AJ_AREA_IN / n);
  return { fontePt: Math.max(AJ_FONTE_PISO, Math.floor(fb * Math.min(1, altura / hb) + 0.5)), alturaIn: altura };
}
// estilo inline de uma tabela com fonte/altura calculadas (20" de slide = 100cqw)
function estiloTabelaAjustada(fontePt, alturaIn) { return "font-size:" + (fontePt * 100 / 72 / 20).toFixed(3) + "cqw;--rh:" + (alturaIn * 5).toFixed(3) + "cqw;"; }
// persistência: rascunho por condomínio no localStorage
function ajustesCarregar() {
  const cid = document.getElementById("f_condominio_id") ? document.getElementById("f_condominio_id").value : "";
  if (AJ_CID === cid) return;
  AJ_CID = cid;
  try { const raw = localStorage.getItem("jetro_ajustes_slides"); const o = raw ? JSON.parse(raw) : null; AJUSTES = o && o.cid === cid ? (o.aj || {}) : {}; } catch (e) { AJUSTES = {}; }
}
function ajustesSalvar() { try { localStorage.setItem("jetro_ajustes_slides", JSON.stringify({ cid: AJ_CID, aj: AJUSTES })); } catch (e) { /* sem storage: só na memória */ } }
const PLANO_LINHAS_GRANDE = 12; // igual a PLANO_LINHAS_GRANDE no backend: poucas linhas => fonte 15pt e linha de 0,55"
const BARRAS_MESES_POR_SLIDE = 12; // igual a BARRAS_MESES_POR_SLIDE em api/generate_pptx.py (12 meses = 1 slide)
const LINHA_MESES_POR_SLIDE = 12; // igual a LINHA_MESES_POR_SLIDE em api/generate_pptx.py
const LINHAS_POR_SLIDE_PLANO = 17; // igual a PLANO_LINHAS_POR_SLIDE em api/generate_pptx.py (13pt, linha de 0,4")

// Mesmo algoritmo do backend (_grupos_plano): um grupo por conta de nível 0 + descendentes; contas sem valor
// nos meses exibidos ficam de fora; grupos de uma linha só são reunidos em "Outras contas".
function gruposPlanoContas(contas, secao, offset, nMeses) {
  let grupos = [];
  contas.filter((c) => c.secao === secao).forEach((c) => {
    const vals = c.valores.slice(offset).map((v) => v || 0);
    while (vals.length < nMeses) vals.push(0);
    vals.length = nMeses;
    if (!vals.some((v) => v)) return;
    const linha = { nome: c.nome, nivel: Math.max(0, c.nivel || 0), negrito: c.natureza === "0", vals };
    if (linha.nivel === 0 || !grupos.length) grupos.push({ nome: c.nome, linhas: [linha] });
    else grupos[grupos.length - 1].linhas.push(linha);
  });
  const soltos = grupos.filter((g) => g.linhas.length === 1);
  if (soltos.length > 1) {
    const saida = [];
    let inserido = false;
    grupos.forEach((g) => {
      if (g.linhas.length === 1) {
        if (!inserido) { saida.push({ nome: "Outras contas", linhas: soltos.reduce((a, s) => a.concat(s.linhas), []) }); inserido = true; }
      } else saida.push(g);
    });
    grupos = saida;
  }
  return grupos;
}
// Colunas de valor da previsão: da planilha (Atual, Previsto) ou as definidas no preenchimento manual.
function nomesColunasPrevisao(prev) { return prev.colunasValores && prev.colunasValores.length ? prev.colunasValores : ["Atual", "Previsto"]; }
function valoresItemPrevisao(it, n) {
  const v = it.valores ? it.valores.slice() : [it.atualRaw, it.previstoRaw];
  while (v.length < n) v.push(null);
  return v;
}
// Igual a _segmentos_previsao no backend: um segmento por grupo + "Resumo" com todo o fechamento na ordem da planilha.
function segmentosPrevisao(itens) {
  const segs = [], resumo = [];
  let atual = null, emFechamento = false;
  itens.forEach((it) => {
    if (it.tipo === "subtotal" && /^total de pagamentos$/i.test(String(it.nome || "").trim())) {
      atual = null;
      emFechamento = true;
      resumo.push(it);
      return;
    }
    if (emFechamento) {
      resumo.push(it);
      return;
    }
    if (it.tipo === "grupo") { atual = { nome: it.nome, itens: [] }; segs.push(atual); }
    else if (atual) { atual.itens.push(it); if (it.tipo === "subtotal") atual = null; }
    else resumo.push(it);
  });
  const out = segs.filter((g) => g.itens.length);
  if (resumo.length) out.push({ nome: "Resumo", itens: resumo });
  return out;
}
const PREVISAO_LINHAS_GRANDE = 12; // igual ao backend: poucas linhas => fonte 24pt e linha de 0,62"
function tituloPlano(rotulo, grupo, sufixo) {
  // Igual a _titulo_plano no backend: sem repetir "Receitas"/"Despesas" quando o nome do grupo já começa com isso.
  const m = grupo.match(/^\s*(receitas?|despesas?)\b[\s:\-\u2013\u2014]*/i);
  let resto = (m ? grupo.slice(m[0].length) : grupo).trim();
  if (!resto) return rotulo + sufixo;
  resto = resto.charAt(0).toUpperCase() + resto.slice(1);
  if (resto.length > 46) resto = resto.slice(0, 45).trimEnd() + "\u2026";
  return rotulo + " \u2014 " + resto + sufixo;
}
function cabecalhoCategoria(nome) {
  nome = (nome || "").trim();
  return nome.length <= 30 ? nome : nome.slice(0, 29).trimEnd() + "\u2026";
}
// Igual a _fonte_tabela_meses no backend (pt): maior fonte em que o valor mais longo cabe na coluna do mês.
function fontePlanoMeses(textos, grande) {
  const maxlen = Math.max(0, ...textos.map((t) => String(t).length));
  let base = grande ? 16 : 14;
  if (maxlen > 10) base -= 2; else if (maxlen > 9) base -= 1;
  return Math.max(12, base);
}
// Igual a _linhas_resumo no backend: TODAS as contas do plano entre os saldos e totais.
// Fonte e altura de linha para n linhas (com cabeçalho) entre y e fim (pol.); igual a _dim_tabela no backend.
function dimTabela(n, y, fim, fb, hb) {
  const altura = Math.min(hb, (fim - y) / Math.max(1, n));
  const fonte = altura >= hb * 0.85 ? fb : altura >= hb * 0.65 ? Math.max(12, fb - 2) : 12;
  return { fontePt: fonte, alturaIn: altura };
}
// Linhas do Resumo Financeiro (igual ao backend): contas, Total, título "Grupos de saldo", grupos e Saldo final.
function linhas16B(b16) {
  const cs = b16.contas, L = [];
  cs.forEach((c) => L.push({ tipo: "conta", nome: c.nome, v: [c.saldoAnt, c.creditos, c.debitos, c.saldoFinal] }));
  if (cs.length > 1) { const soma = (k) => cs.reduce((n, c) => n + (c[k] || 0), 0); L.push({ tipo: "total", nome: "Total", v: [soma("saldoAnt"), soma("creditos"), soma("debitos"), soma("saldoFinal")] }); }
  const grupos = b16.grupos || [], sf = b16.saldoFinalGeral;
  if (grupos.length || (sf !== null && sf !== undefined)) {
    L.push({ tipo: "titulo", nome: "Grupos de saldo" });
    grupos.forEach((g) => L.push({ tipo: "conta", nome: g.nome, v: [g.saldoAnt, g.creditos, g.debitos, g.saldoFinal] }));
    // Saldo final = mesma lógica da linha Total das contas: soma de cada coluna dos grupos (o saldo final do documento, se houver, prevalece)
    const somaG = (k) => grupos.reduce((n, g) => n + (g[k] || 0), 0);
    L.push({ tipo: "saldofinal", nome: "Saldo final", v: [somaG("saldoAnt"), somaG("creditos"), somaG("debitos"), sf !== null && sf !== undefined ? sf : somaG("saldoFinal")] });
  }
  return L;
}
function linhasResumo(a11, offset, nMeses, soGrupos) {
  const fat = (arr) => (arr ? arr.slice(offset, offset + nMeses) : new Array(nMeses).fill(null));
  const contas = a11.contasPlano || [];
  const linhas = [["Saldo anterior", fat(a11.saldoAnterior), false, 0, false]];
  let sepProx = true; // a próxima linha abre um novo bloco (Receitas | Despesas | Saldo final)
  [["receitas", a11.totalReceitas, "Total de Receitas"], ["despesas", a11.totalDespesas, "Total de Despesas"]].forEach(([secao, tot, rotulo]) => {
    // 1º nível = filhos diretos de Receitas/Despesas, sem nomes fixos: com uma única raiz (nível 0, natureza 0) é o nível 1; senão, o menor nível
    const daSecao = contas.filter((c) => c.secao === secao);
    const niveis = daSecao.map((c) => c.nivel || 0);
    const raizes = daSecao.filter((c) => (c.nivel || 0) === 0);
    const nivelAlvo = raizes.length === 1 && raizes[0].natureza === "0" && niveis.indexOf(1) >= 0 ? 1 : (niveis.length ? Math.min(...niveis) : 0);
    daSecao.forEach((c) => {
      const vals = c.valores.slice(offset, offset + nMeses).map((v) => v || 0);
      while (vals.length < nMeses) vals.push(0);
      const ehGrupo = c.natureza === "0";
      if (vals.some((v) => v) && (!soGrupos || (c.nivel || 0) === nivelAlvo)) {
        const valsExibir = (!soGrupos && ehGrupo) ? null : vals;
        linhas.push([c.nome, valsExibir, ehGrupo, soGrupos ? 1 : 1 + Math.max(0, c.nivel || 0), sepProx]);
        sepProx = false;
      }
    });
    linhas.push([rotulo, fat(tot), true, 0, sepProx]);
    sepProx = true;
  });
  linhas.push(["Saldo Final", fat(a11.saldoFinal), true, 0, true]);
  return linhas;
}

function calcularDiferencaMediaRateio(prev) {
  const ra = prev && prev.rateioAtual;
  const rp = prev && prev.rateioPrevisto;
  const un = parseInt(document.getElementById("f_po_unidades_total")?.value || "", 10)
    || parseInt(document.getElementById("f_inad_unidades_total")?.value || "", 10)
    || (prev && prev.unidadesTotal)
    || (rp && rp.unidades)
    || (ra && ra.unidades)
    || null;
  const totA = (prev && typeof prev.totalDespesasAtual === "number") ? prev.totalDespesasAtual : (ra && ra.total);
  const totP = (prev && typeof prev.totalDespesasPrevisto === "number") ? prev.totalDespesasPrevisto : (rp && rp.total);
  if (un && un > 0 && typeof totA === "number" && typeof totP === "number") {
    return (totP - totA) / un;
  }

  const itensA = (ra && ra.itens) || [];
  const itensP = (rp && rp.itens) || [];
  if (itensA.length && itensP.length) {
    const diffs = [];
    const usados = new Set();
    itensA.forEach((a, idxA) => {
      const na = String(a.nome || "").trim().toLowerCase();
      let idxP = itensP.findIndex((p, k) => !usados.has(k) && String(p.nome || "").trim().toLowerCase() === na);
      if (idxP < 0 && idxA < itensP.length && !usados.has(idxA)) idxP = idxA;
      if (idxP >= 0) {
        usados.add(idxP);
        const p = itensP[idxP];
        if (typeof a.recibo === "number" && typeof p.recibo === "number") diffs.push(p.recibo - a.recibo);
      }
    });
    if (diffs.length) {
      const ref = itensP.length ? itensP : itensA;
      const ehFracoes = ref.length >= 2 && !ref.every((it) => /^(cota|fundo|[áa]gua|taxa)/i.test(String(it.nome || "").trim()));
      return ehFracoes ? (diffs.reduce((s, d) => s + d, 0) / diffs.length) : diffs.reduce((s, d) => s + d, 0);
    }
  }
  for (const b of [ra, rp]) {
    if (b && b.reajuste !== null && b.reajuste !== undefined && !b.reajusteIsPercent) return b.reajuste;
  }
  return null;
}

function montarLinhasRateioFracao(prev, reajustePct) {
  const ra = prev && prev.rateioAtual;
  const rp = prev && prev.rateioPrevisto;
  const itensA = (ra && ra.itens) || [];
  const itensP = (rp && rp.itens) || [];
  if (!itensA.length && !itensP.length) return [];
  const fator = typeof reajustePct === "number" ? (1 + reajustePct) : null;
  const rows = [];
  if (itensA.length && itensP.length) {
    const usados = new Set();
    itensA.forEach((a, idxA) => {
      const na = String(a.nome || "").trim().toLowerCase();
      let idxP = itensP.findIndex((p, k) => !usados.has(k) && String(p.nome || "").trim().toLowerCase() === na);
      if (idxP < 0 && idxA < itensP.length && !usados.has(idxA)) idxP = idxA;
      const p = idxP >= 0 ? itensP[idxP] : {};
      if (idxP >= 0) usados.add(idxP);
      const va = a.recibo;
      let vp = p.recibo;
      if ((vp === null || vp === undefined) && typeof va === "number" && fator !== null) vp = va * fator;
      const diff = typeof va === "number" && typeof vp === "number" ? (vp - va) : null;
      rows.push([
        p.nome || a.nome || ("Fração " + (idxA + 1)),
        va === null || va === undefined ? "Conforme consumo" : fmtBRL(va),
        vp === null || vp === undefined ? "Conforme consumo" : fmtBRL(vp),
        diff === null ? "—" : fmtBRL(diff),
      ]);
    });
    itensP.forEach((p, k) => {
      if (usados.has(k)) return;
      const vp = p.recibo;
      rows.push([p.nome || ("Fração " + (k + 1)), "—", vp === null || vp === undefined ? "Conforme consumo" : fmtBRL(vp), "—"]);
    });
  } else if (itensA.length) {
    itensA.forEach((a, idxA) => {
      const va = a.recibo;
      const vp = typeof va === "number" && fator !== null ? (va * fator) : null;
      const diff = typeof va === "number" && typeof vp === "number" ? (vp - va) : null;
      rows.push([
        a.nome || ("Fração " + (idxA + 1)),
        va === null || va === undefined ? "Conforme consumo" : fmtBRL(va),
        vp === null ? "—" : fmtBRL(vp),
        diff === null ? "—" : fmtBRL(diff),
      ]);
    });
  } else {
    itensP.forEach((p, idxP) => {
      const vp = p.recibo;
      rows.push([p.nome || ("Fração " + (idxP + 1)), "—", vp === null || vp === undefined ? "Conforme consumo" : fmtBRL(vp), "—"]);
    });
  }
  return rows;
}

function montarLinhasMediaRateio(prev) {
  const un = parseInt(document.getElementById("f_po_unidades_total")?.value || "", 10)
    || parseInt(document.getElementById("f_inad_unidades_total")?.value || "", 10)
    || prev.unidadesTotal
    || (prev.rateioPrevisto && prev.rateioPrevisto.unidades)
    || (prev.rateioAtual && prev.rateioAtual.unidades)
    || null;
  const totA = typeof prev.totalDespesasAtual === "number" ? prev.totalDespesasAtual : (prev.rateioAtual && prev.rateioAtual.total);
  const totP = typeof prev.totalDespesasPrevisto === "number" ? prev.totalDespesasPrevisto : (prev.rateioPrevisto && prev.rateioPrevisto.total);

  if (un && un > 0 && typeof totA === "number" && typeof totP === "number") {
    let pagA = null, pagP = null, garA = null, garP = null, funA = null, funP = null;
    (prev.itensCompletos || []).forEach((it) => {
      const nm = String(it.nome || "").trim();
      const va = it.valores && typeof it.valores[0] === "number" ? it.valores[0] : (typeof it.atualRaw === "number" ? it.atualRaw : null);
      const vp = it.valores && typeof it.valores[1] === "number" ? it.valores[1] : (typeof it.previstoRaw === "number" ? it.previstoRaw : null);
      if (/^total de pagamentos$/i.test(nm)) { pagA = va; pagP = vp; }
      else if (/^(garantidora|inadimpl[êe]ncia)/i.test(nm)) { garA = va; garP = vp; }
      else if (/^fundo de reserva/i.test(nm) && va !== null) { funA = va; funP = vp; }
    });
    const rows = [];
    if (typeof pagA === "number" && typeof pagP === "number") rows.push(["Cota Condominial (média/unid.)", fmtBRL(pagA / un), fmtBRL(pagP / un), fmtBRL((pagP - pagA) / un), false]);
    if (typeof garA === "number" && typeof garP === "number" && (garA > 0 || garP > 0)) rows.push(["Garantidora (média/unid.)", fmtBRL(garA / un), fmtBRL(garP / un), fmtBRL((garP - garA) / un), false]);
    if (typeof funA === "number" && typeof funP === "number" && (funA > 0 || funP > 0)) rows.push(["Fundo de Reserva (média/unid.)", fmtBRL(funA / un), fmtBRL(funP / un), fmtBRL((funP - funA) / un), false]);
    rows.push(["Total médio por unidade (" + un + " unid.)", fmtBRL(totA / un), fmtBRL(totP / un), fmtBRL((totP - totA) / un), true]);
    return rows;
  }

  const ra = prev.rateioAtual, rp = prev.rateioPrevisto;
  const itensA = (ra && ra.itens) || [], itensP = (rp && rp.itens) || [];
  const ref = itensP.length ? itensP : itensA;
  const ehFracoes = ref.length >= 2 && !ref.every((it) => /^(cota|fundo|[áa]gua|taxa)/i.test(String(it.nome || "").trim()));
  if (!ehFracoes && (itensA.length || itensP.length)) {
    const rows = [];
    const nMax = Math.max(itensA.length, itensP.length);
    for (let i = 0; i < nMax; i++) {
      const a = itensA[i] || {}, p = itensP[i] || {};
      const va = a.recibo, vp = p.recibo;
      rows.push([
        p.nome || a.nome || ("Item " + (i + 1)),
        va === null || va === undefined ? "Conforme consumo" : fmtBRL(va),
        vp === null || vp === undefined ? "Conforme consumo" : fmtBRL(vp),
        typeof va === "number" && typeof vp === "number" ? fmtBRL(vp - va) : "Conforme consumo",
        false,
      ]);
    }
    if ((ra && typeof ra.totalRecibo === "number") || (rp && typeof rp.totalRecibo === "number")) {
      const ta = ra && ra.totalRecibo, tp = rp && rp.totalRecibo;
      rows.push(["TOTAL POR UNIDADE", fmtBRL(ta), fmtBRL(tp), typeof ta === "number" && typeof tp === "number" ? fmtBRL(tp - ta) : "—", true]);
    }
    return rows;
  }
  return [];
}

/* ---------- Monta a lista de slides a partir do que já foi preenchido ---------- */
function buildSlideSpecs() {
  const specs = [];
  const condominio = (document.getElementById("f_condominio").value.trim() || "NOME DO CONDOMÍNIO");
  const tipo = document.getElementById("f_tipo").value;
  const dataExtenso = dateToExtensoPT(document.getElementById("f_data").value);
  const local = document.getElementById("f_local").value.trim();
  const conv1 = fmtHoraBR(document.getElementById("f_conv1").value);
  const conv2 = fmtHoraBR(document.getElementById("f_conv2").value);
  // Na lista "Ordem do dia" o item de Assuntos aparece abreviado; o texto guardado na pauta não muda.
  const pauta = getPautaItemsComRole().map((it) => (it.role === "assuntos" ? TEXTO_PAUTA_ASSUNTOS : it.texto));

  specs.push({ type: "capa", label: "Capa", condominio, tipo, dataExtenso });
  const paginasPauta = paginarPauta(pauta);
  paginasPauta.forEach((itensPagina, i) => {
    specs.push({
      type: "pauta",
      label: paginasPauta.length > 1 ? "Pauta (" + (i + 1) + "/" + paginasPauta.length + ")" : "Pauta",
      tipo, dataExtenso, local, conv1, conv2,
      pauta: itensPagina,
      pagina: i + 1,
      totalPaginas: paginasPauta.length,
      tamanhoPt: tamanhoFontePauta(itensPagina.map((it) => it.texto), i === 0),
    });
  });
  if (!paginasPauta.length) specs.push({ type: "pauta", label: "Pauta", tipo, dataExtenso, local, conv1, conv2, pauta: [], pagina: 1, totalPaginas: 1, tamanhoPt: PAUTA_FONT_BASE });

  // Segue a ordem da pauta, igual ao backend: cada item gera um divisor + os slides do seu "role".
  getPautaItemsComRole().forEach((item, idx0) => {
    const numero = idx0 + 1; // igual ao backend: todo divisor leva o número do item da pauta
    if (item.role !== "assuntos") specs.push({ type: "divider", label: "Divisor", titulo: numero + ". " + item.texto }); // Assuntos Gerais: slide único (numerado abaixo)

    if (item.role === "prestacao") {
      const prefixoAj = item.extraId ? item.extraId + ":" : "";
      const a11 = item.a11 || PARSED.a11;
      if (a11) {
        const todosMeses = a11.months;
        const p11Item = periodoMeses(todosMeses);
        const nChunks = Math.max(1, Math.ceil(todosMeses.length / 12));
        for (let ch = 0; ch < nChunks; ch++) {
          const offset = ch * 12;
          const months = todosMeses.slice(offset, offset + 12);
          const sufCol = nChunks > 1 ? " — " + months[0] + " a " + months[months.length - 1] : "";
          const sl = (arr) => (arr ? arr.slice(offset, offset + months.length) : months.map(() => null));
          if (a11.contasPlano && a11.contasPlano.length) {
            [[true, "Prestação de contas \u2014 Resumo", "resumo-grupos"], [false, "Prestação de contas", "resumo"]].forEach(([soGrupos, tituloBase, chaveBase]) => {
              const chave = prefixoAj + chaveBase;
              const linhasRes = linhasResumo(a11, offset, months.length, soGrupos);
              const pg = paginarTabela(linhasRes.length, "meses", chave);
              pg.faixas.forEach(([i0, i1], pr) => {
                const lote = linhasRes.slice(i0, i1);
                const grande = lote.length + 1 <= PLANO_LINHAS_GRANDE;
                const tit = tituloBase + sufCol + (pg.faixas.length > 1 ? " (" + (pr + 1) + "/" + pg.faixas.length + ")" : "");
                const lay = layoutTabela("meses", lote.length);
                const valsTexto = [].concat(...lote.filter((l) => l[1] !== null).map((l) => l[1].map((v) => fmtBRL(v))));
                const fontePt = Math.min(lay.fontePt, fontePlanoMeses(valsTexto, grande));
                specs.push({
                  type: "tabela11a", label: tit, titulo: tit, periodo: p11Item, months, rows: lote, grande, fontePt, alturaIn: lay.alturaIn,
                  aj: { chave, tipo: "meses", pagina: pr, paginas: pg.faixas.length, linhasPagina: lote.length, L: pg.L, manual: pg.manual, nTotal: linhasRes.length },
                });
              });
            });
          } else {
            specs.push({
              type: "tabela11a", label: "Prestação de contas" + sufCol, titulo: "Prestação de contas" + sufCol, periodo: p11Item, months,
              rows: [["Saldo anterior", sl(a11.saldoAnterior), false, 0], ["Total de Receitas", sl(a11.totalReceitas), true, 0], ["Total de Despesas", sl(a11.totalDespesas), true, 0], ["Saldo Final", sl(a11.saldoFinal), true, 0]],
            });
          }
        }
        const months = todosMeses;
        const sl = (arr) => (arr ? arr.slice() : months.map(() => null));
        // Barras: 6 meses por slide (igual ao backend) pra o rótulo de 14pt caber; mesma escala nas duas páginas
        {
          const rec = sl(a11.totalReceitas).map((v) => v || 0), desp = sl(a11.totalDespesas).map((v) => v || 0);
          const refMax = Math.max(0, ...rec, ...desp);
          const nPag = Math.ceil(months.length / BARRAS_MESES_POR_SLIDE);
          for (let pb = 0; pb < nPag; pb++) {
            const i0 = pb * BARRAS_MESES_POR_SLIDE, i1 = i0 + BARRAS_MESES_POR_SLIDE;
            const suf = nPag > 1 ? " (" + (pb + 1) + "/" + nPag + ")" : "";
            specs.push({
              type: "chart", kind: "bars", label: "Receitas x Despesas" + suf,
              periodo: p11Item, months: months.slice(i0, i1), refMax,
              series: [rec.slice(i0, i1), desp.slice(i0, i1)],
              colors: [PV_COLORS.chartBlue, PV_COLORS.chartRed],
              legend: ["Receitas", "Despesas"],
              titulo: "Prestação de contas — Receitas x Despesas" + suf,
            });
          }
        }
        const slidesLinha = (valores, titulo, nomeSerie) => {
          const vals = sl(valores).map((v) => v || 0), refMax = Math.max(0, ...vals);
          const nPag = Math.ceil(months.length / LINHA_MESES_POR_SLIDE);
          for (let pl = 0; pl < nPag; pl++) {
            const i0 = pl * LINHA_MESES_POR_SLIDE, i1 = i0 + LINHA_MESES_POR_SLIDE;
            const suf = nPag > 1 ? " (" + (pl + 1) + "/" + nPag + ")" : "";
            specs.push({ type: "chart", kind: "line", label: nomeSerie + suf, periodo: p11Item, months: months.slice(i0, i1), refMax, series: [vals.slice(i0, i1)], colors: [PV_COLORS.pptxPurple], legend: [nomeSerie], titulo: titulo + suf });
          }
        };
        if (a11.energiaEletrica) slidesLinha(a11.energiaEletrica, "Despesas com Energia Elétrica", "Energia Elétrica");
        if (a11.aguaEsgotoDespesa) slidesLinha(a11.aguaEsgotoDespesa, "Despesas com Água / Esgoto", "Água / Esgoto");
      } else {
        specs.push({ type: "placeholder", label: "Prestação de contas", titulo: "Prestação de contas", msg: "Busque ou envie o relatório 11A pra ver essa seção" });
      }

      const ehCopiaPrest = !!item.extraId;
      const inadObj = ehCopiaPrest
        ? item.inadimplencia
        : (["f_inad_organica", "f_inad_garantidora", "f_inad_total", "f_inad_valor"].some((id) => document.getElementById(id).value.trim())
            ? {
                valor: document.getElementById("f_inad_valor").value.trim(),
                cobrancas: document.getElementById("f_inad_cobrancas").value.trim(),
                unidades: document.getElementById("f_inad_unidades").value.trim(),
                unidades_total: document.getElementById("f_inad_unidades_total").value.trim(),
                organica: document.getElementById("f_inad_organica").value.trim(),
                garantidora: document.getElementById("f_inad_garantidora").value.trim(),
                total: document.getElementById("f_inad_total").value.trim(),
              }
            : null);

      if (inadObj) specs.push({
        type: "inad", label: "Inadimplência",
        valor: inadObj.valor || "",
        cobrancas: inadObj.cobrancas || "",
        unidades: inadObj.unidades || "",
        unidadesTotal: inadObj.unidades_total || "",
        organica: inadObj.organica || "",
        garantidora: inadObj.garantidora || "",
        total: inadObj.total || "",
      });

      const b16Obj = ehCopiaPrest ? item.b16 : PARSED.b16;
      if (b16Obj) {
        const periodo16 = ehCopiaPrest ? (item.periodo16b || "") : document.getElementById("f_16b_periodo").value.trim();
        const p16Fmt = periodo16B(periodo16);
        const linhas16 = linhas16B(b16Obj);
        const nPag16 = Math.max(1, Math.ceil(linhas16.length / 14));
        for (let pg16 = 0; pg16 < nPag16; pg16++) {
          const rot = "Resumo Financeiro" + (nPag16 > 1 ? " (" + (pg16 + 1) + "/" + nPag16 + ")" : "");
          const lote16 = linhas16.slice(pg16 * 14, (pg16 + 1) * 14);
          const dim16 = dimTabela(lote16.length + 1, 2.2, 9.3, 24, 0.85);
          specs.push({ type: "b16", label: rot, titulo: rot, periodo: p16Fmt, linhas: lote16, fontePt: dim16.fontePt, alturaIn: dim16.alturaIn });
        }
        const MAX_LANCAMENTOS_POR_CONTA = 150;
        const blocosM = [];
        b16Obj.contas.forEach((conta) => {
          if (!conta.movimentos || !conta.movimentos.length) return;
          blocosM.push({ chave: prefixoAj + "lanc:" + slugTabela(conta.nome), nome: conta.nome, itens: conta.movimentos.slice(0, MAX_LANCAMENTOS_POR_CONTA), excedente: Math.max(0, conta.movimentos.length - MAX_LANCAMENTOS_POR_CONTA) });
        });
        fundirBlocos(blocosM, "lanc").forEach((un) => {
          const pg = paginarTabela(un.itens.length, "lanc", un.chave);
          const exced = un.partes.filter((pt) => pt.excedente);
          pg.faixas.forEach(([i0, i1], pagina) => {
            const lote = un.itens.slice(i0, i1);
            const sufixoPagina = pg.faixas.length > 1 ? " (" + (pagina + 1) + "/" + pg.faixas.length + ")" : "";
            const nota = (exced.length && pagina === pg.faixas.length - 1)
              ? "\n\n" + exced.map((pt) => "+ " + pt.excedente + " lançamento(s) adicional(is)" + (un.fundido ? " de " + pt.nome : "") + " não exibidos por limite de tamanho.").join("\n")
              : "";
            const lay = layoutTabela("lanc", lote.length);
            const nomeRot = juntarNomes(un.nomes, "contas");
            specs.push({
              type: "b16mov", label: "Lançamentos " + nomeRot, titulo: "Lançamentos \u2014 " + nomeRot + sufixoPagina, periodo: p16Fmt, movimentos: lote, nota, fontePt: lay.fontePt, alturaIn: lay.alturaIn,
              aj: { chave: un.chave, ultimaChave: un.ultimaChave, familia: "lanc", tipo: "lanc", pagina, paginas: pg.faixas.length, linhasPagina: lote.length, L: pg.L, manual: pg.manual, nTotal: un.itens.length, fundido: un.fundido, nomes: un.nomes },
            });
          });
        });
      }

    } else if (item.role === "previsao") {
      const prefixoAj = item.extraId ? item.extraId + ":" : "";
      const prevObj = item.previsao || PARSED.prev;
      if (prevObj) {
        const LINHAS_POR_SLIDE_PO = LINHAS_POR_TABELA;
        const itens = prevObj.itensCompletos && prevObj.itensCompletos.length ? prevObj.itensCompletos : prevObj.categories.map((c) => ({ tipo: "subtotal", nome: c.nome, atualRaw: c.atual, previstoRaw: c.previsto }));
        const blocosPv = segmentosPrevisao(itens).map((seg) => ({ chave: prefixoAj + "previsao:" + slugTabela(seg.nome), nome: seg.nome, itens: seg.itens }));
        fundirBlocos(blocosPv, "previsao").forEach((un) => {
          const pg = paginarTabela(un.itens.length, "previsao", un.chave);
          pg.faixas.forEach(([i0, i1], pagina) => {
            const lote = un.itens.slice(i0, i1);
            const sufixoPagina = pg.faixas.length > 1 ? " (" + (pagina + 1) + "/" + pg.faixas.length + ")" : "";
            const nome0 = un.nomes[0];
            const nomeT = nome0.length <= 40 ? nome0 : nome0.slice(0, 39).trimEnd() + "\u2026";
            const titulo = un.fundido ? "Previsão \u2014 " + juntarNomes(un.nomes, "grupos") + sufixoPagina : "Previsão \u2014 " + nomeT + sufixoPagina;
            const lay = layoutTabela("previsao", lote.length);
            specs.push({
              type: "previsao", label: titulo, titulo, itens: lote, fontePt: lay.fontePt, alturaIn: lay.alturaIn, colunas: nomesColunasPrevisao(prevObj), cab: un.fundido ? "Descrição" : cabecalhoCategoria(nome0),
              aj: { chave: un.chave, ultimaChave: un.ultimaChave, familia: "previsao", tipo: "previsao", pagina, paginas: pg.faixas.length, linhasPagina: lote.length, L: pg.L, manual: pg.manual, nTotal: un.itens.length, fundido: un.fundido, nomes: un.nomes },
            });
          });
        });
        // Rateio dividido em 2 slides: 1) Média por unidade; 2) Por fração
        const prevR = prevObj;
        if (prevR.rateioAtual || prevR.rateioPrevisto) {
          const cards = [];
          let reajustePct = null;
          [prevR.rateioAtual, prevR.rateioPrevisto].forEach((b) => {
            if (b && b.reajuste !== null && b.reajuste !== undefined && b.reajusteIsPercent) reajustePct = b.reajuste;
          });
          if (reajustePct === null && prevR.rateioAtual && prevR.rateioPrevisto && prevR.rateioAtual.total > 0 && prevR.rateioPrevisto.total !== null) {
            reajustePct = (prevR.rateioPrevisto.total / prevR.rateioAtual.total) - 1;
          }
          if (reajustePct === null && typeof prevR.totalDespesasAtual === "number" && prevR.totalDespesasAtual > 0 && typeof prevR.totalDespesasPrevisto === "number") {
            reajustePct = (prevR.totalDespesasPrevisto / prevR.totalDespesasAtual) - 1;
          }
          if (reajustePct !== null) cards.push(["Reajuste do rateio", (reajustePct * 100).toFixed(2).replace(".", ",") + "%"]);

          const diffMedia = calcularDiferencaMediaRateio(prevR);
          if (diffMedia !== null) cards.push(["Diferença média por unidade", "R$ " + fmtBRL(diffMedia)]);

          if ((prevR.totalDespesasAtual !== null && prevR.totalDespesasAtual !== undefined) || (prevR.totalDespesasPrevisto !== null && prevR.totalDespesasPrevisto !== undefined)) {
            cards.push(["Total de despesas (atual \u2192 previsto)", fmtBRL(prevR.totalDespesasAtual) + " \u2192 " + fmtBRL(prevR.totalDespesasPrevisto)]);
          }

          const linhasMedia = montarLinhasMediaRateio(prevR);
          if (linhasMedia.length) {
            specs.push({ type: "rateio_media", label: "Rateio — Média por unidade", titulo: "Previsão orçamentária — Rateio (Média por unidade)", linhas: linhasMedia, cards });
          }

          const refItens = (prevR.rateioPrevisto && prevR.rateioPrevisto.itens) || (prevR.rateioAtual && prevR.rateioAtual.itens) || [];
          const ehFracoes = refItens.length >= 2 && !refItens.every((it) => /^(cota|fundo|[áa]gua|taxa)/i.test(String(it.nome || "").trim()));
          if (ehFracoes || !linhasMedia.length) {
            const linhasFrac = montarLinhasRateioFracao(prevR, reajustePct);
            if (linhasFrac.length) {
              specs.push({ type: "rateio", label: "Rateio — Por fração", titulo: "Previsão orçamentária — Rateio por fração", linhas: linhasFrac, cards });
            }
          }
        }
      } else {
        specs.push({ type: "placeholder", label: "Previsão Orçamentária", titulo: "Previsão orçamentária", msg: "Envie a planilha de previsão orçamentária pra ver essa seção" });
      }
    } else if (item.role === "assuntos") {
      // Slide único no estilo divisor, só com o título fixo, sem número (igual ao backend).
      specs.push({ type: "divider", label: "Assuntos gerais", titulo: numero + ". Assuntos Gerais" });
    } else if (item.eleicao && eleicaoDados(item.eleicao)) {
      especificarSlidesEleicao(eleicaoDados(item.eleicao)).forEach((sp) => specs.push(sp));
    } else {
      if (item.conteudo) {
        const paginasB = paginarBullets(item.conteudo.split("\n").map((l) => l.trim()).filter(Boolean));
        paginasB.forEach((pg, k) => {
          const suf = paginasB.length > 1 ? " (" + (k + 1) + "/" + paginasB.length + ")" : "";
          specs.push({ type: "bullets", label: item.texto + suf, titulo: item.texto + suf, itens: pg.itens, tam: pg.tam });
        });
      }
      if (item.imagens && item.imagens.length) {
        item.imagens.forEach((img, idxImg) => {
          const sufImg = item.imagens.length > 1 ? " (" + (idxImg + 1) + "/" + item.imagens.length + ")" : "";
          const tit = item.texto + sufImg;
          specs.push({ type: "imagem", label: tit, titulo: tit, dataUrl: img.dataUrl });
        });
      }
    }
  });

  // Período do relatório no cabeçalho dos slides que vêm dele (sem sobrescrever cópias que já têm seu próprio período)
  {
    const a11p = PARSED.a11, p11 = a11p ? periodoMeses(a11p.months) : "";
    const p16 = PARSED.b16 ? periodo16B(document.getElementById("f_16b_periodo").value.trim()) : "";
    specs.forEach((sp) => {
      if ((sp.type === "tabela11a" || sp.type === "chart") && !sp.periodo && p11) sp.periodo = p11;
      else if (sp.type === "b16mov" && !sp.periodo && p16) sp.periodo = p16;
      else if (sp.type === "inad" && !sp.periodo) sp.periodo = "Posição em " + dataHojeBR();
    });
  }
  specs.push({
    type: "encerramento", label: "Encerramento",
    tel: document.getElementById("f_tel").value.trim(),
    whats: document.getElementById("f_whats").value.trim(),
    email: document.getElementById("f_email").value.trim(),
    site: document.getElementById("f_site").value.trim(),
    endereco: document.getElementById("f_endereco").value.trim(),
  });

  return specs;
}

// Primeira linha do endereço em negrito, igual ao .pptx (o modelo deixa a 1ª linha em negrito).
function pvEnderecoHtml(endereco) {
  const partes = endereco.split(String.fromCharCode(10));
  return "<b>" + pvEsc(partes[0]) + "</b>" + (partes.length > 1 ? pvEsc(String.fromCharCode(10) + partes.slice(1).join(String.fromCharCode(10))) : "");
}

// Divisores: maior fonte (60pt → 24pt, de 2 em 2) em que o título cabe na caixa de 15" x 5", sem quebrar palavra no meio.
// Igual a _tam_divisor em api/generate_pptx.py.
function tamDivisor(texto) {
  const util = 15.0 - 0.5, altura = 5.0 - 0.2;
  const lc = (ch) => (ch === " " ? 0.28 : (ch !== ch.toLowerCase() && ch === ch.toUpperCase()) || /\d/.test(ch) ? 0.74 : 0.62);
  const palavras = String(texto || "").split(/\s+/).filter(Boolean);
  for (let tam = DIVISOR_PT_BASE; tam >= 24; tam -= 2) {
    const larg = (p) => [...p].reduce((n, c) => n + lc(c), 0) * tam / 72;
    const esp = lc(" ") * tam / 72;
    if (!palavras.every((p) => larg(p) <= util)) continue;
    let linhas = 1, atual = 0;
    palavras.forEach((p) => {
      const w = larg(p);
      if (atual === 0) atual = w;
      else if (atual + esp + w <= util) atual += esp + w;
      else { linhas++; atual = w; }
    });
    if (linhas * tam * 1.25 / 72 <= altura) return tam;
  }
  return 24;
}
const DIVISOR_PT_BASE = 60;

// Período do relatório no cabeçalho (igual a add_periodo no backend): "Período: 11/2025 a 10/2026".
function mesNum(label) {
  const m = /^\s*([A-Za-zçÇ]{3})[a-z]*\s*\/\s*(\d{4})\s*$/.exec(label || "");
  const i = m ? MESES_ABBR.findIndex((x) => x.toLowerCase() === m[1].toLowerCase()) : -1;
  return i >= 0 ? pad2(i + 1) + "/" + m[2] : (label || "");
}
function periodoMeses(meses) { return meses && meses.length ? "Período: " + mesNum(meses[0]) + " a " + mesNum(meses[meses.length - 1]) : ""; }
function dataHojeBR() { const d = new Date(); return pad2(d.getDate()) + "/" + pad2(d.getMonth() + 1) + "/" + d.getFullYear(); }
function periodo16B(txt) {
  const m = /^\s*De\s+(.+?)\s+at[ée]\s+(.+?)\s*$/i.exec(txt || "");
  return m ? "Período: " + m[1] + " a " + m[2] : (txt && txt.trim() ? "Período: " + txt.trim() : "");
}
function renderSlideHtml(spec) {
  const html = renderSlideHtmlBase(spec);
  if (!spec.periodo) return html;
  const i = html.lastIndexOf("</div>");
  return html.slice(0, i) + '<div class="pv-periodo">' + pvEsc(spec.periodo) + "</div>" + html.slice(i);
}

/* ---------- Renderiza um slide (versão mini, estilo igual ao .pptx) ---------- */
function renderSlideHtmlBase(spec) {
  switch (spec.type) {
        case "capa": {
      // Usa o fundo real do slide 1 do modelo (renderizado sem texto) + texto sobreposto
      // nas coordenadas exatas do .pptx original — mesma hierarquia do modelo: tipo em
      // destaque (negrito grande) e o nome do condomínio logo abaixo (menor).
      // Fonte menor pra textos longos (nome de condomínio grande, tipo customizado), pra nunca cortar.
      function tamCapa(txt, base) {
        if (txt.length <= 20) return base;
        if (txt.length <= 35) return base * 0.78;
        if (txt.length <= 50) return base * 0.62;
        return base * 0.5;
      }
      const tamTipo = tamCapa(spec.tipo, 4.1);
      const tamCond = tamCapa(spec.condominio, 3.6);
      return (
        '<div class="pv-slide pv-capa2">' +
        '<div class="pv-t1" style="font-size:' + tamTipo + 'cqw;">' + pvEsc(spec.tipo) + '</div>' +
        '<div class="pv-t2" style="font-size:' + tamCond + 'cqw;">' + pvEsc(spec.condominio) + '</div>' +
        '</div>'
      );
    }
        case "pauta": {
      const linhasInfo = linhasEdital(spec.dataExtenso, spec.local, spec.conv1, spec.conv2);
      const n = spec.pauta.length;
      const primeira = spec.pagina === 1;
      // Proporções fiéis ao .pptx (slide de 20" de largura): 22pt ≈ 1,53cqw, 24pt ≈ 1,67cqw, 18pt ≈ 1,25cqw,
      // espaçamento entre itens de 10pt ≈ 0,69cqw, e a mesma fonte do modelo (Calibri) — assim a prévia
      // quebra linha e ocupa altura como o PowerPoint, e a capacidade da paginação vale pros dois.
      const fonte = "Calibri,Carlito,'Segoe UI',sans-serif";
      const escala = (spec.tamanhoPt || PAUTA_FONT_BASE) / PAUTA_FONT_BASE;
      const start = n ? spec.pauta[0].numero : 1;
      const items = n
        ? spec.pauta.map((p) => '<li style="margin-bottom:' + (0.69 * escala).toFixed(2) + 'cqw;">' + pvEsc(p.texto) + "</li>").join("")
        : '<li class="pv-empty">Nenhum item na pauta ainda</li>';
      const cabecalho = primeira ? "Itens a serem discutidos" : "Ordem do dia";
      const topo = primeira
        ? '<div class="pv-small-title" style="font-size:1.67cqw;margin-bottom:0.8cqw;">Edital de Convocação — ' + pvEsc(spec.tipo) + '</div>' +
          linhasInfo.map((l) => '<div class="pv-muted" style="font-size:1.25cqw;">' + pvEsc(l) + '</div>').join("") +
          '<div class="pv-small-title" style="font-size:1.53cqw;margin-top:2.5cqw;">Ordem do dia:</div>'
        : "";
      return (
        '<div class="pv-slide">' + pvHeaderBar(cabecalho) + '<div class="pv-body" style="top:12.4%;font-family:' + fonte + ';">' + topo +
        '<ol class="pv-pauta-list" start="' + start + '" style="font-size:' + (1.53 * escala).toFixed(3) + 'cqw;line-height:1.2;padding-left:' + (2.2 * escala).toFixed(2) + 'cqw;margin-top:1.2cqw;">' + items + '</ol></div></div>'
      );
    }
       case "divider": {
      // Usa o fundo real do slide 2 do modelo (curvas + coroa já fazem parte da imagem).
      // Fonte única em todos os divisores (igual a DIVISOR_PT no backend); títulos longos só quebram em mais linhas.
      const tamPt = tamDivisor(spec.titulo); // auto-fit: diminui até caber (igual a _tam_divisor no backend)
      const tam = (tamPt * 100 / 72 / 20).toFixed(2);
      return (
        '<div class="pv-slide pv-divisor2">' +
        '<div class="pv-t1" style="font-size:' + tam + 'cqw;">' + pvEsc(spec.titulo) + '</div>' +
        '</div>'
      );
    }
    case "tabela11a": {
      // Layout fixo: coluna de rótulo larga e uma coluna de largura igual por mês, valores sem quebra.
      const n = spec.months.length;
      const comSub = spec.rows.length > 4; // com contas do plano: tabela compacta a partir do topo (igual ao backend)
      const wRot = comSub ? 24 : 15, wMes = ((100 - wRot) / n).toFixed(3);
      const estilo = spec.fontePt ? ' style="' + estiloTabelaAjustada(spec.fontePt, spec.alturaIn) + '"' : "";
      let t = '<table class="pv-table fin12' + (comSub ? " plano12 ajuste" : "") + '"' + estilo + '><colgroup><col style="width:' + wRot + '%;">' + spec.months.map(() => '<col style="width:' + wMes + '%;">').join("") + '</colgroup>' +
        "<tr><th>" + (comSub ? "Categoria" : "") + "</th>" + spec.months.map((m) => "<th>" + pvEsc(m) + "</th>").join("") + "</tr>";
      spec.rows.forEach(([label, arr, bold, nivel, sep]) => {
        const celulas = arr === null
          ? spec.months.map(() => "<td></td>").join("")
          : arr.map((v) => "<td" + (bold ? ' class="bold"' : "") + ">" + fmtBRL(v) + "</td>").join("");
        t += "<tr" + (sep ? ' class="sep"' : "") + "><td" + (bold ? ' class="bold"' : "") + (nivel ? ' style="padding-left:' + (0.5 + nivel * 0.9) + 'cqw;"' : "") + ">" + pvEsc(label) + "</td>" + celulas + "</tr>";
      });
      t += "</table>";
      return '<div class="pv-slide"><div class="pv-ct">' + pvEsc(spec.titulo || "Prestação de contas") + '</div><div class="pv-body" style="' + (comSub ? "" : "top:32%;") + '">' + t + "</div></div>";
    }

    case "rateio_media": {
      let t = '<table class="pv-table mid"><tr><th>Descrição</th><th>Atual (R$)</th><th>Previsto (R$)</th><th>Diferença (R$)</th></tr>';
      spec.linhas.forEach(([desc, atual, previsto, diff, bold]) => {
        const cls = bold ? ' class="bold"' : "";
        t += "<tr><td" + cls + ">" + pvEsc(desc) + "</td><td" + cls + ">" + pvEsc(atual) + "</td><td" + cls + ">" + pvEsc(previsto) + "</td><td" + cls + ">" + pvEsc(diff) + "</td></tr>";
      });
      t += "</table>";
      const cards = spec.cards.length
        ? '<div class="pv-cards-row" style="position:absolute;left:0;right:0;top:' + ((7.9 - 1.55) * 100 / 20).toFixed(2) + 'cqw;gap:' + (0.4 * 100 / 20) + 'cqw;">' +
          spec.cards.map(([lbl, val]) => '<div class="pv-stat-card"><div class="val">' + pvEsc(val) + '</div><div class="lbl">' + pvEsc(lbl) + "</div></div>").join("") + "</div>"
        : "";
      return '<div class="pv-slide">' + pvHeaderBar(spec.titulo) + '<div class="pv-body pv-rateio" style="top:' + (1.55 * 100 / 11.25).toFixed(2) + '%;">' + t + cards + "</div></div>";
    }

    case "rateio": {
      const cqw = (pt) => (pt * 100 / 72 / 20).toFixed(3) + "cqw";
      let t = '<p style="margin:0 0 .8cqw;font-weight:700;font-size:' + cqw(22) + ';color:#1D1339;">Rateio atual x Rateio Previsto</p>' +
        '<table class="pv-table mid"><tr><th>Fração</th><th>Atual (R$)</th><th>Previsto (R$)</th><th>Diferença (R$)</th></tr>';
      (spec.linhas || []).forEach(([frac, atual, previsto, diff]) => {
        t += "<tr><td>" + pvEsc(frac) + "</td><td>" + pvEsc(atual) + "</td><td>" + pvEsc(previsto) + "</td><td>" + pvEsc(diff) + "</td></tr>";
      });
      t += "</table>";
      const cards = spec.cards.length
        ? '<div class="pv-cards-row" style="position:absolute;left:0;right:0;top:' + ((7.9 - 1.55) * 100 / 20).toFixed(2) + 'cqw;gap:' + (0.4 * 100 / 20) + 'cqw;">' +
          spec.cards.map(([lbl, val]) => '<div class="pv-stat-card"><div class="val">' + pvEsc(val) + '</div><div class="lbl">' + pvEsc(lbl) + "</div></div>").join("") + "</div>"
        : "";
      return '<div class="pv-slide">' + pvHeaderBar(spec.titulo) + '<div class="pv-body pv-rateio" style="top:' + (1.55 * 100 / 11.25).toFixed(2) + '%;">' + t + cards + "</div></div>";
    }
    case "imagem": {
      return '<div class="pv-slide">' + pvHeaderBar(spec.titulo) +
        '<div class="pv-body" style="display:flex;align-items:center;justify-content:center;top:14.5%;bottom:6%;">' +
        '<img src="' + spec.dataUrl + '" style="max-width:100%;max-height:100%;object-fit:contain;border-radius:0.4cqw;box-shadow:0 0.2cqw 0.8cqw rgba(0,0,0,0.12);">' +
        "</div></div>";
    }
    case "plano": {
      const n = spec.months.length;
      const wRot = 24, wMes = ((100 - wRot) / n).toFixed(3);
      let t = '<table class="pv-table fin12 plano12 ajuste" style="' + estiloTabelaAjustada(spec.fontePt, spec.alturaIn) + '"><colgroup><col style="width:' + wRot + '%;">' + spec.months.map(() => '<col style="width:' + wMes + '%;">').join("") + "</colgroup><tr><th>" + pvEsc(spec.cab) + "</th>" +
        spec.months.map((m) => "<th>" + pvEsc(m) + "</th>").join("") + "</tr>";
      spec.linhas.forEach((l) => {
        const cls = l.negrito ? ' class="bold"' : "";
        t += "<tr><td" + cls + ' style="padding-left:' + (0.5 + l.nivel * 0.9) + 'cqw;">' + pvEsc(l.nome) + "</td>" + l.vals.map((v) => "<td" + cls + ">" + fmtBRL(v) + "</td>").join("") + "</tr>";
      });
      t += "</table>";
      return '<div class="pv-slide">' + pvHeaderBar(spec.titulo) + '<div class="pv-body">' + t + "</div></div>";
    }
    case "chart": {
      const svg = pvChartSvg(spec.kind, spec.months, spec.series, spec.colors, spec.legend, spec.refMax);
      return '<div class="pv-slide">' + pvHeaderBar(spec.titulo) + '<div class="pv-body pv-chartbody">' + svg + "</div></div>";
    }
    case "inad": {
      // % = unidades inadimplentes ÷ total de unidades × 100; sem o total, cai no % digitado manualmente
      const pct = pctInadimplencia(spec.unidades, spec.unidadesTotal) || spec.total || "";
      const unidadesTxt = spec.unidades ? (spec.unidadesTotal ? spec.unidades + " de " + spec.unidadesTotal : spec.unidades) : "—";
      const cards = [
        ["Acumulado", spec.valor ? "R$ " + spec.valor : "—"],
        ["Cobranças", spec.cobrancas || "—"],
        ["Unidades inadimplentes", unidadesTxt],
        ["% de inadimplência", pct ? pct + "%" : "—"],
      ];
      let html = '<div class="pv-cards-row">' + cards.map(([lbl, val]) => '<div class="pv-stat-card"><div class="val">' + pvEsc(val) + '</div><div class="lbl">' + lbl + "</div></div>").join("") + "</div>";
      html += '<table class="pv-table inad"><colgroup><col style="width:77%;"><col style="width:23%;"></colgroup><tr><th>Indicador</th><th>Valor</th></tr>';
      html += "<tr><td class=\"bold\">% Inadimplência orgânica</td><td class=\"bold\">" + (spec.organica ? spec.organica + "%" : "—") + "</td></tr>";
      html += "<tr><td class=\"bold\">% Inadimplência garantidora</td><td class=\"bold\">" + (spec.garantidora ? spec.garantidora + "%" : "—") + "</td></tr>";
      html += "<tr><td class=\"bold\">% Inadimplência total (unidades inadimplentes ÷ total de unidades)</td><td class=\"bold\">" + (pct ? pct + "%" : "—") + "</td></tr>";
      html += "</table>";
      return '<div class="pv-slide"><div class="pv-ct">Inadimplência</div><div class="pv-body" style="top:14.2%;">' + html + "</div></div>";
    }
    case "b16": {
      let t = '<table class="pv-table mid ajuste" style="' + estiloTabelaAjustada(spec.fontePt, spec.alturaIn) + '"><tr><th>Conta</th><th>Saldo anterior</th><th>Créditos</th><th>Débitos</th><th>Saldo final</th></tr>';
      spec.linhas.forEach((l) => {
        if (l.tipo === "titulo") { t += '<tr><td colspan="5" class="bold" style="background:#efeafa;">' + pvEsc(l.nome) + "</td></tr>"; return; }
        const neg = l.tipo === "total" || l.tipo === "saldofinal";
        const c = (v, ultimo) => "<td" + (neg || ultimo ? ' class="bold"' : "") + ">" + (v === null || v === undefined ? "" : fmtBRL(v)) + "</td>";
        t += "<tr><td" + (neg ? ' class="bold"' : "") + ">" + pvEsc(l.nome) + "</td>" + c(l.v[0]) + c(l.v[1]) + c(l.v[2]) + c(l.v[3], true) + "</tr>";
      });
      t += "</table>";
      return '<div class="pv-slide">' + pvHeaderBar(spec.titulo) + '<div class="pv-body" style="top:19.5%;">' + t + "</div></div>";
    }
        case "b16mov": {
      let t = '<table class="pv-table compact ajuste" style="' + estiloTabelaAjustada(spec.fontePt, spec.alturaIn) + '"><tr><th>Data</th><th>Descrição</th><th>Tipo</th><th>Valor</th></tr>';
      spec.movimentos.forEach((m) => {
        if (m.sep) { t += '<tr><td colspan="4" class="bold" style="background:#efeafa;">' + pvEsc(m.sep) + "</td></tr>"; return; }
        t += "<tr><td>" + pvEsc(m.data) + "</td><td>" + pvEsc(cortarTexto(m.descricao, 80)) + "</td><td>" + (m.tipo === "C" ? "Crédito" : "Débito") + "</td><td>" + fmtBRL(m.valor) + "</td></tr>";
      });
      t += "</table>";
      const nota = spec.nota ? '<p class="pv-muted" style="margin-top:2%;">' + pvEsc(spec.nota) + "</p>" : "";
      return '<div class="pv-slide">' + pvHeaderBar(spec.titulo) + '<div class="pv-body">' + t + nota + "</div></div>";
    }
    case "eleicao": {
      const k = spec.cards.length, gap = 0.4, w = (17.4 - gap * (k - 1)) / k;
      const res = spec.kind === "res";
      const cards = spec.cards.map((c) => {
        const tam = res ? tamNomesCard(c.nomes, w, spec.alturaIn - 0.95 - 0.3, [36, 32, 28, 26, 24, 22, 20]) : tamNomesCard(c.nomes, w, spec.alturaIn - 0.95 - 0.3);
        const nomes = c.nomes.map((n) => "<div" + (res ? ' class="res"' : "") + ">" + (res ? "" : "\u2022 ") + pvEsc(n) + "</div>").join("");
        return '<div class="pv-el-card" style="height:' + (spec.alturaIn * 5).toFixed(2) + 'cqw;"><div class="pv-el-h' + (res ? " ouro" : "") + '">' + pvEsc(c.cargo) + '</div><div class="pv-el-b' + (res ? " c" : "") + '" style="font-size:' + (tam * 100 / 72 / 20).toFixed(3) + 'cqw;">' + nomes + "</div></div>";
      }).join("");
      const linhas = (spec.linhas || []).map((l) => '<div class="pv-el-l' + (l.forte ? " forte" : "") + '">' + pvEsc(l.txt) + "</div>").join("");
      return '<div class="pv-slide">' + pvHeaderBar(spec.titulo) + '<div class="pv-body" style="overflow:visible;"><div class="pv-el-cards">' + cards + "</div>" + linhas + "</div></div>";
    }
    case "bullets": {
      const itens = spec.itens.map((l) => "<li>" + pvEsc(l) + "</li>").join("");
      return '<div class="pv-slide">' + pvHeaderBar(spec.titulo) + '<div class="pv-body"><ul class="pv-bullets" style="font-size:' + (spec.tam * 100 / 72 / 20).toFixed(3) + 'cqw;">' + itens + "</ul></div></div>";
    }
    case "previsao": {
      const colsPO = spec.colunas || ["Atual", "Previsto"];
      const nColsPO = 1 + colsPO.length;
      let t = '<table class="pv-table compact ajuste" style="' + estiloTabelaAjustada(spec.fontePt, spec.alturaIn) + '"><tr><th>' + pvEsc(spec.cab || "Descrição") + '</th>' + colsPO.map((n) => "<th>" + pvEsc(n) + "</th>").join("") + "</tr>";
      spec.itens.forEach((it) => {
        if (it.tipo === "grupo") {
          t += '<tr><td colspan="' + nColsPO + '" class="bold" style="background:#efeafa;">' + pvEsc(it.nome) + "</td></tr>";
        } else {
          const cls = it.tipo === "subtotal" ? ' class="bold"' : "";
          t += "<tr><td" + cls + ">" + pvEsc(it.nome) + "</td>" + valoresItemPrevisao(it, colsPO.length).slice(0, colsPO.length).map((v) => "<td" + cls + ">" + pvEsc(formatPOValue(v)) + "</td>").join("") + "</tr>";
        }
      });
      t += "</table>";
      return '<div class="pv-slide">' + pvHeaderBar(spec.titulo) + '<div class="pv-body">' + t + "</div></div>";
    }
        case "encerramento": {
      function tamContato(txt, maxChars, base) {
        base = base || 2.7;
        if (!txt || txt.length <= maxChars) return base;
        return Math.max(1.7, base * (maxChars / txt.length));
      }
      const endereco = (spec.endereco || "").replace(/\s*-\s*/g, "\n").replace(/,\s*/g, "\n");
      function campo(cls, valor, maxChars, base) {
        if (!valor) return "";
        return '<div class="pv-c ' + cls + '" style="font-size:' + tamContato(valor, maxChars, base) + 'cqw;">' + pvEsc(valor) + '</div>';
      }
      return (
        '<div class="pv-slide pv-encerramento2">' +
        '<div class="pv-c-titulo" style="font-size:3.06cqw;line-height:1.36;">Agradecemos a todos<br>pela presença!<div class="pv-c-subtitulo" style="font-size:1.94cqw;">Nossos canais de atendimento</div></div>' +
        campo("pv-c-whats", spec.whats, 16, 2.3) +
        campo("pv-c-tel", spec.tel, 16, 2.3) +
        campo("pv-c-site", spec.site, 26, 2.4) +
        campo("pv-c-email", spec.email, 26, 2.7) +
        (endereco ? '<div class="pv-c-end">' + pvEnderecoHtml(endereco) + '</div>' : "") +
        '</div>'
      );
    }
    case "placeholder":
    default:
      return (
        '<div class="pv-slide">' + pvHeaderBar(spec.titulo || "") + '' +
        '<div class="pv-body" style="display:flex;align-items:center;justify-content:center;">' +
        '<span class="pv-empty">' + pvEsc(spec.msg || "Sem dados ainda") + '</span></div></div>'
      );
  }
}



// Prévia simplificada: só a lista de miniaturas, cada uma já mostra o slide inteiro —
// sem visor grande redundante em cima, sem contador de página, sem setas.
let pvDebounceTimer = null;
let pvFsSpecs = [];
let pvFsIndex = 0;

// Prévia simplificada: só a lista de miniaturas, cada uma já mostra o slide inteiro —
// sem visor grande redundante em cima, sem contador de página, sem setas.
// Clicar numa miniatura abre ela em tela cheia (com setas pra navegar entre os slides).
let pvDragIdx = null;
function refreshPreview() {
  ajustesCarregar();
  const specs = buildSlideSpecs();
  pvFsSpecs = specs;
  const thumbsEl = document.getElementById("pvThumbs");
  thumbsEl.innerHTML = "";
  specs.forEach((spec, i) => {
    const wrap = document.createElement("div");
    wrap.className = "pv-thumb-wrap";
    wrap.innerHTML =
      '<div class="pv-thumb-item">' + renderSlideHtml(spec) + '</div>' +
      '<div class="pv-thumb-label">' + (i + 1) + ". " + pvEsc(spec.label) + '</div>' + ajusteControlesHtml(spec, specs, i, false);
    const item = wrap.querySelector(".pv-thumb-item");
    item.addEventListener("click", () => openPvFullscreen(i));
    if (spec.aj) {
      // arrastar um slide de tabela sobre o vizinho para juntá-los
      item.draggable = true;
      item.addEventListener("dragstart", (e) => { pvDragIdx = i; wrap.classList.add("pv-dragging"); try { e.dataTransfer.setData("text/plain", "slide" + i); e.dataTransfer.effectAllowed = "move"; } catch (err) { /* ok */ } });
      item.addEventListener("dragend", () => { pvDragIdx = null; document.querySelectorAll(".pv-drop-ok,.pv-drop-no,.pv-dragging").forEach((el) => el.classList.remove("pv-drop-ok", "pv-drop-no", "pv-dragging")); });
    }
    wrap.addEventListener("dragover", (e) => {
      if (pvDragIdx === null || pvDragIdx === i) return;
      const r = validarJuncao(specs, pvDragIdx, i);
      wrap.classList.toggle("pv-drop-ok", r.ok);
      wrap.classList.toggle("pv-drop-no", !r.ok);
      wrap.dataset.msg = r.ok ? "Soltar para juntar" : (r.motivo || "Não dá para juntar");
      if (r.ok) { e.preventDefault(); e.dataTransfer.dropEffect = "move"; }
    });
    wrap.addEventListener("dragleave", (e) => { if (!wrap.contains(e.relatedTarget)) wrap.classList.remove("pv-drop-ok", "pv-drop-no"); });
    wrap.addEventListener("drop", (e) => {
      if (pvDragIdx === null) return;
      e.preventDefault();
      const r = validarJuncao(specs, pvDragIdx, i);
      const de = pvDragIdx;
      pvDragIdx = null;
      if (r.ok) aplicarJuncao(r, specs[Math.min(de, i)], specs[Math.max(de, i)]);
    });
    thumbsEl.appendChild(wrap);
  });
  if (!document.getElementById("pvFsOverlay").classList.contains("hidden")) {
    pvFsIndex = Math.min(pvFsIndex, pvFsSpecs.length - 1);
    renderPvFullscreen();
  }
}

// Pode juntar os slides i e j? Devolve {ok, op, ...} ou {ok:false, motivo}.
// Duas operações distintas: "pagina" (duas páginas da mesma tabela => mais linhas por slide) e
// "bloco" (dois grupos/contas vizinhos => um slide só com título "A + B").
function validarJuncao(specs, i, j) {
  if (i === j || i < 0 || j < 0 || i >= specs.length || j >= specs.length) return { ok: false };
  const a = specs[Math.min(i, j)], b = specs[Math.max(i, j)];
  if (Math.abs(i - j) !== 1) return { ok: false, motivo: "Só dá para juntar slides vizinhos." };
  if (!a.aj || !b.aj) return { ok: false, motivo: "Este slide não é uma tabela ajustável." };
  if (a.aj.familia !== b.aj.familia) return { ok: false, motivo: "São de tipos diferentes (não dá para misturar)." };
  const cap = capLinhasTabela(a.aj.tipo);
  if (a.aj.chave === b.aj.chave) {
    const soma = a.aj.linhasPagina + b.aj.linhasPagina;
    if (soma > cap) return { ok: false, motivo: "Juntos dariam " + soma + " linhas, mais que o máximo legível (" + cap + ")." };
    return { ok: true, op: "pagina", chave: a.aj.chave, linhas: Math.max(soma, a.aj.L) };
  }
  if (a.aj.pagina !== a.aj.paginas - 1 || b.aj.pagina !== 0) return { ok: false, motivo: "Junte só os slides das pontas: o último de um grupo com o primeiro do seguinte." };
  const sep = SEPS_POR_BLOCO[a.aj.tipo];
  const total = (a.aj.fundido ? a.aj.nTotal : a.aj.nTotal + sep) + (b.aj.fundido ? b.aj.nTotal : b.aj.nTotal + sep);
  if (total > cap) return { ok: false, motivo: "Juntos dariam " + total + " linhas, mais que o máximo legível (" + cap + ")." };
  return { ok: true, op: "bloco", guest: b.aj.chave, pred: a.aj.ultimaChave, chaveUnidade: a.aj.chave };
}
let AJ_DESFAZER = null, AJ_TOAST_T = null;
function ajToast(msg, desfazer) {
  let el = document.getElementById("pvToast");
  if (!el) { el = document.createElement("div"); el.id = "pvToast"; el.className = "pv-toast"; document.body.appendChild(el); }
  el.innerHTML = '<span></span>' + (desfazer ? '<button type="button" id="pvToastDesfazer">Desfazer</button>' : "");
  el.firstChild.textContent = msg;
  el.hidden = false;
  clearTimeout(AJ_TOAST_T); AJ_TOAST_T = setTimeout(() => { el.hidden = true; }, 9000);
  const bt = document.getElementById("pvToastDesfazer");
  if (bt) bt.addEventListener("click", () => { if (AJ_DESFAZER) { AJUSTES = JSON.parse(AJ_DESFAZER); AJ_DESFAZER = null; ajustesSalvar(); el.hidden = true; refreshPreview(); } });
}
function aplicarJuncao(r, a, b) {
  AJ_DESFAZER = JSON.stringify(AJUSTES);
  let chaveAlvo = r.chave || r.chaveUnidade;
  if (r.op === "pagina") AJUSTES[r.chave] = Object.assign({}, AJUSTES[r.chave], { linhas: r.linhas });
  else AJUSTES[r.guest] = Object.assign({}, AJUSTES[r.guest], { juntarCom: r.pred });
  ajustesSalvar();
  refreshPreview();
  const nomeA = (a.aj.nomes || [a.label]).join(" + "), nomeB = (b.aj.nomes || [b.label]).join(" + ");
  ajToast(r.op === "bloco" ? "Juntou " + nomeB + " em " + nomeA + "." : "Juntou as duas páginas (" + r.linhas + " linhas por slide).", true);
  if (!document.getElementById("pvFsOverlay").classList.contains("hidden")) {
    const k = pvFsSpecs.findIndex((x) => x.aj && x.aj.chave === chaveAlvo);
    if (k >= 0) { pvFsIndex = k; renderPvFullscreen(); }
  }
}

// Barra de ajuste de uma tabela: linhas por slide (+/−), separar o que foi juntado, restaurar.
// Na tela cheia (cheia=true) também tem "juntar com o slide anterior/seguinte" (lá não dá para arrastar).
function ajusteControlesHtml(spec, specs, idx, cheia) {
  const a = spec.aj;
  if (!a) return "";
  const cap = capLinhasTabela(a.tipo);
  const podeMais = a.L < cap && a.L < a.nTotal;
  const podeMenos = a.L > 1 && a.paginas < a.nTotal;
  const dado = ' data-aj-idx="' + idx + '"';
  const btn = (acao, rotulo, ok, titulo) => '<button type="button" data-aj="' + acao + '"' + dado + (ok ? "" : " disabled") + ' title="' + pvEsc(titulo || "") + '">' + rotulo + "</button>";
  let h = '<div class="pv-ajuste' + (cheia ? " pv-ajuste-cheia" : "") + '">' +
    '<span class="pv-aj-t" title="Quantas linhas cabem em cada slide desta tabela. Mais linhas = menos slides e fonte menor (mínimo 12pt).">Linhas/slide</span>' +
    btn("menos", "&minus;", podeMenos, "Menos linhas por slide") + '<b>' + a.L + '</b>' + btn("mais", "+", podeMais, "Mais linhas por slide (fonte menor)");
  if (cheia) {
    const ant = validarJuncao(specs, idx, idx - 1), prox = validarJuncao(specs, idx, idx + 1);
    h += btn("juntar-ant", "Juntar com o anterior", ant.ok, ant.ok ? "Junta este slide ao slide anterior" : (ant.motivo || "Não há slide anterior para juntar"));
    h += btn("juntar-prox", "Juntar com o próximo", prox.ok, prox.ok ? "Junta o slide seguinte a este" : (prox.motivo || "Não há slide seguinte para juntar"));
  }
  if (a.fundido) h += btn("separar", "Separar", true, "Desfaz a última junção deste slide");
  if (a.manual) h += btn("reset", "Restaurar linhas", true, "Voltar à paginação automática");
  if (!cheia && !a.fundido && !a.manual) h += '<span class="pv-aj-dica" title="Arraste esta miniatura sobre a vizinha para juntar as duas">⠿ arraste para juntar</span>';
  return h + '<span class="pv-aj-f">' + spec.fontePt + 'pt</span></div>';
}
function ajusteAcao(acao, idx) {
  const specs = pvFsSpecs, spec = specs[idx];
  if (!spec || !spec.aj) return;
  const a = spec.aj;
  AJ_DESFAZER = JSON.stringify(AJUSTES);
  const toastAntigo = document.getElementById("pvToast");
  if (toastAntigo && !acao.startsWith("juntar")) { toastAntigo.hidden = true; AJ_DESFAZER = null; } // o "Desfazer" do aviso vale só para a junção que ele anunciou
  let alvo = a.chave;
  if (acao === "mais") AJUSTES[a.chave] = Object.assign({}, AJUSTES[a.chave], { linhas: a.L + 1 });
  else if (acao === "menos") AJUSTES[a.chave] = Object.assign({}, AJUSTES[a.chave], { linhas: a.L - 1 });
  else if (acao === "reset") { const x = Object.assign({}, AJUSTES[a.chave]); delete x.linhas; if (Object.keys(x).length) AJUSTES[a.chave] = x; else delete AJUSTES[a.chave]; }
  else if (acao === "separar") { const x = Object.assign({}, AJUSTES[a.ultimaChave]); delete x.juntarCom; if (Object.keys(x).length) AJUSTES[a.ultimaChave] = x; else delete AJUSTES[a.ultimaChave]; }
  else if (acao === "juntar-ant" || acao === "juntar-prox") {
    const j = acao === "juntar-ant" ? idx - 1 : idx + 1;
    const r = validarJuncao(specs, idx, j);
    if (!r.ok) return;
    aplicarJuncao(r, specs[Math.min(idx, j)], specs[Math.max(idx, j)]);
    return;
  } else return;
  ajustesSalvar();
  refreshPreview();
  if (!document.getElementById("pvFsOverlay").classList.contains("hidden")) {
    const k = pvFsSpecs.findIndex((x) => x.aj && x.aj.chave === alvo);
    if (k >= 0) { pvFsIndex = k; renderPvFullscreen(); }
  }
}
document.addEventListener("click", (e) => {
  const b = e.target.closest("button[data-aj]");
  if (!b || b.disabled) return;
  e.stopPropagation();
  ajusteAcao(b.dataset.aj, parseInt(b.dataset.ajIdx, 10));
});

function renderPvFullscreen() {
  const spec = pvFsSpecs[pvFsIndex];
  document.getElementById("pvFsSlide").innerHTML = renderSlideHtml(spec);
  document.getElementById("pvFsLabel").textContent = (pvFsIndex + 1) + " / " + pvFsSpecs.length + " — " + spec.label;
  document.getElementById("pvFsAjuste").innerHTML = ajusteControlesHtml(spec, pvFsSpecs, pvFsIndex, true);
  document.getElementById("pvFsPrev").disabled = pvFsIndex === 0;
  document.getElementById("pvFsNext").disabled = pvFsIndex === pvFsSpecs.length - 1;
}
function openPvFullscreen(idx) {
  pvFsIndex = idx;
  renderPvFullscreen();
  document.getElementById("pvFsOverlay").classList.remove("hidden");
}
function closePvFullscreen() {
  document.getElementById("pvFsOverlay").classList.add("hidden");
}
document.getElementById("pvFsClose").addEventListener("click", closePvFullscreen);
document.getElementById("pvFsOverlay").addEventListener("click", (e) => {
  if (e.target.id === "pvFsOverlay") closePvFullscreen();
});
document.getElementById("pvFsPrev").addEventListener("click", () => {
  if (pvFsIndex > 0) { pvFsIndex--; renderPvFullscreen(); }
});
document.getElementById("pvFsNext").addEventListener("click", () => {
  if (pvFsIndex < pvFsSpecs.length - 1) { pvFsIndex++; renderPvFullscreen(); }
});
document.addEventListener("keydown", (e) => {
  if (document.getElementById("pvFsOverlay").classList.contains("hidden")) return;
  if (e.key === "Escape") closePvFullscreen();
  else if (e.key === "ArrowLeft" && pvFsIndex > 0) { pvFsIndex--; renderPvFullscreen(); }
  else if (e.key === "ArrowRight" && pvFsIndex < pvFsSpecs.length - 1) { pvFsIndex++; renderPvFullscreen(); }
});

function schedulePreviewRefresh() {
  clearTimeout(pvDebounceTimer);
  pvDebounceTimer = setTimeout(refreshPreview, 180);
}
document.addEventListener("input", schedulePreviewRefresh);
document.addEventListener("change", schedulePreviewRefresh);
refreshPreview();
/* ===================== Navegação por abas (sidebar, desktop) ===================== */
function switchPanel(n) {
  document.querySelectorAll(".panel").forEach((p) => p.classList.toggle("active", p.dataset.panel === String(n)));
  document.querySelectorAll(".sidebar nav button[data-panel]").forEach((b) => b.classList.toggle("active", b.dataset.panel === String(n)));
  if (String(n) === "9") buildResumo();
}
// Delegação: os botões dos tópicos são criados dinamicamente; abas indisponíveis não respondem ao clique.
document.getElementById("sidebarNav").addEventListener("click", (e) => {
  const btn = e.target.closest("button[data-panel]");
  if (!btn || btn.classList.contains("nav-disabled")) return;
  switchPanel(btn.dataset.panel);
  btn.blur(); // o clique não deve manter a sidebar aberta: ela fecha assim que o mouse sai
});

/* ===================== Resumo (última aba) ===================== */
function statusHtml(ok, textoOk, textoPendente, opcional) {
  if (ok) return '<span class="resumo-status ok">✓ ' + textoOk + '</span>';
  return '<span class="resumo-status ' + (opcional ? "muted" : "pending") + '">— ' + textoPendente + '</span>';
}

function buildResumo() {
  const condominio = document.getElementById("f_condominio").value.trim() || "—";
  const tipo = document.getElementById("f_tipo").value;
  const dataExtenso = dateToExtensoPT(document.getElementById("f_data").value) || "—";
  const local = document.getElementById("f_local").value.trim() || "—";
  const conv1 = fmtHoraBR(document.getElementById("f_conv1").value) || "—";
  const conv2 = fmtHoraBR(document.getElementById("f_conv2").value) || "—";
  const itensPauta = getPautaItemsComRole();
  const pauta = itensPauta.map((it) => it.texto);
  const temPrestacao = itensPauta.some((it) => it.role === "prestacao");
  const temPrevisao = itensPauta.some((it) => it.role === "previsao");
  const temAssuntos = itensPauta.some((it) => it.role === "assuntos");
  const topicos = itensPauta.filter((it) => it.role === "topico");
  const foraDaPauta = '<span class="resumo-status muted">— fora da pauta (não será gerado)</span>';
  const tel = document.getElementById("f_tel").value.trim() || "—";
  const whats = document.getElementById("f_whats").value.trim() || "—";
  const email = document.getElementById("f_email").value.trim() || "—";
  const site = document.getElementById("f_site").value.trim() || "—";
  const endereco = document.getElementById("f_endereco").value.trim() || "—";

  let html = "";

  html += '<div class="resumo-section"><h4>Dados da assembleia</h4>';
  html += '<div class="resumo-row"><span class="k">Condomínio</span><span class="v">' + condominio + '</span></div>';
  html += '<div class="resumo-row"><span class="k">Tipo</span><span class="v">' + tipo + '</span></div>';
  html += '<div class="resumo-row"><span class="k">Data</span><span class="v">' + dataExtenso + '</span></div>';
  html += '<div class="resumo-row"><span class="k">Local</span><span class="v">' + local + '</span></div>';
  html += '<div class="resumo-row"><span class="k">1ª / 2ª convocação</span><span class="v">' + conv1 + ' / ' + conv2 + '</span></div>';
  html += '</div>';

  html += '<div class="resumo-section"><h4>Pauta (' + pauta.length + ' itens)</h4>';
  pauta.forEach((p, i) => (html += '<p class="resumo-pauta">' + (i + 1) + '. ' + pvEsc(p) + '</p>'));
  html += '</div>';

  html += '<div class="resumo-section"><h4>Relatórios e dados financeiros</h4>';
  html += '<div class="resumo-row"><span class="k">Relatório 11A</span><span class="v">' + (temPrestacao ? statusHtml(!!PARSED.a11, "enviado", "pendente") : foraDaPauta) + '</span></div>';
  html += '<div class="resumo-row"><span class="k">Previsão Orçamentária</span><span class="v">' + (temPrevisao ? statusHtml(!!PARSED.prev, "enviada", "pendente") : foraDaPauta) + '</span></div>';
  html += '<div class="resumo-row"><span class="k">Relatório 16B</span><span class="v">' + (temPrestacao ? statusHtml(!!PARSED.b16, "preenchido", "opcional — não preenchido", true) : foraDaPauta) + '</span></div>';
  html += '</div>';

  html += '<div class="resumo-section"><h4>Inadimplência</h4>';
  if (temPrestacao) {
    html += '<div class="resumo-row"><span class="k">Valor acumulado</span><span class="v">' + (document.getElementById("f_inad_valor").value.trim() || "—") + '</span></div>';
    html += '<div class="resumo-row"><span class="k">Cobranças / Unidades</span><span class="v">' + (document.getElementById("f_inad_cobrancas").value.trim() || "—") + ' / ' + (document.getElementById("f_inad_unidades").value.trim() || "—") + (document.getElementById("f_inad_unidades_total").value.trim() ? " (de " + document.getElementById("f_inad_unidades_total").value.trim() + ")" : "") + '</span></div>';
    html += '<div class="resumo-row"><span class="k">% orgânica / garantidora / total</span><span class="v">' + (document.getElementById("f_inad_organica").value.trim() || "—") + ' / ' + (document.getElementById("f_inad_garantidora").value.trim() || "—") + ' / ' + (document.getElementById("f_inad_total").value.trim() || "—") + '</span></div>';
  } else {
    html += '<div class="resumo-row"><span class="k">Inadimplência</span><span class="v">' + foraDaPauta + '</span></div>';
  }
  html += '</div>';

  if (topicos.length || temAssuntos) {
    html += '<div class="resumo-section"><h4>Tópicos e assuntos gerais</h4>';
    topicos.forEach((t) => {
      const qtdAnexos = (t.imagens && t.imagens.length) || 0;
      const temAlgo = !!(t.conteudo || qtdAnexos);
      const rotuloOk = t.conteudo && qtdAnexos
        ? "com texto e " + qtdAnexos + " anexo(s)"
        : (qtdAnexos ? qtdAnexos + " slide(s) anexado(s)" : "com conteúdo");
      html += '<div class="resumo-row"><span class="k">' + pvEsc(t.texto) + '</span><span class="v">' + statusHtml(temAlgo, rotuloOk, "sem conteúdo — só o slide divisor", true) + '</span></div>';
    });
    if (temAssuntos) {
      html += '<div class="resumo-row"><span class="k">Assuntos gerais</span><span class="v"><span class="resumo-status muted">— slide só com o título</span></span></div>';
    }
    html += '</div>';
  }

  html += '<div class="resumo-section"><h4>Canais de atendimento</h4>';
  html += '<div class="resumo-row"><span class="k">Telefone / WhatsApp</span><span class="v">' + tel + ' / ' + whats + '</span></div>';
  html += '<div class="resumo-row"><span class="k">E-mail</span><span class="v">' + email + '</span></div>';
  html += '<div class="resumo-row"><span class="k">Site</span><span class="v">' + site + '</span></div>';
  html += '<div class="resumo-row"><span class="k">Endereço</span><span class="v">' + endereco + '</span></div>';
  html += '</div>';

  document.getElementById("resumoWrap").innerHTML = html;
}

/* ===== Canais de atendimento: campos travados por padrão (valores padrão da Jetro) ===== */
(function () {
  const ids = ["f_tel", "f_whats", "f_email", "f_site", "f_endereco"];
  const btn = document.getElementById("btnTravaCanais");
  const body = document.getElementById("canaisBody");
  function aplicar(travado) {
    ids.forEach((id) => { document.getElementById(id).disabled = travado; });
    body.classList.toggle("canais-travados", travado);
    btn.textContent = travado ? "🔒 Destravar edição" : "🔓 Travar novamente";
  }
  btn.addEventListener("click", () => aplicar(!document.getElementById(ids[0]).disabled));
  aplicar(true);
})();

/* ===== 16B: envio de arquivo (PDF ou planilha) como alternativa a colar o texto ===== */
// Linhas da planilha como texto no padrão brasileiro. Lê o VALOR da célula (não o texto formatado, que muda
// com o formato do Excel: "12,500.50", "12500.50"...): números com casas decimais ou formato de dinheiro viram
// "12.500,50"; inteiros comuns (códigos, anos) ficam como estão.
function planilha16BParaTexto(ws) {
  if (!ws || !ws["!ref"]) return "";
  const rg = XLSX.utils.decode_range(ws["!ref"]);
  const linhas = [];
  for (let r = rg.s.r; r <= rg.e.r; r++) {
    const partes = [];
    const cels = [];
    for (let c = rg.s.c; c <= rg.e.c; c++) cels.push(ws[XLSX.utils.encode_cell({ r, c })]);
    // linha de tabela = rótulo (texto) + 2 ou mais números: todos os números da linha são valores em R$
    const nNum = cels.filter((x) => x && x.t === "n").length;
    const temRotulo = cels.some((x) => x && x.t === "s" && String(x.v).trim());
    const linhaDeValores = temRotulo && nNum >= 2;
    for (const cel of cels) {
      if (!cel || cel.v === undefined || cel.v === null || cel.v === "") continue;
      if (cel.t === "n") {
        const z = String(cel.z || "");
        const dinheiro = linhaDeValores || !Number.isInteger(cel.v) || /0[.,]0/.test(z) || /R\$|\$/.test(z);
        partes.push(dinheiro ? cel.v.toLocaleString("pt-BR", { minimumFractionDigits: 2, maximumFractionDigits: 2 }) : String(cel.v));
      } else {
        partes.push(String(cel.w !== undefined ? cel.w : cel.v).trim());
      }
    }
    const linha = partes.filter(Boolean).join(" ");
    if (linha) linhas.push(linha);
  }
  return linhas.join("\n");
}

// Converte o arquivo em texto (uma linha por linha da tabela) e reaproveita parse16B.
async function arquivoParaTexto16B(file) {
  const nome = file.name.toLowerCase();
  if (/\.(xlsx|xls|csv)$/.test(nome)) {
    const wb = await readWorkbook(file);
    // Procura a aba que tem a tabela do 16B (a 1ª aba nem sempre é ela).
    const abas = wb.SheetNames.map((n) => wb.Sheets[n]);
    const textos = abas.map((ws) => planilha16BParaTexto(ws));
    const idx = textos.findIndex((t) => /saldo/i.test(t) && /\d,\d{2}\b/.test(t));
    return textos[idx >= 0 ? idx : 0];
  }
  if (/\.pdf$/.test(nome)) {
    await carregarPdfJs();
    const buf = await file.arrayBuffer();
    const pdf = await window.pdfjsLib.getDocument({ data: buf }).promise;
    let texto = "";
    for (let p = 1; p <= pdf.numPages; p++) {
      const content = await (await pdf.getPage(p)).getTextContent();
      // agrupa os pedaços de texto pela posição vertical (mesma linha visual) e ordena da esquerda pra direita
      const porLinha = {};
      content.items.forEach((it) => {
        if (!it.str.trim()) return;
        const y = Math.round(it.transform[5] / 3);
        (porLinha[y] = porLinha[y] || []).push({ x: it.transform[4], s: it.str.trim() });
      });
      Object.keys(porLinha).map(Number).sort((a, b) => b - a).forEach((y) => {
        texto += porLinha[y].sort((a, b) => a.x - b.x).map((i) => i.s).join(" ") + "\n";
      });
    }
    if (texto.trim().length < 20) throw new Error("Este PDF parece ser uma imagem escaneada (não tem texto selecionável). Envie a versão em planilha (.xlsx).");
    return texto;
  }
  throw new Error("Formato não suportado. Envie um PDF, .xlsx, .xls ou .csv.");
}

/* ===== PDF -> grade de linhas/colunas (11A e Previsão): reaproveita os parsers de planilha ===== */
// Lê o texto do PDF com a posição de cada pedaço, agrupa em linhas visuais e descobre as colunas pelo alinhamento dos números
// e pelo cabeçalho. Devolve um workbook (1 aba) igual ao de uma planilha: parse11A / parsePrevisao / poImportar3Col funcionam como estão.
// layout "11a": coluna 0 = rótulo, depois os meses. layout "previsao": coluna 0 = código (1.2), 1 = descrição, depois os valores.
const PDF_NUM_RE = /^\(?-?\s*(R\$\s*)?\d{1,3}(\.\d{3})*(,\d+)?\)?%?$|^\(?-?\s*(R\$\s*)?\d+([.,]\d+)?\)?%?$/;
function pdfNumero(txt) {
  let t = String(txt).trim();
  if (!PDF_NUM_RE.test(t)) return null;
  const neg = /^\(.*\)$/.test(t) || /^-/.test(t);
  const pct = /%$/.test(t);
  t = t.replace(/[()R$%\s-]/g, "");
  if (/,/.test(t)) t = t.replace(/\./g, "").replace(",", ".");
  else if (/^\d{1,3}(\.\d{3})+$/.test(t)) t = t.replace(/\./g, "");
  const n = parseFloat(t);
  if (isNaN(n)) return null;
  return (neg ? -n : n) / (pct ? 100 : 1);
}
async function pdfParaGrade(file, layout) {
  await carregarPdfJs();
  const buf = await file.arrayBuffer();
  const pdf = await window.pdfjsLib.getDocument({ data: buf }).promise;
  const linhas = []; // cada linha: lista de células {x, r, h, t}, páginas em sequência
  let nItens = 0;
  for (let pg = 1; pg <= pdf.numPages; pg++) {
    const content = await (await pdf.getPage(pg)).getTextContent();
    const itens = content.items.filter((it) => it.str && it.str.trim()).map((it) => ({ x: it.transform[4], y: it.transform[5], w: it.width || 0, h: Math.abs(it.transform[3]) || it.height || 10, t: it.str.trim().normalize("NFC") }));
    nItens += itens.length;
    itens.sort((a, b) => b.y - a.y || a.x - b.x);
    const porLinha = [];
    itens.forEach((it) => {
      const ult = porLinha[porLinha.length - 1];
      if (ult && Math.abs(ult.y - it.y) <= Math.max(2, it.h * 0.45)) ult.its.push(it);
      else porLinha.push({ y: it.y, its: [it] });
    });
    porLinha.forEach((l) => {
      l.its.sort((a, b) => a.x - b.x);
      const cels = [];
      l.its.forEach((it) => { // pedaços colados viram uma célula só; lacunas maiores separam células
        const c = cels[cels.length - 1];
        if (c && it.x - c.r >= -1 && it.x - c.r < it.h * 0.12) { c.t += it.t; c.r = it.x + it.w; } // só pedaços praticamente encostados (cada célula do Excel já vem como um item)
        else cels.push({ x: it.x, r: it.x + it.w, h: it.h, t: it.t });
      });
      linhas.push(cels);
    });
  }
  if (nItens < 5) throw new Error("Este PDF parece ser uma imagem escaneada (não tem texto selecionável). Envie a versão em planilha (.xlsx).");

  const todas = [];
  linhas.forEach((cs, li) => cs.forEach((c) => { c.li = li; todas.push(c); }));
  const cerquilhas = todas.filter((c) => /^#{3,}$/.test(c.t)).length;
  if (cerquilhas) throw new Error("O PDF tem " + cerquilhas + " célula(s) com ##### (a coluna ficou estreita demais no Excel e o valor não foi impresso). Alargue as colunas, gere o PDF de novo ou envie a planilha (.xlsx).");
  const ehNum = (c) => pdfNumero(c.t) !== null;
  // coluna do rótulo = posição (esquerda) mais comum entre os textos que não são número
  const cont = {};
  todas.filter((c) => !ehNum(c)).forEach((c) => { const k = Math.round(c.x / 2) * 2; cont[k] = (cont[k] || 0) + 1; });
  const ks = Object.keys(cont).map(Number).sort((a, b) => cont[b] - cont[a] || a - b);
  if (!ks.length) throw new Error("Não consegui identificar as colunas do PDF.");
  const rotuloX = ks[0];
  const comCodigo = layout === "previsao";
  const ehCodigo = (c) => comCodigo && ehNum(c) && c.x < rotuloX - 3;
  // linha do cabeçalho: previsão => a que tem "Descrição"; 11A => a com 3+ meses (Mmm/AAAA)
  const mesRe = /^[A-Za-zçÇ]{3}\/\d{4}$/;
  const iCab = linhas.findIndex((cs) => layout === "previsao" ? cs.some((c) => /^descri[çc][ãa]o$/i.test(c.t)) : cs.filter((c) => mesRe.test(c.t)).length >= 3);
  if (iCab < 0) throw new Error(layout === "previsao" ? 'Não encontrei o cabeçalho ("Descrição", "Atual", "Previsto"...) no PDF.' : "Não encontrei a linha com os meses (ex: Dez/2025) no PDF do 11A.");
  const cab = linhas[iCab];
  // região principal = da linha após o cabeçalho até o título do rateio (as tabelas de rateio têm colunas próprias)
  let iFimPrincipal = linhas.length;
  for (let i = iCab + 1; i < linhas.length; i++) {
    if (linhas[i].some((c) => c.x <= rotuloX + 3 && /^(rateio\b|previs[ãa]o\b.*(-|rateio))/i.test(c.t))) { iFimPrincipal = i; break; }
  }
  const principal = (c) => c.li > iCab && c.li < iFimPrincipal;
  // colunas de valor: números agrupados pela borda direita (alinhados à direita compartilham a borda)
  const nums = todas.filter((c) => ehNum(c) && !ehCodigo(c) && c.x >= rotuloX - 3).sort((a, b) => a.r - b.r);
  const grupos = [];
  nums.forEach((c) => { const g = grupos[grupos.length - 1]; if (g && c.r - g.fim <= 6) { g.fim = c.r; g.cels.push(c); } else grupos.push({ fim: c.r, cels: [c] }); });
  let cols = grupos.map((g) => ({ r: g.cels.reduce((a, c) => a + c.r, 0) / g.cels.length, esq: Math.min(...g.cels.map((c) => c.x)), princ: g.cels.filter(principal).length, n: g.cels.length }));
  // colunas centralizadas geram vários grupos de borda direita; colunas diferentes nunca se sobrepõem, então grupos que se sobrepõem são a mesma coluna
  cols.sort((a, b) => a.esq - b.esq);
  cols = cols.reduce((acc, k) => {
    const u = acc[acc.length - 1];
    if (u && k.esq < u.r - 0.5) { u.r = Math.max(u.r, k.r); u.princ += k.princ; u.n += k.n; } else acc.push(Object.assign({}, k));
    return acc;
  }, []);
  const sobrepoe = (c, k) => Math.min(c.r, k.r) - Math.max(c.x, k.esq);
  const dist = (c, k) => Math.abs((c.x + c.r) / 2 - (k.esq + k.r) / 2);
  const melhorCol = (c, lista) => {
    let melhor = null, bestO = -Infinity, bestD = Infinity;
    lista.forEach((k) => { const o = sobrepoe(c, k), d = dist(c, k); if (o > bestO + 1e-6 || (Math.abs(o - bestO) < 1e-6 && d < bestD)) { bestO = o; bestD = d; melhor = k; } });
    return melhor;
  };
  const colsPrinc = () => cols.filter((k) => k.princ > 0);
  // cabeçalho que não cai sobre nenhuma coluna de números da região principal cria a sua própria coluna (ex.: "Descrição p/ Cálculo")
  cab.forEach((c) => {
    if (c.x <= rotuloX + 3) return;
    if (!colsPrinc().some((k) => sobrepoe(c, k) > 0)) cols.push({ r: c.r, esq: c.x, princ: 1 });
  });
  cols.sort((a, b) => a.r - b.r);
  cols.forEach((k, i) => { k.i = i; });
  const limiteDesc = colsPrinc().length ? Math.min(...colsPrinc().map((k) => k.esq)) : Infinity; // a partir daqui começam os valores
  const off = comCodigo ? 2 : 1;
  const assinatura = (cs) => cs.map((c) => c.t.toLowerCase()).join("|");
  const sigCab = assinatura(cab);
  const aoa = [];
  linhas.forEach((cs, i) => {
    if (i > iCab && assinatura(cs) === sigCab) return; // cabeçalho repetido nas páginas seguintes
    const lin = new Array(off + cols.length).fill(null);
    cs.forEach((c) => {
      const num = pdfNumero(c.t);
      if (ehCodigo(c)) { lin[0] = Number.isInteger(num) && !/[.,]/.test(c.t) ? num : c.t; return; }
      // texto único na linha (ignorando o código): título centralizado/mesclado (rateio, nome de grupo) vale como rótulo
      const soTexto = cs.filter((x) => !ehCodigo(x));
      const tituloSolto = layout === "previsao" && soTexto.length === 1 && num === null && i !== iCab;
      if (!cols.length || tituloSolto || c.x <= rotuloX + 3 || (i !== iCab && num === null && c.x < limiteDesc - 2 && lin[off - 1] === null)) { // rótulo (o 1º texto da linha; os seguintes vão para a coluna mais próxima)
        lin[off - 1] = lin[off - 1] ? lin[off - 1] + " " + c.t : c.t;
        return;
      }
      const k = melhorCol(c, i === iCab ? colsPrinc() : cols);
      if (!k) return;
      lin[off + k.i] = num !== null ? num : c.t;
    });
    const ant = aoa[aoa.length - 1];
    const comecaMinusculo = typeof lin[off - 1] === "string" && /^[a-zà-ú(]/.test(lin[off - 1]);
    if (comCodigo && ant && lin[0] === null && comecaMinusculo && i > iCab + 1 && typeof ant[0] !== "undefined" && ant[0] !== null && typeof ant[off - 1] === "string") {
      const antSemValor = ant.slice(off).every((x) => x === null);
      const linSoRotulo = lin.every((x, j) => j === off - 1 || x === null);
      if (linSoRotulo && !antSemValor) { ant[off - 1] += " " + lin[off - 1]; return; } // rótulo longo quebrado: 2ª linha sem valores
      if (antSemValor) { // rótulo longo quebrado: a 1ª linha só tem código e texto, os valores vêm na 2ª
        ant[off - 1] += " " + lin[off - 1];
        for (let j = off; j < lin.length; j++) if (lin[j] !== null) ant[j] = lin[j];
        return;
      }
    }
    if (lin.some((x) => x !== null)) aoa.push(lin);
  });
  const ws = XLSX.utils.aoa_to_sheet(aoa);
  const wb = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(wb, ws, "PDF");
  return wb;
}
// Planilha (SheetJS) ou PDF: devolve sempre um workbook para os parsers.
async function lerArquivoTabela(file, layout) {
  if (/\.pdf$/i.test(file.name)) {
    if (!PDF_GRADE_HABILITADO) throw new Error("Neste relatório o envio em PDF ainda não está disponível. Envie a planilha (.xlsx ou .xls).");
    return pdfParaGrade(file, layout);
  }
  return readWorkbook(file);
}

let pdfJsPromise = null;
function carregarPdfJs() {
  if (window.pdfjsLib) return Promise.resolve();
  if (pdfJsPromise) return pdfJsPromise;
  pdfJsPromise = new Promise((resolve, reject) => {
    const sc = document.createElement("script");
    sc.src = "https://cdn.jsdelivr.net/npm/pdfjs-dist@3.11.174/build/pdf.min.js";
    sc.onload = () => {
      window.pdfjsLib.GlobalWorkerOptions.workerSrc = "https://cdn.jsdelivr.net/npm/pdfjs-dist@3.11.174/build/pdf.worker.min.js";
      resolve();
    };
    sc.onerror = () => { pdfJsPromise = null; reject(new Error("Não consegui carregar o leitor de PDF (sem internet?).")); };
    document.head.appendChild(sc);
  });
  return pdfJsPromise;
}

setupDropzone("dz16b", "file16b", async (file) => {
  const dz = document.getElementById("dz16b");
  dz._mostrar(file.name);
  try {
    const texto = await arquivoParaTexto16B(file);
    const data = parse16B(texto);
    PARSED.b16 = data;
    document.getElementById("f_16b_periodo").value = data.periodo || ""; // período lido do próprio relatório
    dz.classList.add("ok");
    avisarCondominio("status16b", nomeCondominioDoDocumento(texto.split("\n")));
    setStatus("status16b", true, (data.contas.length + data.grupos.length) + " linhas reconhecidas no arquivo");
    renderPreview16B(data);
    if (typeof refreshPreview === "function") refreshPreview();
  } catch (err) {
    PARSED.b16 = null;
    dz.classList.remove("ok");
    avisarCondominio("status16b", null);
    reportError("status16b", "Não consegui ler o 16B", "Erro: " + err.message, "Confira se é o relatório 16B da Superlógica (PDF ou planilha). Se o PDF for uma imagem escaneada, tente enviar a versão em planilha.");
    document.getElementById("preview16b").innerHTML = "";
  }
}, () => {
  PARSED.b16 = null;
  document.getElementById("preview16b").innerHTML = "";
  document.getElementById("f_16b_periodo").value = "";
  document.getElementById("status16b").textContent = "";
  document.getElementById("status16b").className = "status-line";
  avisarCondominio("status16b", null);
});

document.getElementById("poAbaSel").addEventListener("change", (e) => {
  if (!PO_WB) return;
  try {
    if (abaEh3Col(PO_WB, e.target.value)) { poAplicarImport3Col(PO_WB, e.target.value); avisarCondominio("statuspo", nomeCondominioDoDocumento(linhasIniciaisDaAba(PO_WB, e.target.value))); return; }
    document.getElementById("poManual").hidden = true;
    PO_MANUAL.importado = false;
    const data = parsePrevisao(PO_WB, e.target.value);
    PARSED.prev = data;
    const unPlanilha = (data.rateioPrevisto && data.rateioPrevisto.unidades) || (data.rateioAtual && data.rateioAtual.unidades);
    const elUnPo = document.getElementById("f_po_unidades_total");
    if (unPlanilha && elUnPo && !elUnPo.value) elUnPo.value = String(unPlanilha);
    avisarCondominio("statuspo", nomeCondominioDoDocumento(linhasIniciaisDaAba(PO_WB, e.target.value)));
    setStatus("statuspo", true, 'lendo aba "' + data.sheetUsed + '"' + (data.condominioNome ? " (" + data.condominioNome + ")" : "") + " — " + data.categories.length + " categorias, total previsto " + fmtBRL(data.totalDespesasPrevisto, true));
    renderPreviewPO(data);
    if (typeof refreshPreview === "function") refreshPreview();
  } catch (err) {
    reportError("statuspo", "Não consegui ler essa aba", "Erro: " + err.message, "Escolha outra aba da planilha no seletor.");
  }
});


/* ---------- Sidebar: nome de cada aba (mostrado ao abrir com o mouse) ---------- */
// data-nome = nome limpo do item (sem o aviso "indisponível..."); acompanha abas dinâmicas e renomeações.
(function sidebarNomes() {
  const nav = document.getElementById("sidebarNav");
  if (!nav) return;
  const sync = () => nav.querySelectorAll("button[data-panel]").forEach((b) => {
    const n = String(b.dataset.labelOriginal || b.dataset.label || "").split(" \u2014 indispon")[0];
    if (b.dataset.nome !== n) b.dataset.nome = n;
  });
  sync();
  new MutationObserver(sync).observe(nav, { subtree: true, childList: true, attributes: true, attributeFilter: ["data-label"] });
})();

(function syncUnidadesCondominio() {
  const elInad = document.getElementById("f_inad_unidades_total");
  const elPo = document.getElementById("f_po_unidades_total");
  if (!elInad || !elPo) return;
  elInad.addEventListener("input", () => { if (!elPo.dataset.editado) elPo.value = elInad.value; });
  elPo.addEventListener("input", () => { elPo.dataset.editado = elPo.value ? "1" : ""; });
})();

(function setupDuplicacaoSecoes() {
  const btnPrest = document.getElementById("btnDupPrestacao");
  if (btnPrest) btnPrest.addEventListener("click", () => duplicarSecaoEstrutural("prestacao"));

  const btnPrev = document.getElementById("btnDupPrevisao");
  if (btnPrev) btnPrev.addEventListener("click", () => duplicarSecaoEstrutural("previsao"));

  const prevOnSelect = window.onCondominioSelected;
  window.onCondominioSelected = function (id) {
    if (typeof prevOnSelect === "function") prevOnSelect(id);
    document.querySelectorAll(".ext-11a-buscar, .ext-16b-buscar, .ext-inad-buscar").forEach((b) => { b.disabled = !id; });
  };
})();
