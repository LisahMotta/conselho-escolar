// ===== Imports =====
import express from "express";
import cors from "cors";
import multer from "multer";
import XLSX from "xlsx";
import ExcelJS from "exceljs";
import PDFDocument from "pdfkit";

// pdf-parse (compatível com ESM / Node 22)
import * as pdfParseLib from "pdf-parse";
const pdfParse = pdfParseLib.default || pdfParseLib;

// Suporte a codepages para .xls antigo
import * as cpexcel from "xlsx/dist/cpexcel.full.mjs";
(XLSX).set_cptable?.(cpexcel);

// ===== App =====
const app = express();
app.use(cors());
app.use(express.json());
const upload = multer({ storage: multer.memoryStorage() });

// ===== Regras =====
const MEDIA_MINIMA = 5.0;
const PRESENCA_MINIMA_PCT = 75.0;

// ===================== Helpers =====================
const norm  = (s) => (s ?? "").toString().trim();
const strip = (s) => norm(s).normalize("NFD").replace(/[\u0300-\u036f]/g, "").toLowerCase();

function toFloatSafe(v){
  if (v === null || v === undefined || v === "") return NaN;
  if (typeof v === "string") v = v.replace(",", ".");
  const n = Number(v);
  return Number.isFinite(n) ? n : NaN;
}
function isNumericLike(v) {
  if (v === null || v === undefined) return false;
  const s = String(v).replace(",", ".").trim();
  if (s === "" || /[a-z]/i.test(s)) return false;
  const n = Number(s);
  return Number.isFinite(n);
}

// ===================== Mapeamento de cabeçalhos =====================
// "aluno" = NOME do estudante (não situação!)
const HEADER_REGEX = {
  aluno: [
    /\bnome\b/i,
    /\bnome\s*do\s*aluno\b/i,
    /\baluno\(a\)\b/i,
    /\bnome\s*aluno\b/i,
    /\bn[º°]?\s*\/\s*nome\b/i,
    /\bestudante\b/i,
  ],
  situacao: [
    /\bsitua[cç][aã]o\b/i,
    /\bsitua[cç][aã]o\s*do\s*aluno\b/i,
    /\bstatus\b/i, /\bresultado\b/i, /\bcondi[cç][aã]o\b/i,
    /\bobserva[cç][aã]o(?:es)?\b/i
  ],
  nota_final: [
    /^\s*n\s*$/i, /^\s*m\s*$/i, /^\s*mf\s*$/i,
    /\bnota\b/i, /\bm[eé]dia\b/i, /\bm[eé]dia\s*final\b/i, /\bresultado\s*final\b/i
  ],
  faltas_pct: [
    /^\s*f\s*$/i, /\bfreq(?:u[eê]ncia)?\b/i, /\bpresen[çc]a\b/i, /\b%/i,
    /\baproveitamento\b/i, /\bfreq\.?\b/i
  ],
  faltas: [ /\bfaltas?\b/i, /\baus[eê]ncias?\b/i, /\bn[ºo]?\s*faltas?\b/i ],
  total_aulas: [
    /\btotal\s*de?\s*aulas\b/i, /\baulas\s*(?:dadas|ministradas|previstas)\b/i,
    /\bcarga\s*hor[aá]ria\b/i, /\bc\.?h\.?\b/i, /\bch\b/i, /\btotal\b/i
  ],
  turma: [ /\bturma\b/i, /\bs[eé]rie\b/i, /\bano\b/i, /\bs[eé]rie\/?turma\b/i ],
  ra: [ /\bra\b/i, /\bmatr[ií]cula\b/i, /\bregistro\b/i, /\br\.?a\.?\b/i, /\brm\b/i ],
  disciplina: [ /\bdisciplina\b/i, /\bcomponente\b/i, /\bmat[eé]ria\b/i, /\bcomponente\s*curricular\b/i ],
  ac: [ /^\s*ac\s*$/i, /^\s*a\.?c\.?\s*$/i ],
  numero: [ /\bn[úu]mero\s*(do)?\s*aluno\b/i, /^\s*n[º°]\s*$/i, /\bno\b/i, /\bn\.?\s*aluno\b/i, /\bn[úu]mero\b/i ]
};

function criarMapaCabecalhoRegex(headers) {
  const map = {};
  for (const h of headers) {
    const s = h || "";
    let mapped = false;
    for (const [dest, patterns] of Object.entries(HEADER_REGEX)) {
      if (patterns.some((re) => re.test(s))) { map[h] = dest; mapped = true; break; }
    }
    if (mapped) continue;

    // Fallback leve
    const sh = strip(s);
    if (sh.includes("nome") || sh.includes("aluno") || sh.includes("estudante")) { map[h] = "aluno"; continue; }
    if (sh.startsWith("situa") || sh==="status" || sh.startsWith("result")) { map[h] = "situacao"; continue; }
    if (/^n$|^m$|^mf$/.test(sh) || sh.includes("nota") || sh.includes("media")) { map[h] = "nota_final"; continue; }
    if (sh==="f" || sh.includes("freq") || sh.includes("presenca") || sh==="%" ) { map[h] = "faltas_pct"; continue; }
    if (sh.includes("faltas") || sh.includes("ausencias")) { map[h] = "faltas"; continue; }
    if (sh.includes("carga") || sh.includes("ch") || sh.includes("aulas") || sh==="total") { map[h] = "total_aulas"; continue; }
    if (sh.includes("turma") || sh.includes("serie") || sh.includes("ano")) { map[h] = "turma"; continue; }
    if (sh.includes("disciplina") || sh.includes("componente") || sh.includes("materia")) { map[h] = "disciplina"; continue; }
    if (sh.includes("matricula") || sh==="ra" || sh.includes("registro")) { map[h] = "ra"; continue; }
    if (sh.includes("numero") || /^n[º°]$/.test(sh)) { map[h] = "numero"; continue; }
  }
  return map;
}

function normalizarLinhas(rows) {
  if (!rows.length) return [];
  const headers = Object.keys(rows[0] ?? {});
  const mapa = criarMapaCabecalhoRegex(headers);
  return rows.map((r) => {
    const o = {};
    for (const [orig, val] of Object.entries(r)) {
      const dest = mapa[orig];
      o[dest || orig] = val;
    }
    return o;
  });
}

// --- Reparo quando "aluno" está com situação (Ativo/Transferido/...) ---
const SITUACAO_KEYWORDS = [
  "ativo","aprovado","retido","reprovado","promovido",
  "transferido","transferida","transferencia","transferência",
  "remanejamento","remanejado","remanejada",
  "nao comparecimento","não comparecimento","nao compareceu","não compareceu"
];

function isStatusWord(v) {
  const s = strip(v || "");
  if (!s) return false;
  return SITUACAO_KEYWORDS.some(k => s.includes(k));
}
function looksLikeName(v) {
  const s = norm(v);
  return /^[A-Za-zÀ-ÿ'`´^~.\- ]{5,}$/.test(s) && /\s/.test(s);
}

// tenta linha a linha achar a melhor célula de NOME
function findNameInRow(row) {
  const skipKeys = new Set([
    "aluno","situacao","nota_final","faltas","faltas_pct","total_aulas","ac","turma","disciplina","ra","numero","__arquivo","__aba"
  ]);
  for (const [k, v] of Object.entries(row)) {
    if (skipKeys.has(k)) continue;
    const sv = norm(v);
    if (!sv) continue;
    if (isNumericLike(sv)) continue;
    if (isStatusWord(sv)) continue;
    if (looksLikeName(sv)) return sv;
  }
  // tentativas extras: colunas comuns
  for (const key of ["nome","nome_aluno","aluno(a)","estudante"]) {
    if (row[key]) {
      const sv = norm(row[key]);
      if (looksLikeName(sv)) return sv;
    }
  }
  return null;
}

function repairAlunoIfItIsSituacao(rows) {
  if (!rows.length) return rows;

  // 1) Heurística global: muitos "aluno" com status?
  const sample = rows.slice(0, Math.min(rows.length, 50));
  const statusLikeCount = sample.filter(r => isStatusWord(r.aluno)).length;
  const shouldRepairGlobal = statusLikeCount >= Math.ceil(sample.length * 0.3); // mais agressivo (30%+)

  if (!shouldRepairGlobal) return rows;

  // 2) Reparo linha-a-linha
  return rows.map(r => {
    const novo = { ...r };
    if (isStatusWord(novo.aluno)) {
      // guarda a situação se não existir
      if (!novo.situacao) novo.situacao = novo.aluno;
      const candidato = findNameInRow(novo);
      if (candidato) novo.aluno = candidato;
    } else if (!looksLikeName(novo.aluno)) {
      // aluno não parece nome → tenta achar
      const candidato = findNameInRow(novo);
      if (candidato) {
        if (isStatusWord(novo.aluno) && !novo.situacao) novo.situacao = novo.aluno;
        novo.aluno = candidato;
      }
    }
    return novo;
  });
}

// ===================== Leitura Excel “smart” =====================
function sheetToJsonSmart(sheet){
  const A = XLSX.utils.sheet_to_json(sheet, { header: 1, raw: false, defval: "" });
  if (!A.length) return [];
  const scoreTokens = [
    /nome/i, /aluno/i, /estudante/i,
    /^\s*n\s*$/i, /^\s*m\s*$/i, /^\s*mf\s*$/i, /\bnota\b/i, /\bm[eé]dia\b/i,
    /^\s*f\s*$/i, /\bfreq/i, /presen[çc]a/i, /%/,
    /faltas/i, /aus[eê]ncias/i,
    /total/i, /aulas/i, /carga\s*hor[aá]ria/i, /\bc\.?h\.?\b/i
  ];
  let headerRow = -1, bestScore = -1;
  for (let r = 0; r < Math.min(200, A.length); r++){
    const row = A[r].map(c => String(c ?? ""));
    const nonEmpty = row.filter(c => c.trim() !== "").length;
    if (nonEmpty < 2) continue;
    const score = row.reduce((acc, cell) => acc + (scoreTokens.some(re => re.test(cell)) ? 1 : 0), 0);
    if (score > bestScore) { bestScore = score; headerRow = r; }
    const hasNome = row.some(c => /nome|aluno|estudante/i.test(c));
    const hasNota = row.some(c => /^\s*(n|m|mf)\s*$/i.test(c) || /nota|m[eé]dia/i.test(c));
    const hasF    = row.some(c => /^\s*f\s*$/i.test(c) || /freq|presen[çc]a|%/i.test(c));
    if (hasNome && hasNota && hasF) { headerRow = r; break; }
  }
  if (headerRow === -1) headerRow = 0;

  const headers = A[headerRow].map(h => String(h ?? "").trim());
  const dataRows = A.slice(headerRow + 1).filter(row => row.some(cell => String(cell ?? "").trim() !== ""));
  const rows = dataRows.map(arr => {
    const o = {};
    headers.forEach((h, idx) => { o[h] = idx < arr.length ? arr[idx] : ""; });
    return o;
  });
  return normalizarLinhas(rows);
}

// ===================== Inferência / Validação =====================
function autoInferColumns(rows) {
  if (!rows.length) return rows;
  const headers = Object.keys(rows[0]);
  const prof = headers.map(h => {
    let textCnt=0, numCnt=0, in01=0, in0100=0, nonEmpty=0;
    for (const r of rows) {
      const v = r[h];
      if (v !== "" && v !== null && v !== undefined) {
        nonEmpty++;
        if (isNumericLike(v)) {
          numCnt++;
          const x = toFloatSafe(v);
          if (x >= 0 && x <= 10) in01++;
          if (x >= 0 && x <= 100) in0100++;
        } else {
          const s = String(v).trim(); if (s.length >= 3) textCnt++;
        }
      }
    }
    return { h, textCnt, numCnt, in01, in0100, nonEmpty };
  });
  const candNome = [...prof].sort((a,b)=> (b.textCnt - a.textCnt) || (b.nonEmpty - a.nonEmpty))[0];
  const candNota = [...prof].sort((a,b)=> (b.in01 - a.in01) || (b.numCnt - a.numCnt))[0];
  const candFreq = [...prof].filter(p=>!candNota || p.h!==candNota.h).sort((a,b)=> (b.in0100 - a.in0100) || (b.nonEmpty - a.nonEmpty))[0];

  const ren = {};
  if (!headers.includes("aluno") && candNome) ren[candNome.h] = "aluno";
  if (!headers.includes("nota_final") && candNota) ren[candNota.h] = "nota_final";
  if (!headers.includes("faltas_pct") && !headers.includes("faltas") && !headers.includes("total_aulas") && candFreq) ren[candFreq.h] = "faltas_pct";

  if (!Object.keys(ren).length) return rows;
  return rows.map(r => {
    const o = { ...r };
    for (const [from,to] of Object.entries(ren)) if (o[to] === undefined) o[to] = o[from];
    return o;
  });
}
function validarColunas(rows){
  if (!rows.length) throw new Error("Nenhuma linha detectada após a leitura.");
  const cols = new Set(Object.keys(rows[0]));
  const tem = (c)=> cols.has(c);
  const faltaNome = !tem("aluno");
  const faltaNota = !tem("nota_final");
  const temF_pct  = tem("faltas_pct");
  const temF_abs  = tem("faltas") && tem("total_aulas");
  const faltaFreq = !(temF_pct || temF_abs || tem("faltas"));
  const faltas = [];
  if (faltaNome) faltas.push("nome do aluno");
  if (faltaNota) faltas.push("nota_final (N/M/MF/Nota/Média)");
  if (faltaFreq) faltas.push("faltas_pct OU faltas + total_aulas (F/Frequência/Presença)");
  if (faltas.length) throw new Error("Faltam colunas essenciais: " + faltas.join(", "));
}

// ===================== Exclusões por situação/status =====================
const EXCLUDE_KEYWORDS = [
  "nao comparecimento","não comparecimento","nao compareceu","não compareceu",
  "transferido","transferida","transferencia","transferência",
  "remanejado","remanejada","remanejamento"
];
function isExcludedRow(row) {
  if (row.situacao) {
    const s = strip(row.situacao);
    if (s && EXCLUDE_KEYWORDS.some(kw => s.includes(kw))) return true;
  }
  for (const [k, v] of Object.entries(row)) {
    const ks = strip(k);
    if (/(obs|observa[cç][aã]o|resultado|status|ac)/.test(ks)) {
      const vs = strip(v);
      if (EXCLUDE_KEYWORDS.some(kw => vs.includes(kw))) return true;
    }
  }
  return false;
}

// ===================== Ciclo =====================
function inferirCiclo(turmaRaw) {
  const t = strip(turmaRaw || "");
  if (/\b(ensino\s*medio|ensino\s*m[eé]dio|em)\b/.test(t)) return "Ensino Médio";
  const serie = t.match(/\b([123])\s*[ªa]?\s*s[eé]rie\b/);
  if (serie) return "Ensino Médio";
  const ano = t.match(/\b([1-9])\s*[ºo°]?\s*ano\b/);
  if (ano) {
    const n = parseInt(ano[1], 10);
    if (n >= 1 && n <= 5) return "Anos Iniciais";
    if (n >= 6 && n <= 9) return "Anos Finais";
  }
  const num = t.match(/\b([1-9])\b/);
  if (num) {
    const n = parseInt(num[1], 10);
    if (n >= 1 && n <= 5) return "Anos Iniciais";
    if (n >= 6 && n <= 9) return "Anos Finais";
  }
  return "Indefinido";
}

// ===================== Cálculo principal =====================
function calcularRetencao(data, origem = {}){
  data = autoInferColumns(data);

  // Corrige quando "aluno" veio com SITUAÇÃO (global + linha-a-linha)
  data = repairAlunoIfItIsSituacao(data);

  // Remove transferidos / não comparecimento / remanejamento
  data = data.filter(r => !isExcludedRow(r));

  validarColunas(data);

  const out = [];
  for (const row of data){
    const alunoNome = norm(row.aluno);
    const nota  = toFloatSafe(row.nota_final);

    // faltas em %
    let faltas_pct = NaN;
    if (row.faltas_pct !== undefined) {
      const v = toFloatSafe(row.faltas_pct);
      if (Number.isFinite(v)) faltas_pct = (v > 50 ? (100 - v) : v); // presença% -> faltas%
    } else if (row.faltas !== undefined && row.total_aulas !== undefined) {
      const f = toFloatSafe(row.faltas), t = toFloatSafe(row.total_aulas);
      if (Number.isFinite(f) && Number.isFinite(t) && t>0) faltas_pct = (f/t)*100;
    } else if (row.faltas !== undefined) {
      const v = toFloatSafe(row.faltas);
      if (Number.isFinite(v)) faltas_pct = (v <= 1) ? v*100 : v;
    }

    const presenca_pct = Number.isFinite(faltas_pct) ? 100 - faltas_pct : NaN;
    const criterio_nota = Number.isFinite(nota) ? (nota < MEDIA_MINIMA) : false;
    const criterio_freq = Number.isFinite(presenca_pct) ? (presenca_pct < PRESENCA_MINIMA_PCT) : false;
    const retencao = !!(criterio_nota || criterio_freq);

    const partes = [];
    if (criterio_nota) partes.push(`nota ${Number.isFinite(nota)?nota.toFixed(2):"NA"} < ${MEDIA_MINIMA}`);
    if (criterio_freq) partes.push(`presença ${Number.isFinite(presenca_pct)?presenca_pct.toFixed(1):"NA"}% < ${PRESENCA_MINIMA_PCT}%`);

    const registro = {
      aluno: alunoNome,
      nota_final: Number.isFinite(nota)? Number(nota.toFixed(2)) : null,
      presenca_pct: Number.isFinite(presenca_pct)? Number(presenca_pct.toFixed(1)) : null,
      faltas_pct: Number.isFinite(faltas_pct)? Number(faltas_pct.toFixed(1)) : null,
      retencao,
      motivo: partes.length ? partes.join(" e ") : "OK",
      ciclo: inferirCiclo(row.turma),
      situacao: row.situacao ?? undefined,
      ac: row.ac ?? undefined,
      ...origem
    };
    for (const extra of ["turma","ra","disciplina","numero"]) {
      if (row[extra] !== undefined) registro[extra] = row[extra];
    }
    out.push(registro);
  }

  out.sort((a,b)=> a.retencao===b.retencao ? a.aluno.localeCompare(b.aluno) : (a.retencao? -1 : 1));
  return out;
}

// ===================== PDF parsing (texto) =====================
function pdfMetaFromLines(lines) {
  let turma, disciplina;
  for (const l of lines.slice(0, 40)) {
    const s = l.replace(/\s+/g, " ").trim();
    const mTurma = s.match(/\b(Turma|S[eé]rie\/?Turma|S[eé]rie)\s*[:\-]\s*(.+)$/i);
    const mDisc = s.match(/\b(Disciplina|Componente Curricular)\s*[:\-]\s*(.+)$/i);
    if (mTurma && !turma) turma = mTurma[2].trim();
    if (mDisc && !disciplina) disciplina = mDisc[2].trim();
  }
  return { turma, disciplina };
}
function parsePdfToRows(pdfText) {
  const rawLines = pdfText.split(/\r?\n/).map(l => l.trim()).filter(Boolean);
  const lines = rawLines.map(l => l.replace(/\u00A0/g, " "));
  const meta = pdfMetaFromLines(lines);
  const out = [];

  let headerIdx = -1;
  for (let i = 0; i < Math.min(lines.length, 200); i++) {
    const s = strip(lines[i]);
    if (/\bnome\b/.test(s) && (/\b(n|nota|m|mf)\b/.test(s)) && (/\b(f|freq|frequencia|presen[cç]a|%\b)/.test(s))) {
      headerIdx = i; break;
    }
  }
  if (headerIdx === -1) headerIdx = lines.findIndex(l => /\bnome\b/i.test(l) || /\baluno\b/i.test(l) || /\bestudante\b/i.test(l));
  if (headerIdx === -1) return out;

  const headParts = lines[headerIdx].split(/\s{2,}/).map(x => x.trim()).filter(Boolean);
  const headStripped = headParts.map(strip);
  const idxNome = headStripped.findIndex(h => ["nome","aluno","estudante"].includes(h));
  const idxN = headStripped.findIndex(h => ["n","nota","média","media","mf","m"].includes(h));
  const idxF = headStripped.findIndex(h => h === "f" || h.startsWith("freq") || h.includes("frequencia"));
  const idxAC = headStripped.findIndex(h => h === "ac" || h === "a.c." || h === "a c");

  for (let i = headerIdx + 1; i < lines.length; i++) {
    const L = lines[i];
    if (/^-{3,}|^={3,}|^Página\b|^OBS\b|^Observa(ç|c)ões\b/i.test(L)) break;

    const parts = L.split(/\s{2,}/).map(x => x.trim()).filter(Boolean);
    if (!parts.length) continue;

    if (idxNome >= 0 && parts.length >= 2) {
      const nome = parts[idxNome] ?? parts[0];
      const nRaw = idxN >= 0 ? parts[idxN] : undefined;
      const fRaw = idxF >= 0 ? parts[idxF] : undefined;
      const acRaw = idxAC >= 0 ? parts[idxAC] : undefined;
      if (!nome || /^\d+$/.test(nome)) continue;

      out.push({
        aluno: nome,
        nota_final: nRaw ?? "",
        faltas: fRaw ?? "",
        ac: acRaw ?? "",
        turma: meta.turma,
        disciplina: meta.disciplina
      });
    } else {
      const nums = (L.match(/-?\d+(?:[.,]\d+)?/g) || []).map(s => s);
      const nome = L.replace(/-?\d+(?:[.,]\d+)?/g, " ").replace(/\s{2,}/g, " ").trim();
      if (nome && nums.length) {
        const nRaw = nums[0];
        const fRaw = nums[1] ?? nums[0];
        out.push({ aluno: nome, nota_final: nRaw, faltas: fRaw, turma: meta.turma, disciplina: meta.disciplina });
      }
    }
  }
  return out;
}

// ===================== Rotas =====================
app.get("/", (req,res)=>{
  res.type("text").send("API Retenção Escolar ativa. POST /upload para enviar .xls/.xlsx/.csv/.pdf");
});

app.post("/upload", upload.array("arquivos", 12), async (req, res) => {
  try{
    const allRows = [];
    let excluidos = 0;

    for (const f of req.files){
      const name = (f.originalname || "").toLowerCase();

      if (name.endsWith(".pdf")) {
        const pdfData = await pdfParse(f.buffer);
        const pdfRows = parsePdfToRows(pdfData.text);
        if (pdfRows.length) {
          const normRows = normalizarLinhas(pdfRows);
          excluidos += normRows.filter(isExcludedRow).length;
          const calculado = calcularRetencao(normRows, { __arquivo: f.originalname, __aba: "PDF" });
          allRows.push(...calculado);
        }
      } else {
        const wb = XLSX.read(f.buffer, { type: "buffer", cellDates: true });
        for (const sheetName of wb.SheetNames){
          const sheet = wb.Sheets[sheetName];
          const normRows = sheetToJsonSmart(sheet);
          if (!normRows.length) continue;

          // (Opcional) log no servidor pra diagnosticar cabeçalhos
          try { console.log(`>>> [${f.originalname}] Aba: ${sheetName}`); console.log(">>> Colunas detectadas:", Object.keys(normRows[0])); } catch {}

          excluidos += normRows.filter(isExcludedRow).length;
          const calculado = calcularRetencao(normRows, { __arquivo: f.originalname, __aba: sheetName });
          allRows.push(...calculado);
        }
      }
    }

    if (!allRows.length) {
      return res.json({ error: "Nenhum dado processado. Verifique cabeçalhos (Nome, Nota, Frequência/F, Total Aulas) ou se o PDF contém texto." });
    }

    const total = allRows.length;
    const qtd_ret = allRows.filter(r=>r.retencao).length;
    const qtd_ok  = total - qtd_ret;
    const cont_nota  = allRows.filter(r=>/nota /.test(r.motivo) && !/presença /.test(r.motivo)).length;
    const cont_freq  = allRows.filter(r=>/presença /.test(r.motivo) && !/nota /.test(r.motivo)).length;
    const cont_ambos = allRows.filter(r=>/nota /.test(r.motivo) && /presença /.test(r.motivo)).length;

    const candidatos = allRows
      .filter(r => r.retencao)
      .map(r =>
        (r.ra && String(r.ra).trim()) ||
        (r.numero && String(r.numero).trim()) ||
        r.aluno
      )
      .filter(Boolean);

    app.set("lastRows", allRows);
    res.json({
      total, qtd_ret, qtd_ok, cont_nota, cont_freq, cont_ambos,
      excluidos_count: excluidos,
      candidatos,
      rows: allRows
    });
  }catch(e){
    res.json({ error: e.message || "Falha ao processar arquivos." });
  }
});

app.get("/download-retencao", async (req,res)=>{
  const rows = app.get("lastRows") || [];
  if(!rows.length) return res.status(400).send("Nenhum resultado disponível. Faça upload primeiro.");

  const workbook = new ExcelJS.Workbook();
  const geral = workbook.addWorksheet("Geral");
  const ret = workbook.addWorksheet("Retencao");

  const cols = Object.keys(rows[0]);
  geral.columns = cols.map(c=>({ header:c, key:c }));
  ret.columns   = cols.map(c=>({ header:c, key:c }));

  rows.forEach(r=>geral.addRow(r));
  rows.filter(r=>r.retencao).forEach(r=>ret.addRow(r));

  const buffer = await workbook.xlsx.writeBuffer();
  res.setHeader("Content-Type","application/vnd.openxmlformats-officedocument.spreadsheetml.sheet");
  res.setHeader("Content-Disposition","attachment; filename=relatorio_retencao.xlsx");
  res.send(Buffer.from(buffer));
});

app.get("/download-pdf", (req,res)=>{
  const rows = app.get("lastRows") || [];
  if(!rows.length) return res.status(400).send("Nenhum resultado disponível. Faça upload primeiro.");

  const total   = rows.length;
  const qtd_ret = rows.filter(r=>r.retencao).length;
  const qtd_ok  = total - qtd_ret;

  res.setHeader("Content-Type","application/pdf");
  res.setHeader("Content-Disposition","attachment; filename=relatorio_retencao.pdf");

  const doc = new PDFDocument({ size: "A4", margin: 40 });
  doc.pipe(res);

  doc.fontSize(18).text("Relatório de Retenção Escolar", { align: "center" });
  doc.moveDown(0.5);
  doc.fontSize(10).fillColor("#666").text(`Gerado em ${new Date().toLocaleString()}`, { align: "center" });
  doc.moveDown(1.2);

  doc.fillColor("#000").fontSize(12);
  doc.text(`Total de alunos: ${total}`);
  doc.text(`Em risco (retenção): ${qtd_ret}`);
  doc.text(`OK: ${qtd_ok}`);
  doc.moveDown(0.8);

  const cols = Object.keys(rows[0]);
  const showCols = cols.slice(0, 8);
  doc.fontSize(10).fillColor("#000").text("Amostra de resultados:", { underline: true });
  doc.moveDown(0.5);

  doc.font("Helvetica-Bold");
  doc.text(showCols.join(" | "));
  doc.font("Helvetica");
  doc.moveDown(0.2);

  rows.slice(0, 40).forEach(r=>{
    const line = showCols.map(c=>{
      const v = r[c];
      if (c === "retencao") return v ? "SIM" : "NÃO";
      return (v==null?"":String(v));
    }).join(" | ");
    doc.text(line);
  });

  if (rows.length > 40) {
    doc.moveDown(0.5);
    doc.text(`... (${rows.length - 40} linhas ocultas no PDF)`, { italic: true });
  }

  doc.end();
});

// ===== Start =====
const PORT = process.env.PORT || 3000;
app.listen(PORT, ()=>console.log("Servidor na porta " + PORT));
