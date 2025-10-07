// ===== Imports =====
import express from "express";
import cors from "cors";
import multer from "multer";
import XLSX from "xlsx";
import ExcelJS from "exceljs";
import PDFDocument from "pdfkit";

// Import correto do pdf-parse em ES Modules
import pdfParse from "pdf-parse/lib/pdf-parse.js";

const pdfParse = pdfParseLib;


// Suporte a codepages (CP1252 etc.) para .xls antigo
import * as cpexcel from "xlsx/dist/cpexcel.full.mjs";
(XLSX).set_cptable?.(cpexcel);

// ===== App & middlewares =====
const app = express();
app.use(cors());
app.use(express.json());

const upload = multer({ storage: multer.memoryStorage() });

// ===== Regras =====
const MEDIA_MINIMA = 5.0;
const PRESENCA_MINIMA_PCT = 75.0;

// ===== Mapeamento de colunas (aliases) =====
// Cobre “N M F AC” e variações comuns em mapões
const MAPEAMENTO_COLUNAS = {
  aluno: [
    "aluno","nome","nome_aluno","estudante","aluno(a)","Aluno","Nome",
    "nome do aluno","aluno - nome","aluno/servidor"
  ],
  nota_final: [
    "nota_final","media","média","nota","media_final","média final","Média Final","Nota",
    "resultado final","mf","nota anual","nota bimestral","média anual","N","n"
  ],
  faltas_pct: [
    "faltas_%","faltas_percentual","%_faltas","percentual_faltas","faltas_pct","% faltas",
    "frequencia","frequência","frequencia (%)","frequência (%)","frequencia %","presença %","% presença","freq. (%)","freq (%)",
    "% faltas (F)","F (%)"
  ],
  faltas: [
    "faltas","qtde_faltas","ausencias","ausências","Ausências","Faltas",
    "faltas totais","faltas (nº)","nº faltas","nº de faltas","total de faltas","F","f"
  ],
  total_aulas: [
    "total_aulas","carga_horaria","total_de_aulas","carga_horária",
    "aulas previstas","aulas_previstas","Aulas Previstas","Total",
    "c.h.","ch","carga horária","aulas dadas","aulas ministradas","total aulas","Aulas Dadas"
  ],
  turma: ["turma","série","serie","classe","ano","Turma","série/turma","série turma"],
  ra: ["ra","registro","matricula","matrícula","RA","r.a.","rm","r.m."],
  disciplina: ["disciplina","componente","materia","matéria","Disciplina","componente curricular"],
  ac: ["ac","A.C.","A C","AC"] // opcional
};

// ===== Helpers =====
const norm  = (s) => (s ?? "").toString().trim();
const strip = (s) => norm(s).normalize("NFD").replace(/[\u0300-\u036f]/g, "").toLowerCase();

function toFloatSafe(v){
  if (v === null || v === undefined || v === "") return NaN;
  if (typeof v === "string") v = v.replace(",", ".");
  const n = Number(v);
  return Number.isFinite(n) ? n : NaN;
}

function criarMapaCabecalho(headers){
  const map = {};
  const low = headers.map(strip);
  const aliases = Object.fromEntries(
    Object.entries(MAPEAMENTO_COLUNAS).map(([dest, arr]) => [dest, arr.map(strip)])
  );
  for (let i = 0; i < headers.length; i++){
    for (const [dest, al] of Object.entries(aliases)){
      if (al.includes(low[i])) { map[headers[i]] = dest; break; }
    }
  }
  return map;
}

// encontra linha de header mesmo com título antes; tolerante
function sheetToJsonSmart(sheet){
  const A = XLSX.utils.sheet_to_json(sheet, { header: 1, raw: false, defval: "" });
  if (!A.length) return [];

  const wanted = Object.values(MAPEAMENTO_COLUNAS).flat().map(strip);
  let headerRow = -1, bestRow = -1, bestHits = -1;

  for (let r = 0; r < Math.min(200, A.length); r++){
    const row = A[r];
    const clean = row.map(strip);
    const nonEmpty = row.filter(c => norm(c) !== "").length;
    const hits = clean.filter(c => wanted.includes(c)).length;
    if (hits > bestHits) { bestHits = hits; bestRow = r; }
    if (nonEmpty >= 2 && hits >= 1) { headerRow = r; break; }
  }
  if (headerRow === -1) headerRow = bestRow >= 0 ? bestRow : 0;

  const headers = A[headerRow].map(h => norm(h));
  const dataRows = A.slice(headerRow + 1).filter(row => row.some(cell => norm(cell) !== ""));

  const rows = dataRows.map(arr => {
    const o = {};
    headers.forEach((h, idx) => { o[h] = idx < arr.length ? arr[idx] : ""; });
    return o;
  });

  return normalizarLinhas(rows);
}

function normalizarLinhas(rows){
  if (!rows.length) return [];
  const headers = Object.keys(rows[0]);
  const mapa = criarMapaCabecalho(headers);
  return rows.map(r=>{
    const obj = {};
    for (const [orig,val] of Object.entries(r)){
      const dest = mapa[orig];
      obj[dest || orig] = val;
    }
    return obj;
  });
}

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

// Validação flexível das colunas essenciais (depois do mapeamento)
function validarColunas(matrizNormalizada) {
  if (!matrizNormalizada.length) throw new Error("Nenhuma linha detectada após a leitura.");

  const colunas = new Set(Object.keys(matrizNormalizada[0]));
  const tem = (c) => colunas.has(c);

  const faltaAluno = !tem("aluno");
  const faltaNota  = !tem("nota_final"); // “N” já mapeia para nota_final
  const temF_pct   = tem("faltas_pct");
  const temF_abs   = tem("faltas") && tem("total_aulas");
  const faltaFreq  = !(temF_pct || temF_abs || tem("faltas")); // aceita apenas “faltas” (heurística adiante)

  const faltas = [];
  if (faltaAluno) faltas.push("aluno");
  if (faltaNota)  faltas.push("nota_final (ou N)");
  if (faltaFreq)  faltas.push("faltas_pct OU faltas + total_aulas (ou F)");

  if (faltas.length) {
    throw new Error(
      "Faltam colunas essenciais: " + faltas.join(", ") +
      ". Aceitos (exemplos): Aluno | N/Nota | F/Frequência% | Total de Aulas."
    );
  }
}

// ===== Cálculo =====
function calcularRetencao(data, origem = {}){
  validarColunas(data);

  const out = [];
  for (const row of data){
    const aluno = norm(row.aluno);
    const nota  = toFloatSafe(row.nota_final);

    let faltas_pct = NaN;
    if (row.faltas_pct !== undefined) {
      const v = toFloatSafe(row.faltas_pct);
      // se >50, interpretamos como PRESENÇA %, então %faltas = 100 - v
      if (Number.isFinite(v)) faltas_pct = (v > 50 ? (100 - v) : v);
    } else if (row.faltas !== undefined && row.total_aulas !== undefined) {
      const f = toFloatSafe(row.faltas), t = toFloatSafe(row.total_aulas);
      if (Number.isFinite(f) && Number.isFinite(t) && t>0) faltas_pct = (f/t)*100;
    } else if (row.faltas !== undefined) {
      const v = toFloatSafe(row.faltas);
      if (Number.isFinite(v)) {
        if (v <= 1) faltas_pct = v * 100;      // 0,12 -> 12%
        else if (v <= 100) faltas_pct = v;     // 25 -> 25%
      }
    }

    const presenca_pct = Number.isFinite(faltas_pct) ? 100 - faltas_pct : NaN;
    const criterio_nota = Number.isFinite(nota) ? (nota < MEDIA_MINIMA) : false;
    const criterio_freq = Number.isFinite(presenca_pct) ? (presenca_pct < PRESENCA_MINIMA_PCT) : false;
    const retencao = !!(criterio_nota || criterio_freq);

    const partes = [];
    if (criterio_nota) partes.push(`nota ${Number.isFinite(nota)?nota.toFixed(2):"NA"} < ${MEDIA_MINIMA}`);
    if (criterio_freq) partes.push(`presença ${Number.isFinite(presenca_pct)?presenca_pct.toFixed(1):"NA"}% < ${PRESENCA_MINIMA_PCT}%`);

    const registro = {
      aluno,
      nota_final: Number.isFinite(nota)? Number(nota.toFixed(2)) : null,
      presenca_pct: Number.isFinite(presenca_pct)? Number(presenca_pct.toFixed(1)) : null,
      faltas_pct: Number.isFinite(faltas_pct)? Number(faltas_pct.toFixed(1)) : null,
      retencao,
      motivo: partes.length ? partes.join(" e ") : "OK",
      ciclo: inferirCiclo(row.turma),
      ac: row.ac ?? undefined,
      ...origem
    };

    for (const extra of ["turma","ra","disciplina"]) {
      if (row[extra] !== undefined) registro[extra] = row[extra];
    }
    out.push(registro);
  }

  out.sort((a,b)=> a.retencao===b.retencao ? a.aluno.localeCompare(b.aluno) : (a.retencao? -1 : 1));
  return out;
}

// ===== PDF parsing =====
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

  // achar cabeçalho com Aluno e (N|Nota) e (F|Frequência)
  let headerIdx = -1;
  for (let i = 0; i < Math.min(lines.length, 200); i++) {
    const s = strip(lines[i]);
    if (/\baluno\b/.test(s) && (/\b(n|nota)\b/.test(s)) && (/\b(f|freq|frequencia|frequencia \(\%\))\b/.test(s))) {
      headerIdx = i;
      break;
    }
  }
  if (headerIdx === -1) {
    headerIdx = lines.findIndex(l => /\baluno\b/i.test(l));
  }
  if (headerIdx === -1) return out;

  const headParts = lines[headerIdx].split(/\s{2,}/).map(x => x.trim()).filter(Boolean);
  const headStripped = headParts.map(strip);

  const idxAluno = headStripped.findIndex(h => h === "aluno" || h === "nome");
  const idxN = headStripped.findIndex(h => h === "n" || h === "nota" || h === "média" || h === "media" || h === "mf");
  const idxF = headStripped.findIndex(h => h === "f" || h.startsWith("freq") || h.startsWith("% f") || h.includes("frequencia"));
  const idxAC = headStripped.findIndex(h => h === "ac" || h === "a.c." || h === "a c");

  for (let i = headerIdx + 1; i < lines.length; i++) {
    const L = lines[i];
    if (/^-{3,}|^={3,}|^Página\b|^OBS\b|^Observa(ç|c)ões\b/i.test(L)) break;

    const parts = L.split(/\s{2,}/).map(x => x.trim()).filter(Boolean);
    if (!parts.length) continue;

    if (idxAluno >= 0 && parts.length >= 2) {
      const aluno = parts[idxAluno] ?? parts[0];
      const nRaw  = idxN >= 0 ? parts[idxN] : undefined;
      const fRaw  = idxF >= 0 ? parts[idxF] : undefined;
      const acRaw = idxAC >= 0 ? parts[idxAC] : undefined;

      if (!aluno || /^\d+$/.test(aluno)) continue;

      const row = {
        aluno,
        nota_final: nRaw ?? "",
        faltas: fRaw ?? "",
        ac: acRaw ?? "",
        turma: meta.turma,
        disciplina: meta.disciplina
      };
      out.push(row);
    } else {
      const nums = (L.match(/-?\d+(?:[.,]\d+)?/g) || []).map(s => s);
      const nome = L.replace(/-?\d+(?:[.,]\d+)?/g, " ").replace(/\s{2,}/g, " ").trim();
      if (nome && nums.length) {
        const nRaw = nums[0];
        const fRaw = nums[1] ?? nums[0];
        out.push({
          aluno: nome,
          nota_final: nRaw,
          faltas: fRaw,
          turma: meta.turma,
          disciplina: meta.disciplina
        });
      }
    }
  }

  return out;
}

// ===== Rotas =====
app.get("/", (req,res)=>{
  res.type("text").send("API Retenção Escolar ativa. Use POST /upload para enviar .xls/.xlsx/.csv/.pdf");
});

// Upload e processamento (XLS/XLSX/CSV + PDF)
app.post("/upload", upload.array("arquivos", 12), async (req, res) => {
  try{
    const allRows = [];

    for (const f of req.files){
      const name = (f.originalname || "").toLowerCase();

      if (name.endsWith(".pdf")) {
        // --- PDF ---
        const pdfData = await pdfParse(f.buffer);

        const pdfRows = parsePdfToRows(pdfData.text);
        if (pdfRows.length) {
          const normRows = normalizarLinhas(pdfRows);
          const calculado = calcularRetencao(normRows, { __arquivo: f.originalname, __aba: "PDF" });
          allRows.push(...calculado);
        }
      } else {
        // --- Excel / CSV ---
        const wb = XLSX.read(f.buffer, { type: "buffer", cellDates: true });
        for (const sheetName of wb.SheetNames){
          const sheet = wb.Sheets[sheetName];
          const normRows = sheetToJsonSmart(sheet);
          if (!normRows.length) continue;
          const calculado = calcularRetencao(normRows, { __arquivo: f.originalname, __aba: sheetName });
          allRows.push(...calculado);
        }
      }
    }

    if (!allRows.length) {
      return res.json({ error: "Nenhum dado processado. Verifique cabeçalhos (Aluno, N/Nota, F/Frequência, Total Aulas) ou se o PDF contém texto (não imagem)." });
    }

    // KPIs gerais
    const total = allRows.length;
    const qtd_ret = allRows.filter(r=>r.retencao).length;
    const qtd_ok  = total - qtd_ret;
    const cont_nota  = allRows.filter(r=>/nota /.test(r.motivo) && !/presença /.test(r.motivo)).length;
    const cont_freq  = allRows.filter(r=>/presença /.test(r.motivo) && !/nota /.test(r.motivo)).length;
    const cont_ambos = allRows.filter(r=>/nota /.test(r.motivo) && /presença /.test(r.motivo)).length;

    app.set("lastRows", allRows);
    res.json({ total, qtd_ret, qtd_ok, cont_nota, cont_freq, cont_ambos, rows: allRows });
  }catch(e){
    res.json({ error: e.message || "Falha ao processar arquivos." });
  }
});

// Excel (todas as linhas + apenas retenção)
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

// PDF simples (KPIs + primeiras linhas)
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

  // KPIs
  doc.fillColor("#000").fontSize(12);
  doc.text(`Total de alunos: ${total}`);
  doc.text(`Em risco (retenção): ${qtd_ret}`);
  doc.text(`OK: ${qtd_ok}`);
  doc.moveDown(0.8);

  // Tabela simplificada (limita 40 linhas)
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
