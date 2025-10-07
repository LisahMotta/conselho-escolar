import React, { useMemo, useRef, useState, FormEvent } from "react";
import axios from "axios";
import { ResponsiveContainer, BarChart, Bar, XAxis, YAxis, Tooltip } from "recharts";

// Funciona em Vite; tem fallback seguro pro localhost.
const API_URL =
  (typeof import.meta !== "undefined" &&
    (import.meta as any)?.env?.VITE_API_URL) ||
  "http://localhost:3000";

type Registro = {
  aluno: string;
  nota_final: number | null;
  faltas_pct: number | null;
  presenca_pct: number | null;
  retencao: boolean;
  motivo: string;
  ciclo?: string;
  turma?: string;
  disciplina?: string;
  ra?: string | number;
  numero?: string | number;
  __arquivo?: string;
  __aba?: string;
};

type UploadResp = {
  total: number;
  qtd_ret: number;
  qtd_ok: number;
  cont_nota: number;
  cont_freq: number;
  cont_ambos: number;
  excluidos_count?: number;
  candidatos?: (string | number)[];
  rows: Registro[];
};

export default function App() {
  const fileRef = useRef<HTMLInputElement | null>(null);
  const [rows, setRows] = useState<Registro[]>([]);
  const [turma, setTurma] = useState<string>("todas");
  const [disciplina, setDisciplina] = useState<string>("todas");
  const [ciclo, setCiclo] = useState<string>("todos");
  const [kpis, setKpis] = useState<Omit<UploadResp, "rows"> | null>(null);

  const turmas = useMemo(
    () =>
      ["todas", ...Array.from(new Set(rows.map((r) => r.turma).filter(Boolean))) as string[]].sort(),
    [rows]
  );
  const disciplinas = useMemo(
    () =>
      ["todas", ...Array.from(new Set(rows.map((r) => r.disciplina).filter(Boolean))) as string[]].sort(),
    [rows]
  );
  const ciclos = useMemo(
    () => ["todos", ...Array.from(new Set(rows.map((r) => r.ciclo).filter(Boolean))) as string[]],
    [rows]
  );

  const filtrado = useMemo(
    () =>
      rows.filter(
        (r) =>
          (ciclo === "todos" || (r.ciclo ?? "") === ciclo) &&
          (turma === "todas" || (r.turma ?? "") === turma) &&
          (disciplina === "todas" || (r.disciplina ?? "") === disciplina)
      ),
    [rows, ciclo, turma, disciplina]
  );

  const metrics = useMemo(() => {
    const total = filtrado.length;
    const qtd_ret = filtrado.filter((r) => r.retencao).length;
    const qtd_ok = total - qtd_ret;
    const cont_nota = filtrado.filter((r) => /nota /.test(r.motivo) && !/presença /.test(r.motivo)).length;
    const cont_freq = filtrado.filter((r) => /presença /.test(r.motivo) && !/nota /.test(r.motivo)).length;
    const cont_ambos = filtrado.filter((r) => /nota /.test(r.motivo) && /presença /.test(r.motivo)).length;
    return { total, qtd_ret, qtd_ok, cont_nota, cont_freq, cont_ambos };
  }, [filtrado]);

  async function handleUpload(e: FormEvent) {
    e.preventDefault();
    const files = fileRef.current?.files || [];
    if (!files.length) {
      alert("Selecione arquivos .xlsx/.xls/.csv/.pdf");
      return;
    }
    const fd = new FormData();
    Array.from(files).forEach((f) => fd.append("arquivos", f));

    const { data } = await axios.post<UploadResp>(`${API_URL}/upload`, fd, {
      headers: { "Content-Type": "multipart/form-data" },
    });

    if ((data as any)?.error) {
      alert((data as any).error);
      return;
    }

    setRows(data.rows || []);
    setKpis({
      total: data.total,
      qtd_ret: data.qtd_ret,
      qtd_ok: data.qtd_ok,
      cont_nota: data.cont_nota,
      cont_freq: data.cont_freq,
      cont_ambos: data.cont_ambos,
      excluidos_count: data.excluidos_count,
      candidatos: data.candidatos,
    });
    setTurma("todas");
    setDisciplina("todas");
    setCiclo("todos");
  }

  async function baixarExcel() {
    const resp = await axios.get(`${API_URL}/download-retencao`, { responseType: "blob" });
    const a = document.createElement("a");
    a.href = URL.createObjectURL(resp.data);
    a.download = "relatorio_retencao.xlsx";
    a.click();
  }
  async function baixarPdf() {
    const resp = await axios.get(`${API_URL}/download-pdf`, { responseType: "blob" });
    const a = document.createElement("a");
    a.href = URL.createObjectURL(resp.data);
    a.download = "relatorio_retencao.pdf";
    a.click();
  }

  const cols = useMemo(() => (filtrado[0] ? Object.keys(filtrado[0]) : []), [filtrado]);

  return (
    <main style={{ maxWidth: 1100, margin: "24px auto", padding: "0 16px", color: "#e7eaf6" }}>
      {/* Header */}
      <div style={{ display: "flex", gap: 12, alignItems: "center" }}>
        <div
          style={{
            width: 36,
            height: 36,
            borderRadius: 12,
            background: "#334155",
            display: "grid",
            placeItems: "center",
            fontWeight: 700,
          }}
        >
          R
        </div>
        <div style={{ flex: 1 }}>
          <div style={{ fontWeight: 700, fontSize: 18 }}>Conselho Bimestral</div>
          <div style={{ opacity: 0.7, fontSize: 12 }}>EE Profª Malba Thereza Ferraz Campaner</div>
        </div>
        <button onClick={baixarExcel} style={btnOutline}>
          Baixar Excel
        </button>
        <button onClick={baixarPdf} style={btnOutline}>
          Baixar PDF
        </button>
      </div>

      {/* Controles */}
      <form onSubmit={handleUpload} style={card}>
        <div style={{ display: "flex", gap: 8, flexWrap: "wrap", alignItems: "center" }}>
          <input type="file" multiple ref={fileRef} />
          <button type="submit" style={btn}>
            Processar
          </button>
          <select value={ciclo} onChange={(e) => setCiclo(e.target.value)} style={select}>
            {ciclos.map((c) => (
              <option key={c} value={c}>
                {c}
              </option>
            ))}
          </select>
          <select value={turma} onChange={(e) => setTurma(e.target.value)} style={select}>
            {turmas.map((t) => (
              <option key={t} value={t}>
                {t}
              </option>
            ))}
          </select>
          <select value={disciplina} onChange={(e) => setDisciplina(e.target.value)} style={select}>
            {disciplinas.map((d) => (
              <option key={d} value={d}>
                {d}
              </option>
            ))}
          </select>
        </div>
        <p style={{ marginTop: 8, opacity: 0.7, fontSize: 12 }}>
          Envie .xlsx /.xls /.csv /.pdf. Regras: média &lt; 5.0 ou presença &lt; 75% sinaliza retenção.
        </p>
      </form>

      {/* KPIs */}
      <div style={{ display: "grid", gridTemplateColumns: "repeat(4,minmax(0,1fr))", gap: 12 }}>
        <Kpi title="Total de alunos" value={metrics.total} />
        <Kpi title="Em risco (retenção)" value={metrics.qtd_ret} tone="bad" />
        <Kpi title="OK" value={metrics.qtd_ok} tone="ok" />
        <Kpi title="Somatório de alertas" value={metrics.cont_nota + metrics.cont_freq + metrics.cont_ambos} />
      </div>

      {/* Info de excluídos e candidatos */}
      {kpis?.excluidos_count ? (
        <div style={{ ...card, marginTop: 12 }}>
          <b>{kpis.excluidos_count}</b> registro(s) ignorado(s) por status (não comparecimento / transferido).
        </div>
      ) : null}

      {kpis?.candidatos?.length ? (
        <div style={{ ...card, marginTop: 12 }}>
          <h3 style={{ marginTop: 0 }}>Candidatos à retenção (RA → Nº → Nome)</h3>
          <ul style={{ columns: 2, gap: 24, margin: 0 }}>
            {kpis.candidatos.map((id, i) => (
              <li key={i} style={{ marginBottom: 6 }}>
                {id}
              </li>
            ))}
          </ul>
        </div>
      ) : null}

      {/* Gráfico */}
      <div style={{ ...card, marginTop: 12 }}>
        <h3 style={{ marginTop: 0 }}>Distribuição</h3>
        <div style={{ height: 280 }}>
          <ResponsiveContainer width="100%" height="100%">
            <BarChart
              data={[
                { label: "Retenção", val: metrics.qtd_ret },
                { label: "OK", val: metrics.qtd_ok },
              ]}
            >
              <XAxis stroke="#9aa3c0" dataKey="label" />
              <YAxis stroke="#9aa3c0" allowDecimals={false} />
              <Tooltip />
              <Bar dataKey="val" />
            </BarChart>
          </ResponsiveContainer>
        </div>
      </div>

      {/* Tabela */}
      <div style={{ ...card, marginTop: 12 }}>
        <h3 style={{ marginTop: 0 }}>Resultado</h3>
        <div style={{ overflow: "auto" }}>
          <table style={{ minWidth: "100%", borderCollapse: "collapse", fontSize: 14 }}>
            <thead>
              <tr>
                {cols.map((c) => (
                  <th key={c} style={th}>
                    {c}
                  </th>
                ))}
              </tr>
            </thead>
            <tbody>
              {filtrado.map((r, i) => (
                <tr key={i}>
                  {cols.map((c) => (
                    <td key={c} style={td}>
                      {c === "retencao" ? (
                        <span style={r.retencao ? pillBad : pillOk}>{r.retencao ? "SIM" : "NÃO"}</span>
                      ) : (
                        (r as any)[c] ?? ""
                      )}
                    </td>
                  ))}
                </tr>
              ))}
            </tbody>
          </table>
        </div>
        {!filtrado.length && <p style={{ opacity: 0.7, marginTop: 8 }}>Envie uma planilha para visualizar os resultados.</p>}
      </div>
    </main>
  );
}

/* UI helpers inline */
function Kpi({ title, value, tone }: { title: string; value: number; tone?: "ok" | "bad" }) {
  const bg = tone === "ok" ? "#063" : tone === "bad" ? "#5a1020" : "#1f2937";
  return (
    <div style={{ ...card, background: bg }}>
      <div style={{ opacity: 0.8, fontSize: 12 }}>{title}</div>
      <div style={{ fontSize: 28, fontWeight: 700 }}>{value}</div>
    </div>
  );
}

const card: React.CSSProperties = {
  border: "1px solid #263041",
  borderRadius: 12,
  padding: 12,
  background: "#0f172a",
};
const btn: React.CSSProperties = {
  padding: "8px 12px",
  borderRadius: 10,
  border: "1px solid #334155",
  background: "#1f2937",
  color: "#fff",
  cursor: "pointer",
};
const btnOutline: React.CSSProperties = { ...btn, background: "transparent" };
const select: React.CSSProperties = {
  padding: "6px 8px",
  borderRadius: 8,
  background: "#0b1220",
  color: "#e7eaf6",
  border: "1px solid #334155",
};
const th: React.CSSProperties = {
  border: "1px solid #263041",
  padding: "6px 8px",
  textAlign: "left",
  background: "#0b1220",
  position: "sticky",
  top: 0,
};
const td: React.CSSProperties = { border: "1px solid #263041", padding: "6px 8px" };
const pillBase: React.CSSProperties = { display: "inline-block", padding: "2px 8px", borderRadius: 999 };
const pillOk: React.CSSProperties = { ...pillBase, background: "#113d2a" };
const pillBad: React.CSSProperties = { ...pillBase, background: "#5a1020" };
