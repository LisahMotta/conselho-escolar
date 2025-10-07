import React, { useMemo, useRef, useState } from "react";
import axios from "axios";
import { ResponsiveContainer, BarChart, Bar, XAxis, YAxis, Tooltip } from "recharts";

const API_URL = import.meta.env.VITE_API_URL || "http://localhost:3000";

type Registro = {
  aluno: string;
  nota_final: number | null;
  faltas_pct: number | null;
  presenca_pct: number | null;
  retencao: boolean;
  motivo: string;
  turma?: string;
  disciplina?: string;
  ra?: string;
  ciclo?: string; // <-- novo
};

export default function App() {
  const fileRef = useRef<HTMLInputElement | null>(null);
  const [rows, setRows] = useState<Registro[]>([]);
  const [ciclo, setCiclo] = useState<string>("todos");
  const [turma, setTurma] = useState<string>("todas");
  const [disciplina, setDisciplina] = useState<string>("todas");

  // listas para selects
  const ciclos = useMemo(() => {
    const base = Array.from(new Set((rows as any[]).map(r => r.ciclo).filter(Boolean))) as string[];
    const ordem = ["Anos Iniciais", "Anos Finais", "Ensino Médio", "Indefinido"];
    const unicos = base.sort((a, b) => ordem.indexOf(a) - ordem.indexOf(b));
    return ["todos", ...unicos];
  }, [rows]);

  const turmas = useMemo(
    () => ["todas", ...((Array.from(new Set(rows.map(r => r.turma).filter(Boolean))) as string[]).sort())],
    [rows]
  );

  const disciplinas = useMemo(
    () => ["todas", ...((Array.from(new Set(rows.map(r => r.disciplina).filter(Boolean))) as string[]).sort())],
    [rows]
  );

  // agregação por ciclo (para KPIs)
  const porCiclo = useMemo(() => {
    const acc: Record<string, { total: number; ret: number; ok: number }> = {};
    for (const r of rows as any[]) {
      const c = r.ciclo || "Indefinido";
      acc[c] ??= { total: 0, ret: 0, ok: 0 };
      acc[c].total++;
      r.retencao ? acc[c].ret++ : acc[c].ok++;
    }
    return acc;
  }, [rows]);

  // filtro principal (inclui ciclo)
  const filtrado = useMemo(
    () =>
      (rows as any[]).filter(
        r =>
          (ciclo === "todos" || (r.ciclo ?? "") === ciclo) &&
          (turma === "todas" || (r.turma ?? "") === turma) &&
          (disciplina === "todas" || (r.disciplina ?? "") === disciplina)
      ) as Registro[],
    [rows, ciclo, turma, disciplina]
  );

  // métricas gerais do conjunto filtrado
  const metrics = useMemo(() => {
    const total = filtrado.length;
    const qtd_ret = filtrado.filter(r => r.retencao).length;
    const qtd_ok = total - qtd_ret;
    const cont_nota = filtrado.filter(r => /nota /.test(r.motivo) && !/presença /.test(r.motivo)).length;
    const cont_freq = filtrado.filter(r => /presença /.test(r.motivo) && !/nota /.test(r.motivo)).length;
    const cont_ambos = filtrado.filter(r => /nota /.test(r.motivo) && /presença /.test(r.motivo)).length;
    return { total, qtd_ret, qtd_ok, cont_nota, cont_freq, cont_ambos };
  }, [filtrado]);

  // upload
  async function handleUpload(e: React.FormEvent<HTMLFormElement>) {
    e.preventDefault();
    const files = fileRef.current?.files || [];
    if (!files.length) return alert("Selecione arquivos .xlsx/.xls/.csv");
    const fd = new FormData();
    Array.from(files).forEach(f => fd.append("arquivos", f));
    const { data } = await axios.post(`${API_URL}/upload`, fd, { headers: { "Content-Type": "multipart/form-data" } });
    if (data?.error) return alert(data.error);
    setRows(data?.rows ?? []);
    setCiclo("todos");
    setTurma("todas");
    setDisciplina("todas");
  }

  // downloads
  async function baixarExcel() {
    const url = `${API_URL}/download-retencao?turma=${encodeURIComponent(turma)}&disciplina=${encodeURIComponent(
      disciplina
    )}`;
    const resp = await axios.get(url, { responseType: "blob" });
    const a = document.createElement("a");
    a.href = URL.createObjectURL(resp.data);
    a.download = "relatorio_retencao.xlsx";
    a.click();
  }

  async function baixarPdf() {
    const url = `${API_URL}/download-pdf?turma=${encodeURIComponent(turma)}&disciplina=${encodeURIComponent(
      disciplina
    )}`;
    const resp = await axios.get(url, { responseType: "blob" });
    const a = document.createElement("a");
    a.href = URL.createObjectURL(resp.data);
    a.download = "relatorio_retencao.pdf";
    a.click();
  }

  const cols = useMemo(() => (filtrado[0] ? (Object.keys(filtrado[0]) as (keyof Registro)[]) : []), [filtrado]);

  return (
    <div className="wrapper">
      {/* HERO */}
      <div className="hero">
        <div className="logo">R</div>
        <div style={{ flex: 1 }}>
          <div className="title">Conselho Bimestral</div>
          <p className="subtitle">EE Profª Malba Thereza Ferraz Campaner</p>
        </div>
        <div className="row">
          <button className="btn-outline" type="button" onClick={baixarExcel}>
            Baixar Excel
          </button>
          <button className="btn-outline" type="button" onClick={baixarPdf}>
            Baixar PDF
          </button>
        </div>
      </div>

      {/* CONTROLES */}
      <form onSubmit={handleUpload} className="card">
        <div className="row">
          <input type="file" multiple ref={fileRef} />
          <button className="btn" type="submit">
            Processar
          </button>

          {/* NOVO: seletor de ciclo */}
          <select value={ciclo} onChange={e => setCiclo(e.target.value)}>
            {ciclos.map(c => (
              <option key={c} value={c}>
                {c}
              </option>
            ))}
          </select>

          <select value={turma} onChange={e => setTurma(e.target.value)}>
            {turmas.map(t => (
              <option key={t} value={t}>
                {t}
              </option>
            ))}
          </select>
          <select value={disciplina} onChange={e => setDisciplina(e.target.value)}>
            {disciplinas.map(d => (
              <option key={d} value={d}>
                {d}
              </option>
            ))}
          </select>
        </div>
        <p className="muted" style={{ marginTop: 8 }}>
          Envie .xlsx / .xls / .csv. Regras: média &lt; 5.0 ou presença &lt; 75% sinaliza retenção.
        </p>
      </form>

      {/* KPIS GERAIS */}
      <div className="kpis">
        <div className="kpi">
          <h4>Total de alunos</h4>
          <div className="v">{metrics.total}</div>
        </div>
        <div className="kpi bad">
          <h4>Em risco (retenção)</h4>
          <div className="v">{metrics.qtd_ret}</div>
        </div>
        <div className="kpi ok">
          <h4>OK</h4>
          <div className="v">{metrics.qtd_ok}</div>
        </div>
        <div className="kpi">
          <h4>Somatório de alertas</h4>
          <div className="v">{metrics.cont_nota + metrics.cont_freq + metrics.cont_ambos}</div>
        </div>
      </div>

      {/* KPIS POR CICLO */}
      <div className="card">
        <h3>KPIs por Ciclo</h3>
        <div className="kpis" style={{ gridTemplateColumns: "repeat(3, minmax(0,1fr))" }}>
          {Object.entries(porCiclo).map(([nome, k]) => (
            <div key={nome} className="kpi">
              <h4>{nome}</h4>
              <div className="v">{k.total}</div>
              <p className="muted">
                Em risco: <b style={{ color: "#ff9aa4" }}>{k.ret}</b> • OK:{" "}
                <b style={{ color: "#62f2a1" }}>{k.ok}</b>
              </p>
            </div>
          ))}
        </div>
      </div>

      {/* GRÁFICO */}
      <div className="card">
        <h3>Distribuição</h3>
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

      {/* TABELA */}
      <div className="card">
        <h3>Resultado</h3>
        <div className="table-wrap">
          <table>
            <thead>
              <tr>{cols.map(c => <th key={String(c)}>{c}</th>)}</tr>
            </thead>
            <tbody>
              {filtrado.map((r, i) => (
                <tr key={i}>
                  {cols.map(c => (
                    <td key={String(c)}>
                      {c === "retencao" ? (
                        <span className={`pill ${r[c] ? "bad" : "ok"}`}>{r[c] ? "SIM" : "NÃO"}</span>
                      ) : (
                        ((r as any)[c] as any) ?? ""
                      )}
                    </td>
                  ))}
                </tr>
              ))}
            </tbody>
          </table>
        </div>
        {!filtrado.length && (
          <p className="muted" style={{ marginTop: 8 }}>
            Envie uma planilha para visualizar os resultados.
          </p>
        )}
      </div>
    </div>
  );
}
