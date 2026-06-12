import { useState, useCallback, useEffect } from "react";
import * as XLSX from "xlsx";
import { BarChart, Bar, LineChart, Line, XAxis, YAxis, Tooltip, ResponsiveContainer, ReferenceLine, Cell } from "recharts";

const C = {
  laranja: "#F97316", laranjaLight: "#FED7AA",
  verde: "#16A34A", verdeLight: "#BBF7D0",
  vermelho: "#DC2626", vermelhoLight: "#FEE2E2",
  amarelo: "#CA8A04", amareloLight: "#FEF08A",
  azul: "#2563EB", azulLight: "#DBEAFE",
  cinzaFundo: "#F8F7F4", cinzaCard: "#FFFFFF",
  cinzaBorda: "#E5E3DF", cinzaTexto: "#6B7280", texto: "#1C1917",
};

const pct = (v) => (typeof v === "number" ? (v * 100).toFixed(1) + "%" : "—");
const brl = (v) => (typeof v === "number" ? "R$ " + v.toLocaleString("pt-BR", { maximumFractionDigits: 0 }) : "—");
const num = (v) => (typeof v === "number" ? v.toFixed(0) : "—");
const dias = (v) => (typeof v === "number" ? v.toFixed(1) + "d" : "—");

function parseVal(raw) {
  if (raw === null || raw === undefined || raw === "-" || raw === "-%") return null;
  if (typeof raw === "number") return raw;
  if (typeof raw === "string") {
    const m = raw.match(/COMPUTED_VALUE[^,]+,([\-0-9.]+)\)/);
    if (m) return parseFloat(m[1]);
  }
  return null;
}

// ── PARCEIROS config ─────────────────────────────────────────────────────────
const PARCEIROS_CFG = [
  { nome: "Safari",       slaRow: 8,  atrasosRow: 9,  agingRow: 90, csatRow: 131, giroRow: 174 },
  { nome: "Saldão",       slaRow: 10, atrasosRow: 11, agingRow: 91, csatRow: 132, giroRow: 175 },
  { nome: "Movel Service",slaRow: 12, atrasosRow: 13, agingRow: 92, csatRow: 133, giroRow: 176 },
  { nome: "Real Moweis",  slaRow: 14, atrasosRow: 15, agingRow: 93, csatRow: 134, giroRow: 177 },
  { nome: "Tarcis",       slaRow: 16, atrasosRow: 17, agingRow: 94, csatRow: 135, giroRow: null },
  { nome: "LOGME",        slaRow: 18, atrasosRow: 19, agingRow: 95, csatRow: 136, giroRow: 178 },
  { nome: "KMAN",         slaRow: 20, atrasosRow: 21, agingRow: 96, csatRow: 137, giroRow: 179 },
  { nome: "Outeletro",    slaRow: 22, atrasosRow: 23, agingRow: 97, csatRow: 138, giroRow: 180 },
  { nome: "Ponto Mix",    slaRow: 24, atrasosRow: 25, agingRow: 98, csatRow: 139, giroRow: 181 },
  { nome: "ORC",          slaRow: 26, atrasosRow: 27, agingRow: 99, csatRow: 140, giroRow: 182 },
  { nome: "Ebenezer",     slaRow: 28, atrasosRow: 29, agingRow: 100,csatRow: 141, giroRow: 183 },
];

function parseWeekly(data) {
  const wb = XLSX.read(data, { type: "array" });
  const ws = wb.Sheets[wb.SheetNames[0]];
  const rows = XLSX.utils.sheet_to_json(ws, { header: 1, defval: null });
  const weekRow = rows[1];

  const weeks2026 = [];
  for (let c = 56; c <= 107; c++) {
    if (weekRow[c] !== null && weekRow[c] !== undefined)
      weeks2026.push({ col: c, num: parseInt(weekRow[c]) });
  }

  const get = (rowIdx, col) => parseVal(rows[rowIdx]?.[col]);
  const KEY_ROWS = [4, 128, 89, 49];
  const filledWeeks = weeks2026.filter((w) =>
    KEY_ROWS.some((r) => { const v = get(r, w.col); return v !== null && typeof v === "number"; })
  );
  if (filledWeeks.length === 0) throw new Error("Nenhuma semana com dados encontrada.");

  const lastWk = filledWeeks[filledWeeks.length - 1];
  const prevWk = filledWeeks.length > 1 ? filledWeeks[filledWeeks.length - 2] : lastWk;
  const lastCol = lastWk.col;
  const prevCol = prevWk.col;

  const history = (rowIdx, n = 4) =>
    filledWeeks.slice(-n).map((w) => ({ week: `W${w.num}`, val: get(rowIdx, w.col) }));

  const parceiros = PARCEIROS_CFG.map((p) => ({
    ...p,
    sla:      get(p.slaRow, lastCol),
    slaPrev:  get(p.slaRow, prevCol),
    atrasos:  get(p.atrasosRow, lastCol),
    atrasosH: filledWeeks.slice(-4).map(w => ({ week: `W${w.num}`, val: get(p.atrasosRow, w.col) })),
    aging:    get(p.agingRow, lastCol),
    agingH:   filledWeeks.slice(-4).map(w => ({ week: `W${w.num}`, val: get(p.agingRow, w.col) })),
    csat:     get(p.csatRow, lastCol),
    csatPrev: get(p.csatRow, prevCol),
    csatH:    filledWeeks.slice(-4).map(w => ({ week: `W${w.num}`, val: get(p.csatRow, w.col) })),
    giro:     p.giroRow ? get(p.giroRow, lastCol) : null,
  }));

  return {
    currentWeek: lastWk.num,
    slaGeral:     get(4, lastCol),   slaGeralPrev:     get(4, prevCol),
    agendamento:  get(49, lastCol),  agendamentoPrev:  get(49, prevCol),
    aderencia:    get(69, lastCol),  aderenciaPrev:    get(69, prevCol),
    agingMedio:   get(89, lastCol),  agingMedioPrev:   get(89, prevCol),
    mediaColeta:  get(193, lastCol),
    slaCliente:   get(194, lastCol), slaClientePrev:   get(194, prevCol),
    csat:         get(128, lastCol), csatPrev:         get(128, prevCol),
    taxaRespostaCsat: get(127, lastCol),
    // Financeiro Devolução
    rentDev:      get(152, lastCol), rentDevPrev:  get(152, prevCol),
    abrangDev:    get(154, lastCol),
    notasDev:     get(155, lastCol),
    valorDev:     get(156, lastCol),
    estoqueDev:   get(157, lastCol),
    // Financeiro Reversa
    rentRev:      get(153, lastCol), rentRevPrev:  get(153, prevCol),
    abrangRev:    get(158, lastCol),
    notasRev:     get(159, lastCol),
    valorRev:     get(160, lastCol),
    estoqueRev:   get(161, lastCol),
    // PNC
    pctCDs:       get(196, lastCol),
    lotesPrec:    get(197, lastCol),
    receitaPNC:   get(198, lastCol),
    rentPNC:      get(199, lastCol),
    // Históricos
    histSla:          history(4),
    histCsat:         history(128),
    histAgendamento:  history(49),
    histAderencia:    history(69),
    histRentDev:      history(152),
    histRentRev:      history(153),
    histAbrangRev:    history(158),
    parceiros,
    filledWeeks,
  };
}

// ── Componentes base ─────────────────────────────────────────────────────────
function KpiCard({ label, value, prev, format, meta, invertido, semana }) {
  const diff = value !== null && prev !== null ? value - prev : null;
  const isGood = diff === null ? null : invertido ? diff < 0 : diff > 0;
  const atMeta = meta !== undefined && value !== null ? (invertido ? value <= meta : value >= meta) : null;
  return (
    <div style={{ background: C.cinzaCard, border: `1px solid ${C.cinzaBorda}`, borderRadius: 12, padding: "16px 20px", display: "flex", flexDirection: "column", gap: 4 }}>
      <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center" }}>
        <span style={{ fontSize: 11, fontWeight: 600, textTransform: "uppercase", letterSpacing: "0.08em", color: C.cinzaTexto }}>{label}</span>
        {semana && <span style={{ fontSize: 10, color: C.laranja, fontWeight: 700, background: C.laranjaLight, borderRadius: 20, padding: "1px 8px" }}>W{semana}</span>}
      </div>
      <div style={{ display: "flex", alignItems: "baseline", gap: 8 }}>
        <span style={{ fontSize: 28, fontWeight: 700, color: atMeta === false ? C.vermelho : atMeta === true ? C.verde : C.texto }}>
          {value === null ? "—" : format(value)}
        </span>
        {diff !== null && (
          <span style={{ fontSize: 13, fontWeight: 600, color: isGood ? C.verde : C.vermelho }}>
            {diff > 0 ? "▲" : "▼"} {format(Math.abs(diff))}
          </span>
        )}
      </div>
      {meta !== undefined && (
        <span style={{ fontSize: 11, color: C.cinzaTexto }}>
          Meta: {format(meta)}
          {atMeta !== null && <span style={{ marginLeft: 6, color: atMeta ? C.verde : C.vermelho, fontWeight: 600 }}>{atMeta ? "✓" : "✗"}</span>}
        </span>
      )}
    </div>
  );
}

function MiniLine({ data, color = C.laranja, meta }) {
  return (
    <ResponsiveContainer width="100%" height={56}>
      <LineChart data={data} margin={{ top: 4, right: 4, left: 4, bottom: 4 }}>
        {meta && <ReferenceLine y={meta} stroke={C.cinzaBorda} strokeDasharray="3 3" />}
        <Line type="monotone" dataKey="val" stroke={color} strokeWidth={2} dot={false} />
        <Tooltip formatter={(v) => v !== null ? (v < 2 ? pct(v) : v.toFixed(1)) : "—"} labelFormatter={l => l}
          contentStyle={{ fontSize: 11, background: C.texto, color: "#fff", border: "none", borderRadius: 6 }} />
      </LineChart>
    </ResponsiveContainer>
  );
}

function SecHead({ children }) {
  return <p style={{ fontSize: 12, fontWeight: 700, textTransform: "uppercase", letterSpacing: "0.08em", color: C.cinzaTexto, marginBottom: 12, marginTop: 4 }}>{children}</p>;
}

function Card({ children, style }) {
  return <div style={{ background: C.cinzaCard, border: `1px solid ${C.cinzaBorda}`, borderRadius: 12, padding: "20px 24px", ...style }}>{children}</div>;
}

// ── ABA: VISÃO GERAL ─────────────────────────────────────────────────────────
function TabOverview({ d }) {
  const w = d.currentWeek;
  return (
    <div style={{ display: "flex", flexDirection: "column", gap: 24, paddingTop: 28 }}>
      <div style={{ display: "grid", gridTemplateColumns: "repeat(auto-fit, minmax(175px, 1fr))", gap: 12 }}>
        <KpiCard label="SLA Geral" value={d.slaGeral} prev={d.slaGeralPrev} format={pct} meta={0.86} semana={w} />
        <KpiCard label="CSAT (4-5)" value={d.csat} prev={d.csatPrev} format={pct} meta={0.85} semana={w} />
        <KpiCard label="Agendamento" value={d.agendamento} prev={d.agendamentoPrev} format={pct} meta={0.95} semana={w} />
        <KpiCard label="Aderência" value={d.aderencia} prev={d.aderenciaPrev} format={pct} meta={0.95} semana={w} />
        <KpiCard label="Aging Médio de Problema" value={d.agingMedio} prev={d.agingMedioPrev} format={dias} meta={5} invertido semana={w} />
        <KpiCard label="SLA Cliente (15 dias)" value={d.slaCliente} prev={d.slaClientePrev} format={pct} semana={w} />
        <KpiCard label="Média Dias Coleta" value={d.mediaColeta} prev={null} format={dias} semana={w} />
        <KpiCard label="Taxa Resp. CSAT" value={d.taxaRespostaCsat} prev={null} format={pct} semana={w} />
      </div>

      <div style={{ display: "grid", gridTemplateColumns: "1fr 1fr", gap: 16 }}>
        {[
          { title: "SLA Geral", hist: d.histSla, meta: 0.86, color: C.laranja },
          { title: "CSAT (4-5)", hist: d.histCsat, meta: 0.85, color: C.verde },
          { title: "Agendamento Sistêmico", hist: d.histAgendamento, meta: 0.95, color: "#6366F1" },
          { title: "Aderência ao Agendamento", hist: d.histAderencia, meta: 0.95, color: "#0EA5E9" },
        ].map((g, i) => (
          <Card key={i}>
            <SecHead>{g.title} — últimas {g.hist.length} semanas</SecHead>
            <MiniLine data={g.hist} color={g.color} meta={g.meta} />
            <div style={{ display: "flex", justifyContent: "space-between", marginTop: 4 }}>
              {g.hist.map((h, j) => (
                <span key={j} style={{ fontSize: 11, color: C.cinzaTexto }}>{h.week}: <strong style={{ color: h.val !== null && h.val < g.meta ? C.vermelho : C.texto }}>{h.val !== null ? pct(h.val) : "—"}</strong></span>
              ))}
            </div>
          </Card>
        ))}
      </div>

      <Card>
        <SecHead>SLA por Parceiro — W{w}</SecHead>
        <ResponsiveContainer width="100%" height={200}>
          <BarChart data={d.parceiros.filter(p => p.sla !== null).sort((a, b) => a.sla - b.sla)} margin={{ top: 4, right: 4, left: -20, bottom: 44 }}>
            <XAxis dataKey="nome" tick={{ fontSize: 10 }} angle={-35} textAnchor="end" interval={0} />
            <YAxis tickFormatter={v => (v * 100).toFixed(0) + "%"} domain={[0, 1]} tick={{ fontSize: 11 }} />
            <Tooltip formatter={v => pct(v)} contentStyle={{ fontSize: 12, background: C.texto, color: "#fff", border: "none", borderRadius: 6 }} />
            <ReferenceLine y={0.86} stroke={C.vermelho} strokeDasharray="4 2" />
            <Bar dataKey="sla" radius={[4, 4, 0, 0]}>
              {d.parceiros.filter(p => p.sla !== null).sort((a, b) => a.sla - b.sla).map((p, i) => (
                <Cell key={i} fill={p.sla < 0.86 ? C.vermelho : p.sla < 0.92 ? C.amarelo : C.verde} />
              ))}
            </Bar>
          </BarChart>
        </ResponsiveContainer>
      </Card>
    </div>
  );
}

// ── ABA: PARCEIROS ───────────────────────────────────────────────────────────
function TabParceiros({ d }) {
  const [view, setView] = useState("sla");
  const views = [
    { id: "sla",     label: "SLA" },
    { id: "csat",    label: "CSAT" },
    { id: "aging",   label: "Aging" },
    { id: "atrasos", label: "Atrasos" },
  ];
  const w = d.currentWeek;

  return (
    <div style={{ display: "flex", flexDirection: "column", gap: 16, paddingTop: 28 }}>
      <div style={{ display: "flex", gap: 6 }}>
        {views.map(v => (
          <button key={v.id} onClick={() => setView(v.id)} style={{
            padding: "6px 16px", borderRadius: 20, border: `1px solid ${view === v.id ? C.laranja : C.cinzaBorda}`,
            background: view === v.id ? C.laranja : C.cinzaCard, color: view === v.id ? "#fff" : C.texto,
            cursor: "pointer", fontSize: 13, fontWeight: 600,
          }}>{v.label}</button>
        ))}
      </div>

      {view === "sla" && (
        <Card>
          <SecHead>SLA por Parceiro — W{w} (meta 86%)</SecHead>
          <table style={{ width: "100%", borderCollapse: "collapse", fontSize: 13 }}>
            <thead><tr style={{ borderBottom: `2px solid ${C.cinzaBorda}` }}>
              {["Parceiro", "SLA", "vs Ant.", "Atrasos Abertos", "Status"].map(h => (
                <th key={h} style={{ padding: "8px 12px", textAlign: h === "Parceiro" ? "left" : "center", color: C.cinzaTexto, fontWeight: 600, fontSize: 11, textTransform: "uppercase" }}>{h}</th>
              ))}
            </tr></thead>
            <tbody>
              {[...d.parceiros].filter(p => p.sla !== null).sort((a, b) => a.sla - b.sla).map((p, i) => {
                const diff = p.sla !== null && p.slaPrev !== null ? p.sla - p.slaPrev : null;
                const crit = p.sla < 0.86;
                return (
                  <tr key={i} style={{ borderBottom: `1px solid ${C.cinzaBorda}`, background: crit ? C.vermelhoLight + "55" : "transparent" }}>
                    <td style={{ padding: "10px 12px", fontWeight: 600 }}>{p.nome}</td>
                    <td style={{ padding: "10px 12px", textAlign: "center" }}>
                      <span style={{ display: "inline-block", padding: "2px 10px", borderRadius: 20, fontWeight: 700, background: crit ? C.vermelhoLight : C.verdeLight, color: crit ? C.vermelho : C.verde }}>{pct(p.sla)}</span>
                    </td>
                    <td style={{ padding: "10px 12px", textAlign: "center", color: diff === null ? C.cinzaTexto : diff < 0 ? C.vermelho : C.verde, fontWeight: 600 }}>
                      {diff === null ? "—" : `${diff > 0 ? "+" : ""}${(diff * 100).toFixed(1)}pp`}
                    </td>
                    <td style={{ padding: "10px 12px", textAlign: "center", color: p.atrasos > 10 ? C.vermelho : C.texto, fontWeight: p.atrasos > 10 ? 700 : 400 }}>{p.atrasos ?? "—"}</td>
                    <td style={{ padding: "10px 12px", textAlign: "center", fontSize: 18 }}>{crit ? "🔴" : p.sla >= 0.95 ? "🟢" : "🟡"}</td>
                  </tr>
                );
              })}
            </tbody>
          </table>
        </Card>
      )}

      {view === "csat" && (
        <Card>
          <SecHead>CSAT por Parceiro — W{w} (meta 85%)</SecHead>
          <table style={{ width: "100%", borderCollapse: "collapse", fontSize: 13 }}>
            <thead><tr style={{ borderBottom: `2px solid ${C.cinzaBorda}` }}>
              {["Parceiro", "CSAT", "vs Ant.", ...d.parceiros[0].csatH.map(h => h.week)].map(h => (
                <th key={h} style={{ padding: "8px 12px", textAlign: h === "Parceiro" ? "left" : "center", color: C.cinzaTexto, fontWeight: 600, fontSize: 11, textTransform: "uppercase" }}>{h}</th>
              ))}
            </tr></thead>
            <tbody>
              {d.parceiros.filter(p => p.csat !== null).sort((a, b) => (a.csat ?? 1) - (b.csat ?? 1)).map((p, i) => {
                const diff = p.csat !== null && p.csatPrev !== null ? p.csat - p.csatPrev : null;
                const crit = p.csat !== null && p.csat < 0.85;
                return (
                  <tr key={i} style={{ borderBottom: `1px solid ${C.cinzaBorda}`, background: crit ? C.vermelhoLight + "44" : "transparent" }}>
                    <td style={{ padding: "9px 12px", fontWeight: 600 }}>{p.nome}</td>
                    <td style={{ padding: "9px 12px", textAlign: "center" }}>
                      <span style={{ display: "inline-block", padding: "2px 10px", borderRadius: 20, fontWeight: 700, background: crit ? C.vermelhoLight : C.verdeLight, color: crit ? C.vermelho : C.verde }}>{pct(p.csat)}</span>
                    </td>
                    <td style={{ padding: "9px 12px", textAlign: "center", color: diff === null ? C.cinzaTexto : diff < 0 ? C.vermelho : C.verde, fontWeight: 600 }}>
                      {diff === null ? "—" : `${diff > 0 ? "+" : ""}${(diff * 100).toFixed(1)}pp`}
                    </td>
                    {p.csatH.map((h, j) => (
                      <td key={j} style={{ padding: "9px 12px", textAlign: "center", fontSize: 12, color: h.val !== null && h.val < 0.85 ? C.vermelho : C.cinzaTexto }}>
                        {h.val !== null ? pct(h.val) : "—"}
                      </td>
                    ))}
                  </tr>
                );
              })}
            </tbody>
          </table>
        </Card>
      )}

      {view === "aging" && (
        <Card>
          <SecHead>Aging Médio de Problema por Parceiro — W{w}</SecHead>
          <table style={{ width: "100%", borderCollapse: "collapse", fontSize: 13 }}>
            <thead><tr style={{ borderBottom: `2px solid ${C.cinzaBorda}` }}>
              {["Parceiro", "Aging Atual", ...d.parceiros[0].agingH.map(h => h.week)].map(h => (
                <th key={h} style={{ padding: "8px 12px", textAlign: h === "Parceiro" ? "left" : "center", color: C.cinzaTexto, fontWeight: 600, fontSize: 11, textTransform: "uppercase" }}>{h}</th>
              ))}
            </tr></thead>
            <tbody>
              {d.parceiros.filter(p => p.aging !== null).sort((a, b) => (b.aging ?? 0) - (a.aging ?? 0)).map((p, i) => {
                const crit = p.aging > 7;
                return (
                  <tr key={i} style={{ borderBottom: `1px solid ${C.cinzaBorda}`, background: crit ? C.amareloLight + "55" : "transparent" }}>
                    <td style={{ padding: "9px 12px", fontWeight: 600 }}>{p.nome}</td>
                    <td style={{ padding: "9px 12px", textAlign: "center" }}>
                      <span style={{ fontWeight: 700, color: crit ? C.amarelo : C.verde }}>{dias(p.aging)}</span>
                    </td>
                    {p.agingH.map((h, j) => (
                      <td key={j} style={{ padding: "9px 12px", textAlign: "center", fontSize: 12, color: h.val !== null && h.val > 7 ? C.amarelo : C.cinzaTexto }}>
                        {h.val !== null ? dias(h.val) : "—"}
                      </td>
                    ))}
                  </tr>
                );
              })}
            </tbody>
          </table>
        </Card>
      )}

      {view === "atrasos" && (
        <Card>
          <SecHead>Atrasos em Aberto por Parceiro — W{w}</SecHead>
          <table style={{ width: "100%", borderCollapse: "collapse", fontSize: 13 }}>
            <thead><tr style={{ borderBottom: `2px solid ${C.cinzaBorda}` }}>
              {["Parceiro", "Atrasos", ...d.parceiros[0].atrasosH.map(h => h.week)].map(h => (
                <th key={h} style={{ padding: "8px 12px", textAlign: h === "Parceiro" ? "left" : "center", color: C.cinzaTexto, fontWeight: 600, fontSize: 11, textTransform: "uppercase" }}>{h}</th>
              ))}
            </tr></thead>
            <tbody>
              {d.parceiros.filter(p => p.atrasos !== null).sort((a, b) => (b.atrasos ?? 0) - (a.atrasos ?? 0)).map((p, i) => {
                const crit = p.atrasos > 15;
                return (
                  <tr key={i} style={{ borderBottom: `1px solid ${C.cinzaBorda}`, background: crit ? C.vermelhoLight + "44" : "transparent" }}>
                    <td style={{ padding: "9px 12px", fontWeight: 600 }}>{p.nome}</td>
                    <td style={{ padding: "9px 12px", textAlign: "center" }}>
                      <span style={{ fontWeight: 700, color: crit ? C.vermelho : p.atrasos > 5 ? C.amarelo : C.verde }}>{p.atrasos}</span>
                    </td>
                    {p.atrasosH.map((h, j) => (
                      <td key={j} style={{ padding: "9px 12px", textAlign: "center", fontSize: 12, color: h.val !== null && h.val > 15 ? C.vermelho : C.cinzaTexto }}>
                        {h.val !== null ? h.val : "—"}
                      </td>
                    ))}
                  </tr>
                );
              })}
            </tbody>
          </table>
        </Card>
      )}
    </div>
  );
}

// ── ABA: FINANCEIRO ──────────────────────────────────────────────────────────
function TabFinanceiro({ d }) {
  const w = d.currentWeek;
  return (
    <div style={{ display: "flex", flexDirection: "column", gap: 24, paddingTop: 28 }}>
      <SecHead>Devolução CD — W{w}</SecHead>
      <div style={{ display: "grid", gridTemplateColumns: "repeat(auto-fit, minmax(175px, 1fr))", gap: 12 }}>
        <KpiCard label="Rentabilidade Devolução" value={d.rentDev} prev={d.rentDevPrev} format={pct} meta={0.505} semana={w} />
        <KpiCard label="Abrangência Devolução" value={d.abrangDev} prev={null} format={pct} semana={w} />
        <KpiCard label="Notas Coletadas" value={d.notasDev} prev={null} format={num} semana={w} />
        <KpiCard label="Valor Recebido" value={d.valorDev} prev={null} format={brl} semana={w} />
        <KpiCard label="Estoque Evitado" value={d.estoqueDev} prev={null} format={brl} semana={w} />
      </div>

      <div style={{ display: "grid", gridTemplateColumns: "1fr 1fr", gap: 16 }}>
        <Card>
          <SecHead>Rentabilidade Devolução — últimas {d.histRentDev.length} semanas (meta 50,5%)</SecHead>
          <MiniLine data={d.histRentDev} color={C.azul} meta={0.505} />
          <div style={{ display: "flex", justifyContent: "space-between", marginTop: 4 }}>
            {d.histRentDev.map((h, i) => (
              <span key={i} style={{ fontSize: 11, color: C.cinzaTexto }}>{h.week}: <strong style={{ color: h.val !== null && h.val < 0.505 ? C.vermelho : C.texto }}>{h.val !== null ? pct(h.val) : "—"}</strong></span>
            ))}
          </div>
        </Card>
        <Card>
          <SecHead>Rentabilidade Reversa — últimas {d.histRentRev.length} semanas (meta 18%)</SecHead>
          <MiniLine data={d.histRentRev} color={C.laranja} meta={0.18} />
          <div style={{ display: "flex", justifyContent: "space-between", marginTop: 4 }}>
            {d.histRentRev.map((h, i) => (
              <span key={i} style={{ fontSize: 11, color: C.cinzaTexto }}>{h.week}: <strong style={{ color: h.val !== null && h.val < 0.18 ? C.vermelho : C.texto }}>{h.val !== null ? pct(h.val) : "—"}</strong></span>
            ))}
          </div>
        </Card>
      </div>

      <SecHead>Reversa — W{w}</SecHead>
      <div style={{ display: "grid", gridTemplateColumns: "repeat(auto-fit, minmax(175px, 1fr))", gap: 12 }}>
        <KpiCard label="Rentabilidade Reversa" value={d.rentRev} prev={d.rentRevPrev} format={pct} meta={0.18} semana={w} />
        <KpiCard label="Abrangência Reversa" value={d.abrangRev} prev={null} format={pct} meta={0.8} semana={w} />
        <KpiCard label="Notas Coletadas" value={d.notasRev} prev={null} format={num} semana={w} />
        <KpiCard label="Valor Recebido" value={d.valorRev} prev={null} format={brl} semana={w} />
        <KpiCard label="Estoque Evitado" value={d.estoqueRev} prev={null} format={brl} semana={w} />
      </div>

      <SecHead>PNC — Produtos Não Conformes — W{w}</SecHead>
      <div style={{ display: "grid", gridTemplateColumns: "repeat(auto-fit, minmax(175px, 1fr))", gap: 12 }}>
        <KpiCard label="% CDs com Lotes Disponíveis" value={d.pctCDs} prev={null} format={pct} semana={w} />
        <KpiCard label="Lotes Precificados" value={d.lotesPrec} prev={null} format={num} semana={w} />
        <KpiCard label="Receita Total PNC" value={d.receitaPNC} prev={null} format={brl} semana={w} />
        <KpiCard label="Rentabilidade PNC" value={d.rentPNC} prev={null} format={pct} meta={0.17} semana={w} />
      </div>
    </div>
  );
}

// ── ABA: ALERTAS ─────────────────────────────────────────────────────────────
function TabAlertas({ d }) {
  const META_SLA = 0.86;
  const alerts = [];

  // ── Semana atual ──
  if (d.slaGeral !== null && d.slaGeral < META_SLA)
    alerts.push({ tipo: "danger", secao: "Semana atual", msg: `SLA Geral abaixo da meta: ${pct(d.slaGeral)} (meta 86%)` });

  d.parceiros.filter(p => p.sla !== null && p.sla < META_SLA).forEach(p =>
    alerts.push({ tipo: "danger", secao: "Semana atual", msg: `${p.nome}: SLA crítico em ${pct(p.sla)} — ${p.atrasos ?? 0} atrasos` }));

  d.parceiros.filter(p => p.sla !== null && p.slaPrev !== null && p.sla - p.slaPrev < -0.05 && p.sla >= META_SLA).forEach(p =>
    alerts.push({ tipo: "warn", secao: "Semana atual", msg: `${p.nome}: queda de ${((p.sla - p.slaPrev) * 100).toFixed(1)}pp vs semana anterior` }));

  d.parceiros.filter(p => p.csat !== null && p.csat < 0.85).forEach(p =>
    alerts.push({ tipo: "warn", secao: "Semana atual", msg: `${p.nome}: CSAT baixo — ${pct(p.csat)} (meta 85%)` }));

  d.parceiros.filter(p => p.aging !== null && p.aging > 7).forEach(p =>
    alerts.push({ tipo: "warn", secao: "Semana atual", msg: `${p.nome}: aging médio de problema em ${dias(p.aging)} — acima de 7 dias` }));

  d.parceiros.filter(p => p.atrasos !== null && p.atrasos > 20).forEach(p =>
    alerts.push({ tipo: "danger", secao: "Semana atual", msg: `${p.nome}: ${p.atrasos} atrasos em aberto — volume crítico` }));

  if (d.aderencia !== null && d.aderencia < 0.9)
    alerts.push({ tipo: "warn", secao: "Semana atual", msg: `Aderência ao agendamento: ${pct(d.aderencia)} — abaixo da meta (95%)` });

  if (d.agingMedio !== null && d.agingMedio > 5)
    alerts.push({ tipo: "warn", secao: "Semana atual", msg: `Aging médio de problema geral: ${dias(d.agingMedio)} — acima de 5 dias` });

  if (d.csat !== null && d.csat < 0.85)
    alerts.push({ tipo: "warn", secao: "Semana atual", msg: `CSAT geral abaixo da meta: ${pct(d.csat)} (meta 85%)` });

  if (d.rentDev !== null && d.rentDev < 0.505)
    alerts.push({ tipo: "warn", secao: "Semana atual", msg: `Rentabilidade Devolução abaixo da meta: ${pct(d.rentDev)} (meta 50,5%)` });

  if (d.rentRev !== null && d.rentRev < 0.18)
    alerts.push({ tipo: "warn", secao: "Semana atual", msg: `Rentabilidade Reversa abaixo da meta: ${pct(d.rentRev)} (meta 18%)` });

  if (d.abrangRev !== null && d.abrangRev < 0.8)
    alerts.push({ tipo: "warn", secao: "Semana atual", msg: `Abrangência Reversa abaixo de 80%: ${pct(d.abrangRev)}` });

  if (d.rentPNC !== null && d.rentPNC < 0.17)
    alerts.push({ tipo: "warn", secao: "Semana atual", msg: `Rentabilidade PNC abaixo da meta: ${pct(d.rentPNC)} (meta 17%)` });

  // ── Tendência 4 semanas ──
  const slaH = d.histSla.filter(h => h.val !== null);
  if (slaH.length >= 3 && slaH.slice(-3).every((h, i, arr) => i === 0 || h.val < arr[i-1].val)) {
    const diff = slaH[slaH.length-1].val - slaH[slaH.length-3].val;
    alerts.push({ tipo: "warn", secao: "Tendência 4 semanas", msg: `SLA Geral em queda por 3 semanas consecutivas (${(diff*100).toFixed(1)}pp acumulado)` });
  }

  const adherH = d.histAderencia.filter(h => h.val !== null);
  if (adherH.filter(h => h.val < 0.9).length >= 3)
    alerts.push({ tipo: "warn", secao: "Tendência 4 semanas", msg: `Aderência abaixo de 90% em ${adherH.filter(h => h.val < 0.9).length} das últimas ${adherH.length} semanas — padrão persistente` });

  const csatH = d.histCsat.filter(h => h.val !== null);
  if (csatH.length >= 3 && csatH.slice(-3).every((h, i, arr) => i === 0 || h.val < arr[i-1].val))
    alerts.push({ tipo: "warn", secao: "Tendência 4 semanas", msg: `CSAT em queda por 3 semanas consecutivas` });

  d.parceiros.filter(p => p.sla !== null && p.slaPrev !== null && p.sla < META_SLA && p.slaPrev < META_SLA).forEach(p =>
    alerts.push({ tipo: "danger", secao: "Tendência 4 semanas", msg: `${p.nome}: SLA abaixo de 86% por 2 semanas seguidas` }));

  d.parceiros.filter(p => p.csat !== null && p.csatPrev !== null && p.csat < 0.85 && p.csatPrev < 0.85).forEach(p =>
    alerts.push({ tipo: "warn", secao: "Tendência 4 semanas", msg: `${p.nome}: CSAT abaixo de 85% por 2 semanas seguidas` }));

  // Rentabilidade devolução abaixo nas últimas 3
  const rentDevH = d.histRentDev.filter(h => h.val !== null);
  if (rentDevH.filter(h => h.val < 0.505).length >= 3)
    alerts.push({ tipo: "warn", secao: "Tendência 4 semanas", msg: `Rentabilidade Devolução abaixo da meta em ${rentDevH.filter(h=>h.val<0.505).length} das últimas ${rentDevH.length} semanas` });

  if (alerts.length === 0)
    alerts.push({ tipo: "ok", secao: "Semana atual", msg: "Todos os indicadores principais dentro do esperado esta semana." });

  const secoes = [...new Set(alerts.map(a => a.secao))];

  return (
    <div style={{ display: "flex", flexDirection: "column", gap: 20, paddingTop: 28 }}>
      {secoes.map(secao => (
        <div key={secao}>
          <SecHead>{secao === "Semana atual" ? "⚡ Semana atual" : "📈 Tendência últimas 4 semanas"}</SecHead>
          <div style={{ display: "flex", flexDirection: "column", gap: 6 }}>
            {alerts.filter(a => a.secao === secao).map((a, i) => (
              <div key={i} style={{
                display: "flex", alignItems: "flex-start", gap: 10, padding: "10px 14px",
                borderRadius: 8, border: `1px solid ${a.tipo === "danger" ? C.vermelho : a.tipo === "warn" ? C.amarelo : C.verde}`,
                background: a.tipo === "danger" ? C.vermelhoLight : a.tipo === "warn" ? C.amareloLight : C.verdeLight,
              }}>
                <span style={{ fontSize: 16, flexShrink: 0 }}>{a.tipo === "danger" ? "🚨" : a.tipo === "warn" ? "⚠️" : "✅"}</span>
                <span style={{ fontSize: 13, fontWeight: 500, color: C.texto }}>{a.msg}</span>
              </div>
            ))}
          </div>
        </div>
      ))}
    </div>
  );
}

// ── ABA: TEXTO WEEKLY ────────────────────────────────────────────────────────
function TabTexto({ d }) {
  const META_SLA = 0.86;
  const w = d.currentWeek;
  const criticos = d.parceiros.filter(p => p.sla !== null && p.sla < META_SLA);
  const bons = d.parceiros.filter(p => p.sla !== null && p.sla >= 0.95);
  const diffSla = d.slaGeral !== null && d.slaGeralPrev !== null ? d.slaGeral - d.slaGeralPrev : null;
  const slaH = d.histSla.filter(h => h.val !== null);

  const tendenciaSla = () => {
    if (slaH.length < 2) return "";
    const diff = slaH[slaH.length-1].val - slaH[0].val;
    if (Math.abs(diff) < 0.005) return `SLA estável nas últimas ${slaH.length} semanas.`;
    return `SLA com tendência de **${diff > 0 ? "alta" : "queda"}** nas últimas ${slaH.length} semanas (${diff > 0 ? "+" : ""}${(diff*100).toFixed(1)}pp, de ${pct(slaH[0].val)} para ${pct(slaH[slaH.length-1].val)}).`;
  };

  const planoAcao = () => {
    const acoes = [];
    criticos.forEach(p => {
      if (p.atrasos > 20) acoes.push(`Reunião urgente com **${p.nome}** — ${p.atrasos} notas em atraso.`);
      else acoes.push(`Monitorar **${p.nome}** — SLA em ${pct(p.sla)}, acionar plano de recuperação.`);
    });
    d.parceiros.filter(p => p.atrasos !== null && p.atrasos > 20).filter(p => !criticos.includes(p)).forEach(p =>
      acoes.push(`${p.nome}: ${p.atrasos} atrasos em aberto — priorizar acionamento.`));
    d.parceiros.filter(p => p.csat !== null && p.csat < 0.85).forEach(p =>
      acoes.push(`${p.nome}: CSAT em ${pct(p.csat)} — analisar motivos de notas 1-3.`));
    if (d.aderencia !== null && d.aderencia < 0.9) acoes.push("Reforçar aderência ao agendamento com parceiros ofensores.");
    if (d.agingMedio !== null && d.agingMedio > 5) acoes.push("Revisar casos com aging > 5 dias — telefones inválidos e duplicidades.");
    if (d.rentDev !== null && d.rentDev < 0.505) acoes.push(`Rentabilidade Devolução em ${pct(d.rentDev)} — abaixo da meta de 50,5%, investigar mix de lotes.`);
    if (d.rentRev !== null && d.rentRev < 0.18) acoes.push(`Rentabilidade Reversa em ${pct(d.rentRev)} — abaixo da meta de 18%.`);
    if (slaH.length >= 3 && slaH.slice(-3).every((h, i, arr) => i === 0 || h.val < arr[i-1].val))
      acoes.push("SLA em queda por 3 semanas consecutivas — investigar causa raiz.");
    if (acoes.length === 0) acoes.push("Manter acompanhamento padrão. Nenhuma ação urgente identificada.");
    return acoes;
  };

  const linhas = [
    `SLA Geral W${w}: **${pct(d.slaGeral)}**${diffSla !== null ? ` (${diffSla > 0 ? "+" : ""}${(diffSla*100).toFixed(1)}pp vs semana anterior)` : ""} — ${d.slaGeral >= META_SLA ? "acima" : "abaixo"} da meta de 86%.`,
    tendenciaSla(),
    d.agendamento ? `Agendamento sistêmico: **${pct(d.agendamento)}** | Aderência: **${pct(d.aderencia)}**${d.aderencia < 0.9 ? " ⚠️ abaixo de 95%." : "."}` : "",
    criticos.length ? `Parceiros críticos: ${criticos.map(p => `**${p.nome}** (${pct(p.sla)}, ${p.atrasos} atrasos)`).join(", ")}.` : "Nenhum parceiro crítico esta semana.",
    bons.length ? `Destaques positivos: ${bons.map(p => `${p.nome} (${pct(p.sla)})`).join(", ")}.` : "",
    d.rentDev ? `Rentabilidade Devolução: **${pct(d.rentDev)}** (meta 50,5%) | Reversa: **${pct(d.rentRev)}** (meta 18%).` : "",
    d.abrangRev ? `Abrangência Reversa: **${pct(d.abrangRev)}** (meta 80%).` : "",
    d.rentPNC ? `PNC — Rentabilidade: **${pct(d.rentPNC)}** (meta 17%) | Receita: **${brl(d.receitaPNC)}**.` : "",
  ].filter(Boolean);

  return (
    <div style={{ display: "flex", flexDirection: "column", gap: 20, paddingTop: 28 }}>
      <Card>
        <SecHead>📊 Resumo W{w}</SecHead>
        <div style={{ display: "flex", flexDirection: "column", gap: 8 }}>
          {linhas.map((l, i) => (
            <p key={i} style={{ fontSize: 14, lineHeight: 1.6, color: C.texto, margin: 0 }}
              dangerouslySetInnerHTML={{ __html: l.replace(/\*\*(.*?)\*\*/g, '<strong>$1</strong>') }} />
          ))}
        </div>
      </Card>

      <Card>
        <SecHead>📅 Últimas {d.histSla.length} Semanas</SecHead>
        <table style={{ width: "100%", borderCollapse: "collapse", fontSize: 13 }}>
          <thead><tr style={{ borderBottom: `2px solid ${C.cinzaBorda}` }}>
            {["Indicador", ...d.histSla.map(h => h.week)].map((h, i) => (
              <th key={i} style={{ padding: "6px 10px", textAlign: i === 0 ? "left" : "center", color: C.cinzaTexto, fontWeight: 600, fontSize: 11 }}>{h}</th>
            ))}
          </tr></thead>
          <tbody>
            {[
              { label: "SLA Geral", hist: d.histSla, fmt: pct, meta: 0.86 },
              { label: "CSAT", hist: d.histCsat, fmt: pct, meta: 0.85 },
              { label: "Agendamento", hist: d.histAgendamento, fmt: pct, meta: 0.95 },
              { label: "Aderência", hist: d.histAderencia, fmt: pct, meta: 0.95 },
              { label: "Rent. Devolução", hist: d.histRentDev, fmt: pct, meta: 0.505 },
              { label: "Rent. Reversa", hist: d.histRentRev, fmt: pct, meta: 0.18 },
              { label: "Abrang. Reversa", hist: d.histAbrangRev, fmt: pct, meta: 0.8 },
            ].map((row, ri) => (
              <tr key={ri} style={{ borderBottom: `1px solid ${C.cinzaBorda}` }}>
                <td style={{ padding: "7px 10px", fontWeight: 600, color: C.texto }}>{row.label}</td>
                {row.hist.map((h, hi) => {
                  const atMeta = h.val !== null ? h.val >= row.meta : null;
                  const isLast = hi === row.hist.length - 1;
                  return (
                    <td key={hi} style={{ padding: "7px 10px", textAlign: "center", fontWeight: isLast ? 700 : 400, color: atMeta === false ? C.vermelho : atMeta === true ? C.verde : C.cinzaTexto, background: isLast ? C.cinzaFundo : "transparent" }}>
                      {h.val !== null ? row.fmt(h.val) : "—"}
                    </td>
                  );
                })}
              </tr>
            ))}
          </tbody>
        </table>
      </Card>

      <div style={{ background: "#FFF7ED", border: `1px solid ${C.laranjaLight}`, borderRadius: 10, padding: "16px 20px" }}>
        <SecHead style={{ color: C.laranja }}>🎯 Plano de Ação Sugerido</SecHead>
        <ul style={{ margin: 0, paddingLeft: 18, display: "flex", flexDirection: "column", gap: 6 }}>
          {planoAcao().map((a, i) => (
            <li key={i} style={{ fontSize: 14, color: C.texto, lineHeight: 1.5 }}
              dangerouslySetInnerHTML={{ __html: a.replace(/\*\*(.*?)\*\*/g, '<strong>$1</strong>') }} />
          ))}
        </ul>
      </div>
    </div>
  );
}

// ── APP PRINCIPAL ─────────────────────────────────────────────────────────────
// ── Encode/decode com compressão ────────────────────────────────────────────
async function encodeData(data) {
  const json = JSON.stringify(data);
  const bytes = new TextEncoder().encode(json);
  const cs = new CompressionStream("deflate");
  const writer = cs.writable.getWriter();
  writer.write(bytes);
  writer.close();
  const compressed = await new Response(cs.readable).arrayBuffer();
  const b64 = btoa(String.fromCharCode(...new Uint8Array(compressed)));
  return b64.replace(/\+/g, "-").replace(/\//g, "_").replace(/=/g, "");
}

async function decodeData(encoded) {
  try {
    const b64 = encoded.replace(/-/g, "+").replace(/_/g, "/");
    const binary = atob(b64);
    const bytes = new Uint8Array(binary.length).map((_, i) => binary.charCodeAt(i));
    const ds = new DecompressionStream("deflate");
    const writer = ds.writable.getWriter();
    writer.write(bytes);
    writer.close();
    const decompressed = await new Response(ds.readable).arrayBuffer();
    const json = new TextDecoder().decode(decompressed);
    return JSON.parse(json);
  } catch { return null; }
}

async function getDataFromURL() {
  const params = new URLSearchParams(window.location.search);
  const d = params.get("d");
  if (d) return await decodeData(d);
  return null;
}

// ── App Principal ─────────────────────────────────────────────────────────────
export default function App() {
  const [data, setData] = useState(null);
  const [loading, setLoading] = useState(false);
  const [tab, setTab] = useState("overview");
  const [copied, setCopied] = useState(false);
  const [fromURL, setFromURL] = useState(false);

  // Carregar dados da URL ao montar
  useEffect(() => {
    const params = new URLSearchParams(window.location.search);
    if (params.get("d")) {
      setLoading(true);
      getDataFromURL().then(d => {
        if (d) { setData(d); setFromURL(true); }
        setLoading(false);
      });
    }
  }, []);

  const onFile = useCallback((e) => {
    const file = e.target.files[0];
    if (!file) return;
    setLoading(true);
    const reader = new FileReader();
    reader.onload = (ev) => {
      try {
        const parsed = parseWeekly(new Uint8Array(ev.target.result));
        setData(parsed);
        // Limpar params da URL ao subir novo arquivo
        window.history.replaceState({}, "", window.location.pathname);
        setFromURL(false);
        
      }
      catch (err) { alert("Erro ao ler o arquivo.\n" + err.message); }
      setLoading(false);
    };
    reader.readAsArrayBuffer(file);
  }, []);

  const onDrop = useCallback((e) => {
    e.preventDefault();
    const file = e.dataTransfer.files[0];
    if (file) onFile({ target: { files: [file] } });
  }, [onFile]);

  const exportLink = useCallback(async () => {
    if (!data) return;
    setCopied("loading");
    try {
      const encoded = await encodeData(data);
      const url = `${window.location.origin}${window.location.pathname}?d=${encoded}`;
      await navigator.clipboard.writeText(url);
    } catch (e) {
      console.error(e);
    }
    setCopied("done");
    setTimeout(() => setCopied(false), 3000);
  }, [data]);

  const tabs = [
    { id: "overview",    label: "Visão Geral" },
    { id: "parceiros",   label: "Parceiros" },
    { id: "financeiro",  label: "Financeiro" },
    { id: "alertas",     label: "Alertas" },
    { id: "texto",       label: "Texto Weekly" },
  ];

  return (
    <div style={{ minHeight: "100vh", background: C.cinzaFundo, fontFamily: "'Inter','Segoe UI',sans-serif", color: C.texto }}>
      <div style={{ background: C.laranja, padding: "0 32px" }}>
        <div style={{ maxWidth: 1100, margin: "0 auto", display: "flex", alignItems: "center", justifyContent: "space-between", height: 56 }}>
          <div style={{ display: "flex", alignItems: "center", gap: 12 }}>
            <span style={{ fontSize: 22 }}>🪵</span>
            <span style={{ color: "#fff", fontWeight: 700, fontSize: 16 }}>Weekly Parça</span>
            {data && <span style={{ color: "#fff9", fontSize: 13 }}>— W{data.currentWeek} / 2026</span>}
          </div>
          <div style={{ display: "flex", gap: 8, alignItems: "center" }}>
            {data && (
              <button onClick={exportLink} style={{
                background: copied ? "rgba(255,255,255,0.35)" : "rgba(255,255,255,0.15)",
                border: "1px solid rgba(255,255,255,0.3)", borderRadius: 8,
                padding: "6px 14px", cursor: "pointer", color: "#fff", fontSize: 13, fontWeight: 600,
                transition: "background 0.2s",
              }}>
                {copied === "loading" ? "⏳ Gerando link..." : copied === "done" ? "✓ Link copiado!" : "🔗 Compartilhar link"}
              </button>
            )}
            <label style={{ background: "rgba(255,255,255,0.15)", border: "1px solid rgba(255,255,255,0.3)", borderRadius: 8, padding: "6px 14px", cursor: "pointer", color: "#fff", fontSize: 13, fontWeight: 600 }}>
              {data ? "📂 Trocar arquivo" : "📂 Subir Weekly (.xlsx)"}
              <input type="file" accept=".xlsx" onChange={onFile} style={{ display: "none" }} />
            </label>
          </div>
        </div>
      </div>

      <div style={{ maxWidth: 1100, margin: "0 auto", padding: "0 32px" }}>
        {!data && !loading && (
          <div onDrop={onDrop} onDragOver={e => e.preventDefault()}
            style={{ marginTop: 48, border: `2px dashed ${C.cinzaBorda}`, borderRadius: 16, padding: "64px 32px", textAlign: "center", cursor: "pointer", background: C.cinzaCard }}
            onClick={() => document.querySelector('input[type=file]').click()}>
            <div style={{ fontSize: 48, marginBottom: 16 }}>📊</div>
            <p style={{ fontSize: 20, fontWeight: 700, marginBottom: 8 }}>Arraste ou clique para subir a Weekly Parça</p>
            <p style={{ fontSize: 14, color: C.cinzaTexto }}>Arquivo .xlsx — o dashboard detecta automaticamente a última semana preenchida</p>
          </div>
        )}

        {loading && (
          <div style={{ marginTop: 80, textAlign: "center" }}>
            <div style={{ fontSize: 40, marginBottom: 16 }}>⏳</div>
            <p style={{ fontSize: 16, color: C.cinzaTexto }}>Lendo e calculando indicadores...</p>
          </div>
        )}

        {data && (
          <>
            {fromURL && (
              <div style={{ marginTop: 16, padding: "10px 16px", background: C.azulLight, border: `1px solid ${C.azul}`, borderRadius: 8, fontSize: 13, color: C.azul, fontWeight: 500 }}>
                📎 Você está visualizando um dashboard compartilhado — W{data.currentWeek}/2026. Para atualizar, suba um novo arquivo xlsx.
              </div>
            )}
            <div style={{ display: "flex", gap: 4, marginTop: 16, borderBottom: `2px solid ${C.cinzaBorda}` }}>
              {tabs.map(t => (
                <button key={t.id} onClick={() => setTab(t.id)} style={{
                  padding: "10px 20px", border: "none", background: "transparent", cursor: "pointer",
                  fontSize: 14, fontWeight: tab === t.id ? 700 : 500,
                  color: tab === t.id ? C.laranja : C.cinzaTexto,
                  borderBottom: tab === t.id ? `2px solid ${C.laranja}` : "2px solid transparent",
                  marginBottom: -2,
                }}>{t.label}</button>
              ))}
            </div>

            {tab === "overview"   && <TabOverview    d={data} />}
            {tab === "parceiros"  && <TabParceiros   d={data} />}
            {tab === "financeiro" && <TabFinanceiro  d={data} />}
            {tab === "alertas"    && <TabAlertas     d={data} />}
            {tab === "texto"      && <TabTexto       d={data} />}

            <div style={{ height: 40 }} />
          </>
        )}
      </div>
    </div>
  );
}
