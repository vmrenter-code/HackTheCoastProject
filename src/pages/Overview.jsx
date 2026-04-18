import { useState } from "react";
import { skuData } from "../data/mockData";
import {
  BarChart, Bar, XAxis, YAxis, Tooltip, ResponsiveContainer, Legend,
} from "recharts";

const STATUS_CONFIG = {
  critical: { label: "Critical", color: "#ef4444" },
  warning:  { label: "Warning",  color: "#f59e0b" },
  ok:       { label: "OK",       color: "#22c55e" },
};

export default function Overview() {
  const [selected, setSelected] = useState(null);
  const [filter, setFilter] = useState("all");

  const filtered = filter === "all" ? skuData : skuData.filter((s) => s.status === filter);

  const chartData = skuData.map((s) => ({
    name: s.name.split(" ").slice(0, 2).join(" "),
    "Raw Demand":     s.rawDemand,
    "Cleaned Demand": s.cleanedDemand,
  }));

  return (
    <div className="page">
      <div className="stat-grid">
        <StatCard label="Total SKUs"        value={skuData.length} />
        <StatCard label="Critical Reorders" value={skuData.filter(s=>s.status==="critical").length} accent="#ef4444" />
        <StatCard label="Avg. Demand Lift"  value={`+${Math.round(((skuData.reduce((a,s)=>a+s.cleanedDemand,0)/skuData.reduce((a,s)=>a+s.rawDemand,0))-1)*100)}%`} accent="#6ee7b7" />
        <StatCard label="Hidden Lost Sales" value={skuData.reduce((a,s)=>a+s.lostSales,0).toLocaleString()} />
      </div>

      <div className="card">
        <div className="card-header">
          <h2>Raw vs. Cleaned Demand Signal</h2>
          <p className="card-sub">Cleaned demand reveals suppressed volume hidden in raw sales data</p>
        </div>
        <ResponsiveContainer width="100%" height={240}>
          <BarChart data={chartData} barGap={2}>
            <XAxis dataKey="name" tick={{ fontSize: 11, fill: "#94a3b8" }} />
            <YAxis tick={{ fontSize: 11, fill: "#94a3b8" }} />
            <Tooltip contentStyle={{ background: "#1e293b", border: "1px solid #334155", borderRadius: 8, color: "#f1f5f9" }} />
            <Legend />
            <Bar dataKey="Raw Demand"     fill="#334155" radius={[4,4,0,0]} />
            <Bar dataKey="Cleaned Demand" fill="#6ee7b7" radius={[4,4,0,0]} />
          </BarChart>
        </ResponsiveContainer>
      </div>

      <div className="card">
        <div className="card-header row-between">
          <h2>SKU Detail</h2>
          <div className="filter-pills">
            {["all","critical","warning","ok"].map((f) => (
              <button key={f} className={`pill ${filter===f?"active":""}`} onClick={() => setFilter(f)}>
                {f.charAt(0).toUpperCase()+f.slice(1)}
              </button>
            ))}
          </div>
        </div>
        <table className="data-table">
          <thead>
            <tr>
              <th>SKU</th><th>Name</th><th>Channel</th>
              <th>Raw Demand</th><th>Cleaned</th><th>Lost Sales</th>
              <th>Days Left</th><th>Status</th>
            </tr>
          </thead>
          <tbody>
            {filtered.map((s) => (
              <tr key={s.id} className={`table-row ${selected===s.id?"selected":""}`} onClick={() => setSelected(selected===s.id?null:s.id)}>
                <td className="mono">{s.id}</td>
                <td>{s.name}</td>
                <td><span className="channel-tag">{s.channel}</span></td>
                <td className="mono">{s.rawDemand.toLocaleString()}</td>
                <td className="mono clean">{s.cleanedDemand.toLocaleString()}</td>
                <td className="mono warn">{s.lostSales > 0 ? `+${s.lostSales.toLocaleString()}` : "—"}</td>
                <td className="mono">{s.daysUntilStockout}d</td>
                <td>
                  <span className="status-badge" style={{ background: STATUS_CONFIG[s.status].color+"22", color: STATUS_CONFIG[s.status].color }}>
                    {STATUS_CONFIG[s.status].label}
                  </span>
                </td>
              </tr>
            ))}
          </tbody>
        </table>
      </div>
    </div>
  );
}

function StatCard({ label, value, accent }) {
  return (
    <div className="stat-card">
      <div className="stat-value" style={accent ? { color: accent } : {}}>{value}</div>
      <div className="stat-label">{label}</div>
    </div>
  );
}