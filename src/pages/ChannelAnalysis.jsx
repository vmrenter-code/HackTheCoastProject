import { channelData, monthlyDemand } from "../data/mockData";
import {
  LineChart, Line, XAxis, YAxis, Tooltip, ResponsiveContainer, Legend,
} from "recharts";

const CHANNEL_NOTES = {
  "American Market": "Planogram-driven. Demand is predictable but resets seasonally. Watch for shelf resets in Q1 and Q3.",
  "Health Food":     "Steady baseline via UNFI/KeHE. Promotional spikes obscure true demand — cleaned signal shows ~14% more volume.",
  "Asian Market":    "Opportunistic and reactive. Spikes are sudden with little lead time. Maintain higher safety stock for top SKUs.",
  "eCom":            "Long-tail, growing. Small per-SKU volumes but fastest-growing channel. Promo volume is significant relative to size.",
};

export default function ChannelAnalysis() {
  return (
    <div className="page">
      <div className="channel-grid">
        {channelData.map((c) => (
          <div key={c.channel} className="card channel-card">
            <div className="channel-name">{c.channel}</div>
            <div className="channel-stats">
              <div>
                <div className="channel-stat-val raw">{c.rawDemand.toLocaleString()}</div>
                <div className="channel-stat-lbl">Raw Demand</div>
              </div>
              <div className="channel-arrow">→</div>
              <div>
                <div className="channel-stat-val clean">{c.cleanedDemand.toLocaleString()}</div>
                <div className="channel-stat-lbl">Cleaned Demand</div>
              </div>
              <div>
                <div className="channel-stat-val">{c.skus}</div>
                <div className="channel-stat-lbl">SKUs</div>
              </div>
            </div>
            <div className="lift-bar-wrap">
              <div className="lift-bar" style={{ width: `${Math.min(100, Math.round((c.cleanedDemand / c.rawDemand - 1) * 100 * 5))}%` }} />
              <span className="lift-label">+{Math.round((c.cleanedDemand / c.rawDemand - 1) * 100)}% signal lift</span>
            </div>
            <p className="channel-note">{CHANNEL_NOTES[c.channel]}</p>
          </div>
        ))}
      </div>

      <div className="card">
        <div className="card-header">
          <h2>12-Month Demand Trend</h2>
          <p className="card-sub">Cleaned demand consistently reveals more volume than raw sales data captures</p>
        </div>
        <ResponsiveContainer width="100%" height={260}>
          <LineChart data={monthlyDemand}>
            <XAxis dataKey="month" tick={{ fontSize: 11, fill: "#94a3b8" }} />
            <YAxis tick={{ fontSize: 11, fill: "#94a3b8" }} />
            <Tooltip contentStyle={{ background: "#1e293b", border: "1px solid #334155", borderRadius: 8, color: "#f1f5f9" }} />
            <Legend />
            <Line type="monotone" dataKey="raw"     stroke="#475569" strokeWidth={2} dot={false} name="Raw Demand" />
            <Line type="monotone" dataKey="cleaned" stroke="#6ee7b7" strokeWidth={2} dot={false} name="Cleaned Demand" />
          </LineChart>
        </ResponsiveContainer>
      </div>
    </div>
  );
}