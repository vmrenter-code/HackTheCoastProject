import { useState } from "react";
import Overview from "./pages/Overview";
import ReorderAlerts from "./pages/ReorderAlerts";
import ChannelAnalysis from "./pages/ChannelAnalysis";
import DraftPOs from "./pages/DraftPos";

const NAV = [
  { id: "overview",  label: "SKU Overview",         icon: "⬡" },
  { id: "alerts",    label: "Reorder Alerts",        icon: "◈" },
  { id: "channels",  label: "Channel Analysis",      icon: "◎" },
  { id: "pos",       label: "Draft Purchase Orders", icon: "◻" },
];

export default function App() {
  const [active, setActive] = useState("overview");

  return (
    <div className="app-shell">
      <aside className="sidebar">
        <div className="sidebar-brand">
          <span className="brand-mark">PoP</span>
          <span className="brand-sub">Demand Intelligence</span>
        </div>
        <nav className="sidebar-nav">
          {NAV.map((item) => (
            <button
              key={item.id}
              className={`nav-item ${active === item.id ? "active" : ""}`}
              onClick={() => setActive(item.id)}
            >
              <span className="nav-icon">{item.icon}</span>
              <span className="nav-label">{item.label}</span>
            </button>
          ))}
        </nav>
        <div className="sidebar-footer">
          <div className="data-badge">
            <span className="dot live" />
            Live Data Feed
          </div>
        </div>
      </aside>

      <main className="main-content">
        <header className="top-bar">
          <div className="page-title">
            {NAV.find((n) => n.id === active)?.label}
          </div>
          <div className="top-bar-right">
            <span className="timestamp">
              Last updated: {new Date().toLocaleString()}
            </span>
          </div>
        </header>
        <div className="page-body">
          {active === "overview"  && <Overview />}
          {active === "alerts"    && <ReorderAlerts />}
          {active === "channels"  && <ChannelAnalysis />}
          {active === "pos"       && <DraftPOs />}
        </div>
      </main>
    </div>
  );
}