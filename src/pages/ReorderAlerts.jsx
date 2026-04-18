import { skuData } from "../data/mockData";

const sorted = [...skuData].sort((a, b) => a.daysUntilStockout - b.daysUntilStockout);

export default function ReorderAlerts() {
  const critical = sorted.filter((s) => s.status === "critical");
  const warning  = sorted.filter((s) => s.status === "warning");
  const ok       = sorted.filter((s) => s.status === "ok");

  return (
    <div className="page">
      <div className="alert-intro card">
        <div className="alert-intro-icon">⚠</div>
        <div>
          <h2>{critical.length} SKUs require immediate action</h2>
          <p className="card-sub">Based on current inventory levels, lead times, and cleaned demand run rates.</p>
        </div>
      </div>
      <AlertGroup title="Critical — Order Now" items={critical} color="#ef4444" />
      <AlertGroup title="Warning — Order Within 2 Weeks" items={warning} color="#f59e0b" />
      <AlertGroup title="On Track" items={ok} color="#22c55e" />
    </div>
  );
}

function AlertGroup({ title, items, color }) {
  if (!items.length) return null;
  return (
    <div className="card">
      <div className="card-header">
        <h2 style={{ color }}>{title}</h2>
      </div>
      <div className="alert-list">
        {items.map((s) => (
          <div key={s.id} className="alert-row">
            <div className="alert-left">
              <div className="alert-name">{s.name}</div>
              <div className="alert-meta">{s.id} · {s.channel} · Supplier: {s.supplier}</div>
            </div>
            <div className="alert-metrics">
              <Metric label="Days Left"      value={`${s.daysUntilStockout}d`} color={color} />
              <Metric label="In Stock"       value={s.currentInventory.toLocaleString()} />
              <Metric label="Reorder Point"  value={s.reorderPoint.toLocaleString()} />
              <Metric label="Lead Time"      value={`${s.leadTimeDays}d`} />
              {s.suggestedOrderQty > 0 && (
                <Metric label="Suggested Order" value={s.suggestedOrderQty.toLocaleString()} color="#6ee7b7" />
              )}
            </div>
          </div>
        ))}
      </div>
    </div>
  );
}

function Metric({ label, value, color }) {
  return (
    <div className="metric-chip">
      <div className="metric-val" style={color ? { color } : {}}>{value}</div>
      <div className="metric-lbl">{label}</div>
    </div>
  );
}