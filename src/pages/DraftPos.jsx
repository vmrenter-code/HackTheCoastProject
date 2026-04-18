import { draftPOs } from "../data/mockData";

export default function DraftPOs() {
  return (
    <div className="page">
      <div className="card">
        <div className="card-header">
          <h2>Draft Purchase Orders</h2>
          <p className="card-sub">
            System-generated recommendations based on cleaned demand, current inventory,
            and supplier lead times. Buyers should review before submitting.
          </p>
        </div>
        <table className="data-table">
          <thead>
            <tr>
              <th>SKU</th>
              <th>Product</th>
              <th>Supplier</th>
              <th>Qty to Order</th>
              <th>Order By</th>
              <th>Est. Arrival</th>
              <th>Action</th>
            </tr>
          </thead>
          <tbody>
            {draftPOs.map((po) => (
              <tr key={po.sku} className="table-row">
                <td className="mono">{po.sku}</td>
                <td>{po.name}</td>
                <td>{po.supplier}</td>
                <td className="mono clean">{po.qty.toLocaleString()}</td>
                <td className="mono warn">{po.orderBy}</td>
                <td className="mono">{po.estimatedArrival}</td>
                <td>
                  <button className="action-btn">Approve PO</button>
                </td>
              </tr>
            ))}
          </tbody>
        </table>
        <div className="po-disclaimer">
          ⓘ Quantities are calculated from cleaned demand run rates and include a 15% safety
          stock buffer. Verify with your supplier before submitting.
        </div>
      </div>
    </div>
  );
}