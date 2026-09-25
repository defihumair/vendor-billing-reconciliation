from pathlib import Path
import numpy as np
import pandas as pd

HEADS = [
    ("General cargo handling", "CBM", 85, ["Warehouse A", "Warehouse B", "Warehouse C", "Commercial"], (1400, 2800)),
    ("Brand X outbound", "CBM", 40, ["Warehouse C"], (2000, 3200)),
    ("Brand X inbound (manual)", "CBM", 38, ["Warehouse C"], (1200, 2000)),
    ("Brand X inbound (conveyor)", "CBM", 65, ["Warehouse C"], (400, 800)),
    ("Cargo removal & LCL", "CBM", 88, ["Warehouse A", "Warehouse B", "Warehouse C"], (50, 300)),
    ("Cargo shifting", "CBM", 42, ["Warehouse A", "Warehouse B", "Warehouse C"], (20, 550)),
    ("Sorting", "Cartons", 2.5, ["Warehouse A", "Warehouse B", "Warehouse C", "Commercial"], (1500, 7500)),
    ("Brand X sorting", "Cartons", 1.2, ["Warehouse C"], (35000, 52000)),
    ("Sunday working", "CBM", 50, ["Warehouse A", "Warehouse B", "Warehouse C"], (20, 150)),
    ("Labelling", "Cartons", 2.2, ["Warehouse B", "Warehouse C"], (300, 3000)),
]


def build(seed=7, year=2026, first_week=10, weeks=8):
    rng = np.random.default_rng(seed)
    rates = pd.DataFrame([{"billing_head": h, "uom": u, "rate": r} for h, u, r, _, _ in HEADS])
    internal, vendor = [], []
    for week in range(first_week, first_week + weeks):
        invoice_no = f"VA-{year}-W{week:02d}"
        for head, _, rate, facilities, (low, high) in HEADS:
            for facility in facilities:
                qty = int(rng.integers(low, high + 1))
                if head == "Sunday working" and rng.random() < 0.5:
                    qty = 0
                if qty > 0:
                    internal.append({"year": year, "week": week, "facility": facility, "billing_head": head, "qty": qty})
                roll = rng.random()
                if qty == 0:
                    vendor_qty = int(rng.integers(5, 40)) if rng.random() < 0.3 else 0
                elif roll < 0.60:
                    vendor_qty = qty
                elif roll < 0.83:
                    vendor_qty = qty + max(1, int(round(qty * rng.uniform(0.03, 0.15))))
                elif roll < 0.96:
                    vendor_qty = max(0, qty - max(1, int(round(qty * rng.uniform(0.02, 0.08)))))
                else:
                    vendor_qty = 0
                if vendor_qty <= 0:
                    continue
                billed_rate = rate * 1.05 if rng.random() < 0.05 else rate
                vendor.append({
                    "year": year, "week": week, "invoice_no": invoice_no, "facility": facility,
                    "billing_head": head, "qty": vendor_qty, "amount": round(vendor_qty * billed_rate, 2),
                })
    return pd.DataFrame(internal), pd.DataFrame(vendor), rates


def main():
    out = Path(__file__).parent / "sample_data"
    out.mkdir(exist_ok=True)
    internal, vendor, rates = build()
    internal.to_csv(out / "internal_volumes.csv", index=False)
    vendor.to_csv(out / "vendor_invoice.csv", index=False)
    rates.to_csv(out / "rate_card.csv", index=False)
    print(f"Wrote {len(internal)} internal rows, {len(vendor)} vendor rows and {len(rates)} rates to {out}")


if __name__ == "__main__":
    main()
