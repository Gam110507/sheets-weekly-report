"""
Generate a believable order export for a small online store.

This is the RAW tab: what actually lands in a spreadsheet when someone exports
orders from Shopify, WooCommerce or a back office. It is deliberately not tidy,
because the point of the project is that the report survives real input:

  * refunds and cancellations mixed in with sales
  * shipping and discount lines that are not product revenue
  * a few blank customer emails (guest checkout)
  * dates as text in more than one format
  * the occasional duplicated row from a double export

    python3 make_orders.py        # writes data/orders_raw.csv
"""

import math
import random
from pathlib import Path

import numpy as np
import pandas as pd

SEED = 7
random.seed(SEED)
np.random.seed(SEED)

WEEKS = 14
START = pd.Timestamp("2026-05-04")          # a Monday

PRODUCTS = [
    ("Linen Bedding Set",        89.00, 0.16),
    ("Weighted Blanket 7kg",     69.00, 0.13),
    ("Bamboo Pillow Pair",       39.00, 0.12),
    ("Cotton Waffle Throw",      45.00, 0.10),
    ("Silk Pillowcase",          29.00, 0.10),
    ("Mattress Protector",       35.00, 0.08),
    ("Duvet Inner 10.5 tog",     75.00, 0.07),
    ("Fitted Sheet Single",      22.00, 0.07),
    ("Linen Napkin Set of 4",    26.00, 0.06),
    ("Lavender Room Mist",       14.00, 0.06),
    ("Storage Basket Large",     32.00, 0.05),
]
NAMES = [p[0] for p in PRODUCTS]
PRICES = {p[0]: p[1] for p in PRODUCTS}
WEIGHTS = np.array([p[2] for p in PRODUCTS], dtype=float)
WEIGHTS /= WEIGHTS.sum()

CHANNELS = (["Website"] * 62 + ["Instagram"] * 16 + ["Etsy"] * 12
            + ["Amazon"] * 7 + ["Wholesale"] * 3)
COUNTRIES = (["United Kingdom"] * 70 + ["Ireland"] * 8 + ["Germany"] * 7
             + ["France"] * 6 + ["Netherlands"] * 5 + ["Spain"] * 4)

# A gentle upward trend with a promo spike in week 9 and a dip over a quiet
# week, so week-over-week numbers have something real to say.
TREND = [1.00, 1.03, 0.94, 1.08, 1.11, 1.05, 1.17, 1.21,
         1.62, 1.24, 1.19, 1.31, 1.28, 1.36]


def money(x):
    return round(float(x), 2)


rows = []
order_no = 10_000

for w in range(WEEKS):
    week_start = START + pd.Timedelta(days=7 * w)
    n_orders = int(np.random.normal(46, 6) * TREND[w])

    for _ in range(max(n_orders, 5)):
        order_no += 1
        # Weekend-lighter, midweek-heavier, like most consumer stores.
        day = int(np.random.choice(range(7), p=[.17, .17, .16, .15, .14, .11, .10]))
        ts = (week_start + pd.Timedelta(days=day, hours=random.randint(7, 22),
                                        minutes=random.randint(0, 59)))
        channel = random.choice(CHANNELS)
        country = random.choice(COUNTRIES)
        email = "" if random.random() < 0.11 else f"customer{random.randint(1, 900)}@example.com"

        n_lines = 1 + int(np.random.poisson(0.6))
        picks = np.random.choice(len(NAMES), size=min(n_lines, 4),
                                 replace=False, p=WEIGHTS)

        for pi in picks:
            name = NAMES[pi]
            qty = 1 + int(np.random.poisson(0.35))
            unit = PRICES[name]
            # Occasional promo pricing, which is why the report must not assume
            # a fixed price per product.
            if random.random() < 0.09:
                unit = money(unit * random.choice([0.8, 0.85, 0.9]))

            rows.append({
                "Order ID": f"#{order_no}",
                "Date": ts,
                "Line Type": "Product",
                "Item": name,
                "Qty": qty,
                "Unit Price": money(unit),
                "Line Total": money(unit * qty),
                "Channel": channel,
                "Country": country,
                "Customer Email": email,
                "Status": "Fulfilled",
            })

        # Shipping is charged to the customer but is NOT product revenue.
        if random.random() < 0.72:
            ship = random.choice([3.95, 4.95, 5.95, 0.00])
            rows.append({
                "Order ID": f"#{order_no}", "Date": ts, "Line Type": "Shipping",
                "Item": "Shipping", "Qty": 1, "Unit Price": ship,
                "Line Total": ship, "Channel": channel, "Country": country,
                "Customer Email": email, "Status": "Fulfilled",
            })

        # Discounts arrive as negative lines.
        if random.random() < 0.14:
            disc = -money(random.choice([5, 10, 15]))
            rows.append({
                "Order ID": f"#{order_no}", "Date": ts, "Line Type": "Discount",
                "Item": "Promo code", "Qty": 1, "Unit Price": disc,
                "Line Total": disc, "Channel": channel, "Country": country,
                "Customer Email": email, "Status": "Fulfilled",
            })

df = pd.DataFrame(rows)

# ── the messiness ───────────────────────────────────────────────────────────

# Refunds: a copy of a real line, negated, marked Refunded.
refunds = df[df["Line Type"] == "Product"].sample(frac=0.035, random_state=SEED).copy()
refunds["Qty"] *= -1
refunds["Line Total"] = -refunds["Line Total"].abs()
refunds["Status"] = "Refunded"
refunds["Date"] = refunds["Date"] + pd.Timedelta(days=random.randint(2, 9))
df = pd.concat([df, refunds], ignore_index=True)

# Cancelled orders still sit in the export.
idx = df.sample(frac=0.02, random_state=SEED + 1).index
df.loc[idx, "Status"] = "Cancelled"

# A double export duplicates some rows exactly.
df = pd.concat([df, df.sample(frac=0.015, random_state=SEED + 2)], ignore_index=True)

df = df.sort_values("Date").reset_index(drop=True)

# Dates written two ways, because exports are never consistent.
def stamp(ts, i):
    return (ts.strftime("%Y-%m-%d %H:%M") if i % 5 else ts.strftime("%d/%m/%Y %H:%M"))

df["Date"] = [stamp(t, i) for i, t in enumerate(df["Date"])]

Path("data").mkdir(exist_ok=True)
df.to_csv("data/orders_raw.csv", index=False)

prod = df[(df["Line Type"] == "Product") & (df["Status"] == "Fulfilled")]
print(f"wrote data/orders_raw.csv   rows={len(df):,}")
print(f"  weeks                {WEEKS}")
print(f"  distinct orders      {df['Order ID'].nunique():,}")
print(f"  product revenue      {prod['Line Total'].sum():,.2f}")
print(f"  refund lines         {(df['Status'] == 'Refunded').sum()}")
print(f"  cancelled lines      {(df['Status'] == 'Cancelled').sum()}")
print(f"  shipping/discount    {(df['Line Type'] != 'Product').sum()}")
print(f"  guest checkouts      {(df['Customer Email'] == '').sum()}")
