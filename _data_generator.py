import pandas as pd
import numpy as np
from faker import Faker
import random
from datetime import date, timedelta

fake = Faker()
np.random.seed(42)

# ── 1. STOCK PRICE DATA ──────────────────────────────────────────────
# 5 companies ke liye 2 saal ka daily stock data
companies = ['ALPHACORP', 'BETAFIN', 'GAMMAINV', 'DELTATECH', 'EPSILONBK']
sectors   = ['Finance', 'Banking', 'Investment', 'FinTech', 'Banking']

records = []
for ticker, sector in zip(companies, sectors):
    price = random.uniform(100, 500)          # Starting price
    start = date(2022, 1, 1)
    for i in range(730):                       # 2 years daily
        # Random Walk: kal ka price = aaj + thoda upar/neeche
        price = max(10, price * (1 + np.random.normal(0.0003, 0.015)))
        volume = random.randint(100_000, 5_000_000)
        records.append({
            'date': start + timedelta(days=i),
            'ticker': ticker,
            'sector': sector,
            'open':  round(price * random.uniform(0.98, 1.0), 2),
            'high':  round(price * random.uniform(1.0,  1.02), 2),
            'low':   round(price * random.uniform(0.97, 1.0),  2),
            'close': round(price, 2),
            'volume': volume
        })

df_stocks = pd.DataFrame(records)
df_stocks.to_csv('stocks.csv', index=False)
print(f"Stocks: {len(df_stocks)} rows")

# ── 2. QUARTERLY P&L DATA ────────────────────────────────────────────
# Logic: Revenue → COGS → Gross Profit → Operating Exp → Net Profit
pl_records = []
quarters = ['2022-Q1','2022-Q2','2022-Q3','2022-Q4',
            '2023-Q1','2023-Q2','2023-Q3','2023-Q4']

for ticker in companies:
    base_revenue = random.uniform(50_000_000, 500_000_000)  # 5Cr to 50Cr
    for q in quarters:
        # Revenue thoda growth karta hai har quarter
        revenue = base_revenue * random.uniform(0.95, 1.15)
        cogs    = revenue * random.uniform(0.45, 0.60)       # Cost of Goods Sold
        gross_profit = revenue - cogs
        op_expenses  = revenue * random.uniform(0.20, 0.30)  # Salaries, rent, etc.
        ebit         = gross_profit - op_expenses            # Earnings before tax
        tax          = ebit * 0.25                           # 25% tax
        net_profit   = ebit - tax

        pl_records.append({
            'quarter': q,
            'ticker': ticker,
            'revenue': round(revenue, 2),
            'cogs':    round(cogs, 2),
            'gross_profit': round(gross_profit, 2),
            'operating_expenses': round(op_expenses, 2),
            'ebit': round(ebit, 2),
            'net_profit': round(net_profit, 2),
            'gross_margin_pct': round((gross_profit / revenue) * 100, 2),
            'net_margin_pct':   round((net_profit   / revenue) * 100, 2),
        })
        base_revenue = revenue  # Next quarter base = is quarter ka revenue

df_pl = pd.DataFrame(pl_records)
df_pl.to_csv('financials.csv', index=False)
print(f"P&L: {len(df_pl)} rows")