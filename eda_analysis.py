import pandas as pd
import matplotlib.pyplot as plt
import seaborn as sns
from sqlalchemy import create_engine
import urllib
import os

os.makedirs('charts', exist_ok=True)

# ── CONNECTION (SQLAlchemy) ────────────────────────────────────────
params = urllib.parse.quote_plus(
    "DRIVER={ODBC Driver 17 for SQL Server};"
    "SERVER=LAPTOP-LG4BEQ1J\\SQLEXPRESS;"
    "DATABASE=FinanceDB;"
    "Trusted_Connection=yes;"
)
engine = create_engine(f"mssql+pyodbc:///?odbc_connect={params}")
print("✓ SQL Server connected")

# ── DATA LOAD ─────────────────────────────────────────────────────
# 'date' aur 'close' — actual column names jo SSMS ne banaye
df_pl = pd.read_sql("SELECT * FROM vw_pl_summary", engine)

df_stocks = pd.read_sql("""
    SELECT [date], ticker, [close], volume
    FROM stocks
    WHERE [date] >= '2023-01-01'
    ORDER BY [date]
""", engine)

# date column ko datetime mein badlo
df_stocks['date'] = pd.to_datetime(df_stocks['date'])

print(f"✓ P&L rows loaded    : {len(df_pl)}")
print(f"✓ Stock rows loaded  : {len(df_stocks)}")

# ── BASIC ANALYSIS ────────────────────────────────────────────────
print("\n── P&L Summary (Average) ──")
print(df_pl.groupby('ticker')[['revenue','net_profit','net_margin_pct']].mean().round(2))

# ── CHART 1: Revenue Trend ────────────────────────────────────────
plt.figure(figsize=(12, 5))
for ticker in df_pl['ticker'].unique():
    data = df_pl[df_pl['ticker'] == ticker]
    plt.plot(data['quarter'], data['revenue'] / 1e6,
             marker='o', label=ticker, linewidth=2)

plt.title('Quarterly Revenue Trend (₹ Million)', fontsize=14, fontweight='bold')
plt.xlabel('Quarter')
plt.ylabel('Revenue (₹ Million)')
plt.xticks(rotation=45)
plt.legend()
plt.tight_layout()
plt.savefig('charts/revenue_trend.png', dpi=150)
plt.close()
print("✓ Chart 1 saved — revenue_trend.png")

# ── CHART 2: Net Margin Comparison ───────────────────────────────
avg_margin = df_pl.groupby('ticker')['net_margin_pct'].mean().sort_values()

plt.figure(figsize=(8, 5))
bars = plt.barh(avg_margin.index, avg_margin.values,
                color=['#2196F3','#4CAF50','#FF9800','#9C27B0','#F44336'])
plt.bar_label(bars, fmt='%.1f%%', padding=4)
plt.title('Average Net Profit Margin by Company', fontsize=13, fontweight='bold')
plt.xlabel('Net Margin %')
plt.tight_layout()
plt.savefig('charts/net_margin.png', dpi=150)
plt.close()
print("✓ Chart 2 saved — net_margin.png")

# ── CHART 3: P&L Waterfall ────────────────────────────────────────
latest = df_pl[df_pl['quarter'] == '2023-Q4'].set_index('ticker')

fig, ax = plt.subplots(figsize=(10, 5))
x = range(len(latest))
ax.bar(x, latest['revenue']/1e6,      label='Revenue',      color='#2196F3', alpha=0.8)
ax.bar(x, latest['gross_profit']/1e6, label='Gross Profit', color='#4CAF50', alpha=0.8)
ax.bar(x, latest['net_profit']/1e6,   label='Net Profit',   color='#FF9800', alpha=0.8)
ax.set_xticks(list(x))
ax.set_xticklabels(latest.index, rotation=30)
ax.set_title('P&L Comparison — Q4 2023 (₹ Million)', fontsize=13, fontweight='bold')
ax.set_ylabel('₹ Million')
ax.legend()
plt.tight_layout()
plt.savefig('charts/pl_waterfall.png', dpi=150)
plt.close()
print("✓ Chart 3 saved — pl_waterfall.png")

# ── CHART 4: Normalized Stock Performance ────────────────────────
# close column use kar rahe hain (close_price nahi)
pivot = df_stocks.pivot(index='date', columns='ticker', values='close')
pivot_norm = pivot.div(pivot.iloc[0]) * 100   # Base = 100 se normalize

plt.figure(figsize=(12, 5))
for col in pivot_norm.columns:
    plt.plot(pivot_norm.index, pivot_norm[col], label=col, linewidth=1.5)

plt.title('Normalized Stock Performance 2023 (Base = 100)', fontsize=13, fontweight='bold')
plt.xlabel('Date')
plt.ylabel('Indexed Price')
plt.legend()
plt.tight_layout()
plt.savefig('charts/stock_performance.png', dpi=150)
plt.close()
print("✓ Chart 4 saved — stock_performance.png")

print("\n✓ EDA complete! D:\\python\\finance_project\\charts\\ folder check karo")