import pandas as pd
import numpy as np
import matplotlib.pyplot as plt
from sqlalchemy import create_engine, text
import urllib
import warnings
warnings.filterwarnings('ignore')
from sklearn.linear_model import LinearRegression
from statsmodels.tsa.arima.model import ARIMA

# ── CONNECTION ────────────────────────────────────────────────────
params = urllib.parse.quote_plus(
    "DRIVER={ODBC Driver 17 for SQL Server};"
    "SERVER=LAPTOP-LG4BEQ1J\\SQLEXPRESS;"
    "DATABASE=FinanceDB;"
    "Trusted_Connection=yes;"
)
engine = create_engine(f"mssql+pyodbc:///?odbc_connect={params}")
print("✓ SQL Server connected")

# ── DATA LOAD ─────────────────────────────────────────────────────
df = pd.read_sql(
    "SELECT ticker, quarter, revenue FROM financials ORDER BY ticker, quarter",
    engine
)

# Quarter number assign karo per company (1 se 8)
# Logic: ML model numbers samajhta hai, '2022-Q1' nahi
df['q_num'] = df.groupby('ticker').cumcount() + 1

future_quarters = ['2024-Q1', '2024-Q2', '2024-Q3', '2024-Q4']
all_projections = []

print("\n── Forecasting per company ──")

# ── PER COMPANY FORECAST ──────────────────────────────────────────
for ticker in df['ticker'].unique():
    data = df[df['ticker'] == ticker].copy()
    X = data['q_num'].values.reshape(-1, 1)  # Feature matrix
    y = data['revenue'].values               # Target values

    # ── MODEL 1: LINEAR REGRESSION ──────────────────────────────
    # y = m*x + c  →  trend line se future predict karo
    lr = LinearRegression()
    lr.fit(X, y)
    future_X  = np.array([[9], [10], [11], [12]])
    lr_preds  = lr.predict(future_X)

    for q, pred in zip(future_quarters, lr_preds):
        all_projections.append({
            'ticker':        ticker,
            'quarter':       q,
            'projected_rev': round(float(pred), 2),
            'lower_bound':   round(float(pred) * 0.90, 2),
            'upper_bound':   round(float(pred) * 1.10, 2),
            'model_used':    'LinearRegression'
        })
    print(f"  ✓ {ticker} — Linear Regression done")

    # ── MODEL 2: ARIMA(1,1,1) ───────────────────────────────────
    # p=1: 1 lag use karo (pichla quarter)
    # d=1: ek baar difference karo (trend hatao)
    # q=1: 1 moving average term
    try:
        model  = ARIMA(y, order=(1, 1, 1))
        fitted = model.fit()
        fc     = fitted.get_forecast(steps=4)
        preds  = fc.predicted_mean
        ci     = fc.conf_int()

        for i, q in enumerate(future_quarters):
            all_projections.append({
                'ticker':        ticker,
                'quarter':       q,
                'projected_rev': round(float(preds[i]), 2),
                'lower_bound':   round(float(ci.iloc[i, 0]), 2),
                'upper_bound':   round(float(ci.iloc[i, 1]), 2),
                'model_used':    'ARIMA(1,1,1)'
            })
        print(f"  ✓ {ticker} — ARIMA done")

    except Exception as e:
        print(f"  ✗ {ticker} — ARIMA failed: {e}")

# ── SQL SERVER MEIN SAVE ──────────────────────────────────────────
df_proj = pd.DataFrame(all_projections)

with engine.begin() as conn:
    conn.execute(text("DELETE FROM projections"))  # Purana clear karo

    for _, row in df_proj.iterrows():
        conn.execute(text("""
            INSERT INTO projections
                (ticker, quarter, projected_rev, lower_bound, upper_bound, model_used)
            VALUES
                (:ticker, :quarter, :proj, :low, :high, :model)
        """), {
            'ticker':  row['ticker'],
            'quarter': row['quarter'],
            'proj':    row['projected_rev'],
            'low':     row['lower_bound'],
            'high':    row['upper_bound'],
            'model':   row['model_used']
        })

print(f"\n✓ {len(df_proj)} projections saved to SQL Server")

# ── FORECAST CHART ────────────────────────────────────────────────
df_hist = df[['ticker', 'quarter', 'revenue']]
lr_proj = df_proj[df_proj['model_used'] == 'LinearRegression']

fig, axes = plt.subplots(2, 3, figsize=(16, 9))
axes = axes.flatten()

for idx, ticker in enumerate(df['ticker'].unique()):
    ax   = axes[idx]
    hist = df_hist[df_hist['ticker'] == ticker]
    proj = lr_proj[lr_proj['ticker'] == ticker]

    # Historical line
    ax.plot(hist['quarter'], hist['revenue'] / 1e6,
            marker='o', color='#2196F3', label='Actual', linewidth=2)

    # Forecast line
    ax.plot(proj['quarter'], proj['projected_rev'] / 1e6,
            marker='s', color='#FF5722', linestyle='--',
            label='Forecast', linewidth=2)

    # Confidence band (±10%)
    ax.fill_between(proj['quarter'],
                    proj['lower_bound'] / 1e6,
                    proj['upper_bound'] / 1e6,
                    alpha=0.2, color='#FF5722', label='±10% band')

    # Historical aur forecast ke beech vertical line
    ax.axvline(x=7.5, color='gray', linestyle=':', linewidth=1)
    ax.text(7.6, ax.get_ylim()[0], 'Forecast →',
            fontsize=7, color='gray')

    ax.set_title(ticker, fontweight='bold', fontsize=11)
    ax.set_ylabel('₹ Million')
    ax.tick_params(axis='x', rotation=45)
    ax.legend(fontsize=7)
    ax.grid(True, alpha=0.3)

axes[-1].set_visible(False)
plt.suptitle('Revenue Forecast — Next 4 Quarters (2024)',
             fontsize=14, fontweight='bold')
plt.tight_layout()
plt.savefig('charts/forecast.png', dpi=150)
plt.close()
print("✓ Forecast chart saved — charts/forecast.png")

# ── PROJECTION SUMMARY PRINT ──────────────────────────────────────
print("\n── 2024 Revenue Projection Summary (Linear Regression) ──")
summary = lr_proj.groupby('ticker')['projected_rev'].sum() / 1e6
for ticker, rev in summary.items():
    print(f"  {ticker:12s} → ₹{rev:.1f}M (Annual 2024 projected)")