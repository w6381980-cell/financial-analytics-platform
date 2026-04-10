A complete end-to-end financial analytics project built from scratch — covering data generation, SQL database design, exploratory data analysis, ML-based revenue forecasting, automated PDF/Excel reporting, and an interactive Power BI P&L Dashboard.

Domain: Finance & Stock Market
Stack: Python · SQL Server · Power BI
Type: Intermediate-to-Advanced Data Analytics Project


🚀 Features

📦 Synthetic Data Generator — Realistic stock prices & quarterly P&L data using Faker + NumPy
🗄️ SQL Server Database — Normalized schema with Views, Stored Procedures & Window Functions
📈 Exploratory Data Analysis — 4 professional charts (Revenue Trend, Net Margin, P&L, Stock Performance)
🔮 ML Forecasting — Linear Regression + ARIMA(1,1,1) for next 4 quarters revenue projection
📄 Automated Reports — PDF report (ReportLab) + Excel report (openpyxl) with charts & tables
📊 Power BI Dashboard — Interactive P&L Dashboard with KPI cards, slicers & DAX measures


📁 Project Structure
financial-analytics-platform/
│
├── 1_data_generator.py       # Synthetic stock & P&L data generator
├── 2_db_setup.sql            # SQL Server schema, views, stored procedures
├── 3_eda_analysis.py         # Exploratory data analysis + 4 charts
├── 4_projections.py          # Revenue forecasting (LR + ARIMA)
├── 5_report_generator.py     # Auto PDF + Excel report generator
│
├── charts/                   # Auto-generated chart images
│   ├── revenue_trend.png
│   ├── net_margin.png
│   ├── pl_waterfall.png
│   ├── stock_performance.png
│   └── forecast.png
│
├── report.pdf                # Generated financial report
├── report.xlsx               # Generated Excel workbook
├── FinanceDashboard.pbix     # Power BI dashboard file
│
└── README.md

🏢 Dataset Details
TableRowsDescriptionstocks3,650Daily OHLCV data — 5 companies × 2 yearsfinancials40Quarterly P&L — 5 companies × 8 quartersprojections402024 revenue forecast — LR + ARIMA
Companies: ALPHACORP · BETAFIN · DELTATECH · EPSILONBK · GAMMAINV
Period: January 2022 — December 2023 (Historical) + 2024 (Forecast)

🛠️ Tech Stack
LayerTechnologyData GenerationPython · Faker · NumPy · PandasDatabaseSQL Server 2019 · pyodbc · SQLAlchemyAnalysisPandas · Matplotlib · SeabornForecastingScikit-learn · Statsmodels (ARIMA)ReportingReportLab · OpenPyXLDashboardPower BI Desktop · DAX

⚙️ Setup & Installation
Prerequisites

Python 3.10+
SQL Server 2019+ (or SQL Server Express)
Power BI Desktop
ODBC Driver 17 for SQL Server

1. Clone the repository
bashgit clone https://github.com/your-username/financial-analytics-platform.git
cd financial-analytics-platform
2. Install Python dependencies
bashpip install faker pandas numpy pyodbc sqlalchemy statsmodels scikit-learn matplotlib seaborn reportlab openpyxl
3. Run in order
bash# Step 1 — Generate synthetic data
python 1_data_generator.py

# Step 2 — Setup SQL Server database
# Open SSMS → Run 2_db_setup.sql
# Then load CSVs:
# EXEC sp_load_stocks 'C:\path\to\stocks.csv'
# EXEC sp_load_stocks 'C:\path\to\financials.csv'

# Step 3 — Exploratory Data Analysis
python 3_eda_analysis.py

# Step 4 — Revenue Forecasting
python 4_projections.py

# Step 5 — Generate Reports
python 5_report_generator.py

# Step 6 — Open FinanceDashboard.pbix in Power BI Desktop
4. Update connection string
In files 3_eda_analysis.py, 4_projections.py, 5_report_generator.py:
pythonconn_str = (
    "DRIVER={ODBC Driver 17 for SQL Server};"
    "SERVER=YOUR_SERVER_NAME\\SQLEXPRESS;"  # ← change this
    "DATABASE=FinanceDB;"
    "Trusted_Connection=yes;"
)

📊 Power BI Dashboard
KPI Cards
MetricDAX MeasureTotal RevenueSUM(vw_pl_summary[revenue])Total Net ProfitSUM(vw_pl_summary[net_profit])Gross Margin %AVERAGE(vw_pl_summary[gross_margin_pct])Net Margin %AVERAGE(vw_pl_summary[net_margin_pct])YoY Growth %AVERAGE(vw_pl_summary[yoy_growth_pct])
Visuals

📈 Quarterly Revenue Trend (Clustered Bar Chart)
📉 Net Margin % by Company (Line Chart)
🔮 Actual vs Projected Revenue (Line Chart)
📊 Stock Performance — Normalized (Line Chart)
🎛️ Ticker Slicer + Quarter Slicer (Interactive filters)


📈 Sample Results
── P&L Summary (FY 2023 Average) ──
               Revenue      Net Profit   Net Margin
ALPHACORP     ₹445.8M      ₹74.5M       16.74%
BETAFIN       ₹739.7M      ₹126.5M      17.40%
DELTATECH     ₹162.3M      ₹25.8M       16.04%
EPSILONBK     ₹141.4M      ₹22.2M       15.77%
GAMMAINV      ₹296.6M      ₹47.9M       16.00%

── 2024 Revenue Forecast (Linear Regression) ──
BETAFIN       → Highest growth projected
ALPHACORP     → Stable growth trend
GAMMAINV      → Moderate growth

🧠 Concepts Covered
Python         → OOP, data manipulation, file I/O, API calls
SQL            → DDL, DML, Views, Stored Procedures, Window Functions
Machine Learning → Linear Regression, ARIMA time series forecasting
Data Viz       → Matplotlib, Seaborn, Power BI
Reporting      → PDF generation, Excel automation
BI Tools       → DAX measures, data modeling, relationships

📝 License
This project is licensed under the MIT License — see the LICENSE file for details.

👨‍💻 Author
Hardik
Intermediate Data Analyst | Python · SQL · Power BI
