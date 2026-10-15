# MIS ANALYTICS PRO v2.0 - Enterprise ERP Intelligence Software

Advanced Management Information System (MIS) Analytics platform designed for processing, analyzing, and visualizing Excel & ERP reports with automated KPI generation, multi-dimensional breakdowns, dynamic pivot analysis, anomaly detection, and deep data grids.

---

## 🗄️ Database Information
- **Database Engine**: MySQL 8.0 (Service: `MySQL80`)
- **Database Name**: `MIS_DB`
- **Host**: `localhost`
- **User**: `root`
- **Password**: `Rupam@123`
- **Port**: `3306`

---

## 🚀 Live Access
- **Web Application URL**: [http://localhost:5000](http://localhost:5000)
- **Status**: Running and Connected to `MIS_DB`

---

## ⚡ What's New in v3.0 (Advanced Graphs, Report Downloads & Deep Filters)

### 1. 📈 6 Advanced Enterprise Charts Suite
- **Dual-Axis Volume vs Transaction Combo Chart**: Financial Volume (Left Y-Axis Bar) plotted together with Transaction Count (Right Y-Axis Smooth Line).
- **Pareto 80/20 Distribution Analysis**: Ranked category bars with an overlaid cumulative percentage curve and an 80% contribution line.
- **Dimension Share Doughnut Chart**: Slice percentage share with center total label and click-to-filter drill-down.
- **Ranked Top Performers Bar Chart**: Top 10 rankings with hover ticket size and drill-down clicking.
- **Channel / Settlement Mode Distribution**: Breakdown by Cash, Card, UPI, TPA Insurance, etc.
- **Cumulative Revenue Growth Area Curve**: Running accumulation curve over chronological time.
- **1-Click Chart PNG Download**: Every chart has a `📷 PNG` button to download high-resolution chart images.

### 2. 📥 Multi-Format Executive Report Downloads
- **📊 Download Multi-Sheet Excel Report (`.xlsx`)**:
  - `Executive_Summary`: Domain, Financial KPIs, Active Metric, Filter details, AI Findings.
  - `Category_Breakdown`: Department/Category volume, total value, average value, market share %.
  - `Filtered_Records`: Complete records matching current filters.
  - `Audit_Exceptions`: Internal audit flags and anomalies.
- **📄 Download Executive PDF Report (`.pdf`)**:
  - Formatted printable management report with KPIs, observations, category tables, and signature blocks.
- **🖨️ Quick Print View**:
  - Print-optimized CSS stylesheet for standard office printers.

### 3. 🔍 100% Reactive Advanced Filter Engine
- Real-time multi-column rules (`=`, `!=`, `contains`, `>`, `<`).
- Dynamic Metric Switcher (e.g. switch between `NET AMOUNT`, `BILLED AMOUNT`, `DISCOUNT AMOUNT`, `PATIENT AMOUNT PAID`).
- Dynamic Dimension Switcher (e.g. switch between `DEPARTMENT`, `PRIMARY DOCTOR`, `DISCHARGE WARD`, `PANEL GROUP`).
- Direct synchronization with the Data Explorer grid table.

---

## 📁 Sample Excel Files Included
- [Sample_OPD_Billing_MIS.xlsx](file:///d:/Work/MIS%20ANALYSITS/Sample_OPD_Billing_MIS.xlsx) (Hospital OPD Patient Billing Report)
- [Sample_Sales_ERP_MIS.xlsx](file:///d:/Work/MIS%20ANALYSITS/Sample_Sales_ERP_MIS.xlsx) (Enterprise B2B Commercial Sales Report)

---

## 💻 Commands
- **Start Server**: `npm start`
- **Port**: `5000` (Configure in `.env`)
