# Automated EHR Compliance & Data Quality Auditor

An automated Python-driven data auditing pipeline that transforms chaotic, messy Electronic Health Record (EHR) spreadsheet exports into polished, executive-ready Excel compliance dashboards—**without relying on fragile, manual pivot tables.**

This repository serves as an open-source, production-ready framework designed specifically for small-to-mid-size healthcare practices, community clinics, and independent medical groups struggling with "Excel Hell" and manual data compliance tracking.

---

## ⚡ The Problem This Tool Solves
Smaller healthcare facilities lose hours of operational efficiency and bleed revenue due to human errors during patient intake and data entry. Standard EHR reporting modules are notoriously rigid, outputting unreadable `.csv` or `.xlsx` files filled with formatting anomalies. 

Submitting flawed data leads to catastrophic insurance claim rejections, auditing delays, and compliance penalties.

This tool acts as an **automated data clerk and auditor**, instantly evaluating **5 critical healthcare compliance metrics** and generating a presentation-ready diagnostic workbook at a perfect **100% zoom scale viewport** for single-screen review.

---

## 🔍 Automated Compliance Validations
The core pipeline uses `pandas` to isolate data violations and maps them directly to an actionable remediation log:

1. **Invalid/Missing NPI Registry Check:** Flags records where the 10-digit National Provider Identifier (NPI) is missing, truncated, or malformed.
2. **Timeline Discrepancy Auditing:** Identifies human typos where a patient's discharge date is erroneously back-dated before their admission date.
3. **Missing Billing/ICD-10 Codes:** Extracts incomplete patient files missing the critical classification strings required for insurance billing reconciliation.
4. **Duplicate Provider NPI Profiles:** Detects system identity fragmentation where the exact same provider NPI is mapped to varying spelling typographies or multiple internal system IDs.
5. **Future Encounter Date Errors:** Flags logic breaks where admission dates are accidentally recorded in the future relative to the current live processing day.

---

## 🛠️ Architecture & Final Excel Design
The application parses raw files and leverages `openpyxl` to enforce professional corporate graphic design principles:

* **Tab 1: Executive Summary Dashboard**
  * Stripped of default Excel background gridlines to achieve a clean "application window" layout feel.
  * Prominent, centered KPI block cards summarizing high-level operational exposure metrics.
  * Side-by-side execution: A crisp data summary table flanked by an auto-sizing, multi-colored vertical column chart with clear numeric data labels (no overlapping diagonal axis labels).
* **Tab 2: Actionable Data Quality Tally**
  * Keeps native gridlines intact for structured data reading.
  * Freezes top header rows automatically for infinite scroll capabilities.
  * Conditional soft-red background fills to immediately pull the user's eye directly to specific data field failures.
  * Introduces a structural **"Action Required" directive column** telling staff exactly how to fix the issue in the master EHR registry.

---

## 🚀 Getting Started

### Prerequisites
Ensure you have Python 3.8+ installed along with the required processing libraries:

```bash
pip install pandas openpyxl
```

### Installation & Execution
1. Clone this repository to your local system:
   ```bash
   git clone https://github.com
   cd data-quality-log
   ```

2. Generate a fresh, messy baseline EHR file containing intentional compliance errors using the test data script:
   ```bash
   python generate_mock_data.py
   ```

3. Run the primary data auditing pipeline:
   ```bash
   python EHR_Compliance_Audit_Report.py
   ```

4. Open the freshly compiled `EHR_Compliance_Audit_Report.xlsx` file in Microsoft Excel to inspect the visual dashboard.

---

## 📂 Repository File Structure
* `EHR_Compliance_Audit_Report.py` — Core Python data engineering file managing file ingestion, compliance evaluations, and custom openpyxl Excel formatting.
* `generate_mock_data.py` — Developer utility script generating a synthetic `messy_ehr_export.csv` data layout embedded with structural anomalies for integration testing.
* `README.md` — Technical documentation and project overview.

---

## 💼 Custom Integrations & Healthcare Data Consulting
Every single healthcare organization configures its custom EHR fields, clinic location codes, and reporting schedules slightly differently. 

If your facility needs to bridge this exact automated reporting pipeline **directly into your live Athenahealth, eClinicalWorks, NextGen, or Epic system via secure database connections or API extractions**, let's build it.

I specialize in constructing secure, local automation scripts that eliminate spreadsheet manual labor for mid-sized medical practices.

👉 **[Click Here to Schedule a Free 15-Minute Workflow Discovery Call](https://scheduler.zoom.us/jessica-hayden-i79emd/free-discovery-call)**  
Let's talk about connecting this automation tool directly to your clinical database environment.
