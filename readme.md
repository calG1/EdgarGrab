# SECcel 📊

> Instant, one-click `Financial_Report.xlsx` Excel downloads restored to the SEC EDGAR interface.

[![License: MIT](https://img.shields.io/badge/License-MIT-yellow.svg)](https://opensource.org/licenses/MIT)
[![Tampermonkey](https://img.shields.io/badge/Tampermonkey-Compatible-004d40.svg)](https://www.tampermonkey.net/)
[![Platform](https://img.shields.io/badge/Platform-SEC%20EDGAR-blue.svg)](https://www.sec.gov/edgar)

---

## 🧐 The Problem

Recent updates to the SEC EDGAR user interface removed or obscured direct one-click links to download raw `Financial_Report.xlsx` workbooks. Analysts, developers, and researchers are frequently forced into cumbersome interactive web viewers just to copy numbers manually.

However, the SEC's archive servers **still generate and store** these multi-tab Excel files in the background under standardized directory endpoints:

```text
https://www.sec.gov/Archives/edgar/data/<CIK>/<AccessionNoWithoutHyphens>/Financial_Report.xlsx
```

**SECcel** is a lightweight userscript that automatically extracts the filing metadata from your current page, reconstructs the raw archive URL, and triggers an instant download directly in your browser.

---

## ✨ Features

* **Floating Action Button (FAB):** Pins an unobtrusive `📊 Download Financial_Report.xlsx` button to the bottom-right corner whenever you view a single filing or iXBRL report.
* **Inline Search Badges:** Automatically appends a compact `XLS` button next to individual filing entries on company browse and search tables.
* **Smart URL Resolution:** Normalizes CIK numbers (stripping leading zeros) and accession strings (stripping hyphens) on the fly.
* **Instant Feedback & Error Handling:** Provides real-time button status (`⏳ Verifying...` $\rightarrow$ `📥 Downloading...` $\rightarrow$ `✅ Downloaded!`) and throws an explicit toast notification if the filing lacks XBRL data (HTTP 404).
* **Single-Page App (SPA) Aware:** Uses a `MutationObserver` to ensure buttons persist dynamically as you navigate around EDGAR without requiring page reloads.

---

## 🚀 Installation

### 1. Install a Userscript Manager
You need a browser extension capable of executing user scripts. Install one of the following:
* [Tampermonkey](https://www.tampermonkey.net/) *(Recommended)*
* [Violentmonkey](https://violentmonkey.github.io/)

### 2. Install SECcel
Click the direct install link below:

👉 **[Install SECcel (`seccel.user.js`)](https://raw.githubusercontent.com/<YOUR-USERNAME>/<YOUR-REPO>/main/seccel.user.js)**

Tampermonkey will open an installation prompt. Click **Install**.

*(Replace `<YOUR-USERNAME>` and `<YOUR-REPO>` with your actual GitHub repository details.)*

---

## 🛠️ How It Works

When navigating any `sec.gov` page, SECcel executes the following workflow:

```
[ Active EDGAR Tab ]
         │
         ├── 1. Parses window.location / query params for CIK & Accession Number
         ├── 2. Normalizes parameters:
         │        • CIK: '0000320193' ➔ '320193'
         │        • Accession: '0000320193-24-000106' ➔ '000032019324000106'
         │
         ├── 3. Reconstructs Endpoint:
         │        https://www.sec.gov/Archives/edgar/data/320193/000032019324000106/Financial_Report.xlsx
         │
         ├── 4. Dispatches GET request (GM_xmlhttpRequest / Fetch)
         │        ├── Status 200: Converts response to Blob & triggers browser download
         │        └── Status 404: Displays toast notification detailing missing XBRL report
```

---

## 📋 What Filings Contain `Financial_Report.xlsx`?

The SEC only generates `Financial_Report.xlsx` files for submissions that include complete **XBRL / iXBRL** data packages:

| Filing Type | Contains Excel File? | Notes |
| :--- | :---: | :--- |
| **Form 10-K** (Annual Reports) | **Yes** | Complete multi-tab balance sheet, cash flows, income statements, footnotes |
| **Form 10-Q** (Quarterly Reports) | **Yes** | Standard quarterly financial tables |
| **Form 20-F / 40-F** (Foreign Issuers) | **Yes** | Full annual foreign issuer financials |
| **Form 8-K** (Current Reports) | **Conditional** | Only if the 8-K includes Item 2.02 and embedded XBRL tables |
| **Form 3, 4, 5** (Insider Trades) | **No** (404) | Insider filings do not contain full financial statements |
| **Form 13-D / 13-G** (Beneficial Ownership) | **No** (404) | Ownership disclosures do not generate XBRL workbooks |
| **Pre-2009 Filings** | **No** (404) | Filed before the SEC mandated XBRL formatting |

*If you click the download button on a filing that does not produce an Excel report, SECcel will display a clear error notification indicating the 404 status.*

---

## ⚙️ Manual Testing & Debugging

To verify the script manually without browser extensions:

1. Open a filing in EDGAR (e.g., Apple's Form 10-K).
2. Note the **CIK** (`0000320193`) and **Accession Number** (`0000320193-24-000106`).
3. Strip leading zeros from the CIK (`320193`) and hyphens from the accession (`000032019324000106`).
4. Paste the URL into your browser:
   ```text
   https://www.sec.gov/Archives/edgar/data/320193/000032019324000106/Financial_Report.xlsx
   ```
5. If the URL initiates a download, SECcel will trigger it automatically with a single click.

---

## 📄 License

Distributed under the **MIT License**. See [`LICENSE`](./LICENSE) for full details.