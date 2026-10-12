// ==UserScript==
// @name         SECcel: EDGAR Direct Excel Downloader
// @namespace    https://github.com/
// @version      2.2
// @description  Restores direct Financial_Report.xlsx downloads on SEC EDGAR with an intelligent SEC REST API fallback formatted into standard multi-tab financial statements with native currency formatting.
// @author       Cal & Google Gemini
// @match        https://www.sec.gov/*
// @match        https://sec.gov/*
// @require      https://cdn.sheetjs.com/xlsx-latest/package/dist/xlsx.full.min.js
// @grant        GM_xmlhttpRequest
// @connect      www.sec.gov
// @connect      sec.gov
// @connect      data.sec.gov
// @run-at       document-idle
// ==/UserScript==

(function () {
    'use strict';

    // SEC REST API requires a custom User-Agent format: SampleAppName AdminContact@domain.com
    const SEC_USER_AGENT = 'SECcelTool/2.2 admin@example.com';

    // --- Styling Injection ---
    const style = document.createElement('style');
    style.textContent = `
        #sec-excel-fab {
            position: fixed;
            bottom: 24px;
            right: 24px;
            z-index: 9999999;
            background: #107c41;
            color: #ffffff;
            border: 1px solid #0d5e31;
            border-radius: 6px;
            padding: 10px 16px;
            font-family: -apple-system, BlinkMacSystemFont, "Segoe UI", Roboto, sans-serif;
            font-size: 13px;
            font-weight: 600;
            box-shadow: 0 4px 12px rgba(0, 0, 0, 0.25);
            cursor: pointer;
            display: flex;
            align-items: center;
            gap: 8px;
            transition: all 0.2s ease;
        }
        #sec-excel-fab:hover {
            background: #0b5f31;
            box-shadow: 0 6px 16px rgba(0, 0, 0, 0.35);
        }
        #sec-excel-fab:disabled {
            background: #6c757d;
            border-color: #5a6268;
            cursor: not-allowed;
            opacity: 0.85;
        }
        .sec-inline-excel-btn {
            background: #107c41;
            color: #fff !important;
            font-size: 11px;
            font-weight: 600;
            padding: 2px 6px;
            border-radius: 4px;
            margin-left: 6px;
            text-decoration: none !important;
            cursor: pointer;
            border: none;
            display: inline-block;
            vertical-align: middle;
        }
        .sec-inline-excel-btn:hover {
            background: #0b5f31;
        }
        #sec-error-toast {
            position: fixed;
            top: 24px;
            right: 24px;
            z-index: 10000000;
            max-width: 440px;
            background: #b71c1c;
            color: #fff;
            padding: 14px 18px;
            border-radius: 6px;
            font-family: -apple-system, BlinkMacSystemFont, "Segoe UI", Roboto, sans-serif;
            font-size: 12px;
            line-height: 1.5;
            box-shadow: 0 4px 14px rgba(0, 0, 0, 0.3);
            white-space: pre-line;
            transition: opacity 0.3s ease;
        }
        #seccel-modal-backdrop {
            position: fixed;
            top: 0;
            left: 0;
            width: 100vw;
            height: 100vh;
            background: rgba(0, 0, 0, 0.65);
            z-index: 10000001;
            display: flex;
            align-items: center;
            justify-content: center;
            font-family: -apple-system, BlinkMacSystemFont, "Segoe UI", Roboto, sans-serif;
        }
        #seccel-modal-card {
            background: #ffffff;
            color: #212529;
            width: 90%;
            max-width: 480px;
            border-radius: 8px;
            box-shadow: 0 10px 30px rgba(0, 0, 0, 0.4);
            padding: 24px;
            box-sizing: border-box;
        }
        #seccel-modal-card h3 {
            margin: 0 0 10px 0;
            font-size: 16px;
            color: #107c41;
            display: flex;
            align-items: center;
            gap: 8px;
        }
        #seccel-modal-card p {
            margin: 0 0 16px 0;
            font-size: 13px;
            color: #495057;
            line-height: 1.45;
        }
        .seccel-opt-btn {
            width: 100%;
            text-align: left;
            padding: 12px 14px;
            margin-bottom: 10px;
            border-radius: 6px;
            border: 1px solid #ced4da;
            background: #f8f9fa;
            cursor: pointer;
            transition: all 0.15s ease;
        }
        .seccel-opt-btn:hover {
            background: #e9ecef;
            border-color: #adb5bd;
        }
        .seccel-opt-title {
            font-weight: 600;
            font-size: 13px;
            color: #212529;
            margin-bottom: 2px;
        }
        .seccel-opt-desc {
            font-size: 11px;
            color: #6c757d;
        }
        #seccel-modal-close {
            margin-top: 8px;
            width: 100%;
            background: none;
            border: none;
            color: #6c757d;
            font-size: 12px;
            font-weight: 600;
            cursor: pointer;
            padding: 6px;
            text-align: center;
        }
        #seccel-modal-close:hover {
            color: #212529;
            text-decoration: underline;
        }
    `;
    document.head.appendChild(style);

    // --- Toast Notification ---
    function showError(msg) {
        const existing = document.getElementById('sec-error-toast');
        if (existing) existing.remove();

        const toast = document.createElement('div');
        toast.id = 'sec-error-toast';
        toast.textContent = `⚠️ EDGAR Downloader:\n${msg}`;
        document.body.appendChild(toast);

        setTimeout(() => {
            toast.style.opacity = '0';
            setTimeout(() => toast.remove(), 300);
        }, 8000);
    }

    // --- Blob Saver ---
    function saveBlob(blob, filename) {
        const blobUrl = URL.createObjectURL(blob);
        const a = document.createElement('a');
        a.href = blobUrl;
        a.download = filename;
        document.body.appendChild(a);
        a.click();
        a.remove();
        setTimeout(() => URL.revokeObjectURL(blobUrl), 15000);
    }

    // --- Date Formatting Helper ---
    function formatPeriodDate(isoDateStr) {
        if (!isoDateStr) return 'Current';
        const parts = isoDateStr.split('-');
        if (parts.length !== 3) return isoDateStr;
        const months = ['Jan.', 'Feb.', 'Mar.', 'Apr.', 'May', 'Jun.', 'Jul.', 'Aug.', 'Sep.', 'Oct.', 'Nov.', 'Dec.'];
        const m = parseInt(parts[1], 10) - 1;
        const d = parseInt(parts[2], 10);
        const y = parts[0];
        return `${months[m] || parts[1]} ${d}, ${y}`;
    }

    // --- Structured Statement Builder with Currency Formatting ---
    function buildStatementWorksheet(statementTitle, rowsMap, periods) {
        const aoa = [];

        // Row 1: Statement Title
        aoa.push([`${statementTitle} - USD ($)`]);

        // Row 2: Period Headers (Formatted dates)
        const headerRow = ['Line Items', ...periods.map(formatPeriodDate)];
        aoa.push(headerRow);

        // Data Rows: Line item description + formatted numeric values
        for (const conceptKey in rowsMap) {
            const item = rowsMap[conceptKey];
            const row = [item.label];
            let hasValue = false;

            periods.forEach((p) => {
                const val = item.values[p];
                const unit = item.units ? item.units[p] : null;

                if (val !== undefined && val !== null && typeof val === 'number') {
                    // Standard accounting currency: $1,234 ; ($1,234) ; "-"
                    let numFmt = '$#,##0;($#,##0);"-"';

                    if (unit === 'USD') {
                        // Per-share line items need 2 decimal places
                        if (/PerShare|per share/i.test(conceptKey) || /per share/i.test(item.label)) {
                            numFmt = '$#,##0.00;($#,##0.00);"-"';
                        } else {
                            numFmt = '$#,##0;($#,##0);"-"';
                        }
                    } else if (unit === 'USD/shares') {
                        numFmt = '$#,##0.00;($#,##0.00);"-"';
                    } else if (unit === 'shares') {
                        numFmt = '#,##0;(#,##0);"-"';
                    } else if (unit === 'pure') {
                        numFmt = '0.00%';
                    } else {
                        // Default fallback: if in a financial statement and not shares/ratios
                        if (/share|ratio|percent|count/i.test(item.label) || /share|ratio|percent|count/i.test(conceptKey)) {
                            numFmt = '#,##0;(#,##0);"-"';
                        } else {
                            numFmt = '$#,##0;($#,##0);"-"';
                        }
                    }

                    // Push as a typed SheetJS cell object with explicit format
                    row.push({ t: 'n', v: val, z: numFmt });
                    hasValue = true;
                } else if (val !== undefined && val !== null) {
                    row.push(val);
                    hasValue = true;
                } else {
                    row.push('');
                }
            });

            if (hasValue) {
                aoa.push(row);
            }
        }

        const ws = XLSX.utils.aoa_to_sheet(aoa);

        // Set column widths (Wide label column + readable data columns)
        ws['!cols'] = [
            { wch: 55 },
            ...periods.map(() => ({ wch: 18 }))
        ];

        return ws;
    }

    // --- Fallback: SEC Company Facts API Parser ---
    function fetchSecFactsApi(cik, accession, btnElement) {
        if (btnElement) btnElement.innerHTML = '🌐 Querying SEC API...';

        const paddedCik = cik.padStart(10, '0');
        const cleanAcc = accession.replace(/-/g, '');
        const apiUrl = `https://data.sec.gov/api/xbrl/companyfacts/CIK${paddedCik}.json`;

        GM_xmlhttpRequest({
            method: 'GET',
            url: apiUrl,
            headers: {
                'User-Agent': SEC_USER_AGENT,
                Accept: 'application/json'
            },
            onload: (res) => {
                if (res.status !== 200) {
                    showError(`SEC Facts API returned status ${res.status}. Data is unavailable or throttled.`);
                    if (btnElement) btnElement.disabled = false;
                    return;
                }

                try {
                    const data = JSON.parse(res.responseText);
                    const taxonomies = data.facts || {};

                    const periodsSet = new Set();

                    const incomeRows = {};
                    const balanceRows = {};
                    const cashRows = {};
                    const entityRows = {};
                    const allRows = {};

                    const incomePatterns = /Revenue|Sales|CostOf|GrossProfit|OperatingExpenses|OperatingIncome|OperatingLoss|InterestExpense|IncomeTaxExpense|NetIncome|EarningsPerShare/i;
                    const balancePatterns = /Assets|Liabilities|CashAndCashEquivalents|Inventory|Receivables|PropertyPlantAndEquipment|Goodwill|Debt|StockholdersEquity|RetainedEarnings/i;
                    const cashPatterns = /CashProvidedByUsedIn|CashFlow|OperatingActivities|InvestingActivities|FinancingActivities|DepreciationDepletionAndAmortization/i;

                    for (const taxKey in taxonomies) {
                        for (const conceptKey in taxonomies[taxKey]) {
                            const concept = taxonomies[taxKey][conceptKey];
                            const label = concept.label || conceptKey;
                            const units = concept.units || {};

                            for (const unitKey in units) {
                                units[unitKey].forEach((entry) => {
                                    const entryAcc = (entry.accn || '').replace(/-/g, '');
                                    if (entryAcc === cleanAcc) {
                                        const period = entry.end || 'Current';
                                        periodsSet.add(period);

                                        // Register container with both values and unit definitions
                                        const registerEntry = (targetMap) => {
                                            if (!targetMap[conceptKey]) {
                                                targetMap[conceptKey] = { label, values: {}, units: {} };
                                            }
                                            targetMap[conceptKey].values[period] = entry.val;
                                            targetMap[conceptKey].units[period] = unitKey;
                                        };

                                        registerEntry(allRows);

                                        if (taxKey === 'dei') {
                                            registerEntry(entityRows);
                                        } else if (cashPatterns.test(conceptKey)) {
                                            registerEntry(cashRows);
                                        } else if (incomePatterns.test(conceptKey)) {
                                            registerEntry(incomeRows);
                                        } else if (balancePatterns.test(conceptKey)) {
                                            registerEntry(balanceRows);
                                        }
                                    }
                                });
                            }
                        }
                    }

                    // Sort periods descending (most recent date first)
                    const sortedPeriods = Array.from(periodsSet).sort().reverse();

                    if (sortedPeriods.length === 0) {
                        showError(`No XBRL facts were found matching filing accession: ${accession}`);
                        if (btnElement) btnElement.disabled = false;
                        return;
                    }

                    const wb = XLSX.utils.book_new();

                    // 1. Income Statement
                    if (Object.keys(incomeRows).length > 0) {
                        const wsIncome = buildStatementWorksheet('Consolidated Statements Of Income', incomeRows, sortedPeriods);
                        XLSX.utils.book_append_sheet(wb, wsIncome, 'Statements Of Income');
                    }

                    // 2. Balance Sheets
                    if (Object.keys(balanceRows).length > 0) {
                        const wsBalance = buildStatementWorksheet('Consolidated Balance Sheets', balanceRows, sortedPeriods);
                        XLSX.utils.book_append_sheet(wb, wsBalance, 'Balance Sheets');
                    }

                    // 3. Cash Flows
                    if (Object.keys(cashRows).length > 0) {
                        const wsCash = buildStatementWorksheet('Consolidated Statements Of Cash Flows', cashRows, sortedPeriods);
                        XLSX.utils.book_append_sheet(wb, wsCash, 'Statements Of Cash Flows');
                    }

                    // 4. Entity & Document Info
                    if (Object.keys(entityRows).length > 0) {
                        const wsEntity = buildStatementWorksheet('Document and Entity Information', entityRows, sortedPeriods);
                        XLSX.utils.book_append_sheet(wb, wsEntity, 'Entity Info');
                    }

                    // 5. Unfiltered Facts Summary
                    const wsAll = buildStatementWorksheet('All Tagged Filing Facts', allRows, sortedPeriods);
                    XLSX.utils.book_append_sheet(wb, wsAll, 'All Facts');

                    const wbOut = XLSX.write(wb, { bookType: 'xlsx', type: 'array', cellStyles: true });
                    const blob = new Blob([wbOut], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' });
                    saveBlob(blob, `Financial_Report_${cik}_${cleanAcc}_SEC_API.xlsx`);

                    if (btnElement) {
                        btnElement.innerHTML = '✅ API Report Saved!';
                        setTimeout(() => (btnElement.disabled = false), 2500);
                    }
                } catch (e) {
                    showError(`Failed parsing SEC API facts: ${e.message}`);
                    if (btnElement) btnElement.disabled = false;
                }
            },
            onerror: () => {
                showError('Network error connecting to data.sec.gov API.');
                if (btnElement) btnElement.disabled = false;
            }
        });
    }

    // --- Modal Handler ---
    function showSourceSelectionModal(cik, accession, btnElement) {
        const existing = document.getElementById('seccel-modal-backdrop');
        if (existing) existing.remove();

        const backdrop = document.createElement('div');
        backdrop.id = 'seccel-modal-backdrop';

        backdrop.innerHTML = `
            <div id="seccel-modal-card">
                <h3>📊 Excel Report Fallback</h3>
                <p>
                    <strong>Financial_Report.xlsx</strong> was not compiled in the root archive for this filing (HTTP 404).<br><br>
                    You can generate a multi-tab financial report directly from the official <strong>SEC Company Facts REST API</strong>:
                </p>

                <button class="seccel-opt-btn" id="seccel-btn-api">
                    <div class="seccel-opt-title">🌐 Generate from SEC Facts API (.xlsx)</div>
                    <div class="seccel-opt-desc">Extracts, pivots, and formats Income Statements, Balance Sheets, and Cash Flows with native currency formatting.</div>
                </button>

                <button id="seccel-modal-close">Dismiss</button>
            </div>
        `;

        document.body.appendChild(backdrop);

        const closeModal = () => {
            backdrop.remove();
            if (btnElement) btnElement.disabled = false;
        };

        backdrop.querySelector('#seccel-btn-api').addEventListener('click', () => {
            closeModal();
            fetchSecFactsApi(cik, accession, btnElement);
        });

        backdrop.querySelector('#seccel-modal-close').addEventListener('click', closeModal);
    }

    // --- Request & Download Handler ---
    function executeDownload(cik, accession, btnElement) {
        const cleanCik = cik.replace(/^0+/, '') || '0';
        const cleanAcc = accession.replace(/-/g, '');
        const targetUrl = `https://www.sec.gov/Archives/edgar/data/${cleanCik}/${cleanAcc}/Financial_Report.xlsx`;

        const originalText = btnElement ? btnElement.innerHTML : '';
        if (btnElement) {
            btnElement.disabled = true;
            btnElement.innerHTML = '⏳ Verifying...';
        }

        GM_xmlhttpRequest({
            method: 'GET',
            url: targetUrl,
            responseType: 'blob',
            onload: (res) => {
                if (res.status === 200 && res.response) {
                    if (btnElement) btnElement.innerHTML = '📥 Downloading...';
                    saveBlob(res.response, `Financial_Report_${cleanCik}_${cleanAcc}.xlsx`);
                    if (btnElement) {
                        btnElement.innerHTML = '✅ Downloaded!';
                        setTimeout(() => {
                            btnElement.innerHTML = originalText;
                            btnElement.disabled = false;
                        }, 2500);
                    }
                } else if (res.status === 404) {
                    if (btnElement) btnElement.innerHTML = '🔍 Fallback available';
                    showSourceSelectionModal(cleanCik, accession, btnElement);
                } else {
                    if (btnElement) {
                        btnElement.innerHTML = originalText;
                        btnElement.disabled = false;
                    }
                    showError(`Request failed with status ${res.status} (${res.statusText}).`);
                }
            },
            onerror: () => {
                if (btnElement) {
                    btnElement.innerHTML = originalText;
                    btnElement.disabled = false;
                }
                showError(`Network failure while requesting: ${targetUrl}`);
            }
        });
    }

    // --- Parser for Single-Filing Viewers ---
    function getActiveFilingInfo() {
        const fullUrl = decodeURIComponent(window.location.href);

        const match = fullUrl.match(/Archives\/edgar\/data\/(\d+)\/(\d{10}-?\d{2}-?\d{6}|\d{18})/i);
        if (match) {
            return { cik: match[1], accession: match[2] };
        }

        const urlParams = new URLSearchParams(window.location.search);
        const docParam = urlParams.get('doc');
        if (docParam) {
            const docMatch = decodeURIComponent(docParam).match(/Archives\/edgar\/data\/(\d+)\/(\d{10}-?\d{2}-?\d{6}|\d{18})/i);
            if (docMatch) {
                return { cik: docMatch[1], accession: docMatch[2] };
            }
        }

        return null;
    }

    // --- UI Renderers ---
    function renderFloatingButton() {
        const filing = getActiveFilingInfo();
        let btn = document.getElementById('sec-excel-fab');

        if (!filing) {
            if (btn) btn.remove();
            return;
        }

        if (!btn) {
            btn = document.createElement('button');
            btn.id = 'sec-excel-fab';
            btn.innerHTML = '📊 Download Financial Report';
            btn.title = 'Download official Financial_Report.xlsx or generate via SEC Facts API';
            btn.addEventListener('click', () => {
                const currentFiling = getActiveFilingInfo();
                if (currentFiling) {
                    executeDownload(currentFiling.cik, currentFiling.accession, btn);
                }
            });
            document.body.appendChild(btn);
        }
    }

    function renderInlineTableButtons() {
        const targetForms = ['10-K', '10-Q', '10-K/A', '10-Q/A', '20-F', '6-K'];
        const links = document.querySelectorAll('a[href*="/Archives/edgar/data/"]:not([data-sec-excel-processed])');

        links.forEach((a) => {
            const rawHref = decodeURIComponent(a.href);
            const match = rawHref.match(/Archives\/edgar\/data\/(\d+)\/(\d{10}-?\d{2}-?\d{6}|\d{18})/i);

            if (match) {
                a.setAttribute('data-sec-excel-processed', 'true');

                const row = a.closest('tr');
                if (row) {
                    const rowText = row.innerText;
                    const isFinancialFiling = targetForms.some((form) => {
                        const pattern = new RegExp(`\\b${form.replace('/', '\\/')}\\b`, 'i');
                        return pattern.test(rowText);
                    });

                    if (!isFinancialFiling) return;
                }

                const cik = match[1];
                const accession = match[2];

                const inlineBtn = document.createElement('button');
                inlineBtn.className = 'sec-inline-excel-btn';
                inlineBtn.textContent = 'XLS';
                inlineBtn.title = `Download Financial Report for ${accession}`;

                inlineBtn.addEventListener('click', (e) => {
                    e.preventDefault();
                    e.stopPropagation();
                    executeDownload(cik, accession, inlineBtn);
                });

                a.parentNode.insertBefore(inlineBtn, a.nextSibling);
            }
        });
    }

    // --- Lifecycle & DOM Observers ---
    function run() {
        renderFloatingButton();
        renderInlineTableButtons();
    }

    run();

    const observer = new MutationObserver(() => run());
    observer.observe(document.body, { childList: true, subtree: true });

    window.addEventListener('popstate', run);
})();
