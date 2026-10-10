// ==UserScript==
// @name         SEC EDGAR Direct Financial Report (.xlsx) Downloader
// @namespace    https://github.com/
// @version      1.0
// @description  Restores direct Financial_Report.xlsx downloads on SEC EDGAR filing viewers and search tables.
// @author       Cal
// @match        https://www.sec.gov/*
// @match        https://sec.gov/*
// @grant        GM_xmlhttpRequest
// @connect      www.sec.gov
// @connect      sec.gov
// @run-at       document-idle
// ==/UserScript==

(function () {
    'use strict';

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
    `;
    document.head.appendChild(style);

    // --- Toast Notification ---
    function showError(msg) {
        const existing = document.getElementById('sec-error-toast');
        if (existing) existing.remove();

        const toast = document.createElement('div');
        toast.id = 'sec-error-toast';
        toast.textContent = `⚠️ EDGAR Download Error:\n${msg}`;
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

        const handleResponse = (status, statusText, blobData) => {
            if (status === 200 && blobData) {
                if (btnElement) btnElement.innerHTML = '📥 Downloading...';
                saveBlob(blobData, `Financial_Report_${cleanCik}_${cleanAcc}.xlsx`);
                if (btnElement) {
                    btnElement.innerHTML = '✅ Downloaded!';
                    setTimeout(() => {
                        btnElement.innerHTML = originalText;
                        btnElement.disabled = false;
                    }, 2500);
                }
            } else if (status === 404) {
                if (btnElement) btnElement.innerHTML = '❌ Not Found';
                showError(
                    `Financial_Report.xlsx does not exist for this filing (HTTP 404).\n\nTried:\n${targetUrl}\n\nNote: Filings without XBRL data (or non-periodic forms) do not generate this report.`
                );
                setTimeout(() => {
                    if (btnElement) {
                        btnElement.innerHTML = originalText;
                        btnElement.disabled = false;
                    }
                }, 3500);
            } else {
                if (btnElement) btnElement.innerHTML = '❌ HTTP Error';
                showError(`Failed with status ${status} (${statusText}).\nURL: ${targetUrl}`);
                setTimeout(() => {
                    if (btnElement) {
                        btnElement.innerHTML = originalText;
                        btnElement.disabled = false;
                    }
                }, 3500);
            }
        };

        if (typeof GM_xmlhttpRequest !== 'undefined') {
            GM_xmlhttpRequest({
                method: 'GET',
                url: targetUrl,
                responseType: 'blob',
                onload: (res) => handleResponse(res.status, res.statusText, res.response),
                onerror: () => {
                    if (btnElement) {
                        btnElement.innerHTML = '❌ Failed';
                        btnElement.disabled = false;
                    }
                    showError(`Network failure while requesting: ${targetUrl}`);
                }
            });
        } else {
            fetch(targetUrl)
                .then(async (res) => {
                    if (res.ok) {
                        const blob = await res.blob();
                        handleResponse(res.status, res.statusText, blob);
                    } else {
                        handleResponse(res.status, res.statusText, null);
                    }
                })
                .catch((err) => {
                    if (btnElement) {
                        btnElement.innerHTML = '❌ Failed';
                        btnElement.disabled = false;
                    }
                    showError(`Fetch error: ${err.message}\nURL: ${targetUrl}`);
                });
        }
    }

    // --- Parser for Single-Filing Viewers ---
    function getActiveFilingInfo() {
        const fullUrl = decodeURIComponent(window.location.href);

        // Pattern matching: /Archives/edgar/data/<CIK>/<Accession>/
        const match = fullUrl.match(/Archives\/edgar\/data\/(\d+)\/(\d{10}-?\d{2}-?\d{6}|\d{18})/i);
        if (match) {
            return { cik: match[1], accession: match[2] };
        }

        // Fallback: check "doc" query param inside iXBRL viewer
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
            btn.innerHTML = '📊 Download Financial_Report.xlsx';
            btn.title = 'Download raw multi-tab XBRL Excel report';
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
        // Scans listing tables on company search / browse pages
        const links = document.querySelectorAll('a[href*="/Archives/edgar/data/"]:not([data-sec-excel-processed])');

        links.forEach((a) => {
            const rawHref = decodeURIComponent(a.href);
            const match = rawHref.match(/Archives\/edgar\/data\/(\d+)\/(\d{10}-?\d{2}-?\d{6}|\d{18})/i);

            if (match) {
                a.setAttribute('data-sec-excel-processed', 'true');
                const cik = match[1];
                const accession = match[2];

                const inlineBtn = document.createElement('button');
                inlineBtn.className = 'sec-inline-excel-btn';
                inlineBtn.textContent = 'XLS';
                inlineBtn.title = `Download Financial_Report.xlsx for ${accession}`;

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

    // Observe SPA routing and dynamic table rendering
    const observer = new MutationObserver(() => run());
    observer.observe(document.body, { childList: true, subtree: true });

    // Handle history push/pop for EDGAR single-page transitions
    window.addEventListener('popstate', run);
})();
