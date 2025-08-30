// ==UserScript==
// @name         احصائيات الردود - ستار تايمز وكووورة (نسخة متطورة)
// @namespace    http://tampermonkey.net/
// @version      1.0
// @description  يجلب عدد الردود لكل عضوية من عدة مواضيع وتجميعها في Excel. دعم ستايل أسود، اختصارات الإدخال، وأيقونة عائمة.
// @match        https://www.startimes.com/*
// @match        https://forum.kooora.com/*
// @grant        GM_xmlhttpRequest
// @grant        GM_download
// @connect      https://www.startimes.com
// @connect      https://forum.kooora.com
// @require      https://cdnjs.cloudflare.com/ajax/libs/xlsx/0.18.5/xlsx.full.min.js
// ==/UserScript==

(function () {
    'use strict';

    const savedIDs = localStorage.getItem("replyTool_ids") || "";
    const savedVisible = localStorage.getItem("replyTool_visible") !== "false";
    const savedCombine = localStorage.getItem("replyTool_combine") !== "false";

    const box = document.createElement('div');
    box.id = "replyToolBox";
    box.style = `
        position:fixed;
        top:60px;
        left:20px;
        background:#111;
        color:#fff;
        padding:10px;
        border:2px solid #333;
        border-radius:8px;
        z-index:9999;
        width:450px;
        font-family:Tajawal,Arial;
        font-size:14px;
        box-shadow: 0 0 10px #000;
        transition: all 0.3s ease;
        overflow: hidden;
    `;
    box.innerHTML = `
        <button id="toggleBtn" title="إظهار/إخفاء" style="
            position:absolute;
            top:5px;
            left:5px;
            background:#222;
            color:#fff;
            border-radius:50%;
            width:24px;
            height:24px;
            border:none;
            cursor:pointer;
            font-weight:bold;
            font-size:16px;
            z-index:10;
        ">${savedVisible ? "–" : "+"}</button>

        <div id="logoSK" style="
            display:none;
            width:100%;
            height:100%;
            display:flex;
            align-items:center;
            justify-content:center;
            font-size:14px;
            font-weight:bold;
            font-style:italic;
            color:gold;
            font-family:'Georgia',serif;
            pointer-events:none;
            transform: rotate(-30deg);
        ">S&amp;K</div>

        <div id="toolContent" style="display:${savedVisible ? "block" : "none"};">
            <div id="dragHandle" style="text-align:right;cursor:move;color:#999;font-size:18px;">⠿</div>
            <h3 style="margin-top:0;">📊 احصائيات الردود</h3>
            <textarea id="idList" rows="7" style="width:100%;background:#222;color:#fff;border:1px solid #555;border-radius:4px;">${savedIDs}</textarea>
            <label style="display:flex;align-items:center;margin-top:5px;">
                <input type="checkbox" id="combineToggle" ${savedCombine ? "checked" : ""} style="margin-left:5px;" />
                🧮 جمع الإحصائيات
            </label>
            <div id="loadingIndicator" style="display:none;margin-top:5px;color:gold;font-weight:bold;text-align:center;">⏳ جاري المعالجة...</div>
            <button id="calcBtn" style="margin-top:5px;background:#28a745;color:#fff;padding:5px;width:100%;border:none;border-radius:4px;">احسب</button>
            <button id="downloadBtn" style="display:none;margin-top:5px;background:#007bff;color:#fff;padding:5px;width:100%;border:none;border-radius:4px;">📥 تحميل</button>
            <button id="clearBtn" style="margin-top:5px;background:#dc3545;color:#fff;padding:5px;width:100%;border:none;border-radius:4px;">🗑️ مسح الإدخالات</button>
        </div>
    `;
    document.body.appendChild(box);

    const toggleBtn = document.getElementById("toggleBtn");
    const toolContent = document.getElementById("toolContent");
    const logoSK = document.getElementById("logoSK");

    function updateVisibility(state) {
        if (state) {
            toolContent.style.display = "block";
            logoSK.style.display = "none";
            box.style.width = "450px";
            box.style.height = "auto";
            box.style.borderRadius = "8px";
            box.style.padding = "10px";
            toggleBtn.innerText = "–";
        } else {
            toolContent.style.display = "none";
            logoSK.style.display = "flex";
            box.style.width = "48px";
            box.style.height = "48px";
            box.style.borderRadius = "50%";
            box.style.padding = "0";
            toggleBtn.innerText = "+";
        }
        localStorage.setItem("replyTool_visible", state);
    }

    toggleBtn.onclick = () => {
        const isVisible = toolContent.style.display !== "none";
        updateVisibility(!isVisible);
    };

    const dragHandle = document.getElementById("dragHandle");
    let isDragging = false, offsetX = 0, offsetY = 0;

    function startDrag(e) {
        isDragging = true;
        offsetX = e.clientX - box.offsetLeft;
        offsetY = e.clientY - box.offsetTop;
        document.body.style.userSelect = "none";
        document.addEventListener('mousemove', drag);
        document.addEventListener('mouseup', stopDrag);
    }

    function drag(e) {
        if (!isDragging) return;
        box.style.left = (e.clientX - offsetX) + 'px';
        box.style.top = (e.clientY - offsetY) + 'px';
    }

    function stopDrag() {
        isDragging = false;
        document.body.style.userSelect = "";
        document.removeEventListener('mousemove', drag);
        document.removeEventListener('mouseup', stopDrag);
    }

    dragHandle.addEventListener('mousedown', startDrag);
    box.addEventListener('mousedown', (e) => {
        if (toolContent.style.display === 'none') startDrag(e);
    });

    const idInput = document.getElementById('idList');
    idInput.addEventListener("input", () => localStorage.setItem("replyTool_ids", idInput.value));

    const combineToggle = document.getElementById("combineToggle");
    combineToggle.onchange = () => localStorage.setItem("replyTool_combine", combineToggle.checked);

    updateVisibility(savedVisible);

    const clearBtn = document.getElementById('clearBtn');
    clearBtn.onclick = () => {
        localStorage.removeItem("replyTool_ids");
        idInput.value = "";
    };

    const loadingIndicator = document.getElementById("loadingIndicator");
    const calcBtn = document.getElementById("calcBtn");
    const downloadBtn = document.getElementById("downloadBtn");

    function extractTopicId(line) {
        const match = line.match(/t=(\d+)/i);
        if (match) return match[1];
        if (/^\d+$/.test(line)) return line;
        return null;
    }

    function fetchPage(url) {
        return new Promise((resolve, reject) => {
            GM_xmlhttpRequest({
                method: "GET",
                url,
                onload: res => resolve(res.responseText),
                onerror: err => reject(err)
            });
        });
    }

    calcBtn.onclick = async () => {
        loadingIndicator.style.display = "block";
        downloadBtn.style.display = "none";

        const rawLines = idInput.value.trim().split('\n').filter(Boolean);
        const ids = rawLines.map(extractTopicId).filter(Boolean);

        const base = location.hostname.includes("kooora") ? "https://forum.kooora.com" : "https://www.startimes.com";
        const combine = combineToggle.checked;
        const workbook = XLSX.utils.book_new();

        if (!combine) {
            for (const id of ids) {
                loadingIndicator.innerText = `⏳ جاري المعالجة... (الموضوع: ${id})`;
                const url = `${base}/f.aspx?svc=tstats&tstat=${id}&tstatl=n`;
                const html = await fetchPage(url);
                const doc = new DOMParser().parseFromString(html, 'text/html');
                const tds = Array.from(doc.querySelectorAll('td.stats_p'));
                const topicStats = {};
                for (let i = 0; i < tds.length; i += 2) {
                    const nameEl = tds[i].querySelector('font');
                    const linkEl = tds[i].querySelector('a');
                    const countEl = tds[i + 1]?.querySelector('a');
                    if (nameEl && countEl && linkEl) {
                        const name = nameEl.textContent.trim();
                        const url = linkEl.href;
                        const count = parseInt(countEl.textContent.trim());
                        if (!isNaN(count)) {
                            topicStats[name] = topicStats[name] || { count: 0, url };
                            topicStats[name].count += count;
                        }
                    }
                }
                const rows = Object.entries(topicStats)
                    .map(([name, data]) => ({ العضو: { f: `HYPERLINK("${data.url}", "${name}")` }, الردود: data.count }))
                    .sort((a, b) => b.الردود - a.الردود);
                const sheet = XLSX.utils.json_to_sheet(rows);
                XLSX.utils.book_append_sheet(workbook, sheet, `موضوع ${id}`);
            }
        } else {
            const allCombined = {};
            const statsPerMember = {};
            for (const id of ids) {
                loadingIndicator.innerText = `⏳ جاري المعالجة... (الموضوع: ${id})`;
                const url = `${base}/f.aspx?svc=tstats&tstat=${id}&tstatl=n`;
                const html = await fetchPage(url);
                const doc = new DOMParser().parseFromString(html, 'text/html');
                const tds = Array.from(doc.querySelectorAll('td.stats_p'));
                for (let i = 0; i < tds.length; i += 2) {
                    const nameEl = tds[i].querySelector('font');
                    const linkEl = tds[i].querySelector('a');
                    const countEl = tds[i + 1]?.querySelector('a');
                    if (nameEl && countEl && linkEl) {
                        const name = nameEl.textContent.trim();
                        const url = linkEl.href;
                        const count = parseInt(countEl.textContent.trim());
                        if (!isNaN(count)) {
                            if (!allCombined[name]) allCombined[name] = { count: 0, url };
                            allCombined[name].count += count;

                            if (!statsPerMember[name]) statsPerMember[name] = { topics: 0, total: 0, url };
                            statsPerMember[name].topics += 1;
                            statsPerMember[name].total += count;
                        }
                    }
                }
            }

            const sorted = Object.entries(statsPerMember)
            .map(([name, data]) => ({
                العضو: { f: `HYPERLINK("${data.url}", "${name}")` },
                الردود: data.total,
                "عدد المواضيع": data.topics,
                "المتوسط": +(data.total / data.topics).toFixed(2),
                "المتوسط على كل المواضيع": +(data.total / ids.length).toFixed(2)
            }))
                .sort((a, b) => b.الردود - a.الردود);

            const sheet = XLSX.utils.json_to_sheet(sorted);
            XLSX.utils.book_append_sheet(workbook, sheet, 'كل المواضيع');
        }

        const wbout = XLSX.write(workbook, { bookType: 'xlsx', type: 'binary' });
        const buf = new ArrayBuffer(wbout.length);
        const view = new Uint8Array(buf);
        for (let i = 0; i < wbout.length; i++) view[i] = wbout.charCodeAt(i) & 0xff;

        const blob = new Blob([buf], { type: "application/octet-stream" });
        const url = URL.createObjectURL(blob);

        loadingIndicator.style.display = "none";
        downloadBtn.style.display = "inline-block";
        downloadBtn.onclick = () => {
            const a = document.createElement("a");
            a.href = url;
            a.download = "احصائيات_الردود.xlsx";
            a.click();
        };
    };
})();
