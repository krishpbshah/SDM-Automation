/**
 * SDM_LIST.js (Batch Edition)
 *
 * Batch Processor for CA SDM Automation.
 *
 * Features:
 * - CSV Upload & Parsing
 * - Sequential Ticket Processing
 * - User Verification Step per Ticket
 * - Robust Automation Logic (cai_main prioritized)
 */

const express = require('express');
const http = require('http');
const { Server } = require('socket.io');
const { chromium } = require('playwright');
const path = require('path');
const fs = require('fs');
const open = require('open');

// ---- Configuration ----
const PORT = 3000;
const BASE_URL = 'http://servicedesk-web.int.ttc.ca/CAisd/pdmweb.exe';

// ---- Global State ----
let browser = null;
let context = null;
let page = null;
let userAssignee = "Couto, Lucas";

// Batch State
let ticketQueue = [];
let currentTicketIndex = 0;
let isProcessingBatch = false;

// ---- Express & Socket.io Setup ----
const app = express();
const server = http.createServer(app);
const io = new Server(server, { maxHttpBufferSize: 1e8 }); // Allow large uploads

// Serve the Single Page App
app.get('/', (req, res) => {
    res.send(`
<!DOCTYPE html>
<html lang="en">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>SDM Batch Automation</title>
    <style>
        :root { --bg: #0f172a; --card: #1e293b; --text: #f8fafc; --accent: #8b5cf6; --success: #22c55e; --danger: #ef4444; }
        body { font-family: 'Segoe UI', system-ui, sans-serif; background: var(--bg); color: var(--text); margin: 0; padding: 20px; display: flex; justify-content: center; height: 100vh; box-sizing: border-box; }
        .container { display: grid; grid-template-columns: 1fr 1fr; gap: 20px; width: 100%; max-width: 1200px; height: 100%; }
        .panel { background: var(--card); padding: 20px; border-radius: 12px; display: flex; flex-direction: column; gap: 15px; box-shadow: 0 4px 6px -1px rgba(0,0,0,0.5); overflow-y: auto; }
        h2 { margin: 0 0 10px 0; border-bottom: 1px solid #334155; padding-bottom: 10px; color: var(--accent); }
        .status-bar { background: #334155; padding: 10px; border-radius: 6px; text-align: center; font-weight: bold; margin-bottom: 10px; }
        label { font-size: 0.9rem; color: #94a3b8; margin-bottom: 4px; display: block; }
        input, select { width: 100%; padding: 10px; background: #0f172a; border: 1px solid #334155; color: white; border-radius: 6px; margin-bottom: 10px; box-sizing: border-box; }
        button { background: var(--accent); color: white; border: none; padding: 12px; border-radius: 6px; cursor: pointer; font-weight: bold; transition: opacity 0.2s; }
        button:hover { opacity: 0.9; }
        button:disabled { background: #475569; cursor: not-allowed; }
        .row { display: flex; gap: 10px; }
        .log-window { background: #000; font-family: 'Consolas', monospace; padding: 15px; border-radius: 6px; flex-grow: 1; overflow-y: auto; font-size: 0.85rem; white-space: pre-wrap; border: 1px solid #333; }
        .log-entry { margin-bottom: 4px; }
        .log-info { color: #cbd5e1; }
        .log-success { color: var(--success); }
        .log-error { color: var(--danger); }
        .log-warn { color: #f59e0b; }
        .task-list { display: grid; grid-template-columns: repeat(auto-fill, minmax(80px, 1fr)); gap: 10px; }
        .task-item { background: #334155; padding: 10px; border-radius: 4px; text-align: center; cursor: pointer; user-select: none; border: 2px solid transparent; }
        .task-item.selected { border-color: var(--accent); background: #4c1d95; }
        .hidden { display: none; }
        .file-drop { border: 2px dashed #475569; padding: 20px; text-align: center; border-radius: 8px; cursor: pointer; transition: border-color 0.2s; }
        .file-drop:hover { border-color: var(--accent); }
    </style>
</head>
<body>
    <div class="container">
        <div class="panel">
            <h2>Configuration</h2>
            <div class="row">
                <div style="flex:1">
                    <label>First Name</label>
                    <input type="text" id="firstName" placeholder="e.g. Lucas">
                </div>
                <div style="flex:1">
                    <label>Last Name</label>
                    <input type="text" id="lastName" placeholder="e.g. Couto">
                </div>
            </div>
            <button id="saveConfigBtn">Save Configuration</button>

            <h2>Batch Processing</h2>
            <div id="uploadSection">
                <label>Ticket Type (Applied to All)</label>
                <select id="ticketType">
                    <option value="go_cr">Change Request (go_cr)</option>
                    <option value="go_in">Incident (go_in)</option>
                    <option value="go_pr">Problem (go_pr)</option>
                </select>
                <div class="file-drop" onclick="document.getElementById('fileInput').click()">
                    Click to Upload CSV / Text File
                    <input type="file" id="fileInput" accept=".csv,.txt" style="display:none">
                </div>
                <p id="fileName" style="text-align:center; color:#94a3b8; font-size:0.8rem;"></p>
                <button id="startBatchBtn" disabled>Start Batch</button>
            </div>

            <div id="activeTicketSection" class="hidden">
                <div class="status-bar" id="batchStatus">Ticket 1 of 10</div>
                <h3>Current Ticket: <span id="currentTicketNum" style="color:var(--accent)">---</span></h3>
                
                <div id="taskSection" class="hidden">
                    <h4>Select Tasks to Complete</h4>
                    <div id="taskList" class="task-list"></div>
                    <br>
                    <button id="processBtn">Process & Next Ticket</button>
                    <button id="skipBtn" style="background:#475569; margin-top:10px;">Skip Ticket</button>
                </div>
                <div id="loadingMsg" class="hidden" style="text-align:center; color:#94a3b8;">Loading Ticket...</div>
            </div>
        </div>
        <div class="panel">
            <h2>Live Logs</h2>
            <div id="logWindow" class="log-window"></div>
        </div>
    </div>

    <script src="/socket.io/socket.io.js"></script>
    <script>
        const socket = io();
        const logWindow = document.getElementById('logWindow');
        const taskList = document.getElementById('taskList');
        let selectedTasks = new Set();
        let currentBatchSize = 0;

        // --- Logging ---
        socket.on('log', (data) => {
            const div = document.createElement('div');
            div.className = 'log-entry log-' + data.type;
            div.textContent = \`[\${new Date().toLocaleTimeString()}] \${data.msg}\`;
            logWindow.appendChild(div);
            logWindow.scrollTop = logWindow.scrollHeight;
        });

        // --- Config ---
        document.getElementById('saveConfigBtn').addEventListener('click', () => {
            const first = document.getElementById('firstName').value.trim();
            const last = document.getElementById('lastName').value.trim();
            if(!first || !last) return alert('Please enter both names');
            
            socket.emit('set_user', { first, last });
            document.getElementById('startBatchBtn').disabled = false; // Enable if file also ready
            document.getElementById('saveConfigBtn').textContent = 'Saved!';
            setTimeout(() => document.getElementById('saveConfigBtn').textContent = 'Save Configuration', 2000);
        });

        // --- File Upload ---
        document.getElementById('fileInput').addEventListener('change', (e) => {
            const file = e.target.files[0];
            if(!file) return;
            document.getElementById('fileName').textContent = file.name;
            
            const reader = new FileReader();
            reader.onload = (evt) => {
                const content = evt.target.result;
                const lines = content.split(/\\r?\\n/).map(l => l.trim()).filter(l => /^\\d+$/.test(l)); // Simple regex for numbers
                if(lines.length === 0) return alert('No valid ticket numbers found in file');
                
                socket.emit('upload_batch', lines);
                currentBatchSize = lines.length;
                document.getElementById('startBatchBtn').textContent = \`Start Batch (\${lines.length} Tickets)\`;
                document.getElementById('startBatchBtn').disabled = false;
            };
            reader.readAsText(file);
        });

        document.getElementById('startBatchBtn').addEventListener('click', () => {
            const type = document.getElementById('ticketType').value;
            document.getElementById('uploadSection').classList.add('hidden');
            document.getElementById('activeTicketSection').classList.remove('hidden');
            socket.emit('start_batch', { type });
        });

        // --- Batch Flow ---
        socket.on('ticket_ready', (data) => {
            // data = { index, total, ticketNum, tasks }
            document.getElementById('batchStatus').textContent = \`Ticket \${data.index + 1} of \${data.total}\`;
            document.getElementById('currentTicketNum').textContent = data.ticketNum;
            document.getElementById('loadingMsg').classList.add('hidden');
            document.getElementById('taskSection').classList.remove('hidden');
            
            taskList.innerHTML = '';
            selectedTasks.clear();
            
            if(data.tasks.length === 0) {
                 taskList.innerHTML = '<div style="grid-column: 1/-1; text-align:center; color:#94a3b8;">No tasks found.</div>';
            }

            data.tasks.forEach(t => {
                const el = document.createElement('div');
                el.className = 'task-item';
                el.textContent = t;
                el.onclick = () => {
                    if(selectedTasks.has(t)) {
                        selectedTasks.delete(t);
                        el.classList.remove('selected');
                    } else {
                        selectedTasks.add(t);
                        el.classList.add('selected');
                    }
                };
                taskList.appendChild(el);
            });
        });

        socket.on('loading_ticket', () => {
            document.getElementById('taskSection').classList.add('hidden');
            document.getElementById('loadingMsg').classList.remove('hidden');
        });

        socket.on('batch_complete', () => {
            document.getElementById('activeTicketSection').classList.add('hidden');
            document.getElementById('uploadSection').classList.remove('hidden');
            alert('Batch Processing Complete!');
        });

        document.getElementById('processBtn').addEventListener('click', () => {
            if(selectedTasks.size === 0) return alert('Select at least one task (or click Skip)');
            socket.emit('process_current_ticket', Array.from(selectedTasks));
        });

        document.getElementById('skipBtn').addEventListener('click', () => {
            socket.emit('skip_current_ticket');
        });

    </script>
</body>
</html>
    `);
});

// ---- Automation Logic ----

function log(type, msg) {
    const safeType = (type || 'info').toString();
    console.log(`[${safeType.toUpperCase()}] ${msg}`);
    io.emit('log', { type: safeType, msg });
}

async function sleep(ms) { return new Promise(r => setTimeout(r, ms)); }

async function waitSettled(pageLike, timeoutMs = 15000) {
    const start = Date.now();
    try { await pageLike.waitForLoadState?.('domcontentloaded', { timeout: Math.min(5000, timeoutMs) }); } catch { }
    try { await pageLike.waitForLoadState?.('load', { timeout: Math.min(5000, timeoutMs) }); } catch { }
    while (Date.now() - start < timeoutMs) {
        try { if (pageLike.evaluate) { await pageLike.evaluate(() => document.readyState); } break; }
        catch { await sleep(150); }
    }
}

async function initBrowser() {
    if (browser) return;
    log('info', 'Launching Browser...');
    browser = await chromium.launch({ headless: false });
    context = await browser.newContext({ ignoreHTTPSErrors: true, viewport: { width: 1400, height: 900 } });
    page = await context.newPage();

    log('info', `Navigating to ${BASE_URL}`);
    try {
        await page.goto(BASE_URL, { waitUntil: 'domcontentloaded' });
        await waitSettled(page, 2000);
        log('success', 'SDM Loaded. Ready for Batch.');
    } catch (e) {
        log('error', 'Failed to load SDM: ' + e.message);
    }
}

// ---- Socket Events ----

io.on('connection', (socket) => {
    log('info', 'Web Client Connected');
    initBrowser();

    socket.on('set_user', (data) => {
        userAssignee = `${data.last}, ${data.first}`;
        log('success', `Assignee set to: ${userAssignee}`);
    });

    socket.on('upload_batch', (tickets) => {
        ticketQueue = tickets;
        currentTicketIndex = 0;
        log('info', `Batch Uploaded: ${tickets.length} tickets loaded.`);
    });

    socket.on('start_batch', async (data) => {
        socket.data.ticketType = data.type;
        await processNextTicket(socket);
    });

    socket.on('process_current_ticket', async (tasks) => {
        const popup = socket.data.popup;
        const wfFrame = socket.data.wfFrame;
        const ticketNum = ticketQueue[currentTicketIndex];

        if (!popup) return log('error', 'No active ticket popup');

        log('info', `Processing Tasks for Ticket ${ticketNum}: ${tasks.join(', ')}`);

        // --- Process Tasks Loop (Same as SDM_CLI.js) ---
        for (const taskText of tasks) {
            // Status Check
            try {
                const status = await wfFrame.evaluate((text) => {
                    const anchors = Array.from(document.querySelectorAll('a.record, a[href^="javascript:do_default("], tr.jqgrow td:first-child a'));
                    const anchor = anchors.find(a => a.textContent.trim() == text);
                    if (!anchor) return null;
                    const row = anchor.closest('tr');
                    if (!row) return null;
                    const cols = row.querySelectorAll('td');
                    if (cols.length > 5) return cols[5].textContent.trim();
                    return null;
                }, taskText);

                if (status && status.toUpperCase() !== 'PENDING') {
                    log('warn', `Skipping Task ${taskText}: Status is ${status}`);
                    continue;
                }
            } catch (e) { }

            // Process
            log('info', `Starting Task ${taskText}...`);
            let success = false;
            let detailPopup = null;

            for (let attempt = 1; attempt <= 3; attempt++) {
                try {
                    let anchor = await findTaskAnchorAcrossFrames(popup, taskText, wfFrame);
                    if (!anchor) {
                        const invoked = await invokeDoDefaultForTask(wfFrame, taskText);
                        if (!invoked) throw new Error('Task link not found');
                    }

                    detailPopup = await runAndCatchPopup(popup, context, async () => {
                        if (anchor && anchor.locator) await anchor.locator.click({ timeout: 4000 });
                    }, 12000);

                    if (!detailPopup) throw new Error('Detail popup failed to open');

                    await updateTaskDetail(detailPopup, taskText);
                    log('success', `Task ${taskText} Completed!`);
                    success = true;
                    break;
                } catch (e) {
                    log('warn', `Attempt ${attempt} failed: ${e.message}`);
                    if (detailPopup) await detailPopup.close().catch(() => { });
                    detailPopup = null;
                    await sleep(2000);
                }
            }

            if (success && detailPopup) await detailPopup.close().catch(() => { });
            await sleep(1500);
        }
        // -----------------------------------------------

        // Close Ticket Popup
        log('info', `Closing Ticket ${ticketNum}...`);
        await popup.close().catch(() => { });

        // Move to Next
        currentTicketIndex++;
        await processNextTicket(socket);
    });

    socket.on('skip_current_ticket', async () => {
        const popup = socket.data.popup;
        const ticketNum = ticketQueue[currentTicketIndex];
        log('warn', `Skipping Ticket ${ticketNum} by user request.`);

        if (popup) await popup.close().catch(() => { });

        currentTicketIndex++;
        await processNextTicket(socket);
    });
});

async function processNextTicket(socket) {
    if (currentTicketIndex >= ticketQueue.length) {
        log('success', 'Batch Processing Complete!');
        socket.emit('batch_complete');
        return;
    }

    const ticketNum = ticketQueue[currentTicketIndex];
    const ticketType = socket.data.ticketType;

    log('info', `----------------------------------------`);
    log('info', `Loading Ticket ${currentTicketIndex + 1}/${ticketQueue.length}: ${ticketNum}`);
    socket.emit('loading_ticket');

    try {
        // 1. Search & Open Ticket
        const searchFrame = await findFrameWithSelectors(page, ['input[name="searchKey"]'], 30000, 200);
        if (!searchFrame) throw new Error('Search UI not found (Main Page)');

        await searchFrame.fill('input[name="searchKey"]', '');
        await searchFrame.fill('input[name="searchKey"]', ticketNum);

        const sel = await searchFrame.$('#ticket_type');
        if (sel) await searchFrame.selectOption('#ticket_type', ticketType).catch(() => { });

        log('info', 'Clicking Search...');
        const popup = await runAndCatchPopup(page, context, async () => {
            await searchFrame.click('a#imgBtn0, a[name="imgBtn0"]', { timeout: 5000 });
        }, 15000);

        if (!popup) throw new Error('Ticket popup did not open');
        await waitSettled(popup, 3000);

        // 2. Open Workflow Tasks
        const wfFrame = await openWorkflowTasksTab(popup);

        // 3. Discover Tasks
        const tasks = await discoverTasks(wfFrame);
        if (tasks.length === 0) {
            log('warn', 'No tasks found automatically.');
            // We still show the UI so user can skip or see "No tasks"
        }

        // 4. Notify UI
        socket.data.popup = popup;
        socket.data.wfFrame = wfFrame;

        socket.emit('ticket_ready', {
            index: currentTicketIndex,
            total: ticketQueue.length,
            ticketNum: ticketNum,
            tasks: tasks
        });
        log('success', `Ticket ${ticketNum} Ready. Waiting for user selection...`);

    } catch (e) {
        log('error', `Failed to load Ticket ${ticketNum}: ${e.message}`);
        // Auto-skip on failure? Or wait for user to click skip? 
        // Better to wait for user to see the error and click skip to avoid runaway failures.
        socket.emit('ticket_ready', {
            index: currentTicketIndex,
            total: ticketQueue.length,
            ticketNum: ticketNum + " (LOAD FAILED)",
            tasks: []
        });
    }
}

// ---- Robust Helpers (Identical to SDM_CLI.js) ----

async function findFrameWithSelectors(pageOrPopup, selectors, timeoutMs = 30000, pollMs = 200) {
    const start = Date.now();
    while (Date.now() - start < timeoutMs) {
        for (const f of pageOrPopup.frames()) {
            let ok = true;
            for (const sel of selectors) {
                const h = await f.$(sel);
                if (!h) { ok = false; break; }
            }
            if (ok) return f;
        }
        await sleep(pollMs);
    }
    return null;
}

async function runAndCatchPopup(pageOrPopup, context, actionFn, timeout = 15000) {
    const ownerPage = pageOrPopup.page ? pageOrPopup.page() : pageOrPopup;
    const pagesBefore = context.pages().length;
    const pagePopupP = ownerPage.waitForEvent('popup', { timeout }).catch(() => null);
    const ctxPageP = context.waitForEvent('page', { timeout }).catch(() => null);

    await actionFn();

    let popup = await Promise.race([pagePopupP, ctxPageP, sleep(1000).then(() => null)]);

    if (!popup) {
        await sleep(400);
        const pagesAfter = context.pages();
        if (pagesAfter.length > pagesBefore) {
            popup = pagesAfter[pagesAfter.length - 1];
            if (popup === ownerPage) popup = null;
        }
    }

    if (popup) {
        try { await popup.waitForLoadState('domcontentloaded', { timeout: 8000 }); } catch { }
        await waitSettled(popup, 2500);
    }

    return popup;
}

async function getWorkflowTasksFrame(popup, timeoutMs = 15000, pollMs = 200) {
    const deadline = Date.now() + timeoutMs;
    while (Date.now() < deadline) {
        const byName = popup.frame({ name: 'accTab_5_crro_nb_int_iframe_0' });
        if (byName) { try { await byName.waitForLoadState?.('domcontentloaded', { timeout: 3000 }); } catch { } return byName; }
        const frames = popup.frames();
        const match = frames.find(f => (f.url() || '').includes('FACTORY=cr_wf'));
        if (match) { try { await match.waitForLoadState?.('domcontentloaded', { timeout: 3000 }); } catch { } return match; }
        await popup.waitForTimeout(pollMs);
    }
    return null;
}

async function openWorkflowTasksTab(popup) {
    await waitSettled(popup, 6000);
    let lastErr;
    for (let i = 1; i <= 3; i++) {
        try {
            const frame = await findFrameWithSelectors(popup, ['#accrdnHyprlnk1'], 6000, 150)
                || await findFrameWithSelectors(popup, ['#tabHyprlnk1_5'], 6000, 150)
                || popup.mainFrame();

            const accordion = frame.locator('h2#accrdnHyprlnk1, #accrdnHyprlnk1');
            if (await accordion.count()) {
                try { await accordion.scrollIntoViewIfNeeded({ timeout: 2000 }); } catch { }
                await accordion.click({ timeout: 4000 }).catch(() => { });
                await waitSettled(popup, 1200);
            } else {
                const accByText = frame.locator('h2:has-text("Additional Information"), a:has-text("Additional Information")').first();
                if (await accByText.count()) {
                    try { await accByText.scrollIntoViewIfNeeded({ timeout: 2000 }); } catch { }
                    await accByText.click({ timeout: 4000 }).catch(() => { });
                    await waitSettled(popup, 1200);
                }
            }

            const tab = frame.locator('a#tabHyprlnk1_5');
            if (await tab.count()) {
                const cls = (await tab.getAttribute('class')) || '';
                if (!cls.includes('current')) {
                    try { await tab.scrollIntoViewIfNeeded({ timeout: 2000 }); } catch { }
                    await tab.click({ timeout: 4000 });
                }
            } else {
                const tabByText = frame.locator('a:has-text("Workflow Tasks")').first();
                if (!await tabByText.count()) throw new Error('Workflow Tasks tab not found');
                try { await tabByText.scrollIntoViewIfNeeded({ timeout: 2000 }); } catch { }
                await tabByText.click({ timeout: 4000 });
            }
            await waitSettled(popup, 1200);

            const wfFrame = await getWorkflowTasksFrame(popup, 15000, 200);
            if (!wfFrame) throw new Error('Workflow Tasks iframe not found/loaded');

            try {
                await wfFrame.waitForSelector('a[href^="javascript:do_default("], a.record, tr.jqgrow td:first-child a', { timeout: 4000 });
            } catch { }
            return wfFrame;
        } catch (e) {
            lastErr = e;
            await waitSettled(popup, 800);
        }
    }
    throw lastErr || new Error('Failed to open Workflow Tasks tab');
}

async function discoverTasks(wfFrame) {
    return await wfFrame.evaluate(() => {
        const anchors = Array.from(document.querySelectorAll('a.record, a[href^="javascript:do_default("], tr.jqgrow td:first-child a'));
        return [...new Set(anchors.map(a => a.textContent.trim()).filter(t => /^\d+$/.test(t)))].sort();
    });
}

async function findTaskAnchorAcrossFrames(popup, taskText, preferredFrame = null) {
    const exactText = new RegExp(`^\\s*${taskText}\\s*$`);
    const tryInFrame = async (f) => {
        const idGuess = (taskText === '200') ? '#rslnk_0_0' : (taskText === '250') ? '#rslnk_1_0' : null;
        if (idGuess) { const byKnownId = f.locator(idGuess); if (await byKnownId.count()) return { frame: f, locator: byKnownId }; }
        const byText = f.locator('a.record', { hasText: exactText }).first();
        if (await byText.count()) return { frame: f, locator: byText };
        const firstCol = f.locator('tr.jqgrow td:first-child a', { hasText: exactText }).first();
        if (await firstCol.count()) return { frame: f, locator: firstCol };
        const doDef = f.locator('a[href^="javascript:do_default("]', { hasText: exactText }).first();
        if (await doDef.count()) return { frame: f, locator: doDef };
        return null;
    };

    if (preferredFrame) {
        const hit = await tryInFrame(preferredFrame);
        if (hit) return hit;
    } else {
        const wfFrame = popup.frame({ name: 'accTab_5_crro_nb_int_iframe_0' })
            || popup.frames().find(f => (f.url() || '').includes('FACTORY=cr_wf'));
        if (wfFrame) { const hit = await tryInFrame(wfFrame); if (hit) return hit; }
    }

    for (const f of popup.frames()) {
        const hit = await tryInFrame(f);
        if (hit) return hit;
    }
    return null;
}

async function invokeDoDefaultForTask(frame, taskText) {
    const n = await frame.evaluate((text) => {
        const a = Array.from(document.querySelectorAll('a[href^="javascript:do_default("]'))
            .find(x => (x.textContent || '').trim() === String(text));
        if (!a) return null;
        const href = a.getAttribute('href') || '';
        const m = href.match(/do_default\((\d+)\)/);
        return m ? parseInt(m[1], 10) : null;
    }, String(taskText));
    if (n == null) return false;
    await frame.evaluate((row) => { if (typeof window.do_default === 'function') window.do_default(row); }, n);
    return true;
}

async function selectOptionSmart(frame, selector, desiredValue, desiredLabel) {
    const sel = frame.locator(selector);
    try { await sel.waitFor({ state: 'attached', timeout: 5000 }); } catch (e) { throw new Error(`Select not found: ${selector}`); }

    try { await sel.selectOption({ value: desiredValue }); return true; } catch { }
    if (desiredLabel) {
        try { await sel.selectOption({ label: desiredLabel }); return true; } catch { }
        try {
            const opts = await frame.$$eval(selector + ' option', os => os.map(o => ({ v: o.value, t: (o.textContent || '').trim() })));
            const found = opts.find(o => o.t.toLowerCase().includes(desiredLabel.toLowerCase()));
            if (found) { await sel.selectOption({ value: found.v }); return true; }
        } catch { }
    }
    return false;
}

async function clickEditRobust(detailPopup) {
    const start = Date.now();
    log('info', '[DEBUG] Scanning frames for Edit button (Prioritizing cai_main)...');
    while (Date.now() - start < 15000) {
        // 1. Try cai_main explicitly FIRST
        const caiMain = detailPopup.frame({ name: 'cai_main' });
        if (caiMain) {
            const btn = caiMain.locator('a#imgBtn0, a[name="imgBtn0"], a.button:has(span:has-text("Edit"))').first();
            if (await btn.count() && await btn.isVisible()) {
                log('info', `[DEBUG] Found Edit button in prioritized frame 'cai_main'`);
                try {
                    await btn.scrollIntoViewIfNeeded({ timeout: 1500 }).catch(() => { });
                    await sleep(500);
                    await btn.click({ timeout: 6000 });
                    log('success', `[DEBUG] Clicked Edit button in 'cai_main'`);
                    return true;
                } catch (e) {
                    log('warn', `[DEBUG] Standard click failed in 'cai_main': ${e.message}, trying JS click...`);
                    try { await btn.evaluate(e => e.click()); return true; } catch { }
                }
            }
        }

        // 2. Fallback (excluding gobtn as requested)
        for (const f of detailPopup.frames()) {
            if (f.name() === 'cai_main') continue; // Already checked
            if (f.name() === 'gobtn') continue;

            const btn = f.locator('a#imgBtn0, a[name="imgBtn0"], a.button:has(span:has-text("Edit"))').first();
            if (await btn.count() && await btn.isVisible()) {
                log('info', `[DEBUG] Found Edit button in frame '${f.name()}'`);
                try {
                    await btn.scrollIntoViewIfNeeded({ timeout: 1500 }).catch(() => { });
                    await sleep(500);
                    await btn.click({ timeout: 6000 });
                    log('success', `[DEBUG] Clicked Edit button in '${f.name()}'`);
                    return true;
                } catch (e) {
                    try { await btn.evaluate(e => e.click()); return true; } catch { }
                }
            }
        }
        await sleep(500);
    }
    throw new Error("Could not find or click Edit button");
}

async function findMainFrameRobust(detailPopup) {
    const start = Date.now();
    log('info', '[DEBUG] Scanning frames for Assignee/Status fields...');
    while (Date.now() - start < 15000) {
        for (const f of detailPopup.frames()) {
            const hasFields = await f.locator('input[name="assignee_combo_name"], select[name="SET.status"]').first().count();
            if (hasFields) {
                log('info', `[DEBUG] Found Main Frame with fields: '${f.name()}'`);
                return f;
            }
        }
        await sleep(500);
    }
    return null;
}

async function updateTaskDetail(detailPopup, taskText) {
    log('info', `Updating Task ${taskText}...`);

    try {
        log('info', 'Searching for Edit button...');
        await clickEditRobust(detailPopup);
    } catch (e) {
        throw new Error(`Edit failed: ${e.message}`);
    }

    log('info', 'Waiting for page to settle after Edit click...');
    await waitSettled(detailPopup, 2000);

    const mainFrame = await findMainFrameRobust(detailPopup);
    if (!mainFrame) throw new Error("Could not find Main Frame with form fields (Edit mode not active?)");

    try {
        const assignee = mainFrame.locator('input[name="assignee_combo_name"]');
        await assignee.waitFor({ state: 'visible', timeout: 10000 });

        if (await assignee.count()) {
            await assignee.click({ timeout: 2000 }).catch(() => { });
            await assignee.fill(userAssignee, { timeout: 3000 });
            await assignee.press('Enter').catch(() => { });
            await assignee.evaluate(el => el.blur()).catch(() => { });
            await waitSettled(detailPopup, 500);
        } else {
            throw new Error('Assignee field not found');
        }
    } catch (e) {
        throw new Error(`Assignee set error: ${e.message}`);
    }

    try {
        const statusSel = 'select[name="SET.status"]';
        const status = mainFrame.locator(statusSel);
        if (await status.count()) {
            const ok = await selectOptionSmart(mainFrame, statusSel, 'COMP', 'Complete');
            if (!ok) throw new Error('Could not set status to Complete');
            await waitSettled(detailPopup, 300);
        } else {
            throw new Error('Status select not found');
        }
    } catch (e) {
        throw new Error(`Status set error: ${e.message}`);
    }

    try {
        log('info', 'Clicking Save...');
        const start = Date.now();
        let clicked = false;
        while (Date.now() - start < 10000) {
            const caiMain = detailPopup.frame({ name: 'cai_main' });
            if (caiMain) {
                const btn = caiMain.locator('a.button:has(span:has-text("Save")), a:has-text("Save")').first();
                if (await btn.count() && await btn.isVisible()) {
                    try {
                        await btn.scrollIntoViewIfNeeded({ timeout: 1500 }).catch(() => { });
                        await btn.click({ timeout: 2000 });
                        clicked = true;
                        break;
                    } catch (e) {
                        try { await btn.evaluate(e => e.click()); clicked = true; break; } catch { }
                    }
                }
            }

            for (const f of detailPopup.frames()) {
                if (f.name() === 'cai_main') continue;
                if (f.name() === 'gobtn') continue;

                const btn = f.locator('a.button:has(span:has-text("Save")), a:has-text("Save")').first();
                if (await btn.count() && await btn.isVisible()) {
                    try {
                        await btn.scrollIntoViewIfNeeded({ timeout: 1500 }).catch(() => { });
                        await btn.click({ timeout: 2000 });
                        clicked = true;
                        break;
                    } catch (e) {
                        try { await btn.evaluate(e => e.click()); clicked = true; break; } catch { }
                    }
                }
            }
            if (clicked) break;
            await sleep(500);
        }
        if (!clicked) throw new Error("Save button not found/clicked");
        await waitSettled(detailPopup, 1500);
    } catch (e) {
        throw new Error(`Save click error: ${e.message}`);
    }
}

// ---- Start Server ----
server.listen(PORT, () => {
    console.log(`Server running at http://localhost:${PORT}`);
    open(`http://localhost:${PORT}`);
});
