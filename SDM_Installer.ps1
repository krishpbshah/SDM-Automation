$npmrc = @'
strict-ssl=false
registry=http://registry.npmjs.org/
'@
$npmrc | Set-Content -Path ".\.npmrc" -Encoding UTF8

# Set env var for the current session just in case
$env:NODE_TLS_REJECT_UNAUTHORIZED = "0"

$packageJson = @'
{
  "name": "sdm-automation",
  "version": "2.0.0",
  "description": "Web-based Automation script for SDM",
  "main": "SDM_CLI.js",
  "scripts": {
    "start": "node SDM_CLI.js"
  },
  "dependencies": {
    "express": "^4.18.2",
    "socket.io": "^4.7.2",
    "playwright": "^1.40.0",
    "open": "^8.4.2"
  }
}
'@
$packageJson | Set-Content -Path ".\package.json" -Encoding UTF8

$scriptContent = @'
/**
 * SDM_CLI.js (Web Edition)
 *
 * Local Web Server for CA SDM Automation.
 *
 * Features:
 * - Web GUI (Dark Mode)
 * - Dynamic Assignee Name
 * - Real-time Logs via WebSockets
 * - Persistent Browser Session
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
let userAssignee = "Couto, Lucas"; // Default, will be updated by user

// ---- Express & Socket.io Setup ----
const app = express();
const server = http.createServer(app);
const io = new Server(server);

// Serve the Single Page App
app.get('/', (req, res) => {
    res.send(`
<!DOCTYPE html>
<html lang="en">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>SDM Automation</title>
    <style>
        :root { --bg: #0f172a; --card: #1e293b; --text: #f8fafc; --accent: #3b82f6; --success: #22c55e; --danger: #ef4444; }
        body { font-family: 'Segoe UI', system-ui, sans-serif; background: var(--bg); color: var(--text); margin: 0; padding: 20px; display: flex; justify-content: center; height: 100vh; box-sizing: border-box; }
        .container { display: grid; grid-template-columns: 1fr 1fr; gap: 20px; width: 100%; max-width: 1200px; height: 100%; }
        .panel { background: var(--card); padding: 20px; border-radius: 12px; display: flex; flex-direction: column; gap: 15px; box-shadow: 0 4px 6px -1px rgba(0,0,0,0.5); overflow-y: auto; }
        h2 { margin: 0 0 10px 0; border-bottom: 1px solid #334155; padding-bottom: 10px; color: var(--accent); }
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
        .task-item.selected { border-color: var(--accent); background: #1e3a8a; }
        .hidden { display: none; }
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

            <h2>Ticket Control</h2>
            <label>Ticket Number</label>
            <input type="text" id="ticketNum" placeholder="e.g. 123456">
            <label>Ticket Type</label>
            <select id="ticketType">
                <option value="go_cr">Change Request (go_cr)</option>
                <option value="go_in">Incident (go_in)</option>
                <option value="go_pr">Problem (go_pr)</option>
            </select>
            <button id="openTicketBtn" disabled>Open Ticket</button>

            <div id="taskSection" class="hidden">
                <h2>Select Tasks</h2>
                <div id="taskList" class="task-list"></div>
                <br>
                <button id="processBtn">Process Selected Tasks</button>
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
            document.getElementById('openTicketBtn').disabled = false;
            document.getElementById('saveConfigBtn').textContent = 'Saved!';
            setTimeout(() => document.getElementById('saveConfigBtn').textContent = 'Save Configuration', 2000);
        });

        // --- Ticket ---
        document.getElementById('openTicketBtn').addEventListener('click', () => {
            const num = document.getElementById('ticketNum').value.trim();
            const type = document.getElementById('ticketType').value;
            if(!num) return alert('Enter ticket number');
            
            document.getElementById('taskSection').classList.add('hidden');
            taskList.innerHTML = '';
            selectedTasks.clear();
            
            socket.emit('open_ticket', { num, type });
        });

        // --- Tasks ---
        socket.on('tasks_discovered', (tasks) => {
            document.getElementById('taskSection').classList.remove('hidden');
            taskList.innerHTML = '';
            tasks.forEach(t => {
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

        document.getElementById('processBtn').addEventListener('click', () => {
            if(selectedTasks.size === 0) return alert('Select at least one task');
            socket.emit('process_tasks', Array.from(selectedTasks));
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

async function waitSettled(pageLike, timeoutMs = 15000) {
    const start = Date.now();
    try { await pageLike.waitForLoadState?.('domcontentloaded', { timeout: 5000 }); } catch { }
    while (Date.now() - start < timeoutMs) {
        try { if (pageLike.evaluate) { await pageLike.evaluate(() => document.readyState); } break; }
        catch { await new Promise(r => setTimeout(r, 150)); }
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
        log('success', 'SDM Loaded. Ready for Ticket.');
    } catch (e) {
        log('error', 'Failed to load SDM: ' + e.message);
    }
}

// ---- Socket Events ----

io.on('connection', (socket) => {
    log('info', 'Web Client Connected');
    initBrowser(); // Start browser on first connection

    socket.on('set_user', (data) => {
        // Format: "Last, First"
        userAssignee = `${data.last}, ${data.first}`;
        log('success', `Assignee set to: ${userAssignee}`);
    });

    socket.on('open_ticket', async (data) => {
        log('info', `Opening Ticket ${data.num} (${data.type})...`);
        try {
            const searchFrame = await findFrameWithSelectors(page, ['input[name="searchKey"]'], 5000);
            if (!searchFrame) throw new Error('Search UI not found');

            await searchFrame.fill('input[name="searchKey"]', '');
            await searchFrame.fill('input[name="searchKey"]', data.num);
            
            const sel = await searchFrame.$('#ticket_type');
            if (sel) await searchFrame.selectOption('#ticket_type', data.type).catch(()=>{});

            log('info', 'Clicking Search...');
            const popup = await runAndCatchPopup(page, context, async () => {
                await searchFrame.click('a#imgBtn0, a[name="imgBtn0"]', { timeout: 5000 });
            });

            if (!popup) throw new Error('Ticket popup did not open');
            
            log('success', 'Ticket Opened. Loading Workflow Tasks...');
            await waitSettled(popup, 3000);

            // Open Workflow Tasks Tab
            const wfFrame = await openWorkflowTasksTab(popup);
            
            // Discover Tasks
            const tasks = await discoverTasks(wfFrame);
            if(tasks.length === 0) {
                log('warn', 'No tasks found automatically. Adding defaults.');
                tasks.push('200', '250', '300');
            }
            
            // Store popup in socket data for next step
            socket.data.popup = popup;
            socket.data.wfFrame = wfFrame;
            
            socket.emit('tasks_discovered', tasks);
            log('info', `Tasks Discovered: ${tasks.join(', ')}`);

        } catch (e) {
            log('error', e.message);
        }
    });

    socket.on('process_tasks', async (tasks) => {
        const popup = socket.data.popup;
        const wfFrame = socket.data.wfFrame;
        if(!popup) return log('error', 'No active ticket popup');

        log('info', `Processing Tasks: ${tasks.join(', ')}`);

        for (const taskText of tasks) {
            log('info', `Starting Task ${taskText}...`);
            let success = false;
            for(let attempt=1; attempt<=3; attempt++) {
                try {
                    // Find Anchor
                    let anchor = await findTaskAnchorAcrossFrames(popup, taskText, wfFrame);
                    if(!anchor) {
                        // Try do_default
                        const invoked = await invokeDoDefaultForTask(wfFrame, taskText);
                        if(!invoked) throw new Error('Task link not found');
                    }

                    // Click
                    const detailPopup = await runAndCatchPopup(popup, context, async () => {
                        if(anchor && anchor.locator) await anchor.locator.click({timeout: 4000});
                    }, 10000);

                    if(!detailPopup) throw new Error('Detail popup failed to open');
                    
                    await updateTaskDetail(detailPopup, taskText);
                    log('success', `Task ${taskText} Completed!`);
                    success = true;
                    break;
                } catch (e) {
                    log('warn', `Attempt ${attempt} failed: ${e.message}`);
                    await new Promise(r => setTimeout(r, 2000));
                }
            }
            if(!success) log('error', `Failed to process Task ${taskText}`);
        }
        log('success', 'All requested tasks finished.');
    });
});

// ---- Helpers (Restored from Automation.js) ----

async function findFrameWithSelectors(pageOrPopup, selectors, timeoutMs=5000) {
    const start = Date.now();
    while (Date.now() - start < timeoutMs) {
        for (const f of pageOrPopup.frames()) {
            let ok = true;
            for (const sel of selectors) { if (!await f.$(sel)) { ok = false; break; } }
            if (ok) return f;
        }
        await new Promise(r => setTimeout(r, 200));
    }
    return null;
}

async function runAndCatchPopup(pageOrPopup, context, actionFn, timeout=15000) {
    const ownerPage = pageOrPopup.page ? pageOrPopup.page() : pageOrPopup;
    const popupP = ownerPage.waitForEvent('popup', { timeout }).catch(() => null);
    await actionFn();
    const popup = await popupP;
    if(popup) await waitSettled(popup, 2500);
    return popup;
}

async function openWorkflowTasksTab(popup) {
    await waitSettled(popup, 2000);
    const frame = await findFrameWithSelectors(popup, ['#tabHyprlnk1_5'], 5000) || popup.mainFrame();
    const tab = frame.locator('a#tabHyprlnk1_5');
    if(await tab.count()) await tab.click();
    else {
        const txt = frame.locator('a:has-text("Workflow Tasks")');
        if(await txt.count()) await txt.click();
    }
    await waitSettled(popup, 2000);
    
    for(let i=0; i<10; i++) {
        const f = popup.frames().find(f => (f.url()||'').includes('FACTORY=cr_wf'));
        if(f) return f;
        await new Promise(r => setTimeout(r, 500));
    }
    throw new Error('Workflow Tasks iframe not found');
}

async function discoverTasks(wfFrame) {
    return await wfFrame.evaluate(() => {
        const anchors = Array.from(document.querySelectorAll('a.record, a[href^="javascript:do_default("], tr.jqgrow td:first-child a'));
        return [...new Set(anchors.map(a => a.textContent.trim()).filter(t => /^\d+$/.test(t)))].sort();
    });
}

async function findTaskAnchorAcrossFrames(popup, taskText, preferredFrame) {
    const exactText = new RegExp(`^\\s*${taskText}\\s*$`);
    const tryFrame = async (f) => {
        const l = f.locator('a.record', { hasText: exactText }).first();
        if(await l.count()) return { frame: f, locator: l };
        return null;
    };
    if(preferredFrame) { const h = await tryFrame(preferredFrame); if(h) return h; }
    for(const f of popup.frames()) { const h = await tryFrame(f); if(h) return h; }
    return null;
}

async function invokeDoDefaultForTask(frame, taskText) {
    return await frame.evaluate((text) => {
        const a = Array.from(document.querySelectorAll('a[href^="javascript:do_default("]')).find(x => x.textContent.trim() == text);
        if(!a) return false;
        const m = a.getAttribute('href').match(/do_default\((\d+)\)/);
        if(m) { window.do_default(m[1]); return true; }
        return false;
    }, taskText);
}

// ---- Restored Robust Helpers ----

async function findHeaderFrame(detailPopup) {
    const named = detailPopup.frame({ name: 'cai_header' });
    if (named) return named;
    for (const f of detailPopup.frames()) {
        const hasBtn = await f.locator('a.button:has(span:has-text("Edit")), a.button:has(span:has-text("Save")), a:has-text("Edit"), a:has-text("Save")').first().count();
        if (hasBtn) return f;
    }
    return detailPopup.mainFrame();
}

async function findMainFrame(detailPopup) {
    const named = detailPopup.frame({ name: 'cai_main' });
    if (named) return named;
    for (const f of detailPopup.frames()) {
        const hasFields = await f.locator('input[name="assignee_combo_name"], select[name="SET.status"]').first().count();
        if (hasFields) return f;
    }
    return detailPopup.mainFrame();
}

async function clickHeaderButtonByText(headerFrame, label, timeout = 6000) {
    const btn = headerFrame.locator(`a.button:has(span:has-text("${label}")), a:has-text("${label}")`).first();
    if (!await btn.count()) throw new Error(`${label} button not found in header frame`);
    try { await btn.scrollIntoViewIfNeeded({ timeout: 1500 }); } catch {}
    await btn.click({ timeout });
}

async function selectOptionSmart(frame, selector, desiredValue, desiredLabel) {
    const sel = frame.locator(selector);
    if (!await sel.count()) throw new Error(`Select not found: ${selector}`);
    try { await sel.selectOption({ value: desiredValue }); return true; } catch {}
    if (desiredLabel) {
        try { await sel.selectOption({ label: desiredLabel }); return true; } catch {}
        try {
            const opts = await frame.$$eval(selector + ' option', os => os.map(o => ({ v: o.value, t: (o.textContent || '').trim() })));
            const found = opts.find(o => o.t.toLowerCase().includes(desiredLabel.toLowerCase()));
            if (found) { await sel.selectOption({ value: found.v }); return true; }
        } catch {}
    }
    return false;
}

async function updateTaskDetail(detailPopup, taskText) {
    log('info', `Updating Task ${taskText}...`);
    
    const headerFrame = await findHeaderFrame(detailPopup);
    const mainFrame = await findMainFrame(detailPopup);

    // Edit
    try {
        await clickHeaderButtonByText(headerFrame, 'Edit', 8000);
        await waitSettled(detailPopup, 1200);
        await mainFrame.waitForSelector('input[name="assignee_combo_name"]', { timeout: 6000 });
        await mainFrame.waitForSelector('select[name="SET.status"]', { timeout: 6000 });
    } catch (e) {
        log('warn', `Edit click issue: ${e.message}`);
    }

    // Assignee
    try {
        const assignee = mainFrame.locator('input[name="assignee_combo_name"]');
        if (await assignee.count()) {
            await assignee.click({ timeout: 2000 }).catch(() => {});
            await assignee.fill(userAssignee, { timeout: 3000 });
            await assignee.press('Enter').catch(() => {});
            await assignee.evaluate(el => el.blur()).catch(() => {});
            await waitSettled(detailPopup, 500);
        } else {
            log('warn', `Assignee field not found in main frame.`);
        }
    } catch (e) {
        log('warn', `Assignee set error: ${e.message}`);
    }

    // Status
    try {
        const statusSel = 'select[name="SET.status"]';
        const status = mainFrame.locator(statusSel);
        if (await status.count()) {
            const ok = await selectOptionSmart(mainFrame, statusSel, 'COMP', 'Complete');
            if (!ok) log('warn', `Could not set status to Complete`);
            await waitSettled(detailPopup, 300);
        } else {
            log('warn', `Status select not found in main frame.`);
        }
    } catch (e) {
        log('warn', `Status set error: ${e.message}`);
    }

    // Save
    try {
        await clickHeaderButtonByText(headerFrame, 'Save', 8000);
        await waitSettled(detailPopup, 1500);
    } catch (e) {
        log('warn', `Save click error: ${e.message}`);
    }
}

// ---- Start Server ----
server.listen(PORT, () => {
    console.log(`Server running at http://localhost:${PORT}`);
    open(`http://localhost:${PORT}`);
});
'@ | Set-Content -Path ".\SDM_CLI.js" -Encoding UTF8
