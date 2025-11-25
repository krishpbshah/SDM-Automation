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
 * - Robust Automation Logic (Ported from Automation.js)
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

async function sleep(ms){ return new Promise(r => setTimeout(r, ms)); }

async function waitSettled(pageLike, timeoutMs = 15000) {
    const start = Date.now();
    try { await pageLike.waitForLoadState?.('domcontentloaded', { timeout: Math.min(5000, timeoutMs) }); } catch {}
    try { await pageLike.waitForLoadState?.('load', { timeout: Math.min(5000, timeoutMs) }); } catch {}
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
            const searchFrame = await findFrameWithSelectors(page, ['input[name="searchKey"]'], 30000, 200);
            if (!searchFrame) throw new Error('Search UI not found');

            await searchFrame.fill('input[name="searchKey"]', '');
            await searchFrame.fill('input[name="searchKey"]', data.num);
            
            const sel = await searchFrame.$('#ticket_type');
            if (sel) await searchFrame.selectOption('#ticket_type', data.type).catch(()=>{});

            log('info', 'Clicking Search...');
            const popup = await runAndCatchPopup(page, context, async () => {
                await searchFrame.click('a#imgBtn0, a[name="imgBtn0"]', { timeout: 5000 });
            }, 15000);

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
            log('info', `Checking status for Task ${taskText}...`);
            
            // --- STATUS CHECK ---
            try {
                const status = await wfFrame.evaluate((text) => {
                    const anchors = Array.from(document.querySelectorAll('a.record, a[href^="javascript:do_default("], tr.jqgrow td:first-child a'));
                    const anchor = anchors.find(a => a.textContent.trim() == text);
                    if(!anchor) return null;
                    
                    // Find row
                    const row = anchor.closest('tr');
                    if(!row) return null;
                    
                    // Column 5 (0-indexed) is usually Status
                    const cols = row.querySelectorAll('td');
                    if(cols.length > 5) return cols[5].textContent.trim();
                    return null;
                }, taskText);

                if(status && status.toUpperCase() !== 'PENDING') {
                    log('warn', `Skipping Task ${taskText}: Status is ${status}`);
                    continue;
                } else if(status) {
                    log('info', `Status is ${status}. Processing...`);
                }
            } catch(e) {
                log('warn', `Could not verify status for ${taskText}, proceeding anyway...`);
            }
            // --------------------

            log('info', `Starting Task ${taskText}...`);
            let success = false;
            let detailPopup = null;

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
                    detailPopup = await runAndCatchPopup(popup, context, async () => {
                        if(anchor && anchor.locator) await anchor.locator.click({timeout: 4000});
                    }, 12000);

                    if(!detailPopup) throw new Error('Detail popup failed to open');
                    
                    await updateTaskDetail(detailPopup, taskText);
                    log('success', `Task ${taskText} Completed!`);
                    success = true;
                    break;
                } catch (e) {
                    log('warn', `Attempt ${attempt} failed: ${e.message}`);
                    if(detailPopup) await detailPopup.close().catch(()=>{}); // Close if failed
                    detailPopup = null;
                    await new Promise(r => setTimeout(r, 2000));
                }
            }
            
            if(success && detailPopup) {
                await detailPopup.close().catch(()=>{});
            }
            
            if(!success) log('error', `Failed to process Task ${taskText}`);
            
            // Wait before next task
            await new Promise(r => setTimeout(r, 2000));
        }
        log('success', 'All requested tasks finished.');
    });
});

// ---- Robust Helpers (Ported from Automation.js) ----

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
  const ctxPageP   = context.waitForEvent('page', { timeout }).catch(() => null);

  await actionFn();

  let popup = await Promise.race([ pagePopupP, ctxPageP, sleep(1000).then(() => null) ]);

  if (!popup) {
    await sleep(400);
    const pagesAfter = context.pages();
    if (pagesAfter.length > pagesBefore) {
      popup = pagesAfter[pagesAfter.length - 1];
      if (popup === ownerPage) popup = null;
    }
  }

  if (popup) {
    try { await popup.waitForLoadState('domcontentloaded', { timeout: 8000 }); } catch {}
    await waitSettled(popup, 2500);
  }

  return popup;
}

async function getWorkflowTasksFrame(popup, timeoutMs = 15000, pollMs = 200) {
  const deadline = Date.now() + timeoutMs;
  while (Date.now() < deadline) {
    const byName = popup.frame({ name: 'accTab_5_crro_nb_int_iframe_0' });
    if (byName) { try { await byName.waitForLoadState?.('domcontentloaded', { timeout: 3000 }); } catch {} return byName; }
    const frames = popup.frames();
    const match = frames.find(f => (f.url() || '').includes('FACTORY=cr_wf'));
    if (match) { try { await match.waitForLoadState?.('domcontentloaded', { timeout: 3000 }); } catch {} return match; }
    await popup.waitForTimeout(pollMs);
  }
  return null;
}

async function openWorkflowTasksTab(popup) {
    // Adapted from Automation.js
    await waitSettled(popup, 6000);
    let lastErr;
    for (let i = 1; i <= 3; i++) {
        try {
            const frame = await findFrameWithSelectors(popup, ['#accrdnHyprlnk1'], 6000, 150)
                        || await findFrameWithSelectors(popup, ['#tabHyprlnk1_5'], 6000, 150)
                        || popup.mainFrame();

            // Accordion
            const accordion = frame.locator('h2#accrdnHyprlnk1, #accrdnHyprlnk1');
            if (await accordion.count()) {
                try { await accordion.scrollIntoViewIfNeeded({ timeout: 2000 }); } catch {}
                await accordion.click({ timeout: 4000 }).catch(() => {});
                await waitSettled(popup, 1200);
            } else {
                const accByText = frame.locator('h2:has-text("Additional Information"), a:has-text("Additional Information")').first();
                if (await accByText.count()) {
                    try { await accByText.scrollIntoViewIfNeeded({ timeout: 2000 }); } catch {}
                    await accByText.click({ timeout: 4000 }).catch(() => {});
                    await waitSettled(popup, 1200);
                }
            }

            // Workflow Tasks tab
            const tab = frame.locator('a#tabHyprlnk1_5');
            if (await tab.count()) {
                const cls = (await tab.getAttribute('class')) || '';
                if (!cls.includes('current')) {
                    try { await tab.scrollIntoViewIfNeeded({ timeout: 2000 }); } catch {}
                    await tab.click({ timeout: 4000 });
                }
            } else {
                const tabByText = frame.locator('a:has-text("Workflow Tasks")').first();
                if (!await tabByText.count()) throw new Error('Workflow Tasks tab not found');
                try { await tabByText.scrollIntoViewIfNeeded({ timeout: 2000 }); } catch {}
                await tabByText.click({ timeout: 4000 });
            }
            await waitSettled(popup, 1200);

            // Tasks iframe
            const wfFrame = await getWorkflowTasksFrame(popup, 15000, 200);
            if (!wfFrame) throw new Error('Workflow Tasks iframe not found/loaded');

            try {
                await wfFrame.waitForSelector('a[href^="javascript:do_default("], a.record, tr.jqgrow td:first-child a', { timeout: 4000 });
            } catch {}
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
  try { await sel.waitFor({ state: 'attached', timeout: 5000 }); } catch(e) { throw new Error(`Select not found: ${selector}`); }
  
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

// --- STATELESS ROBUST HELPERS ---

async function clickEditRobust(detailPopup) {
    const start = Date.now();
    while (Date.now() - start < 15000) {
        // Scan all frames for the Edit button
        for (const f of detailPopup.frames()) {
            const btn = f.locator('a#imgBtn0, a[name="imgBtn0"], a.button:has(span:has-text("Edit"))').first();
            if (await btn.count() && await btn.isVisible()) {
                log('info', `Found Edit button in frame '${f.name()}'`);
                try {
                    await btn.click({ timeout: 2000 });
                    return true; // Clicked!
                } catch (e) {
                    log('warn', 'Standard click failed, trying JS click...');
                    try {
                        await btn.evaluate(e => e.click());
                        return true;
                    } catch (e2) {}
                }
            }
        }
        await sleep(500);
    }
    throw new Error("Could not find or click Edit button");
}

async function waitForSaveRobust(detailPopup) {
    const start = Date.now();
    while (Date.now() - start < 20000) {
        for (const f of detailPopup.frames()) {
            const btn = f.locator('a.button:has(span:has-text("Save")), a:has-text("Save")').first();
            if (await btn.count() && await btn.isVisible()) {
                log('info', `Edit mode confirmed (Save button found in '${f.name()}')`);
                return true;
            }
        }
        await sleep(500);
    }
    return false;
}

async function findMainFrameRobust(detailPopup) {
    const start = Date.now();
    while (Date.now() - start < 15000) {
        for (const f of detailPopup.frames()) {
            const hasFields = await f.locator('input[name="assignee_combo_name"], select[name="SET.status"]').first().count();
            if (hasFields) return f;
        }
        await sleep(500);
    }
    return null;
}

async function updateTaskDetail(detailPopup, taskText) {
  log('info', `Updating Task ${taskText}...`);
  
  // 1. Click Edit (Stateless)
  try {
    log('info', 'Searching for Edit button...');
    await clickEditRobust(detailPopup);
  } catch (e) {
      throw new Error(`Edit failed: ${e.message}`);
  }

  // 2. Wait for Save (Stateless)
  log('info', 'Waiting for Edit mode (Save button)...');
  const inEditMode = await waitForSaveRobust(detailPopup);
  
  if (!inEditMode) {
      log('warn', 'Could not confirm Edit mode (Save button missing), proceeding anyway...');
  }

  // 3. Find Main Frame (Stateless)
  const mainFrame = await findMainFrameRobust(detailPopup);
  if (!mainFrame) throw new Error("Could not find Main Frame with form fields");

  // Assignee
  try {
    const assignee = mainFrame.locator('input[name="assignee_combo_name"]');
    await assignee.waitFor({ state: 'visible', timeout: 10000 });
    
    if (await assignee.count()) {
      await assignee.click({ timeout: 2000 }).catch(() => {});
      await assignee.fill(userAssignee, { timeout: 3000 });
      await assignee.press('Enter').catch(() => {});
      await assignee.evaluate(el => el.blur()).catch(() => {});
      await waitSettled(detailPopup, 500);
    } else {
      throw new Error('Assignee field not found');
    }
  } catch (e) {
    throw new Error(`Assignee set error: ${e.message}`);
  }

  // Status
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

  // Save (Stateless)
  try {
    log('info', 'Clicking Save...');
    // Reuse clickEditRobust logic but for Save
    const start = Date.now();
    let clicked = false;
    while (Date.now() - start < 10000) {
        for (const f of detailPopup.frames()) {
            const btn = f.locator('a.button:has(span:has-text("Save")), a:has-text("Save")').first();
            if (await btn.count() && await btn.isVisible()) {
                try {
                    await btn.click({ timeout: 2000 });
                    clicked = true;
                    break;
                } catch (e) {
                     try { await btn.evaluate(e => e.click()); clicked = true; break; } catch {}
                }
            }
        }
        if(clicked) break;
        await sleep(500);
    }
    if(!clicked) throw new Error("Save button not found/clicked");
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
'@ | Set-Content -Path ".\SDM_CLI.js" -Encoding UTF8
