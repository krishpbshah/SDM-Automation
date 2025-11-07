/**
 * ttc-workflow-tasks.js
 *
 * Flow:
 * 1) Open CA SDM -> Global search -> ticket popup
 * 2) Additional Information -> Workflow Tasks (tabHyprlnk1_5) -> wait for cr_wf iframe
 * 3) For each task (e.g., 200, 250): open detail popup
 * 4) In detail popup:
 *    - Click Edit (in header frame; no new popup)
 *    - Set Assignee = "Couto, Lucas" (auto-fill)
 *    - If Status is Pending, set to Complete
 *    - Click Save (in header)
 *    - Dump before/after HTML and screenshots
 *
 * Usage:
 *   node ttc-workflow-tasks.js <SEARCH_TEXT> [ticketType] [task1] [task2] [...]
 *
 * Example:
 *   node ttc-workflow-tasks.js 846349 go_cr 200 250
 */

const { chromium } = require('playwright');
const fs = require('fs');
const path = require('path');

const BASE_URL = 'http://servicedesk-web.int.ttc.ca/CAisd/pdmweb.exe';
const TARGET_ASSIGNEE = 'Couto, Lucas'; // visible combo text (auto-fill)

function save(file, content) { try { fs.writeFileSync(file, content); } catch {} }
async function sleep(ms){ return new Promise(r => setTimeout(r, ms)); }
function safe(name) { return (name || '').replace(/[^\w.-]+/g, '_') || 'unnamed'; }

async function waitSettled(pageLike, timeoutMs = 15000) {
  const start = Date.now();
  try { await pageLike.waitForLoadState?.('domcontentloaded', { timeout: Math.min(5000, timeoutMs) }); } catch {}
  try { await pageLike.waitForLoadState?.('load', { timeout: Math.min(5000, timeoutMs) }); } catch {}
  while (Date.now() - start < timeoutMs) {
    try { if (pageLike.evaluate) { await pageLike.evaluate(() => document.readyState); } break; }
    catch { await sleep(150); }
  }
}

async function contentSafe(page, attempts = 6) {
  for (let i = 0; i < attempts; i++) {
    try { return await page.content(); }
    catch (e) {
      const msg = (e?.message || '').toLowerCase();
      if (msg.includes('navigating')) { await waitSettled(page, 2500); await sleep(200); continue; }
      throw e;
    }
  }
  try { return await page.evaluate(() => document.documentElement.outerHTML); }
  catch { return '<!-- [unavailable] -->'; }
}

async function screenshotSafe(page, outPath, fullPage = true, attempts = 3) {
  for (let i = 0; i < attempts; i++) {
    try { await page.screenshot({ path: outPath, fullPage }); return; }
    catch (e) {
      const msg = (e?.message || '').toLowerCase();
      if (msg.includes('navigating')) { await waitSettled(page, 2500); await sleep(200); continue; }
      throw e;
    }
  }
}

async function dumpPage(page, base, tag) {
  const html = await contentSafe(page);
  save(`${base}_${tag}_Top.html`, html);
  try { await screenshotSafe(page, `${base}_${tag}_Top.png`, true); } catch {}
  let i = 0;
  for (const f of page.frames()) {
    try {
      const fhtml = await f.evaluate(() => document.documentElement.outerHTML);
      save(`${base}_${tag}_Frame_${i}_${safe(f.name())}.html`, fhtml);
    } catch {}
    i++;
  }
}

// ---- Utilities ----

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

// Prefer the Workflow Tasks iframe (accTab_5_crro_nb_int_iframe_0 / FACTORY=cr_wf)
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

// Open accordion + click "Workflow Tasks" tab, then wait for its iframe to load; return the iframe frame
async function openWorkflowTasksTab(popup, {
  settleBeforeMs = 6000,          // wait (>=6s) for first popup to fully render
  settleAfterAccordionMs = 1200,  // short wait after accordion click
  settleAfterTabMs = 1200,        // short wait after tab click
  attempts = 3
} = {}) {
  await waitSettled(popup, settleBeforeMs);

  let lastErr;
  for (let i = 1; i <= attempts; i++) {
    try {
      const frame = await findFrameWithSelectors(popup, ['#accrdnHyprlnk1'], 6000, 150)
                 || await findFrameWithSelectors(popup, ['#tabHyprlnk1_5'], 6000, 150)
                 || popup.mainFrame();

      // Accordion
      const accordion = frame.locator('h2#accrdnHyprlnk1, #accrdnHyprlnk1');
      if (await accordion.count()) {
        try { await accordion.scrollIntoViewIfNeeded({ timeout: 2000 }); } catch {}
        await accordion.click({ timeout: 4000 }).catch(() => {});
        await waitSettled(popup, settleAfterAccordionMs);
      } else {
        const accByText = frame.locator('h2:has-text("Additional Information"), a:has-text("Additional Information")').first();
        if (await accByText.count()) {
          try { await accByText.scrollIntoViewIfNeeded({ timeout: 2000 }); } catch {}
          await accByText.click({ timeout: 4000 }).catch(() => {});
          await waitSettled(popup, settleAfterAccordionMs);
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
      await waitSettled(popup, settleAfterTabMs);

      // Tasks iframe
      const wfFrame = await getWorkflowTasksFrame(popup, 15000, 200);
      if (!wfFrame) throw new Error('Workflow Tasks iframe not found/loaded (accTab_5_crro_nb_int_iframe_0)');

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

/** ===== Detail popup: robust header/main frame resolution ===== **/

async function findHeaderFrame(detailPopup) {
  // Try by common name first
  const named = detailPopup.frame({ name: 'cai_header' });
  if (named) return named;

  // Otherwise, search for a frame that actually contains Edit/Save buttons
  for (const f of detailPopup.frames()) {
    const hasBtn = await f.locator('a.button:has(span:has-text("Edit")), a.button:has(span:has-text("Save")), a:has-text("Edit"), a:has-text("Save")').first().count();
    if (hasBtn) return f;
  }
  // Fallback
  return detailPopup.mainFrame();
}

async function findMainFrame(detailPopup) {
  // Try by common name first
  const named = detailPopup.frame({ name: 'cai_main' });
  if (named) return named;

  // Otherwise find frame that contains fields
  for (const f of detailPopup.frames()) {
    const hasFields = await f.locator('input[name="assignee_combo_name"], select[name="SET.status"]').first().count();
    if (hasFields) return f;
  }
  // Fallback
  return detailPopup.mainFrame();
}

async function clickHeaderButtonByText(headerFrame, label, timeout = 6000) {
  // Use visible text in the <span> (ids like imgBtn0 repeat)
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

async function updateTaskDetail(detailPopup, base, taskText) {
  const headerFrame = await findHeaderFrame(detailPopup);
  const mainFrame   = await findMainFrame(detailPopup);

  // BEFORE dump
  try {
    const htmlBefore = await contentSafe(detailPopup);
    save(`${base}_Task_${taskText}_Before.html`, htmlBefore);
    await screenshotSafe(detailPopup, `${base}_Task_${taskText}_Before.png`, true);
  } catch {}

  // 1) Click Edit (no new popup; page updates in-place)
  try {
    await clickHeaderButtonByText(headerFrame, 'Edit', 8000);
    await waitSettled(detailPopup, 1200);
    // Wait until fields are present/ready
    await mainFrame.waitForSelector('input[name="assignee_combo_name"]', { timeout: 6000 });
    await mainFrame.waitForSelector('select[name="SET.status"]', { timeout: 6000 });
  } catch (e) {
    // Might already be in edit mode—continue
    // console.warn(`[Task ${taskText}] Edit not clicked: ${e.message}`);
  }

  // 2) Set assignee
  try {
    const assignee = mainFrame.locator('input[name="assignee_combo_name"]');
    if (await assignee.count()) {
      await assignee.click({ timeout: 2000 }).catch(() => {});
      await assignee.fill(TARGET_ASSIGNEE, { timeout: 3000 });
      // Trigger CA SDM auto-fill (detailAutofill) and validate
      await assignee.press('Enter').catch(() => {});
      await assignee.evaluate(el => el.blur()).catch(() => {});
      await waitSettled(detailPopup, 500);
    } else {
      console.warn(`[Task ${taskText}] Assignee field not found in main frame.`);
    }
  } catch (e) {
    console.warn(`[Task ${taskText}] Assignee set error: ${e.message}`);
  }

  // 3) Set status Pending -> Complete (value='COMP' or label contains 'Complete')
  try {
    const statusSel = 'select[name="SET.status"]';
    const status = mainFrame.locator(statusSel);
    if (await status.count()) {
      const currVal = await status.inputValue().catch(() => '');
      const currLabel = await status.evaluate(s => s.options[s.selectedIndex]?.text || '').catch(() => '');
      const isPend = (s) => (s || '').toUpperCase().includes('PEND') || (s || '').toLowerCase().includes('pending');

      if (isPend(currVal) || isPend(currLabel)) {
        const ok = await selectOptionSmart(mainFrame, statusSel, 'COMP', 'Complete');
        if (!ok) console.warn(`[Task ${taskText}] Could not set status to Complete`);
        await waitSettled(detailPopup, 300);
      }
    } else {
      console.warn(`[Task ${taskText}] Status select not found in main frame.`);
    }
  } catch (e) {
    console.warn(`[Task ${taskText}] Status set error: ${e.message}`);
  }

  // 4) Click Save (header frame)
  try {
    await clickHeaderButtonByText(headerFrame, 'Save', 8000);
    await waitSettled(detailPopup, 1500);
  } catch (e) {
    console.warn(`[Task ${taskText}] Save click error: ${e.message}`);
  }

  // AFTER dump
  try {
    const htmlAfter = await contentSafe(detailPopup);
    save(`${base}_Task_${taskText}_After.html`, htmlAfter);
    await screenshotSafe(detailPopup, `${base}_Task_${taskText}_After.png`, true);
  } catch {}

  // Keep popup open (or close if you prefer):
  // await detailPopup.close().catch(() => {});
}

// ---- Main flow ----

(async () => {
  const searchText = (process.argv[2] || '').trim();
  const ticketType = (process.argv[3] || '').trim();
  const taskArgs = process.argv.slice(4);
  const tasksToOpen = taskArgs.length ? taskArgs : ['200', '250'];

  if (!searchText) {
    console.error('Usage: node ttc-workflow-tasks.js <SEARCH_TEXT> [ticketType] [task1] [task2] [...]');
    process.exit(1);
  }

  const stamp = new Date().toISOString().replace(/[:.]/g, '-');
  const base = path.join(process.cwd(), `OPEN_${searchText}_${stamp}`);

  const browser = await chromium.launch({ headless: false });
  const context = await browser.newContext({
    ignoreHTTPSErrors: true,
    viewport: { width: 1400, height: 900 },
  });
  const page = await context.newPage();

  // 1) Navigate to SDM
  console.log(`[NAV] ${BASE_URL}`);
  await page.goto(BASE_URL, { waitUntil: 'domcontentloaded' }).catch(()=>{});
  try { await page.waitForLoadState('networkidle', { timeout: 12000 }); } catch {}
  await waitSettled(page, 2000);

  // 2) Find search UI and open result popup
  const searchFrame = await findFrameWithSelectors(page, [
    'input[name="searchKey"]',
    'a#imgBtn0, a[name="imgBtn0"]'
  ], 30000, 200);

  if (!searchFrame) {
    console.error('❌ Search UI not found in any frame.');
    await browser.close(); process.exit(2);
  }
  console.log(`✅ Search UI found in frame: "${searchFrame.name() || '(unnamed)'}"`);

  await searchFrame.fill('input[name="searchKey"]', '');
  await searchFrame.fill('input[name="searchKey"]', searchText);

  if (ticketType) {
    const sel = await searchFrame.$('#ticket_type');
    if (sel) { await searchFrame.selectOption('#ticket_type', ticketType).catch(()=>{}); }
  }

  console.log('[CLICK] Go (imgBtn0)');
  const popup = await runAndCatchPopup(page, context, async () => {
    try { await searchFrame.click('a#imgBtn0, a[name="imgBtn0"]', { timeout: 5000 }); } catch {}
  }, 15000);

  if (!popup) {
    console.error('❌ No popup detected after Go.');
    await browser.close(); process.exit(3);
  }

  console.log('✅ Popup detected. Stabilizing...');
  try { await popup.bringToFront?.(); } catch {}
  await waitSettled(popup, 3000);

  // 3) Open accordion + Workflow Tasks tab, wait for the tasks iframe
  const wfFrame = await openWorkflowTasksTab(popup, {
    settleBeforeMs: 6000,
    settleAfterAccordionMs: 1500,
    settleAfterTabMs: 1500,
    attempts: 3
  });

  // 4) For each requested task, click then EDIT/ASSIGN/STATUS/SAVE in the detail popup
  async function openTaskAndProcess(taskText) {
    console.log(`\n=== TASK ${taskText}: locating anchor in workflow iframe ===`);

    try {
      await wfFrame.waitForSelector('a[href^="javascript:do_default("], a.record, tr.jqgrow td:first-child a', { timeout: 4000 });
    } catch {}

    let anchor = null;
    for (let i = 0; i < 10; i++) {
      anchor = await findTaskAnchorAcrossFrames(popup, taskText, wfFrame);
      if (anchor) break;
      await sleep(250);
    }

    if (!anchor) {
      console.warn('  -> Anchor not found by text/id; trying do_default(n) fallback...');
      let invoked = false;
      if (wfFrame) invoked = await invokeDoDefaultForTask(wfFrame, taskText);
      if (!invoked) {
        for (const f of popup.frames()) {
          invoked = await invokeDoDefaultForTask(f, taskText);
          if (invoked) { anchor = { frame: f, locator: null }; break; }
        }
      }
      if (!invoked) {
        console.warn(`⚠️ Task ${taskText}: could not locate anchor/link in any frame.`);
        return;
      }
    }

    const detailPopup = await runAndCatchPopup(popup, context, async () => {
      if (anchor.locator) {
        try { await anchor.locator.scrollIntoViewIfNeeded({ timeout: 1000 }); } catch {}
        try { await anchor.locator.click({ timeout: 4000 }); } catch {}
      }
    }, 12000);

    if (detailPopup) {
      await waitSettled(detailPopup, 2500);

      // Best-effort title check (non-fatal)
      try {
        const mf = await findMainFrame(detailPopup);
        const ok = await mf.locator('h2:has-text("Workflow Detail"), h2:has-text("Request/Incident/Problem Workflow Detail")').first().count();
        if (!ok) console.warn(`⚠️ Task ${taskText}: Expected title not found (continuing).`);
      } catch {}

      // Perform Edit -> Assignee -> Status -> Save
      await updateTaskDetail(detailPopup, base, taskText);

      console.log(`✅ Task ${taskText} processed (Edit/Assign/Status/Save).`);
    } else {
      console.warn(`  -> No new popup detected after clicking ${taskText}. Stopping here.`);
    }
  }

  for (const t of tasksToOpen) {
    await openTaskAndProcess(String(t));
  }

  console.log(`\n[DONE] Base: ${base}\n`);
  await browser.close();
})();




