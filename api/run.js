// api/run.js
import { chromium } from 'playwright';

// ---- Config ----
const BASE_URL = 'http://servicedesk-web.int.ttc.ca/CAisd/pdmweb.exe';
const TARGET_ASSIGNEE = 'Couto, Lucas';
const DEFAULT_TICKET_TYPE = 'go_cr';
const DEFAULT_TASKS = ['200', '250'];

// ---- Small utilities ----
const sleep = (ms) => new Promise(r => setTimeout(r, ms));
async function waitSettled(pageLike, timeoutMs = 12000) {
  const start = Date.now();
  try { await pageLike.waitForLoadState?.('domcontentloaded', { timeout: Math.min(4000, timeoutMs) }); } catch {}
  try { await pageLike.waitForLoadState?.('load', { timeout: Math.min(4000, timeoutMs) }); } catch {}
  while (Date.now() - start < timeoutMs) {
    try { if (pageLike.evaluate) { await pageLike.evaluate(() => document.readyState); } break; }
    catch { await sleep(120); }
  }
}

async function findFrameWithSelectors(pageOrPopup, selectors, timeoutMs = 20000, pollMs = 150) {
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

async function runAndCatchPopup(pageOrPopup, context, actionFn, timeout = 12000) {
  const ownerPage = pageOrPopup.page ? pageOrPopup.page() : pageOrPopup;
  const pagesBefore = context.pages().length;
  const pagePopupP = ownerPage.waitForEvent('popup', { timeout }).catch(() => null);
  const ctxPageP = context.waitForEvent('page', { timeout }).catch(() => null);
  await actionFn();
  let popup = await Promise.race([ pagePopupP, ctxPageP, sleep(800).then(() => null) ]);
  if (!popup) {
    await sleep(300);
    const pagesAfter = context.pages();
    if (pagesAfter.length > pagesBefore) {
      popup = pagesAfter[pagesAfter.length - 1];
      if (popup === ownerPage) popup = null;
    }
  }
  if (popup) {
    try { await popup.waitForLoadState('domcontentloaded', { timeout: 6000 }); } catch {}
    await waitSettled(popup, 2000);
  }
  return popup;
}

async function getWorkflowTasksFrame(popup, timeoutMs = 12000, pollMs = 150) {
  const deadline = Date.now() + timeoutMs;
  while (Date.now() < deadline) {
    const byName = popup.frame({ name: 'accTab_5_crro_nb_int_iframe_0' });
    if (byName) {
      try { await byName.waitForLoadState?.('domcontentloaded', { timeout: 2500 }); } catch {}
      return byName;
    }
    const frames = popup.frames();
    const match = frames.find(f => (f.url() || '').includes('FACTORY=cr_wf'));
    if (match) {
      try { await match.waitForLoadState?.('domcontentloaded', { timeout: 2500 }); } catch {}
      return match;
    }
    await popup.waitForTimeout(pollMs);
  }
  return null;
}

async function openWorkflowTasksTab(popup, {
  settleBeforeMs = 5000,
  settleAfterAccordionMs = 1000,
  settleAfterTabMs = 1000,
  attempts = 3
} = {}) {
  await waitSettled(popup, settleBeforeMs);
  let lastErr;
  for (let i = 1; i <= attempts; i++) {
    try {
      const frame = (await findFrameWithSelectors(popup, ['#accrdnHyprlnk1'], 6000, 120))
                 || (await findFrameWithSelectors(popup, ['#tabHyprink1_5'], 6000, 120))
                 || popup.mainFrame();

      // Accordion: "Additional Information"
      const accordion = frame.locator('h2#accrdnHyprlnk1, #accrdnHyprlnk1').first();
      if (await accordion.count()) {
        try { await accordion.scrollIntoViewIfNeeded({ timeout: 1500 }); } catch {}
        await accordion.click({ timeout: 3500 }).catch(() => {});
        await waitSettled(popup, settleAfterAccordionMs);
      } else {
        const accByText = frame.locator('h2:has-text("Additional Information"), a:has-text("Additional Information")').first();
        if (await accByText.count()) {
          try { await accByText.scrollIntoViewIfNeeded({ timeout: 1500 }); } catch {}
          await accByText.click({ timeout: 3500 }).catch(() => {});
          await waitSettled(popup, settleAfterAccordionMs);
        }
      }

      // Workflow Tasks tab
      const tab = frame.locator('a#tabHyprink1_5').first();
      if (await tab.count()) {
        const cls = (await tab.getAttribute('class')) || '';
        if (!cls.includes('current')) {
          try { await tab.scrollIntoViewIfNeeded({ timeout: 1500 }); } catch {}
          await tab.click({ timeout: 3500 });
        }
      } else {
        const tabByText = frame.locator('a:has-text("Workflow Tasks")').first();
        if (!await tabByText.count()) throw new Error('Workflow Tasks tab not found');
        try { await tabByText.scrollIntoViewIfNeeded({ timeout: 1500 }); } catch {}
        await tabByText.click({ timeout: 3500 });
      }
      await waitSettled(popup, settleAfterTabMs);

      const wfFrame = await getWorkflowTasksFrame(popup, 12000, 150);
      if (!wfFrame) throw new Error('Workflow Tasks iframe not found/loaded');

      try {
        await wfFrame.waitForSelector('a[href^="javascript:do_default("], a.record, tr.jqgrow td:first-child a', { timeout: 3500 });
      } catch {}
      return wfFrame;
    } catch (e) {
      lastErr = e;
      await waitSettled(popup, 700);
    }
  }
  throw lastErr || new Error('Failed to open Workflow Tasks tab');
}

async function findTaskAnchorAcrossFrames(popup, taskText, preferredFrame = null) {
  const exactText = new RegExp(`^\\s*${taskText}\\s*$`);
  const tryInFrame = async (f) => {
    const idGuess = (taskText === '200') ? '#rslnk_0_0' : (taskText === '250') ? '#rslnk_1_0' : null;
    if (idGuess) {
      const byKnownId = f.locator(idGuess);
      if (await byKnownId.count()) return { frame: f, locator: byKnownId };
    }
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
    if (wfFrame) {
      const hit = await tryInFrame(wfFrame);
      if (hit) return hit;
    }
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
  await frame.evaluate((row) => {
    if (typeof window.do_default === 'function') window.do_default(row);
  }, n);
  return true;
}

async function findHeaderFrame(detailPopup) {
  const named = detailPopup.frame({ name: 'cai_header' });
  if (named) return named;
  for (const f of detailPopup.frames()) {
    const hasBtn = await f.locator(
      'a.button:has(span:has-text("Edit")), a.button:has(span:has-text("Save")), a:has-text("Edit"), a:has-text("Save")'
    ).first().count();
    if (hasBtn) return f;
  }
  return detailPopup.mainFrame();
}

async function findMainFrame(detailPopup) {
  const named = detailPopup.frame({ name: 'cai_main' });
  if (named) return named;
  for (const f of detailPopup.frames()) {
    const hasFields = await f.locator(
      'input[name="assignee_combo_name"], select[name="SET.status"]'
    ).first().count();
    if (hasFields) return f;
  }
  return detailPopup.mainFrame();
}

async function clickHeaderButtonByText(headerFrame, label, timeout = 6000) {
  const btn = headerFrame.locator(`a.button:has(span:has-text("${label}")), a:has-text("${label}")`).first();
  if (!await btn.count()) throw new Error(`${label} button not found in header frame`);
  try { await btn.scrollIntoViewIfNeeded({ timeout: 1200 }); } catch {}
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
  const headerFrame = await findHeaderFrame(detailPopup);
  const mainFrame = await findMainFrame(detailPopup);

  // 1) Click Edit
  try {
    await clickHeaderButtonByText(headerFrame, 'Edit', 7000);
    await waitSettled(detailPopup, 1000);
    await mainFrame.waitForSelector('input[name="assignee_combo_name"]', { timeout: 5000 });
    await mainFrame.waitForSelector('select[name="SET.status"]', { timeout: 5000 });
  } catch (e) {
    // Might already be in edit mode
  }

  // 2) Set assignee
  try {
    const assignee = mainFrame.locator('input[name="assignee_combo_name"]');
    if (await assignee.count()) {
      await assignee.click({ timeout: 2000 }).catch(() => {});
      await assignee.fill(TARGET_ASSIGNEE, { timeout: 3000 });
      await assignee.press('Enter').catch(() => {});
      await assignee.evaluate(el => el.blur()).catch(() => {});
      await waitSettled(detailPopup, 400);
    }
  } catch (e) {
    // non-fatal
  }

  // 3) Status: Pending -> Complete
  try {
    const statusSel = 'select[name="SET.status"]';
    const status = mainFrame.locator(statusSel);
    if (await status.count()) {
      const currVal = await status.inputValue().catch(() => '');
      const currLabel = await status.evaluate(s => s.options[s.selectedIndex]?.text || '').catch(() => '');
      const isPend = (s) => (s || '').toUpperCase().includes('PEND') || (s || '').toLowerCase().includes('pending');
      if (isPend(currVal) || isPend(currLabel)) {
        const ok = await selectOptionSmart(mainFrame, statusSel, 'COMP', 'Complete');
        if (!ok) { /* leave as-is if cannot change */ }
        await waitSettled(detailPopup, 250);
      }
    }
  } catch (e) {
    // non-fatal
  }

  // 4) Click Save
  try {
    await clickHeaderButtonByText(headerFrame, 'Save', 7000);
    await waitSettled(detailPopup, 1200);
  } catch (e) {
    // non-fatal
  }
}

async function openTaskAndProcess(popup, context, wfFrame, taskText) {
  // Try to locate anchor by text/id, else use do_default fallback
  let anchor = null;
  try {
    await wfFrame.waitForSelector('a[href^="javascript:do_default("], a.record, tr.jqgrow td:first-child a', { timeout: 3500 });
  } catch {}

  for (let i = 0; i < 10; i++) {
    anchor = await findTaskAnchorAcrossFrames(popup, taskText, wfFrame);
    if (anchor) break;
    await sleep(200);
  }

  if (!anchor) {
    let invoked = false;
    if (wfFrame) invoked = await invokeDoDefaultForTask(wfFrame, taskText);
    if (!invoked) {
      for (const f of popup.frames()) {
        invoked = await invokeDoDefaultForTask(f, taskText);
        if (invoked) { anchor = { frame: f, locator: null }; break; }
      }
    }
    if (!invoked) return { task: taskText, status: 'anchor_not_found' };
  }

  const detailPopup = await runAndCatchPopup(popup, context, async () => {
    if (anchor?.locator) {
      try { await anchor.locator.scrollIntoViewIfNeeded({ timeout: 800 }); } catch {}
      try { await anchor.locator.click({ timeout: 3500 }); } catch {}
    }
  }, 10000);

  if (!detailPopup) return { task: taskText, status: 'detail_popup_missing' };

  await waitSettled(detailPopup, 2000);
  await updateTaskDetail(detailPopup, taskText);
  return { task: taskText, status: 'processed' };
}

// ---- Vercel handler ----
export default async function handler(req, res) {
  if (req.method !== 'POST') {
    return res.status(405).json({ error: 'Method not allowed' });
  }

  const { requestNumber, ticketType = DEFAULT_TICKET_TYPE, tasks = DEFAULT_TASKS } = req.body || {};
  if (!requestNumber || String(requestNumber).trim() === '') {
    return res.status(400).json({ error: 'Request number required' });
  }

  const taskList = Array.isArray(tasks) ? tasks.map(String) : DEFAULT_TASKS;

  let browser;
  try {
    browser = await chromium.launch({ headless: true });
    const context = await browser.newContext({
      ignoreHTTPSErrors: true,
      viewport: { width: 1400, height: 900 },
    });
    const page = await context.newPage();

    // 1) Navigate to SDM
    await page.goto(BASE_URL, { waitUntil: 'domcontentloaded' }).catch(()=>{});
    try { await page.waitForLoadState('networkidle', { timeout: 10000 }); } catch {}
    await waitSettled(page, 1500);

    // 2) Find search UI
    const searchFrame = await findFrameWithSelectors(page, [
      'input[name="searchKey"]',
      'a#imgBtn0, a[name="imgBtn0"]'
    ], 20000, 150);

    if (!searchFrame) {
      await browser?.close();
      return res.status(500).json({ error: 'Search UI not found in any frame.' });
    }

    // 3) Fill search and ticket type; click Go
    await searchFrame.fill('input[name="searchKey"]', String(requestNumber));

    const sel = await searchFrame.$('#ticket_type');
    if (sel) { await searchFrame.selectOption('#ticket_type', String(ticketType)).catch(()=>{}); }

    const popup = await runAndCatchPopup(page, context, async () => {
      try { await searchFrame.click('a#imgBtn0, a[name="imgBtn0"]', { timeout: 4000 }); } catch {}
    }, 12000);

    if (!popup) {
      await browser?.close();
      return res.status(500).json({ error: 'No popup detected after Go.' });
    }

    try { await popup.bringToFront?.(); } catch {}
    await waitSettled(popup, 2500);

    // 4) Open Workflow Tasks tab
    const wfFrame = await openWorkflowTasksTab(popup, {
      settleBeforeMs: 5000,
      settleAfterAccordionMs: 1000,
      settleAfterTabMs: 1000,
      attempts: 3
    });

    // 5) Process each task
    const results = [];
    for (const t of taskList) {
      const r = await openTaskAndProcess(popup, context, wfFrame, String(t));
      results.push(r);
    }

    await browser?.close();
    return res.status(200).json({
      success: true,
      requestNumber: String(requestNumber),
      ticketType: String(ticketType),
      results
    });
  } catch (err) {
    try { await browser?.close(); } catch {}
    return res.status(500).json({ error: err.message || 'Unexpected error' });
  }
}
