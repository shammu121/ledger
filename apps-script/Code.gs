/**
 * LEFT TRACKER — expense tracker + to-do list
 * Google Sheet = database · this script = private API · hosted page + iPhone Shortcuts = app
 *
 * 1. Change SECRET below to your own word (letters/numbers only).
 * 2. Run setup() once (creates the sheets, lists and default categories).
 * 3. Deploy → New deployment → Web app → Execute as: Me · Who has access: Anyone.
 * After editing this file later: Deploy → Manage deployments → Edit → Version: New version (keeps the same URL).
 */

const SECRET = 'change-this-key-123';   // <- CHANGE THIS
const TZ = 'Asia/Kolkata';
const CURRENCY = '₹';

const SH_ENTRIES = 'Entries';
const SH_LISTS = 'Lists';
const SH_CATS = 'Categories';
const SH_TASKS = 'Tasks';
const PRIORITIES = ['High', 'Medium', 'Low', 'Later'];

// Lists: name, emoji, budget type ('Monthly' resets every month, 'Total' = one budget for a project)
const DEFAULT_LISTS = [
  ['Personal', '🙂', 'Monthly'],
  ['Business', '💼', 'Monthly'],
  ['Office Interior', '🛠️', 'Total'],
];

// Categories: list, name, emoji, budget (0 = no limit), colour
const DEFAULT_CATEGORIES = [
  ['Personal', 'Food', '🍽️', 4000, '#FCEFC4'],
  ['Personal', 'Groceries', '🛒', 3000, '#D6F5C8'],
  ['Personal', 'Transport', '🚗', 2500, '#C4F0F0'],
  ['Personal', 'Shopping', '🛍️', 3000, '#C9D5FA'],
  ['Personal', 'Clothes', '👕', 1500, '#BFE0FF'],
  ['Personal', 'Going Out', '🍻', 2000, '#F9C9DA'],
  ['Personal', 'Subscriptions', '📺', 1000, '#DCCBFA'],
  ['Personal', 'Personal care', '🧴', 1500, '#FAD3E3'],
  ['Personal', 'Fitness & Sport', '🏋️', 1500, '#C8F2DC'],
  ['Personal', 'Bills', '🧾', 3000, '#FFD9C2'],
  ['Personal', 'Travel', '✈️', 0, '#D3E4FF'],
  ['Personal', 'Other', '📦', 0, '#E3E3E8'],

  ['Business', 'Ads', '📣', 0, '#F9C9DA'],
  ['Business', 'Inventory', '📦', 0, '#FCEFC4'],
  ['Business', 'Shipping', '🚚', 0, '#C4F0F0'],
  ['Business', 'Packaging', '🎁', 0, '#D6F5C8'],
  ['Business', 'Software & Tools', '💻', 0, '#DCCBFA'],
  ['Business', 'Salaries', '👥', 0, '#C9D5FA'],
  ['Business', 'Rent & Bills', '🏢', 0, '#FFD9C2'],
  ['Business', 'Travel', '✈️', 0, '#D3E4FF'],
  ['Business', 'Other', '🧾', 0, '#E3E3E8'],

  ['Office Interior', 'Furniture', '🪑', 0, '#FCEFC4'],
  ['Office Interior', 'Paint & Walls', '🎨', 0, '#F9C9DA'],
  ['Office Interior', 'Electrical & Lights', '💡', 0, '#C4F0F0'],
  ['Office Interior', 'Labour', '👷', 0, '#C9D5FA'],
  ['Office Interior', 'Decor', '🪴', 0, '#D6F5C8'],
  ['Office Interior', 'Materials', '🧱', 0, '#FFD9C2'],
  ['Office Interior', 'Other', '📦', 0, '#E3E3E8'],
];

// Words that point to a category when there is no history yet (matched by category name in any list).
const KEYWORDS = {
  'Food': ['chai', 'tea', 'coffee', 'lunch', 'dinner', 'breakfast', 'snack', 'biryani', 'swiggy', 'zomato', 'juice', 'restaurant', 'shawarma', 'pizza', 'burger'],
  'Groceries': ['milk', 'grocery', 'groceries', 'vegetable', 'veggies', 'fruit', 'rice', 'egg', 'bread', 'supermarket', 'blinkit', 'zepto', 'instamart', 'bigbasket'],
  'Transport': ['petrol', 'diesel', 'fuel', 'uber', 'ola', 'rapido', 'auto', 'taxi', 'bus', 'train', 'metro', 'parking', 'toll', 'fastag'],
  'Shopping': ['amazon', 'flipkart', 'meesho', 'shopping', 'gift'],
  'Clothes': ['shirt', 'jeans', 'tshirt', 't-shirt', 'shoes', 'dress', 'clothes'],
  'Going Out': ['movie', 'cinema', 'party', 'outing', 'bowling', 'concert'],
  'Subscriptions': ['netflix', 'spotify', 'prime', 'hotstar', 'youtube', 'icloud', 'chatgpt', 'claude', 'subscription', 'recharge'],
  'Personal care': ['salon', 'haircut', 'dental', 'doctor', 'medicine', 'pharmacy', 'skincare', 'perfume'],
  'Fitness & Sport': ['gym', 'protein', 'turf', 'football', 'badminton', 'yoga'],
  'Bills': ['electricity', 'kseb', 'wifi', 'internet', 'water', 'gas', 'emi', 'insurance'],
  'Ads': ['meta', 'facebook', 'instagram', 'google ads', 'ads', 'boost', 'influencer'],
  'Shipping': ['courier', 'shiprocket', 'delhivery', 'dtdc', 'shipping', 'freight'],
  'Packaging': ['box', 'boxes', 'tape', 'packaging', 'bubble', 'label'],
  'Software & Tools': ['shopify', 'canva', 'domain', 'hosting', 'app', 'software'],
  'Salaries': ['salary', 'wages', 'staff'],
  'Rent & Bills': ['rent', 'office rent', 'electricity'],
  'Furniture': ['chair', 'table', 'desk', 'sofa', 'shelf', 'cupboard', 'rack'],
  'Paint & Walls': ['paint', 'putty', 'wallpaper', 'mural', 'primer'],
  'Electrical & Lights': ['light', 'led', 'wire', 'switch', 'fan', 'socket', 'electrician'],
  'Labour': ['labour', 'labor', 'carpenter', 'painter', 'mason', 'worker'],
  'Decor': ['plant', 'frame', 'decor', 'rug', 'curtain'],
  'Materials': ['plywood', 'cement', 'tiles', 'glass', 'screws', 'hardware'],
  'Travel': ['flight', 'trip', 'stay', 'airbnb', 'resort', 'hotel'],
};

/* ───────────────────────── SETUP ───────────────────────── */

function setup() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const e = ss.getSheetByName(SH_ENTRIES) || ss.insertSheet(SH_ENTRIES);
  if (e.getLastRow() === 0) {
    e.appendRow(['ID', 'Date', 'Amount', 'List', 'Category', 'Note', 'Created']);
    e.setFrozenRows(1);
    e.getRange('B:B').setNumberFormat('@');
    e.getRange('C:C').setNumberFormat('#,##0.00');
  }
  const l = ss.getSheetByName(SH_LISTS) || ss.insertSheet(SH_LISTS);
  if (l.getLastRow() === 0) {
    l.appendRow(['Name', 'Emoji', 'Budget type', 'Total budget (for Total lists)']);
    l.setFrozenRows(1);
    DEFAULT_LISTS.forEach(r => l.appendRow(r.concat([0])));
  }
  const c = ss.getSheetByName(SH_CATS) || ss.insertSheet(SH_CATS);
  if (c.getLastRow() === 0) {
    c.appendRow(['List', 'Name', 'Emoji', 'Budget', 'Colour']);
    c.setFrozenRows(1);
    DEFAULT_CATEGORIES.forEach(r => c.appendRow(r));
  }
  const t = ss.getSheetByName(SH_TASKS) || ss.insertSheet(SH_TASKS);
  if (t.getLastRow() === 0) {
    t.appendRow(['ID', 'Title', 'Priority', 'Due', 'Note', 'Done', 'Created', 'Completed']);
    t.setFrozenRows(1);
    t.getRange('D:D').setNumberFormat('@');
  }
  const s1 = ss.getSheetByName('Sheet1');
  if (s1 && s1.getLastRow() === 0 && ss.getSheets().length > 1) ss.deleteSheet(s1);
  Logger.log('Setup done. Now deploy as a Web app.');
}

/* ───────────────────────── WEB ENTRY POINTS ───────────────────────── */

function doGet(e) {
  const p = (e && e.parameter) || {};
  const action = p.action || '';
  if (SECRET === 'change-this-key-123') return json_({ ok: false, error: 'Set your own SECRET in Code.gs first' });
  if (p.key !== SECRET) return p.fmt ? text_('Error: wrong key') : json_({ ok: false, error: 'Wrong key' });

  if (action === 'data') return json_(Object.assign({ ok: true }, getData_()));
  // fmt=1 → Shortcut-friendly: lists come back as plain arrays, results as plain text
  const f = !!p.fmt;
  if (action === 'lists') { const r = listsForShortcut_(); return json_(f ? r.lists : r); }
  if (action === 'categories') { const r = categoriesForShortcut_(p.list || '', p.note || ''); return json_(f ? r.categories : r); }
  if (action === 'add') { const r = addFromShortcut_(p); return f ? text_(r.ok ? r.message : 'Error: ' + r.error + '\nReceived → list: ' + (p.list || '(empty)') + ' | note: ' + (p.note || '(empty)') + ' | amount: ' + (p.amount || '(empty)') + ' | category: ' + (p.category || '(empty)') + '\nRaw: ' + (e.queryString || '').replace(/key=[^&]*/, 'key=***').slice(0, 300)) : json_(r); }
  if (action === 'addtask') { const r = addTaskFromShortcut_(p); return f ? text_(r.ok ? r.message : 'Error: ' + r.error) : json_(r); }
  if (action === 'summary') return json_(summary_(resolveList_(p.list, readLists_())));
  return json_({ ok: true, app: 'left-tracker' });
}

// The app page sends { key, op, args } — Shortcuts send { key, list, amount, ... } or { key, type:'task', ... }
const OPS = {
  saveEntry: apiSaveEntry, deleteEntry: apiDeleteEntry, saveSettings: apiSaveSettings,
  saveTask: apiSaveTask, setTaskDone: apiSetTaskDone, deleteTask: apiDeleteTask, clearDone: apiClearDone,
  getData: apiGetData,
};

function doPost(e) {
  let b = {};
  try { b = JSON.parse(e.postData.contents); } catch (err) { b = (e && e.parameter) || {}; }
  if (SECRET === 'change-this-key-123') return json_({ ok: false, error: 'Set your own SECRET in Code.gs first' });
  if (b.key !== SECRET) return json_({ ok: false, error: 'Wrong key' });
  try {
    if (b.op) {
      const fn = OPS[b.op];
      if (!fn) return json_({ ok: false, error: 'Unknown op' });
      return json_(fn.apply(null, [b.key].concat(b.args || [])));
    }
    if (b.type === 'task' || b.action === 'addtask') return json_(addTaskFromShortcut_(b));
    return json_(addFromShortcut_(b));
  } catch (err) {
    return json_({ ok: false, error: String(err && err.message || err) });
  }
}

/* ───────────────────────── SHORTCUT API ───────────────────────── */

// Step 1 of the Shortcut: which list?
function listsForShortcut_() {
  return { ok: true, lists: readLists_().map(l => nb_(l.emoji + ' ' + l.name)) };
}

// Step 4 of the Shortcut: categories of that list, best guess for the note first.
function categoriesForShortcut_(listLabel, note) {
  const lists = readLists_();
  const list = resolveList_(listLabel, lists);
  const cats = readCategories_().filter(c => c.list === list);
  const guess = note ? guessCategory_(note, list, readEntries_(), cats) : '';
  const labels = cats.map(c => nb_(c.emoji + ' ' + c.name));
  if (guess) {
    const i = cats.findIndex(c => c.name === guess);
    if (i > 0) { const g = labels.splice(i, 1)[0]; labels.unshift(g); }
  }
  return { ok: true, list, categories: labels, suggested: guess };
}

function addFromShortcut_(p) {
  const res = saveEntry_({ amount: p.amount, list: p.list, category: p.category, note: p.note, date: p.date });
  if (!res.ok) return res;
  const e = res.entry;
  const s = summary_(e.list);
  s.message = CURRENCY + fmt_(e.amount) + ' · ' + (e.note || e.category) + ' → ' + e.list + ' / ' + e.category + '\n' + s.message;
  s.entry = e;
  return s;
}

function summary_(listName) {
  const lists = readLists_();
  const list = lists.find(l => l.name === listName) || lists[0];
  const cats = readCategories_().filter(c => c.list === list.name);
  const entries = readEntries_().filter(x => x.list === list.name);
  const now = new Date();
  const ym = Utilities.formatDate(now, TZ, 'yyyy-MM');
  const today = Utilities.formatDate(now, TZ, 'yyyy-MM-dd');

  if (list.type === 'Total') {
    const spent = entries.reduce((a, x) => a + x.amount, 0);
    const budget = list.total || cats.reduce((a, c) => a + (c.budget || 0), 0);
    const message = budget
      ? CURRENCY + fmt_(spent) + ' spent of ' + CURRENCY + fmt_(budget) + (spent > budget ? ' · ' + CURRENCY + fmt_(spent - budget) + ' over' : ' · ' + CURRENCY + fmt_(budget - spent) + ' left')
      : CURRENCY + fmt_(spent) + ' spent so far on ' + list.name;
    return { ok: true, list: list.name, spent, budget, left: budget - spent, message };
  }

  const y = +ym.slice(0, 4), m = +ym.slice(5, 7), d = +today.slice(8, 10);
  const daysLeft = new Date(y, m, 0).getDate() - d + 1;
  const budget = cats.reduce((a, c) => a + (c.budget || 0), 0);
  let spent = 0, spentToday = 0;
  entries.forEach(x => {
    if (x.date.slice(0, 7) === ym) spent += x.amount;
    if (x.date === today) spentToday += x.amount;
  });
  const left = budget - spent;
  let message;
  if (!budget) message = CURRENCY + fmt_(spent) + ' spent this month on ' + list.name;
  else if (left >= 0) message = CURRENCY + fmt_(left) + ' left this month · ' + CURRENCY + fmt_(Math.floor(left / daysLeft)) + '/day';
  else message = CURRENCY + fmt_(-left) + ' over budget this month';
  return { ok: true, list: list.name, spent, budget, left, spentToday, daysLeft, message };
}

// Shortcut 2: add a task. Body: { key, type:'task', title, priority, due? }
function addTaskFromShortcut_(p) {
  const res = saveTask_({ title: p.title || p.task || p.note, priority: p.priority, due: p.due, note: p.note && p.title ? p.note : '' });
  if (!res.ok) return res;
  const open = readTasks_().filter(t => !t.done);
  const today = Utilities.formatDate(new Date(), TZ, 'yyyy-MM-dd');
  const dueToday = open.filter(t => t.due && t.due <= today).length;
  const icon = { High: '🔴', Medium: '🟠', Low: '🔵', Later: '⚪' }[res.task.priority];
  return {
    ok: true, task: res.task, open: open.length,
    message: '✓ ' + res.task.title + '\n' + icon + ' ' + res.task.priority + ' · ' + open.length + ' open' + (dueToday ? ' · ' + dueToday + ' due today' : ''),
  };
}

/* ───────────────────────── PAGE API (google.script.run) ───────────────────────── */

function apiSaveTask(key, task) { check_(key); return saveTask_(task); }

function apiSetTaskDone(key, id, done) {
  check_(key);
  return withLock_(() => {
    const sh = sheet_(SH_TASKS);
    const row = findRow_(sh, id);
    if (row < 0) return { ok: false, error: 'Not found' };
    sh.getRange(row, 6, 1, 1).setValue(done ? true : false);
    sh.getRange(row, 8, 1, 1).setValue(done ? new Date() : '');
    return { ok: true };
  });
}

function apiDeleteTask(key, id) {
  check_(key);
  return withLock_(() => {
    const sh = sheet_(SH_TASKS);
    const row = findRow_(sh, id);
    if (row > 0) sh.deleteRow(row);
    return { ok: true };
  });
}

function apiClearDone(key) {
  check_(key);
  return withLock_(() => {
    const sh = sheet_(SH_TASKS);
    const n = sh.getLastRow() - 1;
    if (n < 1) return { ok: true, cleared: 0 };
    const done = sh.getRange(2, 6, n, 1).getValues();
    let cleared = 0;
    for (let i = done.length - 1; i >= 0; i--) if (done[i][0] === true || done[i][0] === 'TRUE') { sh.deleteRow(i + 2); cleared++; }
    return { ok: true, cleared };
  });
}


function apiGetData(key) { check_(key); return Object.assign({ ok: true }, getData_()); }

function apiSaveEntry(key, entry) { check_(key); return saveEntry_(entry); }

function apiDeleteEntry(key, id) {
  check_(key);
  return withLock_(() => {
    const sh = sheet_(SH_ENTRIES);
    const row = findRow_(sh, id);
    if (row < 0) return { ok: false, error: 'Not found' };
    sh.deleteRow(row);
    return { ok: true };
  });
}

// settings = { lists:[{name, emoji, type, total, oldName}], categories:[{list, name, emoji, budget, color, oldName}] }
function apiSaveSettings(key, settings) {
  check_(key);
  return withLock_(() => {
    const lists = (settings.lists || []).filter(l => l && String(l.name || '').trim());
    if (!lists.length) return { ok: false, error: 'Keep at least one list' };
    const listRenames = {};
    lists.forEach(l => { l.name = String(l.name).trim(); if (l.oldName && l.oldName !== l.name) listRenames[l.oldName] = l.name; });
    const names = lists.map(l => l.name);

    const cats = (settings.categories || [])
      .filter(c => c && String(c.name || '').trim())
      .map(c => ({ ...c, list: listRenames[c.list] || c.list, name: String(c.name).trim() }))
      .filter(c => names.includes(c.list));

    writeRows_(sheet_(SH_LISTS), 4, lists.map(l => [l.name, String(l.emoji || '📁'), l.type === 'Total' ? 'Total' : 'Monthly', Math.max(0, Number(l.total) || 0)]));
    writeRows_(sheet_(SH_CATS), 5, cats.map(c => [c.list, c.name, String(c.emoji || '📦'), Math.max(0, Number(c.budget) || 0), c.color || '#E3E3E8']));

    // Carry renames into old entries
    const catRenames = [];
    (settings.categories || []).forEach(c => {
      if (c && c.oldName && c.name && c.oldName !== String(c.name).trim()) catRenames.push([listRenames[c.list] || c.list, c.oldName, String(c.name).trim()]);
    });
    if (Object.keys(listRenames).length || catRenames.length) renameInEntries_(listRenames, catRenames);
    return { ok: true, lists: readLists_(), categories: readCategories_(), entries: readEntries_() };
  });
}

/* ───────────────────────── CORE ───────────────────────── */

function getData_() {
  return {
    entries: readEntries_(),
    lists: readLists_(),
    categories: readCategories_(),
    tasks: readTasks_(),
    currency: CURRENCY,
    today: Utilities.formatDate(new Date(), TZ, 'yyyy-MM-dd'),
  };
}

function saveEntry_(input) {
  const amount = parseAmount_(input.amount);
  if (!(amount > 0)) return { ok: false, error: 'Enter an amount' };
  const lists = readLists_();
  const list = resolveList_(input.list, lists);
  const cats = readCategories_().filter(c => c.list === list);
  const note = String(input.note || '').trim();
  let category = resolveCategory_(input.category, cats);
  if (!category) category = guessCategory_(note, list, readEntries_(), cats)
    || ((cats.find(c => c.name === 'Other') || cats[cats.length - 1] || { name: 'Other' }).name);
  const date = normDate_(input.date) || Utilities.formatDate(new Date(), TZ, 'yyyy-MM-dd');

  return withLock_(() => {
    const sh = sheet_(SH_ENTRIES);
    if (input.id) {
      const row = findRow_(sh, input.id);
      if (row > 0) {
        sh.getRange(row, 2, 1, 5).setValues([[date, amount, list, category, note]]);
        return { ok: true, entry: { id: String(input.id), date, amount, list, category, note, created: Number(input.created) || Date.now() } };
      }
    }
    const id = Utilities.getUuid().slice(0, 8);
    const created = Date.now();
    sh.appendRow([id, date, amount, list, category, note, new Date(created)]);
    return { ok: true, entry: { id, date, amount, list, category, note, created } };
  });
}

function saveTask_(input) {
  const title = String(input.title || '').trim();
  if (!title) return { ok: false, error: 'Enter a task' };
  const pv = String(input.priority || '').toLowerCase();
  const priority = PRIORITIES.find(p => pv.indexOf(p.toLowerCase()) >= 0) || 'Medium';
  const due = normDate_(input.due);
  const note = String(input.note || '').trim();
  return withLock_(() => {
    const sh = sheet_(SH_TASKS);
    if (input.id) {
      const row = findRow_(sh, input.id);
      if (row > 0) {
        sh.getRange(row, 2, 1, 4).setValues([[title, priority, due, note]]);
        return { ok: true, task: { id: String(input.id), title, priority, due, note, done: !!input.done, created: Number(input.created) || Date.now(), completed: Number(input.completed) || 0 } };
      }
    }
    const id = Utilities.getUuid().slice(0, 8);
    const created = Date.now();
    sh.appendRow([id, title, priority, due, note, false, new Date(created), '']);
    return { ok: true, task: { id, title, priority, due, note, done: false, created, completed: 0 } };
  });
}

function readTasks_() {
  const sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SH_TASKS);
  if (!sh) return [];
  const n = sh.getLastRow() - 1;
  if (n < 1) return [];
  return sh.getRange(2, 1, n, 8).getValues()
    .filter(r => r[0] !== '' && String(r[1]).trim())
    .map(r => ({
      id: String(r[0]),
      title: String(r[1]),
      priority: PRIORITIES.indexOf(String(r[2])) >= 0 ? String(r[2]) : 'Medium',
      due: normDate_(r[3]),
      note: String(r[4] || ''),
      done: r[5] === true || r[5] === 'TRUE',
      created: r[6] instanceof Date ? r[6].getTime() : 0,
      completed: r[7] instanceof Date ? r[7].getTime() : 0,
    }));
}

function readEntries_() {
  const sh = sheet_(SH_ENTRIES);
  const n = sh.getLastRow() - 1;
  if (n < 1) return [];
  return sh.getRange(2, 1, n, 7).getValues()
    .filter(r => r[0] !== '' && r[2] !== '')
    .map(r => ({
      id: String(r[0]),
      date: normDate_(r[1]),
      amount: Number(r[2]) || 0,
      list: String(r[3] || ''),
      category: String(r[4] || 'Other'),
      note: String(r[5] || ''),
      created: r[6] instanceof Date ? r[6].getTime() : 0,
    }));
}

function readLists_() {
  const sh = sheet_(SH_LISTS);
  const n = sh.getLastRow() - 1;
  const rows = n < 1 ? DEFAULT_LISTS.map(r => r.concat([0])) : sh.getRange(2, 1, n, 4).getValues();
  return rows.filter(r => String(r[0]).trim())
    .map(r => ({ name: String(r[0]).trim(), emoji: String(r[1] || '📁'), type: String(r[2]) === 'Total' ? 'Total' : 'Monthly', total: Number(r[3]) || 0 }));
}

function readCategories_() {
  const sh = sheet_(SH_CATS);
  const n = sh.getLastRow() - 1;
  const rows = n < 1 ? DEFAULT_CATEGORIES : sh.getRange(2, 1, n, 5).getValues();
  return rows.filter(r => String(r[0]).trim() && String(r[1]).trim())
    .map(r => ({ list: String(r[0]).trim(), name: String(r[1]).trim(), emoji: String(r[2] || '📦'), budget: Number(r[3]) || 0, color: String(r[4] || '#E3E3E8') }));
}

// Accepts "Personal", "🙂 Personal", "personal" — falls back to the first list
function resolveList_(value, lists) {
  const v = un_(value).trim().toLowerCase();
  const hit = v && (lists.find(l => l.name.toLowerCase() === v || (l.emoji + ' ' + l.name).toLowerCase() === v)
    || lists.find(l => v.endsWith(l.name.toLowerCase())));
  return hit ? hit.name : (lists[0] ? lists[0].name : 'Personal');
}

// Accepts "Food", "🍽️ Food", "food"
function resolveCategory_(value, cats) {
  const v = un_(value).trim().toLowerCase();
  if (!v) return '';
  const hit = cats.find(c => c.name.toLowerCase() === v || (c.emoji + ' ' + c.name).toLowerCase() === v)
    || cats.find(c => v.endsWith(c.name.toLowerCase()));
  return hit ? hit.name : '';
}

// 1) same note used before in this list → that category  2) same first word  3) keyword list  4) note prefix
function guessCategory_(note, list, entries, cats) {
  const n = String(note || '').trim().toLowerCase();
  if (!n) return '';
  const names = cats.map(c => c.name);
  const sorted = entries.filter(e => e.list === list && names.includes(e.category))
    .sort((a, b) => (b.created || 0) - (a.created || 0));
  const exact = sorted.find(e => e.note.toLowerCase() === n);
  if (exact) return exact.category;
  const first = n.split(/\s+/)[0];
  if (first.length > 2) {
    const partial = sorted.find(e => e.note.toLowerCase().split(/\s+/)[0] === first);
    if (partial) return partial.category;
  }
  for (const cat in KEYWORDS) {
    if (names.includes(cat) && KEYWORDS[cat].some(k => n.includes(k))) return cat;
  }
  if (n.length >= 3) {
    const pre = sorted.find(e => e.note.toLowerCase().indexOf(n) === 0);
    if (pre) return pre.category;
  }
  return '';
}

function renameInEntries_(listRenames, catRenames) {
  const sh = sheet_(SH_ENTRIES);
  const n = sh.getLastRow() - 1;
  if (n < 1) return;
  const rng = sh.getRange(2, 4, n, 2); // List, Category
  const vals = rng.getValues().map(r => {
    let list = listRenames[r[0]] || r[0], cat = r[1];
    catRenames.forEach(([l, from, to]) => { if (list === l && cat === from) cat = to; });
    return [list, cat];
  });
  rng.setValues(vals);
}

/* ───────────────────────── HELPERS ───────────────────────── */

function sheet_(name) {
  const sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(name);
  if (!sh) throw new Error('Run setup() first — sheet "' + name + '" missing');
  return sh;
}

function writeRows_(sh, width, rows) {
  if (sh.getLastRow() > 1) sh.getRange(2, 1, sh.getLastRow() - 1, width).clearContent();
  if (rows.length) sh.getRange(2, 1, rows.length, width).setValues(rows);
}

function findRow_(sh, id) {
  const n = sh.getLastRow() - 1;
  if (n < 1) return -1;
  const ids = sh.getRange(2, 1, n, 1).getValues();
  for (let i = 0; i < ids.length; i++) if (String(ids[i][0]) === String(id)) return i + 2;
  return -1;
}

function withLock_(fn) {
  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try { return fn(); } finally { lock.releaseLock(); }
}

function check_(key) { if (key !== SECRET) throw new Error('Wrong key'); }

function parseAmount_(v) { return Math.round(parseFloat(String(v == null ? '' : v).replace(/[^0-9.]/g, '')) * 100) / 100; }

function normDate_(v) {
  if (!v) return '';
  if (v instanceof Date) return Utilities.formatDate(v, TZ, 'yyyy-MM-dd');
  const s = String(v).trim();
  const m = s.match(/^(\d{4})-(\d{1,2})-(\d{1,2})/);
  if (m) return m[1] + '-' + ('0' + m[2]).slice(-2) + '-' + ('0' + m[3]).slice(-2);
  const d = new Date(s);
  return isNaN(d) ? '' : Utilities.formatDate(d, TZ, 'yyyy-MM-dd');
}

function fmt_(n) {
  return Math.round(n).toString().replace(/(\d)(?=(\d\d)+\d$)/g, '$1,');
}

// Shortcuts cuts a link at the first normal space, so labels use no-break spaces (look identical)
function nb_(s) { return String(s).replace(/ /g, '\u2800').replace(/&/g, '\uff06'); }
function un_(s) { return String(s || '').replace(/[\u2800\u00a0]/g, ' ').replace(/\uff06/g, '&').replace(/\+/g, ' '); }

function text_(s) {
  return ContentService.createTextOutput(String(s)).setMimeType(ContentService.MimeType.TEXT);
}

function json_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}
