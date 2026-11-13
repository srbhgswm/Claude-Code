/* Pocket Budget — offline budgeting PWA with an Excel (.xlsx) mirror.
 * All data lives on the device (localStorage). The workbook is generated from
 * that data on demand and can be imported back to restore everything. */
'use strict';

/* ---------------------------------------------------------------- constants */

const STORE_KEY = 'pocketBudget.v1';

const DEBT_TYPES = ['Credit Card', 'Student Loan', 'Auto Loan', 'Mortgage', 'Personal Loan', 'Medical', 'Line of Credit', 'Buy Now Pay Later', 'Other'];

const EXPENSE_CATEGORIES = ['Groceries', 'Dining Out', 'Transport', 'Gas', 'Shopping', 'Entertainment', 'Health', 'Personal Care', 'Education', 'Gifts', 'Travel', 'Household', 'Other'];

const BILL_CATEGORIES = ['Rent / Housing', 'Utilities', 'Phone', 'Internet', 'Insurance', 'Subscriptions', 'Transport', 'Childcare', 'Gym', 'Savings Transfer', 'Other'];

const INCOME_SOURCES = ['Salary', 'Paycheck', 'Freelance', 'Side Gig', 'Bonus', 'Refund', 'Gift', 'Investment', 'Other'];

const FREQUENCIES = {
  once:        { label: 'One-time',           perMonth: 0 },
  weekly:      { label: 'Weekly',             perMonth: 52 / 12 },
  biweekly:    { label: 'Every 2 weeks',      perMonth: 26 / 12 },
  semimonthly: { label: 'Twice a month',      perMonth: 2 },
  monthly:     { label: 'Monthly',            perMonth: 1 },
  quarterly:   { label: 'Quarterly',          perMonth: 1 / 3 },
  yearly:      { label: 'Yearly',             perMonth: 1 / 12 },
};
const RECURRING_FREQS = Object.keys(FREQUENCIES).filter(k => k !== 'once');

const CURRENCIES = ['USD', 'EUR', 'GBP', 'INR', 'CAD', 'AUD', 'JPY', 'CHF', 'SGD', 'MXN'];

/* -------------------------------------------------------------------- state */

function emptyState() {
  return { version: 1, profile: null, debts: [], income: [], expenses: [], bills: [], updatedAt: null };
}

let state = load();
let ui = { tab: 'home', month: monthKey(new Date()), moneySeg: 'expenses', onboarding: null };

function load() {
  try {
    const raw = localStorage.getItem(STORE_KEY);
    if (raw) return Object.assign(emptyState(), JSON.parse(raw));
  } catch (e) { console.warn('Could not read saved data', e); }
  return emptyState();
}

function save() {
  state.updatedAt = new Date().toISOString();
  try { localStorage.setItem(STORE_KEY, JSON.stringify(state)); }
  catch (e) { toast('Could not save on this device'); console.error(e); }
}

const uid = () => Date.now().toString(36) + Math.random().toString(36).slice(2, 7);

/* ------------------------------------------------------------------ helpers */

function esc(s) {
  return String(s ?? '').replace(/[&<>"']/g, c => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;' }[c]));
}

function money(n, opts = {}) {
  const cur = state.profile?.currency || 'USD';
  try {
    return new Intl.NumberFormat(undefined, { style: 'currency', currency: cur, maximumFractionDigits: opts.whole ? 0 : 2, minimumFractionDigits: opts.whole ? 0 : 2 }).format(n || 0);
  } catch { return (n || 0).toFixed(2); }
}

const round2 = n => Math.round((Number(n) || 0) * 100) / 100;
const num = v => { const n = parseFloat(String(v).replace(/[^0-9.\-]/g, '')); return isFinite(n) ? n : 0; };

function isoDate(d) {
  const y = d.getFullYear(), m = String(d.getMonth() + 1).padStart(2, '0'), day = String(d.getDate()).padStart(2, '0');
  return `${y}-${m}-${day}`;
}
const today = () => isoDate(new Date());
function monthKey(d) { return isoDate(d).slice(0, 7); }
function parseDate(s) { const [y, m, d] = String(s).split('-').map(Number); return new Date(y, (m || 1) - 1, d || 1); }
function monthLabel(key) { return parseDate(key + '-01').toLocaleDateString(undefined, { month: 'long', year: 'numeric' }); }
function shiftMonth(key, delta) { const d = parseDate(key + '-01'); d.setMonth(d.getMonth() + delta); return monthKey(d); }
function daysInMonth(key) { const d = parseDate(key + '-01'); return new Date(d.getFullYear(), d.getMonth() + 1, 0).getDate(); }
function fmtDate(s) { return parseDate(s).toLocaleDateString(undefined, { month: 'short', day: 'numeric' }); }
function ordinal(n) { const s = ['th', 'st', 'nd', 'rd'], v = n % 100; return n + (s[(v - 20) % 10] || s[v] || s[0]); }

/** Next calendar date (>= from) that falls on `day` of a month, clamped to month length. */
function nextDueDate(day, from = new Date()) {
  if (!day) return null;
  const base = new Date(from.getFullYear(), from.getMonth(), from.getDate());
  for (let i = 0; i < 2; i++) {
    const y = base.getFullYear(), m = base.getMonth() + i;
    const last = new Date(y, m + 1, 0).getDate();
    const d = new Date(y, m, Math.min(day, last));
    if (d >= base) return d;
  }
  return null;
}

/* ------------------------------------------------------------- calculations */

function monthlyAmount(item) {
  const f = FREQUENCIES[item.frequency] || FREQUENCIES.monthly;
  return item.amount * f.perMonth;
}

function debtStats(d) {
  const r = d.apr / 100 / 12;
  const interest = d.balance * r;
  let months = null; // null => never pays off at minimum
  if (d.balance <= 0) months = 0;
  else if (d.minPayment > interest) {
    months = r === 0 ? Math.ceil(d.balance / d.minPayment)
      : Math.ceil(-Math.log(1 - (r * d.balance) / d.minPayment) / Math.log(1 + r));
  }
  const totalInterest = months == null ? null : Math.max(0, d.minPayment * months - d.balance);
  return { interest: round2(interest), months, totalInterest: totalInterest == null ? null : round2(totalInterest) };
}

function summary(key = ui.month) {
  const recurringIncome = state.income.filter(i => i.frequency !== 'once');
  const recurringIncomeMonthly = recurringIncome.reduce((s, i) => s + monthlyAmount(i), 0);
  const oneOffIncome = state.income.filter(i => i.frequency === 'once' && i.date?.startsWith(key)).reduce((s, i) => s + i.amount, 0);
  const totalIncome = recurringIncomeMonthly + oneOffIncome;

  const billsMonthly = state.bills.reduce((s, b) => s + monthlyAmount(b), 0);
  const debtMinimums = state.debts.reduce((s, d) => s + (d.minPayment || 0), 0);
  const committed = billsMonthly + debtMinimums;

  const monthExpenses = state.expenses.filter(e => e.date?.startsWith(key));
  const spent = monthExpenses.reduce((s, e) => s + e.amount, 0);

  const plannedLeft = totalIncome - committed;
  const expendable = plannedLeft - spent;

  const isCurrent = key === monthKey(new Date());
  const daysLeft = isCurrent ? daysInMonth(key) - new Date().getDate() + 1 : null;
  const perDay = isCurrent && daysLeft > 0 ? expendable / daysLeft : null;

  const byCategory = {};
  monthExpenses.forEach(e => { byCategory[e.category] = (byCategory[e.category] || 0) + e.amount; });

  const totalDebt = state.debts.reduce((s, d) => s + d.balance, 0);
  const monthlyInterest = state.debts.reduce((s, d) => s + debtStats(d).interest, 0);
  const weightedApr = totalDebt ? state.debts.reduce((s, d) => s + d.apr * d.balance, 0) / totalDebt : 0;

  return {
    recurringIncomeMonthly, oneOffIncome, totalIncome, billsMonthly, debtMinimums, committed, spent, plannedLeft,
    expendable, daysLeft, perDay, byCategory, totalDebt, monthlyInterest, weightedApr, monthExpenses,
  };
}

function upcoming(days = 14) {
  const now = new Date(); now.setHours(0, 0, 0, 0);
  const limit = new Date(now); limit.setDate(limit.getDate() + days);
  const items = [];
  state.bills.forEach(b => {
    if (!['monthly', 'semimonthly'].includes(b.frequency) || !b.dueDay) return;
    const d = nextDueDate(b.dueDay, now);
    if (d && d <= limit) items.push({ date: d, name: b.name, kind: b.category, amount: b.amount });
  });
  state.debts.forEach(dt => {
    const d = nextDueDate(dt.dueDay, now);
    if (d && d <= limit) items.push({ date: d, name: dt.lender, kind: dt.type, amount: dt.minPayment });
  });
  return items.sort((a, b) => a.date - b.date);
}

/* ------------------------------------------------------------------- render */

const $ = sel => document.querySelector(sel);
const view = $('#view');

function render() {
  const onboarded = !!state.profile && !ui.onboarding;
  $('#tabbar').hidden = !onboarded;
  $('#export-btn').hidden = !onboarded;
  if (!onboarded) return renderOnboarding();

  document.querySelectorAll('.tabbar button').forEach(b => b.classList.toggle('active', b.dataset.tab === ui.tab));
  const titles = { home: 'Overview', money: 'Income & Spending', bills: 'Recurring Bills', debts: 'Debts', profile: 'Profile' };
  $('#page-title').textContent = titles[ui.tab];
  ({ home: renderHome, money: renderMoney, bills: renderBills, debts: renderDebts, profile: renderProfile })[ui.tab]();
}

function monthNav() {
  return `<div class="monthnav">
    <button class="ghost small" data-act="month" data-delta="-1" type="button" aria-label="Previous month">‹ Prev</button>
    <strong>${esc(monthLabel(ui.month))}</strong>
    <button class="ghost small" data-act="month" data-delta="1" type="button" aria-label="Next month">Next ›</button>
  </div>`;
}

function renderHome() {
  const s = summary();
  const cls = s.expendable >= 0 ? 'pos' : 'neg';
  const usedPct = s.plannedLeft > 0 ? Math.min(100, (s.spent / s.plannedLeft) * 100) : 100;
  const cats = Object.entries(s.byCategory).sort((a, b) => b[1] - a[1]);
  const up = upcoming();
  const name = state.profile.name ? `, ${esc(state.profile.name.split(' ')[0])}` : '';

  view.innerHTML = `
    ${monthNav()}
    <section class="card hero">
      <div class="label">Expendable money left${name}</div>
      <div class="big ${cls}">${money(s.expendable)}</div>
      ${s.perDay != null ? `<div class="sub">≈ ${money(Math.max(0, s.perDay))} per day for the next ${s.daysLeft} day${s.daysLeft === 1 ? '' : 's'}</div>` : ''}
      <div class="bar" aria-hidden="true"><span style="width:${usedPct}%;${usedPct >= 100 ? 'background:var(--bad)' : ''}"></span></div>
      <div class="sub" style="margin-top:6px">${money(s.spent)} spent of ${money(Math.max(0, s.plannedLeft))} free to spend</div>
    </section>

    <section class="card">
      <h2>How it adds up</h2>
      <div class="rows">
        <div class="row"><div class="l">Recurring income <div class="s">monthly equivalent</div></div><div class="r pos">+${money(s.recurringIncomeMonthly)}</div></div>
        ${s.oneOffIncome ? `<div class="row"><div class="l">One-time income</div><div class="r pos">+${money(s.oneOffIncome)}</div></div>` : ''}
        <div class="row"><div class="l">Recurring bills</div><div class="r">−${money(s.billsMonthly)}</div></div>
        <div class="row"><div class="l">Debt minimum payments</div><div class="r">−${money(s.debtMinimums)}</div></div>
        <div class="row"><div class="l">Spending this month</div><div class="r">−${money(s.spent)}</div></div>
        <div class="row total"><div class="l">Expendable</div><div class="r ${cls}">${money(s.expendable)}</div></div>
      </div>
      ${s.totalIncome === 0 ? `<p class="note">Add your paycheck under <b>Money → Income</b> to see a real number.</p>` : ''}
    </section>

    <section class="card">
      <h2>Due in the next 14 days</h2>
      ${up.length ? `<div class="rows">${up.map(u => `
        <div class="row"><div class="l"><div class="t">${esc(u.name)}</div><div class="s">${esc(u.kind)} · ${u.date.toLocaleDateString(undefined, { weekday: 'short', month: 'short', day: 'numeric' })}</div></div>
        <div class="r">${money(u.amount)}</div></div>`).join('')}</div>`
        : `<div class="empty">Nothing due soon 🎉</div>`}
    </section>

    ${state.debts.length ? `
    <section class="card">
      <h2>Debt snapshot</h2>
      <div class="rows">
        <div class="row"><div class="l">Total owed</div><div class="r">${money(s.totalDebt)}</div></div>
        <div class="row"><div class="l">Interest accruing / month</div><div class="r neg">${money(s.monthlyInterest)}</div></div>
        <div class="row"><div class="l">Average APR (weighted)</div><div class="r">${s.weightedApr.toFixed(2)}%</div></div>
      </div>
      ${avalancheTip()}
    </section>` : ''}

    <section class="card">
      <h2>Spending by category</h2>
      ${cats.length ? `<div class="rows">${cats.map(([c, v]) => `
        <div class="row"><div class="l" style="flex:1"><div class="t">${esc(c)}</div>
          <div class="bar"><span style="width:${(v / cats[0][1]) * 100}%"></span></div></div>
        <div class="r">${money(v)}</div></div>`).join('')}</div>`
        : `<div class="empty">No expenses logged this month.</div>`}
    </section>

    <div class="actions">
      <button type="button" data-act="add-expense">+ Expense</button>
      <button type="button" class="secondary" data-act="add-income">+ Income</button>
    </div>`;
}

function avalancheTip() {
  const s = summary();
  const extra = s.expendable;
  const target = [...state.debts].filter(d => d.balance > 0).sort((a, b) => b.apr - a.apr)[0];
  if (!target) return '';
  const never = state.debts.filter(d => d.balance > 0 && debtStats(d).months == null);
  return `<p class="note">💡 Pay extra toward <b>${esc(target.lender)}</b> (${target.apr}% APR) first — the highest rate costs you the most.
    ${extra > 0 ? `Putting even part of your ${money(extra, { whole: true })} leftover there saves the most interest.` : ''}</p>
    ${never.length ? `<p class="note neg">⚠︎ Minimum payment doesn't cover interest on: ${never.map(d => esc(d.lender)).join(', ')}. That balance will grow.</p>` : ''}`;
}

function renderMoney() {
  const seg = ui.moneySeg;
  let list = '';
  if (seg === 'expenses') {
    const items = state.expenses.filter(e => e.date?.startsWith(ui.month)).sort((a, b) => b.date.localeCompare(a.date));
    const total = items.reduce((s, e) => s + e.amount, 0);
    list = items.length ? `<div class="rows">${items.map(e => `
      <div class="row tap" data-act="edit-expense" data-id="${e.id}"><div class="l"><div class="t">${esc(e.description || e.category)}</div>
        <div class="s">${fmtDate(e.date)} · ${esc(e.category)}${e.method ? ' · ' + esc(e.method) : ''}</div></div>
        <div class="r">−${money(e.amount)}</div></div>`).join('')}
      <div class="row total"><div class="l">Total</div><div class="r">${money(total)}</div></div></div>`
      : `<div class="empty">No expenses this month.</div>`;
  } else {
    const rec = state.income.filter(i => i.frequency !== 'once');
    const once = state.income.filter(i => i.frequency === 'once' && i.date?.startsWith(ui.month)).sort((a, b) => b.date.localeCompare(a.date));
    list = `
      <h2 style="margin-top:4px">Recurring</h2>
      ${rec.length ? `<div class="rows">${rec.map(i => `
        <div class="row tap" data-act="edit-income" data-id="${i.id}"><div class="l"><div class="t">${esc(i.source)}</div>
          <div class="s">${money(i.amount)} · ${FREQUENCIES[i.frequency].label}</div></div>
          <div class="r pos">+${money(monthlyAmount(i))}<div class="s">/ month</div></div></div>`).join('')}</div>`
        : `<div class="empty">No recurring income yet.</div>`}
      <h2 style="margin-top:12px">One-time in ${esc(monthLabel(ui.month))}</h2>
      ${once.length ? `<div class="rows">${once.map(i => `
        <div class="row tap" data-act="edit-income" data-id="${i.id}"><div class="l"><div class="t">${esc(i.source)}</div>
          <div class="s">${fmtDate(i.date)}${i.note ? ' · ' + esc(i.note) : ''}</div></div>
          <div class="r pos">+${money(i.amount)}</div></div>`).join('')}</div>`
        : `<div class="empty">None logged.</div>`}`;
  }
  view.innerHTML = `
    ${monthNav()}
    <div class="segmented" role="tablist">
      <button type="button" data-act="seg" data-seg="expenses" class="${seg === 'expenses' ? 'active' : ''}">Expenses</button>
      <button type="button" data-act="seg" data-seg="income" class="${seg === 'income' ? 'active' : ''}">Income</button>
    </div>
    <section class="card">${list}</section>
    <button type="button" class="block" data-act="${seg === 'expenses' ? 'add-expense' : 'add-income'}">+ Add ${seg === 'expenses' ? 'expense' : 'income'}</button>`;
}

function renderBills() {
  const bills = [...state.bills].sort((a, b) => (a.dueDay || 99) - (b.dueDay || 99));
  const total = bills.reduce((s, b) => s + monthlyAmount(b), 0);
  view.innerHTML = `
    <section class="card">
      ${bills.length ? `<div class="rows">${bills.map(b => `
        <div class="row tap" data-act="edit-bill" data-id="${b.id}"><div class="l"><div class="t">${esc(b.name)}</div>
          <div class="s">${esc(b.category)} · ${FREQUENCIES[b.frequency].label}${b.dueDay ? ' · due ' + ordinal(b.dueDay) : ''}${b.autopay ? ' · autopay' : ''}</div></div>
          <div class="r">${money(b.amount)}${b.frequency !== 'monthly' ? `<div class="s">${money(monthlyAmount(b))}/mo</div>` : ''}</div></div>`).join('')}
        <div class="row total"><div class="l">Monthly total</div><div class="r">${money(total)}</div></div></div>`
        : `<div class="empty">No recurring bills yet. Add rent, utilities, phone, subscriptions…</div>`}
    </section>
    <button type="button" class="block" data-act="add-bill">+ Add recurring bill</button>`;
}

function renderDebts() {
  const groups = {};
  state.debts.forEach(d => (groups[d.type] = groups[d.type] || []).push(d));
  const s = summary();
  view.innerHTML = `
    ${state.debts.length ? `
    <section class="card">
      <div class="rows">
        <div class="row total"><div class="l">Total owed</div><div class="r">${money(s.totalDebt)}</div></div>
        <div class="row"><div class="l">Minimum payments / month</div><div class="r">${money(s.debtMinimums)}</div></div>
        <div class="row"><div class="l">Interest / month</div><div class="r neg">${money(s.monthlyInterest)}</div></div>
      </div>
      ${avalancheTip()}
    </section>` : ''}
    ${Object.keys(groups).length ? DEBT_TYPES.filter(t => groups[t]).map(t => `
      <section class="card">
        <h2>${esc(t)} <span class="muted" style="text-transform:none">· ${money(groups[t].reduce((a, d) => a + d.balance, 0))}</span></h2>
        <div class="rows">${groups[t].map(d => {
          const st = debtStats(d);
          const nd = nextDueDate(d.dueDay);
          return `<div class="row tap" data-act="edit-debt" data-id="${d.id}"><div class="l"><div class="t">${esc(d.lender)}</div>
            <div class="s">${d.apr}% APR · min ${money(d.minPayment)}${nd ? ' · due ' + nd.toLocaleDateString(undefined, { month: 'short', day: 'numeric' }) : ''}</div>
            <div class="s">${st.months == null ? '<span class="neg">Never paid off at minimum</span>' : st.months === 0 ? 'Paid off' : `Paid off in ~${st.months} mo · ${money(st.totalInterest, { whole: true })} interest`}</div></div>
            <div class="r">${money(d.balance)}<div class="s neg">+${money(st.interest)}/mo</div></div></div>`;
        }).join('')}</div>
      </section>`).join('')
      : `<section class="card"><div class="empty">No debts recorded. Lucky you — or tap below to add one.</div></section>`}
    <button type="button" class="block" data-act="add-debt">+ Add debt</button>`;
}

function renderProfile() {
  const p = state.profile;
  const counts = `${state.debts.length} debts · ${state.bills.length} bills · ${state.income.length} income · ${state.expenses.length} expenses`;
  view.innerHTML = `
    <section class="card">
      <h3>${esc(p.name || 'Me')}</h3>
      <div class="s muted">${esc(p.currency)} · paid ${esc(FREQUENCIES[p.payFrequency]?.label.toLowerCase() || '—')}</div>
      <button type="button" class="secondary block" data-act="edit-profile">Edit profile</button>
    </section>
    <section class="card">
      <h2>Excel workbook</h2>
      <p class="note">Your data is mirrored to an Excel file with sheets for Summary, Profile, Debts, Income, Expenses and Bills (with live formulas).
        Export it to Files, iCloud Drive or OneDrive. Importing a workbook replaces the data in the app — so you can also edit the sheet and bring it back.</p>
      <p class="note">Last change: ${state.updatedAt ? new Date(state.updatedAt).toLocaleString() : '—'}<br>${counts}</p>
      <div class="actions">
        <button type="button" data-act="export">Export .xlsx</button>
        <button type="button" class="secondary" data-act="import">Import .xlsx</button>
      </div>
      <input type="file" id="import-file" accept=".xlsx,application/vnd.openxmlformats-officedocument.spreadsheetml.sheet" hidden>
    </section>
    <section class="card">
      <h2>Danger zone</h2>
      <button type="button" class="danger block" data-act="reset">Erase all data on this device</button>
    </section>
    <p class="note" style="text-align:center">Everything stays on your phone. Nothing is sent anywhere.</p>`;
}

/* --------------------------------------------------------------- onboarding */

function renderOnboarding() {
  $('#page-title').textContent = 'Pocket Budget';
  const step = ui.onboarding?.step || 'welcome';
  if (step === 'welcome') {
    view.innerHTML = `
      <section class="welcome">
        <div class="logo">💸</div>
        <h2>Let's set up your budget</h2>
        <p class="muted">We'll build your profile, list your debts and bills, and show how much money you really have to spend.
          Everything is saved on this phone and mirrored to an Excel workbook.</p>
        <button type="button" class="block" data-act="ob-start">Get started</button>
        <button type="button" class="ghost block" data-act="import">I have a workbook — import it</button>
        <input type="file" id="import-file" accept=".xlsx" hidden>
      </section>`;
    return;
  }
  const lists = {
    debts: { title: 'Step 2 of 3 · Your debts', intro: 'Add each credit card, loan or other debt. You can always add more later.',
      items: state.debts.map(d => `${esc(d.lender)} <span class="chip">${esc(d.type)}</span>`), add: 'add-debt', addLabel: '+ Add a debt', next: 'bills' },
    bills: { title: 'Step 3 of 3 · Recurring bills', intro: 'Rent, utilities, phone, insurance, subscriptions — anything that repeats.',
      items: state.bills.map(b => `${esc(b.name)} · ${money(b.amount)}`), add: 'add-bill', addLabel: '+ Add a bill', next: 'done' },
  }[step];
  view.innerHTML = `
    <div class="steps">${lists.title}</div>
    <section class="card">
      <p class="muted" style="margin-top:0">${lists.intro}</p>
      ${lists.items.length ? `<div class="rows">${lists.items.map(i => `<div class="row"><div class="l">${i}</div><div class="r pos">✓</div></div>`).join('')}</div>` : ''}
      <button type="button" class="secondary block" data-act="${lists.add}">${lists.addLabel}</button>
    </section>
    <button type="button" class="block" data-act="ob-next" data-next="${lists.next}">${lists.items.length ? 'Continue' : 'Skip for now'}</button>`;
}

/* -------------------------------------------------------------------- forms */

const sheet = $('#sheet');
const sheetForm = $('#sheet-form');

function options(list, selected) {
  return list.map(o => {
    const [v, l] = Array.isArray(o) ? o : [o, o];
    return `<option value="${esc(v)}" ${v === selected ? 'selected' : ''}>${esc(l)}</option>`;
  }).join('');
}
const freqOptions = (keys, sel) => options(keys.map(k => [k, FREQUENCIES[k].label]), sel);

function field(label, html) { return `<label class="field"><span>${label}</span>${html}</label>`; }
function input(name, value = '', attrs = '') { return `<input name="${name}" value="${esc(value)}" ${attrs}>`; }
const moneyInput = (name, value, extra = '') => input(name, value === undefined || value === null || value === '' ? '' : value, `inputmode="decimal" placeholder="0.00" ${extra}`);

/** Open a bottom sheet with a form. `onSubmit(data)` returns an error string or nothing. */
function openSheet({ title, body, onSubmit, onDelete }) {
  sheetForm.innerHTML = `
    <h2>${esc(title)}</h2>
    ${body}
    <div class="err" id="form-err"></div>
    <div class="actions">
      <button type="button" class="secondary" value="cancel" data-sheet="cancel">Cancel</button>
      <button type="submit" value="save">Save</button>
    </div>
    ${onDelete ? `<button type="button" class="danger block" data-sheet="delete">Delete</button>` : ''}`;
  sheetForm.onsubmit = ev => {
    ev.preventDefault();
    const fd = Object.fromEntries(new FormData(sheetForm).entries());
    sheetForm.querySelectorAll('input[type=checkbox]').forEach(c => { fd[c.name] = c.checked; });
    const err = onSubmit(fd);
    if (err) { $('#form-err').textContent = err; return; }
    save(); sheet.close(); render();
  };
  sheetForm.onclick = ev => {
    const b = ev.target.closest('[data-sheet]');
    if (!b) return;
    if (b.dataset.sheet === 'cancel') sheet.close();
    if (b.dataset.sheet === 'delete' && confirm('Delete this item?')) { onDelete(); save(); sheet.close(); render(); }
  };
  sheet.showModal();
  sheetForm.scrollTop = 0;
}

function profileForm(isNew) {
  const p = state.profile || { name: '', currency: 'USD', payFrequency: 'biweekly' };
  const body = `
    ${field('Your name', input('name', p.name, 'autocomplete="given-name" placeholder="e.g. Alex"'))}
    <div class="grid2">
      ${field('Currency', `<select name="currency">${options(CURRENCIES, p.currency)}</select>`)}
      ${field('I get paid', `<select name="payFrequency">${freqOptions(RECURRING_FREQS, p.payFrequency)}</select>`)}
    </div>
    ${isNew ? `${field('Take-home pay per paycheck (after tax)', moneyInput('pay', ''))}
      <p class="note">Optional — this becomes a recurring income entry. You can add other income later.</p>` : ''}`;
  openSheet({
    title: isNew ? 'Step 1 of 3 · Your profile' : 'Edit profile',
    body,
    onSubmit: d => {
      if (!d.name.trim()) return 'Please enter your name.';
      const createdAt = state.profile?.createdAt || today();
      state.profile = { name: d.name.trim(), currency: d.currency, payFrequency: d.payFrequency, createdAt };
      if (isNew && num(d.pay) > 0) {
        state.income.push({ id: uid(), source: 'Paycheck', amount: round2(num(d.pay)), frequency: d.payFrequency, date: today(), note: '' });
      }
      if (isNew) ui.onboarding = { step: 'debts' };
    },
  });
}

function debtForm(existing) {
  const d = existing || { type: 'Credit Card', lender: '', balance: '', minPayment: '', dueDay: '', apr: '', notes: '' };
  openSheet({
    title: existing ? 'Edit debt' : 'Add a debt',
    body: `
      ${field('Type', `<select name="type">${options(DEBT_TYPES, d.type)}</select>`)}
      ${field('Name / lender', input('lender', d.lender, 'placeholder="e.g. Chase Sapphire, Sallie Mae"'))}
      <div class="grid2">
        ${field('Amount owed', moneyInput('balance', d.balance))}
        ${field('Interest rate (APR %)', input('apr', d.apr, 'inputmode="decimal" placeholder="e.g. 24.99"'))}
        ${field('Minimum / monthly payment', moneyInput('minPayment', d.minPayment))}
        ${field('Due day of month', input('dueDay', d.dueDay, 'inputmode="numeric" placeholder="1–31"'))}
      </div>
      ${field('Notes', input('notes', d.notes, 'placeholder="optional"'))}`,
    onSubmit: f => {
      if (!f.lender.trim()) return 'Give this debt a name.';
      const due = parseInt(f.dueDay, 10);
      if (f.dueDay && (!(due >= 1 && due <= 31))) return 'Due day must be between 1 and 31.';
      const apr = num(f.apr);
      if (apr < 0 || apr > 100) return 'APR should be between 0 and 100.';
      const rec = { id: d.id || uid(), type: f.type, lender: f.lender.trim(), balance: round2(num(f.balance)), minPayment: round2(num(f.minPayment)), dueDay: due || null, apr, notes: f.notes.trim() };
      upsert(state.debts, rec);
    },
    onDelete: existing && (() => remove(state.debts, d.id)),
  });
}

function incomeForm(existing) {
  const i = existing || { source: 'Paycheck', amount: '', frequency: 'once', date: today(), note: '' };
  openSheet({
    title: existing ? 'Edit income' : 'Log income',
    body: `
      ${field('Source', `<input name="source" list="income-sources" value="${esc(i.source)}"><datalist id="income-sources">${options(INCOME_SOURCES)}</datalist>`)}
      <div class="grid2">
        ${field('Amount', moneyInput('amount', i.amount))}
        ${field('Date', input('date', i.date, 'type="date"'))}
      </div>
      ${field('How often?', `<select name="frequency">${freqOptions(Object.keys(FREQUENCIES), i.frequency)}</select>`)}
      ${field('Note', input('note', i.note, 'placeholder="optional"'))}
      <p class="note">Recurring income counts every month automatically. One-time income counts only in the month of its date.</p>`,
    onSubmit: f => {
      if (num(f.amount) <= 0) return 'Enter an amount greater than zero.';
      if (!f.date) return 'Pick a date.';
      upsert(state.income, { id: i.id || uid(), source: f.source.trim() || 'Income', amount: round2(num(f.amount)), frequency: f.frequency, date: f.date, note: f.note.trim() });
    },
    onDelete: existing && (() => remove(state.income, i.id)),
  });
}

function expenseForm(existing) {
  const defaultDate = ui.month === monthKey(new Date()) ? today() : ui.month + '-01';
  const e = existing || { date: defaultDate, category: 'Groceries', description: '', amount: '', method: '' };
  const methods = ['', 'Debit', 'Cash', ...state.debts.filter(d => d.type === 'Credit Card').map(d => d.lender), 'Credit Card', 'Other'];
  openSheet({
    title: existing ? 'Edit expense' : 'Log expense',
    body: `
      <div class="grid2">
        ${field('Amount', moneyInput('amount', e.amount, 'autofocus'))}
        ${field('Date', input('date', e.date, 'type="date"'))}
      </div>
      ${field('Category', `<select name="category">${options(EXPENSE_CATEGORIES, e.category)}</select>`)}
      ${field('Description', input('description', e.description, 'placeholder="e.g. Trader Joe\'s"'))}
      ${field('Paid with', `<select name="method">${options([...new Set(methods)].map(m => [m, m || '—']), e.method)}</select>`)}`,
    onSubmit: f => {
      if (num(f.amount) <= 0) return 'Enter an amount greater than zero.';
      if (!f.date) return 'Pick a date.';
      upsert(state.expenses, { id: e.id || uid(), date: f.date, category: f.category, description: f.description.trim(), amount: round2(num(f.amount)), method: f.method });
    },
    onDelete: existing && (() => remove(state.expenses, e.id)),
  });
}

function billForm(existing) {
  const b = existing || { name: '', category: 'Utilities', amount: '', frequency: 'monthly', dueDay: '', autopay: false };
  openSheet({
    title: existing ? 'Edit bill' : 'Add recurring bill',
    body: `
      ${field('Name', input('name', b.name, 'placeholder="e.g. Rent, Netflix, Electric"'))}
      ${field('Category', `<select name="category">${options(BILL_CATEGORIES, b.category)}</select>`)}
      <div class="grid2">
        ${field('Amount', moneyInput('amount', b.amount))}
        ${field('Due day of month', input('dueDay', b.dueDay, 'inputmode="numeric" placeholder="1–31"'))}
      </div>
      ${field('How often?', `<select name="frequency">${freqOptions(RECURRING_FREQS, b.frequency)}</select>`)}
      <label class="check"><input type="checkbox" name="autopay" ${b.autopay ? 'checked' : ''}> On autopay</label>`,
    onSubmit: f => {
      if (!f.name.trim()) return 'Give the bill a name.';
      if (num(f.amount) <= 0) return 'Enter an amount greater than zero.';
      const due = parseInt(f.dueDay, 10);
      if (f.dueDay && !(due >= 1 && due <= 31)) return 'Due day must be between 1 and 31.';
      upsert(state.bills, { id: b.id || uid(), name: f.name.trim(), category: f.category, amount: round2(num(f.amount)), frequency: f.frequency, dueDay: due || null, autopay: !!f.autopay });
    },
    onDelete: existing && (() => remove(state.bills, b.id)),
  });
}

function upsert(list, rec) { const i = list.findIndex(x => x.id === rec.id); if (i >= 0) list[i] = rec; else list.push(rec); }
function remove(list, id) { const i = list.findIndex(x => x.id === id); if (i >= 0) list.splice(i, 1); }

/* -------------------------------------------------------------------- Excel */

const SHEETS = {
  Debts:    [['ID', 'id'], ['Type', 'type'], ['Lender', 'lender'], ['Balance', 'balance'], ['APR %', 'apr'], ['Min Payment', 'minPayment'], ['Due Day', 'dueDay'], ['Notes', 'notes']],
  Income:   [['ID', 'id'], ['Date', 'date'], ['Source', 'source'], ['Amount', 'amount'], ['Frequency', 'frequency'], ['Note', 'note']],
  Expenses: [['ID', 'id'], ['Date', 'date'], ['Category', 'category'], ['Description', 'description'], ['Amount', 'amount'], ['Paid With', 'method']],
  Bills:    [['ID', 'id'], ['Name', 'name'], ['Category', 'category'], ['Amount', 'amount'], ['Frequency', 'frequency'], ['Due Day', 'dueDay'], ['Autopay', 'autopay']],
};
const STATE_KEY = { Debts: 'debts', Income: 'income', Expenses: 'expenses', Bills: 'bills' };
// Monthly multiplier per frequency, expressed as an Excel expression on a cell reference.
const freqFormula = ref => `IF(${ref}="weekly",52/12,IF(${ref}="biweekly",26/12,IF(${ref}="semimonthly",2,IF(${ref}="monthly",1,IF(${ref}="quarterly",1/3,IF(${ref}="yearly",1/12,0))))))`;

function buildWorkbook() {
  const XL = window.XLSX;
  const wb = XL.utils.book_new();
  const s = summary(monthKey(new Date()));
  const cm = monthKey(new Date());

  // Summary — formulas so the workbook stays live if edited in Excel/Numbers.
  const summaryRows = [
    ['Pocket Budget — Summary', ''],
    ['Generated', new Date().toLocaleString()],
    ['Month', cm],
    [],
    ['Recurring income (monthly)', round2(s.recurringIncomeMonthly), 'SUMPRODUCT(Income!D2:D1000,Income!G2:G1000)'],
    ['One-time income this month', round2(s.oneOffIncome), `SUMPRODUCT((LEFT(Income!B2:B1000,7)=B3)*(Income!E2:E1000="once")*Income!D2:D1000)`],
    ['Recurring bills (monthly)', round2(s.billsMonthly), 'SUM(Bills!H2:H1000)'],
    ['Debt minimum payments', round2(s.debtMinimums), 'SUM(Debts!F2:F1000)'],
    ['Spent this month', round2(s.spent), `SUMPRODUCT((LEFT(Expenses!B2:B5000,7)=B3)*Expenses!E2:E5000)`],
    ['Expendable this month', round2(s.expendable), 'B5+B6-B7-B8-B9'],
    [],
    ['Total debt', round2(s.totalDebt), 'SUM(Debts!D2:D1000)'],
    ['Interest accruing / month', round2(s.monthlyInterest), 'SUM(Debts!I2:I1000)'],
  ];
  const ws = XL.utils.aoa_to_sheet(summaryRows.map(r => r.slice(0, 2)));
  summaryRows.forEach((r, i) => { if (r[2]) ws[`B${i + 1}`] = { t: 'n', v: r[1], f: r[2], z: '#,##0.00' }; });
  ws['!cols'] = [{ wch: 30 }, { wch: 18 }];
  XL.utils.book_append_sheet(wb, ws, 'Summary');

  // Profile
  const p = state.profile || {};
  const wp = XL.utils.aoa_to_sheet([['Field', 'Value'], ['Name', p.name || ''], ['Currency', p.currency || ''], ['Pay Frequency', p.payFrequency || ''], ['Created', p.createdAt || '']]);
  wp['!cols'] = [{ wch: 16 }, { wch: 24 }];
  XL.utils.book_append_sheet(wb, wp, 'Profile');

  // Data sheets
  for (const [name, cols] of Object.entries(SHEETS)) {
    const rows = state[STATE_KEY[name]].map(r => cols.map(([, k]) => (r[k] ?? '')));
    const extra = { Debts: ['Monthly Interest'], Income: ['Monthly Equivalent'], Bills: ['Monthly Equivalent'] }[name] || [];
    const w = XL.utils.aoa_to_sheet([[...cols.map(c => c[0]), ...extra], ...rows]);
    rows.forEach((_, i) => {
      const n = i + 2;
      if (name === 'Debts')  w[`I${n}`] = { t: 'n', v: debtStats(state.debts[i]).interest, f: `D${n}*E${n}/1200`, z: '#,##0.00' };
      if (name === 'Income') w[`G${n}`] = { t: 'n', v: round2(monthlyAmount(state.income[i])), f: `D${n}*${freqFormula(`E${n}`)}`, z: '#,##0.00' };
      if (name === 'Bills')  w[`H${n}`] = { t: 'n', v: round2(monthlyAmount(state.bills[i])), f: `D${n}*${freqFormula(`E${n}`)}`, z: '#,##0.00' };
    });
    if (rows.length) w['!ref'] = XL.utils.encode_range({ s: { r: 0, c: 0 }, e: { r: rows.length, c: cols.length + extra.length - 1 } });
    w['!cols'] = [{ wch: 12 }, ...cols.slice(1).map(() => ({ wch: 16 })), ...extra.map(() => ({ wch: 18 }))];
    XL.utils.book_append_sheet(wb, w, name);
  }
  return wb;
}

async function exportWorkbook() {
  if (!window.XLSX) return toast('Excel library failed to load');
  const wb = buildWorkbook();
  const data = XLSX.write(wb, { bookType: 'xlsx', type: 'array' });
  const filename = `PocketBudget-${today()}.xlsx`;
  const blob = new Blob([data], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' });
  const file = new File([blob], filename, { type: blob.type });
  // iOS: the share sheet lets you "Save to Files", iCloud Drive, OneDrive, Mail…
  if (navigator.canShare && navigator.canShare({ files: [file] })) {
    try { await navigator.share({ files: [file], title: filename }); return; }
    catch (e) { if (e.name === 'AbortError') return; }
  }
  const a = document.createElement('a');
  a.href = URL.createObjectURL(blob); a.download = filename;
  document.body.appendChild(a); a.click(); a.remove();
  setTimeout(() => URL.revokeObjectURL(a.href), 4000);
  toast('Workbook exported');
}

function importWorkbook(file) {
  const reader = new FileReader();
  reader.onload = () => {
    try {
      const wb = XLSX.read(reader.result, { type: 'array' });
      const next = emptyState();
      const prof = wb.Sheets.Profile ? XLSX.utils.sheet_to_json(wb.Sheets.Profile, { header: 1 }) : [];
      const pv = Object.fromEntries(prof.slice(1).map(r => [r[0], r[1]]));
      next.profile = { name: String(pv.Name || 'Me'), currency: CURRENCIES.includes(pv.Currency) ? pv.Currency : 'USD', payFrequency: FREQUENCIES[pv['Pay Frequency']] ? pv['Pay Frequency'] : 'biweekly', createdAt: String(pv.Created || today()) };
      for (const [name, cols] of Object.entries(SHEETS)) {
        const ws = wb.Sheets[name];
        if (!ws) continue;
        const rows = XLSX.utils.sheet_to_json(ws, { defval: '', raw: true });
        next[STATE_KEY[name]] = rows.filter(r => Object.values(r).some(v => v !== '')).map(r => {
          const o = {};
          cols.forEach(([h, k]) => { o[k] = r[h]; });
          return normalizeRecord(name, o);
        });
      }
      if (!confirm(`Replace everything in the app with this workbook?\n${next.debts.length} debts, ${next.bills.length} bills, ${next.income.length} income, ${next.expenses.length} expenses.`)) return;
      state = next; ui.onboarding = null; save(); ui.tab = 'home'; render();
      toast('Workbook imported');
    } catch (e) { console.error(e); toast('That file could not be read as a Pocket Budget workbook'); }
  };
  reader.readAsArrayBuffer(file);
}

function excelDateToIso(v) {
  if (typeof v === 'number') { const d = new Date(Math.round((v - 25569) * 864e5)); return isoDate(new Date(d.getUTCFullYear(), d.getUTCMonth(), d.getUTCDate())); }
  const s = String(v || '').trim();
  if (/^\d{4}-\d{2}-\d{2}/.test(s)) return s.slice(0, 10);
  const d = new Date(s); return isNaN(d) ? today() : isoDate(d);
}

function normalizeRecord(sheetName, o) {
  o.id = o.id ? String(o.id) : uid();
  const freq = f => (FREQUENCIES[String(f).toLowerCase()] ? String(f).toLowerCase() : 'monthly');
  const day = v => { const n = parseInt(v, 10); return n >= 1 && n <= 31 ? n : null; };
  switch (sheetName) {
    case 'Debts': return { id: o.id, type: DEBT_TYPES.includes(o.type) ? o.type : 'Other', lender: String(o.lender || 'Debt'), balance: round2(num(o.balance)), apr: num(o.apr), minPayment: round2(num(o.minPayment)), dueDay: day(o.dueDay), notes: String(o.notes || '') };
    case 'Income': return { id: o.id, date: excelDateToIso(o.date), source: String(o.source || 'Income'), amount: round2(num(o.amount)), frequency: String(o.frequency).toLowerCase() === 'once' ? 'once' : freq(o.frequency), note: String(o.note || '') };
    case 'Expenses': return { id: o.id, date: excelDateToIso(o.date), category: String(o.category || 'Other'), description: String(o.description || ''), amount: round2(num(o.amount)), method: String(o.method || '') };
    case 'Bills': return { id: o.id, name: String(o.name || 'Bill'), category: String(o.category || 'Other'), amount: round2(num(o.amount)), frequency: freq(o.frequency), dueDay: day(o.dueDay), autopay: o.autopay === true || /^(true|yes|1)$/i.test(String(o.autopay)) };
  }
}

/* ------------------------------------------------------------------- events */

function toast(msg) {
  const t = $('#toast'); t.textContent = msg; t.classList.add('show');
  clearTimeout(toast._t); toast._t = setTimeout(() => t.classList.remove('show'), 2200);
}

const find = (list, id) => list.find(x => x.id === id);

document.addEventListener('click', ev => {
  const tab = ev.target.closest('.tabbar button');
  if (tab) { ui.tab = tab.dataset.tab; render(); window.scrollTo(0, 0); return; }
  const el = ev.target.closest('[data-act]');
  if (!el || el.closest('dialog')) return;
  const id = el.dataset.id;
  switch (el.dataset.act) {
    case 'month': ui.month = shiftMonth(ui.month, +el.dataset.delta); render(); break;
    case 'seg': ui.moneySeg = el.dataset.seg; render(); break;
    case 'add-expense': expenseForm(); break;
    case 'edit-expense': expenseForm(find(state.expenses, id)); break;
    case 'add-income': incomeForm(); break;
    case 'edit-income': incomeForm(find(state.income, id)); break;
    case 'add-bill': billForm(); break;
    case 'edit-bill': billForm(find(state.bills, id)); break;
    case 'add-debt': debtForm(); break;
    case 'edit-debt': debtForm(find(state.debts, id)); break;
    case 'edit-profile': profileForm(false); break;
    case 'ob-start': profileForm(true); break;
    case 'ob-next':
      if (el.dataset.next === 'done') { ui.onboarding = null; ui.tab = 'home'; toast('All set!'); }
      else ui.onboarding = { step: el.dataset.next };
      render(); break;
    case 'export': exportWorkbook(); break;
    case 'import': $('#import-file').click(); break;
    case 'reset':
      if (confirm('Erase all budget data on this device? Export the Excel file first if you want a backup.')) {
        localStorage.removeItem(STORE_KEY); state = emptyState(); ui = { tab: 'home', month: monthKey(new Date()), moneySeg: 'expenses', onboarding: null }; render();
      }
      break;
  }
});

document.addEventListener('change', ev => {
  if (ev.target.id === 'import-file' && ev.target.files[0]) { importWorkbook(ev.target.files[0]); ev.target.value = ''; }
});

$('#export-btn').addEventListener('click', exportWorkbook);
sheet.addEventListener('click', ev => { if (ev.target === sheet) sheet.close(); }); // tap backdrop to dismiss

if ('serviceWorker' in navigator && location.protocol !== 'file:') {
  navigator.serviceWorker.register('sw.js').catch(e => console.warn('SW registration failed', e));
}

render();
