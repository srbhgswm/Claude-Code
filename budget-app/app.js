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
  return { version: 2, profile: null, debts: [], income: [], expenses: [], bills: [], cardPlan: { spending: null }, updatedAt: null };
}

let state = load();
let ui = { tab: 'home', month: monthKey(new Date()), moneySeg: 'expenses', debtSeg: 'all', onboarding: null };
let memo = {}; // per-change cache for the heavier calculations

function load() {
  try {
    const raw = localStorage.getItem(STORE_KEY);
    if (raw) {
      const s = Object.assign(emptyState(), JSON.parse(raw));
      if (!s.cardPlan) s.cardPlan = { spending: null }; // data saved by v1 of the app
      return s;
    }
  } catch (e) { console.warn('Could not read saved data', e); }
  return emptyState();
}

function save() {
  memo = {};
  state.updatedAt = new Date().toISOString();
  try { localStorage.setItem(STORE_KEY, JSON.stringify(state)); }
  catch (e) { toast('Could not save on this device'); console.error(e); }
}
const cached = (key, fn) => (key in memo ? memo[key] : (memo[key] = fn()));

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
function startOfToday() { const d = new Date(); return new Date(d.getFullYear(), d.getMonth(), d.getDate()); }
function addDays(d, n) { return new Date(d.getFullYear(), d.getMonth(), d.getDate() + n); }
function addMonths(d, n) { return new Date(d.getFullYear(), d.getMonth() + n, d.getDate()); }
function dayDiff(a, b) { return Math.round((b - a) / 864e5); }
function clampDate(y, m, day) { return new Date(y, m, Math.min(day, new Date(y, m + 1, 0).getDate())); }
function monthOffset(key) { const [y, m] = key.split('-').map(Number), n = new Date(); return (y - n.getFullYear()) * 12 + (m - 1 - n.getMonth()); }
function offsetLabel(k) { const n = new Date(); return new Date(n.getFullYear(), n.getMonth() + k, 1).toLocaleDateString(undefined, { month: 'short', year: 'numeric' }); }
const shortDate = d => d.toLocaleDateString(undefined, { weekday: 'short', month: 'short', day: 'numeric' });

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

/* ---- paychecks */

/** How the user is paid, or null if paydays haven't been set up yet. */
function payConfig() {
  const p = state.profile; if (!p) return null;
  const f = p.payFrequency;
  if ((f === 'weekly' || f === 'biweekly') && p.payAnchor) return { f, step: f === 'weekly' ? 7 : 14, anchor: parseDate(p.payAnchor) };
  if ((f === 'semimonthly' || f === 'monthly') && p.payDays?.length) return { f, days: f === 'monthly' ? [p.payDays[0]] : p.payDays.slice(0, 2) };
  return null;
}

/** Paychecks in a month. With biweekly pay the first two are "regular"; a third is an extra paycheck with no bills. */
function paydaysInMonth(key) {
  const cfg = payConfig(); if (!cfg) return [];
  const start = parseDate(key + '-01'), end = new Date(start.getFullYear(), start.getMonth() + 1, 1);
  let dates = [];
  if (cfg.step) {
    const n = Math.ceil(dayDiff(cfg.anchor, start) / cfg.step);
    for (let d = addDays(cfg.anchor, n * cfg.step); d < end; d = addDays(d, cfg.step)) dates.push(d);
  } else {
    const seen = new Set();
    dates = cfg.days.map(day => clampDate(start.getFullYear(), start.getMonth(), day))
      .filter(d => !seen.has(+d) && seen.add(+d)).sort((a, b) => a - b);
  }
  return dates.map((date, i) => ({ date, regular: cfg.f !== 'biweekly' || i < 2, index: i }));
}

function paydaysBetween(s, e) {
  const out = [];
  for (let d = new Date(s.getFullYear(), s.getMonth(), 1); d < e; d = new Date(d.getFullYear(), d.getMonth() + 1, 1)) {
    paydaysInMonth(monthKey(d)).forEach(p => { if (p.date >= s && p.date < e) out.push(p); });
  }
  return out;
}

/** Number of paychecks per month that carry bills (biweekly: 2, extras are free). */
function regularPerMonth() {
  return { weekly: 52 / 12, biweekly: 2, semimonthly: 2, monthly: 1 }[state.profile?.payFrequency] || 2;
}

/** Take-home amount of one paycheck: recurring income logged at the profile's pay frequency. */
function paycheckAmount() {
  const f = state.profile?.payFrequency;
  return state.income.filter(i => i.frequency === f).reduce((s, i) => s + i.amount, 0);
}

/** Income in a typical month (biweekly counts two paychecks, not 2.17). */
function planIncomeMonthly() {
  const f = state.profile?.payFrequency;
  return state.income.filter(i => i.frequency !== 'once')
    .reduce((s, i) => s + (i.frequency === f ? i.amount * regularPerMonth() : monthlyAmount(i)), 0);
}

/* ---- debts & credit-card payoff */

const isCard = d => d.type === 'Credit Card';
const openCards = () => state.debts.filter(d => isCard(d) && d.balance > 0);
const billsMonthly = () => state.bills.reduce((s, b) => s + monthlyAmount(b), 0);
const debtMinimums = () => state.debts.filter(d => d.balance > 0).reduce((s, d) => s + (d.minPayment || 0), 0);

/** Everything left after bills, all minimums and everyday spending goes to credit cards. */
function cardBudget() {
  return cached('cardBudget', () => {
    const spending = state.cardPlan?.spending ?? null;
    const cardMins = openCards().reduce((s, d) => s + d.minPayment, 0);
    const available = round2(planIncomeMonthly() - billsMonthly() - debtMinimums() - (spending ?? 0));
    const extra = spending == null ? 0 : Math.max(0, available);
    return { active: spending != null && openCards().length > 0, spending, cardMins, available, extra, total: round2(cardMins + extra) };
  });
}

/**
 * Month-by-month credit card payoff. Every card gets its minimum; with `rollover` the rest of the
 * budget (and the minimums of cards already paid off) goes to the highest-APR card first ("avalanche"),
 * which is the fastest and cheapest way to clear the balances.
 */
function simulateCards(extra, rollover = true) {
  const cards = openCards().map(d => ({ id: d.id, lender: d.lender, apr: d.apr, min: d.minPayment, bal: d.balance, paidOff: null, interest: 0 }));
  const budget = cards.reduce((s, c) => s + c.min, 0) + extra;
  const months = [];
  let totalInterest = 0;
  for (let k = 0; k < 600 && cards.some(c => c.bal > 0.005); k++) {
    const payments = {};
    let remaining = budget;
    const open = cards.filter(c => c.bal > 0.005);
    open.forEach(c => { const i = c.bal * c.apr / 1200; c.bal += i; c.interest += i; totalInterest += i; });
    open.forEach(c => { const p = Math.min(c.min, c.bal); c.bal -= p; payments[c.id] = p; remaining -= p; });
    if (rollover) {
      open.filter(c => c.bal > 0.005).sort((a, b) => b.apr - a.apr || a.bal - b.bal).forEach(c => {
        if (remaining <= 0.005) return;
        const p = Math.min(remaining, c.bal); c.bal -= p; payments[c.id] += p; remaining -= p;
      });
    }
    open.forEach(c => { if (c.bal <= 0.005) { c.bal = 0; if (c.paidOff == null) c.paidOff = k; } });
    Object.keys(payments).forEach(id => { payments[id] = round2(payments[id]); });
    months.push({ k, payments, balances: Object.fromEntries(cards.map(c => [c.id, round2(c.bal)])), total: round2(Object.values(payments).reduce((a, b) => a + b, 0)) });
  }
  const done = cards.every(c => c.bal <= 0.005);
  cards.forEach(c => { c.interest = round2(c.interest); });
  return { months, cards, totalInterest: round2(totalInterest), done, monthsToFree: done ? months.length : null };
}

const cardPlan = () => cached('cardPlan', () => simulateCards(cardBudget().extra));

/** What to pay on a debt in month offset k (0 = this month). Credit cards follow the payoff plan when it's on. */
function debtPaymentFor(d, k) {
  if (d.balance <= 0) return 0;
  if (isCard(d) && cardBudget().active && k >= 0) {
    const m = cardPlan().months[k];
    return m ? (m.payments[d.id] || 0) : 0;
  }
  return d.minPayment || 0;
}

/* ---- paycheck split */

const obligations = k => billsMonthly() + state.debts.reduce((s, d) => s + debtPaymentFor(d, k), 0);
/** Equal share of the month's bills + debt payments that every regular paycheck sets aside. */
const contribution = k => round2(obligations(k) / regularPerMonth());
const sinkingMonthly = () => state.bills.filter(b => ['quarterly', 'yearly'].includes(b.frequency)).reduce((s, b) => s + monthlyAmount(b), 0);

/** Every individual bill/debt payment due in [s, e), with the amount and due date. */
function paymentEvents(s, e) {
  const out = [];
  const push = (date, name, kind, amount) => { if (amount > 0.004 && date >= s && date < e) out.push({ date, name, kind, amount: round2(amount) }); };
  for (let d = new Date(s.getFullYear(), s.getMonth(), 1); d < e; d = new Date(d.getFullYear(), d.getMonth() + 1, 1)) {
    const y = d.getFullYear(), m = d.getMonth(), last = new Date(y, m + 1, 0).getDate(), k = monthOffset(monthKey(d));
    state.bills.forEach(b => {
      if (b.frequency === 'monthly') push(clampDate(y, m, b.dueDay || last), b.name, b.category, b.amount);
      if (b.frequency === 'semimonthly') {
        const d1 = b.dueDay || 1;
        push(clampDate(y, m, d1), b.name, b.category, b.amount);
        push(clampDate(y, m, d1 <= 15 ? d1 + 15 : d1 - 15), b.name, b.category, b.amount);
      }
    });
    state.debts.forEach(dt => push(clampDate(y, m, dt.dueDay || last), dt.lender, dt.type, debtPaymentFor(dt, k)));
  }
  // weekly / every-2-weeks bills are paid as they come, out of each regular paycheck
  const payAsYouGo = state.bills.filter(b => b.frequency === 'weekly' || b.frequency === 'biweekly');
  if (payAsYouGo.length) paydaysBetween(s, e).filter(p => p.regular).forEach(p =>
    payAsYouGo.forEach(b => push(p.date, b.name, b.category, monthlyAmount(b) / regularPerMonth())));
  return out.sort((a, b) => a.date - b.date || b.amount - a.amount);
}

/**
 * Splitting bills evenly means some bills come due before the paycheck that funds them.
 * Simulate a "bills money" balance for 12 months and return how much needs to be in it today
 * so it never goes negative.
 */
function cushionNeeded() {
  return cached('cushion', () => {
    if (!payConfig()) return 0;
    // measured from the first payday of this month, as if the plan started then
    const s = paydaysInMonth(monthKey(new Date()))[0]?.date || startOfToday(), e = addMonths(s, 12);
    const flows = [
      ...paydaysBetween(s, e).filter(p => p.regular).map(p => ({ date: p.date, amt: contribution(monthOffset(monthKey(p.date))) })),
      ...paymentEvents(s, e).map(x => ({ date: x.date, amt: -x.amount })),
    ].sort((a, b) => a.date - b.date || b.amt - a.amt);
    let bal = 0, min = 0;
    flows.forEach(f => { bal += f.amt; min = Math.min(min, bal); });
    return round2(-min);
  });
}

/* ---- monthly overview */

function summary(key = ui.month) {
  const k = monthOffset(key);
  const pays = payConfig() ? paydaysInMonth(key) : null;
  const payFreq = state.profile?.payFrequency;
  const recurringIncomeMonthly = state.income.filter(i => i.frequency !== 'once')
    .reduce((s, i) => s + (pays && i.frequency === payFreq ? i.amount * pays.length : monthlyAmount(i)), 0);
  const oneOffIncome = state.income.filter(i => i.frequency === 'once' && i.date?.startsWith(key)).reduce((s, i) => s + i.amount, 0);
  const totalIncome = recurringIncomeMonthly + oneOffIncome;

  const bills = billsMonthly();
  const debtPayments = state.debts.reduce((s, d) => s + debtPaymentFor(d, k), 0);
  const committed = bills + debtPayments;

  const monthExpenses = state.expenses.filter(e => e.date?.startsWith(key));
  const spent = monthExpenses.reduce((s, e) => s + e.amount, 0);

  const plannedLeft = totalIncome - committed;
  const expendable = plannedLeft - spent;

  const isCurrent = k === 0;
  const daysLeft = isCurrent ? daysInMonth(key) - new Date().getDate() + 1 : null;
  const perDay = isCurrent && daysLeft > 0 ? expendable / daysLeft : null;

  const byCategory = {};
  monthExpenses.forEach(e => { byCategory[e.category] = (byCategory[e.category] || 0) + e.amount; });

  const totalDebt = state.debts.reduce((s, d) => s + d.balance, 0);
  const monthlyInterest = state.debts.reduce((s, d) => s + debtStats(d).interest, 0);
  const weightedApr = totalDebt ? state.debts.reduce((s, d) => s + d.apr * d.balance, 0) / totalDebt : 0;

  return {
    recurringIncomeMonthly, oneOffIncome, totalIncome, billsMonthly: bills, debtPayments, debtMinimums: debtMinimums(), committed, spent, plannedLeft,
    expendable, daysLeft, perDay, byCategory, totalDebt, monthlyInterest, weightedApr, monthExpenses, paychecks: pays ? pays.length : null,
    cardPlanOn: cardBudget().active && k >= 0,
  };
}

function upcoming(days = 14) {
  const s = startOfToday();
  return paymentEvents(s, addDays(s, days + 1));
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
  const titles = { home: 'Overview', plan: 'Paycheck Plan', money: 'Income & Spending', bills: 'Recurring Bills', debts: 'Debts', profile: 'Profile' };
  $('#page-title').textContent = titles[ui.tab];
  ({ home: renderHome, plan: renderPlan, money: renderMoney, bills: renderBills, debts: renderDebts, profile: renderProfile })[ui.tab]();
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
    ${paycheckTeaser()}

    <section class="card">
      <h2>How it adds up</h2>
      <div class="rows">
        <div class="row"><div class="l">Recurring income <div class="s">${s.paychecks != null ? `${s.paychecks} paycheck${s.paychecks === 1 ? '' : 's'} this month` : 'monthly equivalent'}</div></div><div class="r pos">+${money(s.recurringIncomeMonthly)}</div></div>
        ${s.oneOffIncome ? `<div class="row"><div class="l">One-time income</div><div class="r pos">+${money(s.oneOffIncome)}</div></div>` : ''}
        <div class="row"><div class="l">Recurring bills</div><div class="r">−${money(s.billsMonthly)}</div></div>
        <div class="row"><div class="l">Debt payments <div class="s">${s.cardPlanOn ? 'minimums + extra toward credit cards' : 'minimum payments'}</div></div><div class="r">−${money(s.debtPayments)}</div></div>
        <div class="row"><div class="l">Spending this month</div><div class="r">−${money(s.spent)}</div></div>
        <div class="row total"><div class="l">Expendable</div><div class="r ${cls}">${money(s.expendable)}</div></div>
      </div>
      ${s.totalIncome === 0 ? `<p class="note">Add your paycheck under <b>Money → Income</b> to see a real number.</p>` : ''}
    </section>

    <section class="card">
      <h2>Due in the next 14 days</h2>
      ${up.length ? `<div class="rows">${up.map(u => `
        <div class="row"><div class="l"><div class="t">${esc(u.name)}</div><div class="s">${esc(u.kind)} · ${shortDate(u.date)}</div></div>
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
  const never = state.debts.filter(d => d.balance > 0 && debtStats(d).months == null);
  const warn = never.length ? `<p class="note neg">⚠︎ Minimum payment doesn't cover interest on: ${never.map(d => esc(d.lender)).join(', ')}. That balance will grow.</p>` : '';
  if (!openCards().length) return warn;
  const cb = cardBudget();
  if (cb.active) {
    const plan = cardPlan();
    return `<p class="note">💳 Card plan: pay <b>${money(cb.total)}/month</b> on credit cards —
      ${plan.done ? `debt-free by <b>${offsetLabel(plan.monthsToFree - 1)}</b>.` : 'not enough to pay them off yet.'}
      <a href="#" data-act="go-cards">See plan</a></p>${warn}`;
  }
  return `<p class="note">💡 Find out the fastest way to pay off your credit cards. <a href="#" data-act="go-cards">Build my payoff plan</a></p>${warn}`;
}

function paycheckTeaser() {
  if (!payConfig()) return `<section class="card"><h2>Paycheck plan</h2>
    <p class="note" style="margin:0">Tell us your paydays and we'll split every bill evenly across your paychecks.</p>
    <button type="button" class="secondary block" data-act="edit-profile">Set up paydays</button></section>`;
  const s = startOfToday();
  const next = paydaysBetween(s, addDays(s, 40))[0];
  if (!next) return '';
  const k = monthOffset(monthKey(next.date));
  const share = next.regular ? contribution(k) : 0;
  return `<section class="card tap" data-act="go-plan">
    <h2>Next paycheck · ${shortDate(next.date)}</h2>
    <div class="rows">
      <div class="row"><div class="l">Set aside for bills & debts</div><div class="r">${money(share)}</div></div>
      <div class="row total"><div class="l">Yours to spend</div><div class="r pos">${money(paycheckAmount() - share)}</div></div>
    </div>
    <p class="note" style="margin-bottom:0">${next.regular ? 'Tap to see what to pay and when ›' : '🎉 Extra paycheck — no bills come out of this one. Tap for details ›'}</p>
  </section>`;
}

/* ---- Plan tab: every bill split evenly across paychecks, with what to pay and when */

function renderPlan() {
  const cfg = payConfig();
  const freq = state.profile.payFrequency;
  if (!cfg) {
    view.innerHTML = `<section class="card">
      <h2>Set up your paydays</h2>
      <p class="muted" style="margin-top:0">${['quarterly', 'yearly'].includes(freq)
        ? 'Paycheck plans work with weekly, every-2-weeks, twice-a-month or monthly pay. Change how you get paid in your profile.'
        : 'Add your payday so the app can split your bills and debt payments evenly between your paychecks, and tell you what to pay from each one.'}</p>
      <button type="button" class="block" data-act="edit-profile">Set paydays</button></section>`;
    return;
  }
  const key = ui.month, k = monthOffset(key);
  const pays = paydaysInMonth(key);
  const P = paycheckAmount(), T = contribution(k), R = regularPerMonth();
  const total = obligations(k), sink = sinkingMonthly(), cushion = cushionNeeded();
  const nextMonthFirst = paydaysInMonth(shiftMonth(key, 1))[0]?.date;
  const monthStart = parseDate(key + '-01');
  const before = pays.length ? paymentEvents(monthStart, pays[0].date) : [];
  const target = [...openCards()].sort((a, b) => b.apr - a.apr)[0];
  const cb = cardBudget();

  const evRows = evs => evs.map(e => `
    <div class="row"><div class="l"><div class="t">${esc(e.name)}</div><div class="s">${esc(e.kind)} · pay by ${shortDate(e.date)}</div></div>
    <div class="r">${money(e.amount)}</div></div>`).join('');

  view.innerHTML = `
    ${monthNav()}
    <section class="card hero">
      <div class="label">From each ${freq === 'weekly' ? '' : 'regular '}paycheck, set aside</div>
      <div class="big">${money(T)}</div>
      <div class="sub">for bills & debt payments · keep <b class="${P - T >= 0 ? 'pos' : 'neg'}">${money(P - T)}</b> to spend</div>
    </section>
    <section class="card">
      <h2>How the split works</h2>
      <div class="rows">
        <div class="row"><div class="l">Bills this month</div><div class="r">${money(billsMonthly())}</div></div>
        <div class="row"><div class="l">Debt payments${cb.active && k >= 0 ? ' <div class="s">incl. credit card payoff plan</div>' : ''}</div><div class="r">${money(total - billsMonthly())}</div></div>
        <div class="row total"><div class="l">Total ÷ ${freq === 'weekly' ? '4.33 (avg. weekly paychecks)' : R + ' paychecks'}</div><div class="r">${money(T)} each</div></div>
      </div>
      ${P === 0 ? `<p class="note neg">Add your paycheck as recurring income (${esc(FREQUENCIES[freq].label.toLowerCase())}) under Money → Income.</p>` : ''}
      ${P > 0 && P < T ? `<p class="note neg">⚠︎ One paycheck (${money(P)}) is less than its share of the bills.</p>` : ''}
      ${sink > 0 ? `<p class="note">Includes ${money(sink)}/mo saved toward quarterly & yearly bills, so they never land on one paycheck.</p>` : ''}
      <p class="note">Every paycheck puts the same amount into your <b>bills money</b> (a separate savings or checking account works great), and bills are paid out of it on their due dates — so no half of the month is harder than the other.</p>
      ${cushion > 0 ? `<p class="note warn">Keep a one-time cushion of about <b>${money(cushion)}</b> in your bills money. Some bills fall due before the paycheck that pays for them, and this covers that gap so equal paychecks always work.</p>` : ''}
    </section>

    ${before.length ? `<section class="card">
      <h2>Due before your first ${esc(monthLabel(key).split(' ')[0])} paycheck</h2>
      <div class="rows">${evRows(before)}</div>
      <p class="note" style="margin-bottom:0">Covered by last month's paychecks.</p></section>` : ''}

    ${pays.map((p, i) => {
      const next = pays[i + 1]?.date || nextMonthFirst || addDays(p.date, 14);
      const evs = paymentEvents(p.date, next);
      const due = evs.reduce((s, e) => s + e.amount, 0);
      const share = p.regular ? T : 0;
      const diff = round2(share - due);
      return `<section class="card">
        <h2>${shortDate(p.date)} · ${p.regular ? `Paycheck ${i + 1}` : 'Extra paycheck 🎉'}</h2>
        <div class="rows">
          <div class="row"><div class="l">Paycheck</div><div class="r pos">+${money(P)}</div></div>
          <div class="row"><div class="l">Set aside for bills & debts</div><div class="r">−${money(share)}</div></div>
          <div class="row total"><div class="l">Left to spend</div><div class="r ${P - share >= 0 ? 'pos' : 'neg'}">${money(P - share)}</div></div>
        </div>
        <h2 style="margin-top:14px">Pay before ${shortDate(next)}</h2>
        ${evs.length ? `<div class="rows">${evRows(evs)}
          <div class="row total"><div class="l">Due this period</div><div class="r">${money(due)}</div></div></div>
          ${p.regular ? `<p class="note" style="margin-bottom:0">${Math.abs(diff) < 0.01 ? 'Exactly this paycheck’s share.'
            : diff > 0 ? `Keep the other ${money(diff)} you set aside in your bills money — it pays for bills due after your next paycheck.`
            : `${money(-diff)} of this comes from bills money set aside from earlier paychecks.`}</p>`
            : '<p class="note" style="margin-bottom:0">Paid from bills money you already set aside.</p>'}`
          : `<div class="empty">Nothing due in this stretch.</div>`}
        ${!p.regular ? `<p class="note">This paycheck carries no bills.${target ? ` Putting it toward <b>${esc(target.lender)}</b> (${target.apr}% APR) is the fastest way to cut your card debt.` : ' Consider saving it.'}</p>` : ''}
      </section>`;
    }).join('')}
    ${!pays.length ? `<section class="card"><div class="empty">No paychecks this month.</div></section>` : ''}
    <p class="note" style="text-align:center">Pay dates assume you're paid on the scheduled day. Move a payment earlier if a payday falls on a weekend or holiday.</p>`;
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
  const seg = `<div class="segmented" role="tablist">
      <button type="button" data-act="debt-seg" data-seg="all" class="${ui.debtSeg === 'all' ? 'active' : ''}">All debts</button>
      <button type="button" data-act="debt-seg" data-seg="cards" class="${ui.debtSeg === 'cards' ? 'active' : ''}">Credit card payoff</button>
    </div>`;
  if (ui.debtSeg === 'cards') { view.innerHTML = seg + cardPlanView(); return; }
  const groups = {};
  state.debts.forEach(d => (groups[d.type] = groups[d.type] || []).push(d));
  const s = summary();
  view.innerHTML = `
    ${seg}
    ${state.debts.length ? `
    <section class="card">
      <div class="rows">
        <div class="row total"><div class="l">Total owed</div><div class="r">${money(s.totalDebt)}</div></div>
        <div class="row"><div class="l">Minimum payments / month</div><div class="r">${money(debtMinimums())}</div></div>
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

/* ---- Credit card payoff: the largest affordable payment, paid highest-APR first */

function suggestedSpending() {
  const totals = [];
  for (let i = 1; i <= 3; i++) {
    const key = shiftMonth(monthKey(new Date()), -i);
    const t = state.expenses.filter(e => e.date?.startsWith(key)).reduce((s, e) => s + e.amount, 0);
    if (t > 0) totals.push(t);
  }
  return totals.length ? round2(totals.reduce((a, b) => a + b, 0) / totals.length) : null;
}

function cardPlanView() {
  const cards = openCards();
  if (!cards.length) return `<section class="card"><div class="empty">No credit card balances. Add a debt with type “Credit Card” to build a payoff plan.</div></section>`;
  const cb = cardBudget();
  const avg = suggestedSpending();
  const spendingNow = cb.spending ?? avg ?? '';
  const freeBeforeSpending = round2(planIncomeMonthly() - billsMonthly() - debtMinimums());

  const setup = `<section class="card">
    <h2>Money available for credit cards</h2>
    <div class="rows">
      <div class="row"><div class="l">Income in a typical month${state.profile.payFrequency === 'biweekly' ? '<div class="s">2 paychecks — extra paychecks are a bonus</div>' : ''}</div><div class="r pos">+${money(planIncomeMonthly())}</div></div>
      <div class="row"><div class="l">Recurring bills</div><div class="r">−${money(billsMonthly())}</div></div>
      <div class="row"><div class="l">All debt minimums</div><div class="r">−${money(debtMinimums())}</div></div>
      <div class="row"><div class="l">Everyday spending<div class="s">groceries, gas, fun…</div></div><div class="r">−${money(num(spendingNow))}</div></div>
      <div class="row total"><div class="l">Extra for credit cards</div><div class="r ${freeBeforeSpending - num(spendingNow) > 0 ? 'pos' : 'neg'}">${money(Math.max(0, freeBeforeSpending - num(spendingNow)))}</div></div>
    </div>
    <label class="field"><span>How much do you need for everyday spending each month?</span>
      <input id="spending-input" inputmode="decimal" placeholder="${avg != null ? avg.toFixed(2) : '0.00'}" value="${cb.spending ?? ''}"></label>
    <p class="note">${avg != null ? `You've averaged ${money(avg)}/month in logged expenses recently. ` : ''}Everything above that goes to your cards — that's what makes payoff as fast as possible.</p>
    <div class="actions">
      <button type="button" data-act="save-spending">${cb.active ? 'Update plan' : 'Build my payoff plan'}</button>
      ${cb.active ? '<button type="button" class="secondary" data-act="plan-off">Turn off</button>' : ''}
    </div>
  </section>`;

  if (!cb.active) {
    return setup + `<section class="card"><h2>Paying minimums only</h2><div class="rows">${cards.map(d => {
      const st = debtStats(d);
      return `<div class="row"><div class="l"><div class="t">${esc(d.lender)}</div><div class="s">${d.apr}% APR · min ${money(d.minPayment)}</div></div>
        <div class="r">${st.months == null ? '<span class="neg">never</span>' : `${st.months} mo`}<div class="s">${st.totalInterest == null ? '' : money(st.totalInterest, { whole: true }) + ' interest'}</div></div></div>`;
    }).join('')}</div></section>`;
  }

  const plan = cardPlan();
  const base = simulateCards(0, false);
  const saved = base.done ? round2(base.totalInterest - plan.totalInterest) : null;
  const order = [...plan.cards].sort((a, b) => (a.paidOff ?? 1e9) - (b.paidOff ?? 1e9));
  const focus = [...cards].sort((a, b) => b.apr - a.apr || a.balance - b.balance)[0];
  const shown = plan.months.slice(0, 60);

  return `
    <section class="card hero">
      <div class="label">Pay on your credit cards every month</div>
      <div class="big">${money(cb.total)}</div>
      <div class="sub">${money(cb.cardMins)} minimums + ${money(cb.extra)} extra</div>
      <div class="sub" style="margin-top:8px">${plan.done
        ? `Debt-free in <b>${plan.monthsToFree} month${plan.monthsToFree === 1 ? '' : 's'}</b> — by <b>${offsetLabel(plan.monthsToFree - 1)}</b>`
        : '<span class="neg">This isn’t enough to ever pay them off — lower everyday spending or bills.</span>'}</div>
    </section>
    ${cb.extra === 0 ? `<section class="card"><p class="note neg" style="margin:0">⚠︎ After bills, minimums and everyday spending there's nothing extra left (${money(cb.available)}). The plan pays minimums only, rolling each paid-off card's minimum into the next.</p></section>` : ''}

    <section class="card">
      <h2>This month, pay</h2>
      <div class="rows">${cards.map(d => {
        const pay = plan.months[0]?.payments[d.id] || 0;
        const pc = plan.cards.find(c => c.id === d.id);
        return `<div class="row"><div class="l"><div class="t">${esc(d.lender)} ${d.id === focus.id ? '<span class="chip">focus</span>' : ''}</div>
          <div class="s">${d.apr}% APR · due ${d.dueDay ? ordinal(d.dueDay) : 'end of month'} · ${pc.paidOff != null ? `paid off ${offsetLabel(pc.paidOff)}` : 'not paid off'}</div></div>
          <div class="r">${money(pay)}<div class="s">min ${money(d.minPayment)}</div></div></div>`;
      }).join('')}</div>
      <p class="note">Every card gets its minimum, and every extra dollar goes to <b>${esc(focus.lender)}</b> — the highest interest rate. When a card is paid off, its payment rolls to the next one. This is the fastest, cheapest order (the “avalanche” method).</p>
    </section>

    <section class="card">
      <h2>Plan vs. minimums only</h2>
      <div class="rows">
        <div class="row"><div class="l">Debt-free</div><div class="r">${plan.done ? offsetLabel(plan.monthsToFree - 1) : 'never'} <span class="muted">vs</span> ${base.done ? offsetLabel(base.monthsToFree - 1) : 'never'}</div></div>
        <div class="row"><div class="l">Interest paid</div><div class="r">${money(plan.totalInterest, { whole: true })} <span class="muted">vs</span> ${base.done ? money(base.totalInterest, { whole: true }) : '∞'}</div></div>
        ${plan.done && base.done ? `<div class="row total"><div class="l">You save</div><div class="r pos">${money(saved, { whole: true })} · ${base.monthsToFree - plan.monthsToFree} months</div></div>` : ''}
      </div>
    </section>

    <section class="card">
      <h2>Payoff order</h2>
      <div class="rows">${order.map((c, i) => `<div class="row"><div class="l"><div class="t">${i + 1}. ${esc(c.lender)}</div><div class="s">${c.apr}% APR · ${money(c.interest, { whole: true })} interest</div></div>
        <div class="r">${c.paidOff != null ? offsetLabel(c.paidOff) : '—'}</div></div>`).join('')}</div>
    </section>

    <section class="card">
      <details>
        <summary><b>Month-by-month schedule</b></summary>
        <div class="tablewrap"><table>
          <thead><tr><th>Month</th>${cards.map(d => `<th>${esc(d.lender)}</th>`).join('')}<th>Total</th><th>Left owed</th></tr></thead>
          <tbody>${shown.map(m => `<tr><td>${offsetLabel(m.k)}</td>${cards.map(d => `<td>${m.payments[d.id] ? money(m.payments[d.id], { whole: true }) : '—'}</td>`).join('')}
            <td><b>${money(m.total, { whole: true })}</b></td><td>${money(Object.values(m.balances).reduce((a, b) => a + b, 0), { whole: true })}</td></tr>`).join('')}</tbody>
        </table></div>
        ${plan.months.length > shown.length ? `<p class="note">Showing the first 5 years.</p>` : ''}
      </details>
    </section>
    ${state.profile.payFrequency === 'biweekly' ? `<p class="note">💡 Twice a year you get a 3rd paycheck in a month. Sending it to ${esc(focus.lender)} gets you out even sooner.</p>` : ''}
    <p class="note">These amounts are already built into your Paycheck Plan and Overview. Update card balances now and then to keep the plan accurate.</p>
    ${setup}`;
}

function paydayText() {
  const cfg = payConfig();
  if (!cfg) return ' · <span class="warn">paydays not set</span>';
  if (cfg.step) { const n = paydaysBetween(startOfToday(), addDays(startOfToday(), 15))[0]; return n ? ` · next payday ${shortDate(n.date)}` : ''; }
  return ' · on the ' + cfg.days.map(ordinal).join(' & ');
}

function renderProfile() {
  const p = state.profile;
  const counts = `${state.debts.length} debts · ${state.bills.length} bills · ${state.income.length} income · ${state.expenses.length} expenses`;
  view.innerHTML = `
    <section class="card">
      <h3>${esc(p.name || 'Me')}</h3>
      <div class="s muted">${esc(p.currency)} · paid ${esc(FREQUENCIES[p.payFrequency]?.label.toLowerCase() || '—')}${paydayText()}</div>
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
function openSheet({ title, body, onSubmit, onDelete, onChange }) {
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
  sheetForm.onchange = ev => onChange && onChange(ev);
  sheet.showModal();
  sheetForm.scrollTop = 0;
  if (onChange) onChange({ target: sheetForm });
}

function profileForm(isNew) {
  const p = state.profile || { name: '', currency: 'USD', payFrequency: 'biweekly' };
  const days = p.payDays || [1, 15];
  const body = `
    ${field('Your name', input('name', p.name, 'autocomplete="given-name" placeholder="e.g. Alex"'))}
    <div class="grid2">
      ${field('Currency', `<select name="currency">${options(CURRENCIES, p.currency)}</select>`)}
      ${field('I get paid', `<select name="payFrequency">${freqOptions(RECURRING_FREQS, p.payFrequency)}</select>`)}
    </div>
    <div data-freqs="weekly biweekly">
      ${field('A recent or upcoming payday', input('payAnchor', p.payAnchor || '', 'type="date"'))}
      <p class="note">Used to work out all your future paydays.</p>
    </div>
    <div data-freqs="semimonthly" class="grid2">
      ${field('First payday (day of month)', input('payDay1', days[0], 'inputmode="numeric" placeholder="1"'))}
      ${field('Second payday', input('payDay2', days[1] ?? 15, 'inputmode="numeric" placeholder="15"'))}
    </div>
    <div data-freqs="monthly">
      ${field('Payday (day of month)', input('payDayM', days[0], 'inputmode="numeric" placeholder="1"'))}
    </div>
    ${isNew ? `${field('Take-home pay per paycheck (after tax)', moneyInput('pay', ''))}
      <p class="note">Optional — this becomes a recurring income entry. You can add other income later.</p>` : ''}`;
  openSheet({
    title: isNew ? 'Step 1 of 3 · Your profile' : 'Edit profile',
    body,
    onChange: () => {
      const f = sheetForm.elements.payFrequency.value;
      sheetForm.querySelectorAll('[data-freqs]').forEach(el => { el.hidden = !el.dataset.freqs.split(' ').includes(f); });
    },
    onSubmit: d => {
      if (!d.name.trim()) return 'Please enter your name.';
      const f = d.payFrequency;
      const day = v => { const n = parseInt(v, 10); return n >= 1 && n <= 31 ? n : null; };
      let payAnchor = p.payAnchor || null, payDays = p.payDays || null;
      if (f === 'weekly' || f === 'biweekly') payAnchor = d.payAnchor || null;
      if (f === 'semimonthly') {
        const a = day(d.payDay1), b = day(d.payDay2);
        if (!a || !b || a === b) return 'Enter two different paydays between 1 and 31.';
        payDays = [Math.min(a, b), Math.max(a, b)];
      }
      if (f === 'monthly') { const a = day(d.payDayM); if (!a) return 'Enter a payday between 1 and 31.'; payDays = [a]; }
      // keep the paycheck income in step with a change of pay frequency
      if (!isNew && state.profile && state.profile.payFrequency !== f) {
        state.income.forEach(i => { if (i.frequency === state.profile.payFrequency && /pay ?check|salary/i.test(i.source)) i.frequency = f; });
      }
      const createdAt = state.profile?.createdAt || today();
      state.profile = { name: d.name.trim(), currency: d.currency, payFrequency: f, payAnchor, payDays, createdAt };
      if (isNew && num(d.pay) > 0) {
        state.income.push({ id: uid(), source: 'Paycheck', amount: round2(num(d.pay)), frequency: f, date: payAnchor || today(), note: '' });
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
    ['Recurring income this month', round2(s.recurringIncomeMonthly)],
    ['One-time income this month', round2(s.oneOffIncome), `SUMPRODUCT((LEFT(Income!B2:B1000,7)=B3)*(Income!E2:E1000="once")*Income!D2:D1000)`],
    ['Recurring bills (monthly)', round2(s.billsMonthly), 'SUM(Bills!H2:H1000)'],
    [s.cardPlanOn ? 'Debt payments (incl. card payoff plan)' : 'Debt minimum payments', round2(s.debtPayments), s.cardPlanOn ? null : 'SUM(Debts!F2:F1000)'],
    ['Spent this month', round2(s.spent), `SUMPRODUCT((LEFT(Expenses!B2:B5000,7)=B3)*Expenses!E2:E5000)`],
    ['Expendable this month', round2(s.expendable), 'B5+B6-B7-B8-B9'],
    [],
    ['Total debt', round2(s.totalDebt), 'SUM(Debts!D2:D1000)'],
    ['Interest accruing / month', round2(s.monthlyInterest), 'SUM(Debts!I2:I1000)'],
    [],
    ['Set aside per regular paycheck', contribution(0)],
    ['Left to spend per paycheck', round2(paycheckAmount() - contribution(0))],
    ['Bills-money cushion (one-time)', cushionNeeded()],
    ['Credit cards: monthly payment', cardBudget().active ? cardBudget().total : ''],
    ['Credit cards: debt-free by', cardBudget().active && cardPlan().done ? offsetLabel(cardPlan().monthsToFree - 1) : ''],
  ];
  const ws = XL.utils.aoa_to_sheet(summaryRows.map(r => r.slice(0, 2)));
  summaryRows.forEach((r, i) => { if (r[2]) ws[`B${i + 1}`] = { t: 'n', v: r[1], f: r[2], z: '#,##0.00' }; });
  ws['!cols'] = [{ wch: 30 }, { wch: 18 }];
  XL.utils.book_append_sheet(wb, ws, 'Summary');

  // Profile
  const p = state.profile || {};
  const wp = XL.utils.aoa_to_sheet([['Field', 'Value'], ['Name', p.name || ''], ['Currency', p.currency || ''], ['Pay Frequency', p.payFrequency || ''],
    ['Next Payday', p.payAnchor || ''], ['Pay Days', (p.payDays || []).join(',')], ['Everyday Spending', state.cardPlan?.spending ?? ''], ['Created', p.createdAt || '']]);
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

  // Paycheck Plan — next 3 months, one row per payment, grouped by paycheck
  const pp = [['Paycheck Date', 'Paycheck', 'Paycheck Amount', 'Set Aside for Bills', 'Left to Spend', 'Pay', 'Category', 'Pay By', 'Amount']];
  if (payConfig()) {
    const P = paycheckAmount();
    for (let k = 0; k < 3; k++) {
      const key = shiftMonth(monthKey(new Date()), k);
      const pays = paydaysInMonth(key), nextFirst = paydaysInMonth(shiftMonth(key, 1))[0]?.date;
      pays.forEach((p, i) => {
        const share = p.regular ? contribution(k) : 0;
        const head = [isoDate(p.date), p.regular ? `Paycheck ${i + 1}` : 'Extra paycheck', round2(P), share, round2(P - share)];
        const evs = paymentEvents(p.date, pays[i + 1]?.date || nextFirst || addDays(p.date, 14));
        if (!evs.length) pp.push([...head, '(nothing due)', '', '', '']);
        evs.forEach((e, j) => pp.push([...(j ? ['', '', '', '', ''] : head), e.name, e.kind, isoDate(e.date), e.amount]));
      });
    }
  } else pp.push(['Set your paydays in the app to generate this plan.']);
  const wpp = XL.utils.aoa_to_sheet(pp);
  wpp['!cols'] = [12, 14, 15, 18, 13, 22, 16, 12, 12].map(wch => ({ wch }));
  XL.utils.book_append_sheet(wb, wpp, 'Paycheck Plan');

  // Card Payoff — month-by-month avalanche schedule
  const cards = openCards();
  const cp = [['Month', ...cards.flatMap(d => [`${d.lender} Payment`, `${d.lender} Balance`]), 'Total Payment', 'Total Owed']];
  if (cardBudget().active) {
    cardPlan().months.forEach(m => cp.push([offsetLabel(m.k), ...cards.flatMap(d => [m.payments[d.id] || 0, m.balances[d.id]]), m.total,
      round2(Object.values(m.balances).reduce((a, b) => a + b, 0))]));
  } else cp.push(['Build a credit card payoff plan in the app (Debts → Credit card payoff) to generate this schedule.']);
  const wcp = XL.utils.aoa_to_sheet(cp);
  wcp['!cols'] = cp[0].map(() => ({ wch: 16 }));
  XL.utils.book_append_sheet(wb, wcp, 'Card Payoff');
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
      const payDays = String(pv['Pay Days'] ?? '').split(',').map(n => parseInt(n, 10)).filter(n => n >= 1 && n <= 31);
      next.profile = { name: String(pv.Name || 'Me'), currency: CURRENCIES.includes(pv.Currency) ? pv.Currency : 'USD', payFrequency: FREQUENCIES[pv['Pay Frequency']] ? pv['Pay Frequency'] : 'biweekly',
        payAnchor: pv['Next Payday'] ? excelDateToIso(pv['Next Payday']) : null, payDays: payDays.length ? payDays : null, createdAt: String(pv.Created || today()) };
      next.cardPlan = { spending: pv['Everyday Spending'] === undefined || pv['Everyday Spending'] === '' ? null : round2(num(pv['Everyday Spending'])) };
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
    case 'debt-seg': ui.debtSeg = el.dataset.seg; render(); break;
    case 'go-cards': ev.preventDefault(); ui.tab = 'debts'; ui.debtSeg = 'cards'; render(); window.scrollTo(0, 0); break;
    case 'go-plan': ui.tab = 'plan'; render(); window.scrollTo(0, 0); break;
    case 'save-spending': {
      const v = $('#spending-input').value.trim();
      const val = v === '' ? suggestedSpending() : num(v);
      if (val == null || val < 0) { toast('Enter how much you need for everyday spending'); break; }
      state.cardPlan = { spending: round2(val) }; save(); render(); window.scrollTo(0, 0);
      toast('Payoff plan updated'); break;
    }
    case 'plan-off': state.cardPlan = { spending: null }; save(); render(); break;
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
        localStorage.removeItem(STORE_KEY); state = emptyState(); memo = {}; ui = { tab: 'home', month: monthKey(new Date()), moneySeg: 'expenses', debtSeg: 'all', onboarding: null }; render();
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
