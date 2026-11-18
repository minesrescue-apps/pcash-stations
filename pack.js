/**
 * PCash Vs IP – month-end pack engine.
 * Turns the app's entries for one ReconMonth into the MRS petty cash form layout:
 * voucher list (expenses), income block, CF/PC summary and the station funds reconciliation.
 * Pure function: no DOM, no network. Works in the browser (window.PCashPack) and in Node (require).
 */
(function (root) {
  const num = v => { const n = parseFloat(String(v == null ? '' : v).replace(/[^0-9.\-]/g, '')); return isFinite(n) ? n : 0; };
  const r2 = n => Math.round((n + Number.EPSILON) * 100) / 100;
  const has = v => v !== null && v !== undefined && String(v).trim() !== '';
  const dmy = iso => { const p = String(iso || '').split('-'); return p.length === 3 ? p[2] + '.' + p[1] + '.' + p[0] : String(iso || ''); };
  const up = s => String(s || '').toUpperCase();

  /** Kitchen control forms carry the card-machine takings in the note, e.g. "Swipes R10143". */
  function swipesFromNote(note) {
    const m = /swipes?\s*:?\s*R?\s*([0-9][0-9\s,]*(?:\.[0-9]{1,2})?)/i.exec(String(note || ''));
    return m ? num(m[1].replace(/[\s,]/g, '')) : 0;
  }

  function byDateThenCapture(a, b) {
    const k = e => (e.Date || '') + String(e.Voucher || '').padStart(6, '0') + String(e.Line || '').padStart(3, '0') + (e.Created || '');
    return k(a).localeCompare(k(b));
  }

  /**
   * Which section of the voucher list an entry belongs to.
   *  1 = bank side (card swipes, bank fees, bank-side outs)  -> SWIPES TOTAL column
   *  2 = cash box (cash / own-card slips, cash outs)          -> CASH TOTAL column
   *  3 = EFT from the main MRS account                         -> PC/CF EFT PURCHASES column
   */
  function sectionOf(e) {
    if (e.Kind === 'SLIP') return e.Side === 'CARD' ? 1 : e.Side === 'EFT' ? 3 : 2;
    if (e.Kind === 'OUT') return e.Side === 'BANK' ? 1 : 2;
    if (e.Kind === 'IN' && e.Type === 'CASH_EXCHANGE') return e.Direction === 'CASH_IN' ? 1 : 2;
    return 0;
  }

  const OUT_ARTICLE = {
    STAFF_WITHDRAWAL: 'STAFF WITHDRAWAL', CASH_DRAW: 'CASH DRAW (REPAID TO MAIN ACC)',
    REIMBURSE: 'OWN CARD REIMBURSED', BANK_FEE: '', OTHER_OUT: ''
  };

  function expenseLine(e) {
    const a = num(e.Amount), v = num(e.VAT);
    const L = { alloc: e.Alloc || 'PC', firm: up(e.Name), article: up(e.Note), cx: 0, cost: 0, vat: 0, cash: 0, swipe: 0, eft: 0 };
    if (e.Kind === 'SLIP') {
      L.cost = r2(a - v); L.vat = v;
      if (e.Side === 'CARD') L.swipe = a; else if (e.Side === 'EFT') L.eft = a; else L.cash = a;
      if (e.Side === 'OWN') L.article = ('OWN CARD ' + L.article).trim();
    } else if (e.Kind === 'OUT') {
      L.alloc = e.Alloc || 'PC';
      if (e.Type === 'BANK_FEE') { L.firm = (up(e.Name) || 'FNB') + ' SERVICE FEE'; L.article = ''; }
      else L.article = OUT_ARTICLE[e.Type] || L.article;
      if (e.Type === 'STAFF_WITHDRAWAL' || e.Type === 'CASH_DRAW') L.cx = a;
      L.cost = a;
      if (e.Side === 'BANK') L.swipe = a; else L.cash = a;
    } else { // cash exchange: money moves between bank and box
      L.alloc = 'PC'; L.firm = up(e.Name) || 'CASH EXCHANGE'; L.article = 'CASH EXCHANGE';
      L.cx = a; L.cost = a;
      if (e.Direction === 'CASH_IN') L.swipe = a; else L.cash = a;
    }
    return L;
  }

  function incomeLine(e) {
    const a = num(e.Amount);
    const L = { alloc: 'PC', firm: up(e.Name), article: dmy(e.Date), cx: 0, cost: a, vat: 0, total: 0, eft: 0, swipeNC: 0 };
    const bank = e.Side === 'BANK';
    switch (e.Type) {
      case 'KITCHEN_SALES':
        L.alloc = 'CF'; L.firm = 'KITCHEN SALES'; L.swipeNC = swipesFromNote(e.Note); break;
      case 'STORE_SALES':
        L.firm = 'STORES SALES'; L.article = ((e.Note && !/swipe/i.test(e.Note) ? up(e.Note) + ' ' : '') + dmy(e.Date)).trim(); break;
      case 'REPAYMENT':
        L.cx = a; L.article = 'WITHDRAWAL REPAID ' + dmy(e.Date); break;
      case 'TRANSFER':
        L.firm = up(e.Name) || 'TRANSFER FROM MAIN ACC'; L.article = 'TRANSFER ' + dmy(e.Date); break;
      case 'CASH_EXCHANGE':
        L.cx = a; L.firm = up(e.Name) || 'CASH EXCHANGE'; L.article = 'CASH EXCHANGE ' + dmy(e.Date);
        if (e.Direction === 'CASH_IN') { L.total = a; return L; }
        L.eft = a; return L;
      default:
        L.article = ((up(e.Note) + ' ') + dmy(e.Date)).trim();
    }
    if (bank) L.eft = a; else L.total = a;
    return L;
  }

  /**
   * @param entries  all rows from the Entries sheet
   * @param month    'YYYY-MM' (ReconMonth)
   * @param bal      { openBank, openCash, stmtClose, cashCount } numbers (null when unknown)
   */
  function build(entries, month, bal) {
    const live = (entries || []).filter(e => e.ReconMonth === month && e.Status !== 'VOID');

    // ---- expenses: bank side, then cash box, then EFT; date order inside each section
    const exp = live.filter(e => sectionOf(e) > 0)
      .sort((a, b) => (sectionOf(a) - sectionOf(b)) || byDateThenCapture(a, b));
    const expenses = [];
    let vno = 0, lastKey = null;
    exp.forEach(e => {
      // lines of one split slip share a voucher number (same app voucher + same section)
      const key = e.Kind === 'SLIP' && e.Voucher ? 'S' + e.Voucher + '|' + sectionOf(e) : 'E' + e.ID;
      if (key !== lastKey) vno++;
      lastKey = key;
      const L = expenseLine(e);
      L.voucher = vno; L.section = sectionOf(e); L.ids = [e.ID]; L.fileId = e.FileId || ''; L.date = e.Date;
      expenses.push(L);
    });

    // ---- income: kitchen sales first, then everything else; date order inside each
    const inc = live.filter(e => e.Kind === 'IN')
      .sort((a, b) => ((a.Type === 'KITCHEN_SALES' ? 0 : 1) - (b.Type === 'KITCHEN_SALES' ? 0 : 1)) || byDateThenCapture(a, b));
    const income = inc.map((e, i) => Object.assign(incomeLine(e), { no: i + 1, ids: [e.ID], fileId: e.FileId || '', date: e.Date }));

    // ---- totals
    const sum = (rows, k) => r2(rows.reduce((t, r) => t + (r[k] || 0), 0));
    const expT = {}; ['cx', 'cost', 'vat', 'cash', 'swipe', 'eft'].forEach(k => expT[k] = sum(expenses, k));
    expT.cashPlusSwipe = r2(expT.cash + expT.swipe);
    const incT = {}; ['cx', 'cost', 'vat', 'total', 'eft', 'swipeNC'].forEach(k => incT[k] = sum(income, k));

    const by = (rows, alloc, f) => r2(rows.filter(r => r.alloc === alloc).reduce((t, r) => t + f(r), 0));
    // same SUMIF logic as the form: money that moved (cash + swipe + EFT) per PC/CF
    const summary = {
      expCF: by(expenses, 'CF', r => r.cash + r.swipe + r.eft), expPC: by(expenses, 'PC', r => r.cash + r.swipe + r.eft),
      incCF: by(income, 'CF', r => r.total + r.eft + r.swipeNC), incPC: by(income, 'PC', r => r.total + r.eft + r.swipeNC)
    };
    summary.expTotal = r2(summary.expCF + summary.expPC);
    summary.incTotal = r2(summary.incCF + summary.incPC);
    summary.vatClaimable = expT.vat;

    // ---- station funds reconciliation (same arithmetic as the MRS form)
    const b = bal || {};
    const known = v => v !== null && v !== undefined && v !== '' && isFinite(v);
    const recon = {
      openBank: known(b.openBank) ? r2(b.openBank) : null,
      openCash: known(b.openCash) ? r2(b.openCash) : null,
      income: r2(incT.total + incT.eft),
      expenditure: expT.cashPlusSwipe,
      bankClose: known(b.stmtClose) ? r2(b.stmtClose) : null,
      cashClose: known(b.cashCount) ? r2(b.cashCount) : null
    };
    recon.openTotal = recon.openBank === null || recon.openCash === null ? null : r2(recon.openBank + recon.openCash);
    recon.reconClose = recon.openTotal === null ? null : r2(recon.openTotal + recon.income - recon.expenditure);
    recon.actualClose = recon.bankClose === null || recon.cashClose === null ? null : r2(recon.bankClose + recon.cashClose);
    recon.difference = recon.reconClose === null || recon.actualClose === null ? null : r2(recon.reconClose - recon.actualClose);

    recon.control = r2(recon.expenditure - recon.income);
    return { month, expenses, expTotals: expT, income, incTotals: incT, summary, recon };
  }

  /** Cash count by denomination -> total. counts: { '200': 0, '100': 61, ... } */
  const DENOMS = [200, 100, 50, 20, 10, 5, 2, 1, 0.5, 0.2, 0.1, 0.05];
  function countTotal(counts) {
    return r2(DENOMS.reduce((t, d) => t + d * num((counts || {})[String(d)]), 0));
  }

  const api = { build, countTotal, DENOMS, swipesFromNote, sectionOf, dmy };
  if (typeof module !== 'undefined' && module.exports) module.exports = api; else root.PCashPack = api;
})(typeof window !== 'undefined' ? window : this);
