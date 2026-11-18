/**
 * PCash Vs IP – writes a month-end pack (from pack.js) into the MRS petty cash form template.
 * Template: template.xlsx (the October 2026 sheet with the inputs cleared).
 * Works with ExcelJS in the browser (window.ExcelJS) and in Node (require('exceljs')).
 */
(function (root) {
  const MON = ['JAN', 'FEB', 'MAR', 'APR', 'MAY', 'JUN', 'JUL', 'AUG', 'SEP', 'OCT', 'NOV', 'DEC'];
  const MONTH_FULL = ['JANUARY', 'FEBRUARY', 'MARCH', 'APRIL', 'MAY', 'JUNE', 'JULY', 'AUGUST', 'SEPTEMBER', 'OCTOBER', 'NOVEMBER', 'DECEMBER'];
  const EXP_FIRST = 9, EXP_LAST = 86, INC_FIRST = 92, INC_LAST = 118;
  const COUNT_ROWS = { '200': 140, '100': 141, '50': 142, '20': 143, '10': 144, '5': 145, '2': 146, '1': 147, '0.5': 148, '0.2': 149, '0.1': 150, '0.05': 151 };

  function monthInfo(month) {
    const y = +month.slice(0, 4), m = +month.slice(5, 7);
    const last = new Date(Date.UTC(y, m, 0)).getUTCDate();
    return {
      sheetName: 'PC CF ' + MONTH_FULL[m - 1] + ' ' + y,
      closeDate: String(last).padStart(2, '0') + '/' + String(m).padStart(2, '0') + '/' + y,
      openLabel: '1 ' + MON[m - 1] + ' ' + y,
      closeLabel: last + ' ' + MON[m - 1] + ' ' + y
    };
  }

  const blank = v => (v === 0 || v === null || v === undefined || v === '') ? null : v;
  const YELLOW = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFFFFF00' } };
  // CF (Coke fund 6300) lines are highlighted yellow, as on the hand-built forms
  // (ExcelJS shares style objects between cells loaded from the same template style, so give
  //  each cell its own copy before changing the fill, or every look-alike cell turns yellow too)
  function markCF(ws, r) {
    'ABCDEFGHIJ'.split('').forEach(col => {
      const c = ws.getCell(col + r);
      c.style = Object.assign({}, c.style, { fill: YELLOW });
    });
  }

  /**
   * @param wb     ExcelJS workbook loaded from template.xlsx
   * @param pack   result of PCashPack.build()
   * @param counts cash count by denomination, e.g. { '100': 61, '0.5': 56 }
   * @returns list of warnings (empty when everything fitted)
   */
  function fill(wb, pack, counts) {
    const warn = [];
    const ws = wb.worksheets[0];
    const mi = monthInfo(pack.month);
    ws.name = mi.sheetName;

    // Stale cached results from the template must not show: keep the formula, drop the result.
    ws.eachRow({ includeEmpty: false }, row => row.eachCell({ includeEmpty: false }, c => {
      if (c.value && typeof c.value === 'object' && c.value.formula) c.value = { formula: c.value.formula };
      if (typeof c.value === 'string' && c.value.indexOf('{') >= 0)
        c.value = c.value.replace('{OPEN}', mi.openLabel).replace('{CLOSE}', mi.closeLabel);
    }));
    if (wb.calcProperties) wb.calcProperties.fullCalcOnLoad = true; else wb.calcProperties = { fullCalcOnLoad: true };

    ws.getCell('H4').value = mi.closeDate;

    // ---- voucher list
    const maxExp = EXP_LAST - EXP_FIRST + 1;
    if (pack.expenses.length > maxExp) warn.push(pack.expenses.length + ' expense lines but the form has room for ' + maxExp + '. The rest were left off.');
    pack.expenses.slice(0, maxExp).forEach((L, i) => {
      const r = EXP_FIRST + i;
      ws.getCell('A' + r).value = L.voucher;
      ws.getCell('B' + r).value = L.alloc;
      ws.getCell('C' + r).value = L.firm || null;
      ws.getCell('D' + r).value = L.article || null;
      ws.getCell('E' + r).value = blank(L.cx);
      ws.getCell('F' + r).value = L.cost;
      ws.getCell('G' + r).value = blank(L.vat);
      ws.getCell('H' + r).value = blank(L.cash);
      ws.getCell('I' + r).value = blank(L.swipe);
      ws.getCell('J' + r).value = blank(L.eft);
      if (L.alloc === 'CF') markCF(ws, r);
    });

    // ---- income block
    const maxInc = INC_LAST - INC_FIRST + 1;
    if (pack.income.length > maxInc) warn.push(pack.income.length + ' income lines but the form has room for ' + maxInc + '. The rest were left off.');
    pack.income.slice(0, maxInc).forEach((L, i) => {
      const r = INC_FIRST + i;
      ws.getCell('A' + r).value = L.no;
      ws.getCell('B' + r).value = L.alloc;
      ws.getCell('C' + r).value = L.firm || null;
      ws.getCell('D' + r).value = L.article || null;
      ws.getCell('E' + r).value = blank(L.cx);
      ws.getCell('F' + r).value = L.cost;
      // cash takings show in TOTAL even when zero (kitchen days with swipes only)
      ws.getCell('H' + r).value = L.eft ? null : L.total;
      ws.getCell('I' + r).value = blank(L.eft);
      ws.getCell('J' + r).value = blank(L.swipeNC);
      if (L.alloc === 'CF') markCF(ws, r);
    });

    // ---- station funds reconciliation inputs
    const R = pack.recon;
    ws.getCell('F138').value = R.openBank;
    ws.getCell('F139').value = R.openCash;
    ws.getCell('F147').value = R.bankClose;
    Object.keys(COUNT_ROWS).forEach(d => {
      const n = counts && counts[d] !== undefined && counts[d] !== '' ? Number(counts[d]) : 0;
      ws.getCell('I' + COUNT_ROWS[d]).value = isFinite(n) ? n : 0;
    });
    if (R.openBank === null || R.openCash === null) warn.push('Opening bank or cash is missing.');
    if (R.bankClose === null) warn.push('Statement closing balance is missing.');
    return warn;
  }

  function fileName(month) {
    const y = month.slice(0, 4), m = +month.slice(5, 7);
    return 'F' + (m >= 7 ? +y + 1 : y) + ' - HO PETTYCASH ' + MON[m - 1] + y.slice(2) + '.xlsx';
  }

  const api = { fill, monthInfo, fileName, COUNT_ROWS };
  if (typeof module !== 'undefined' && module.exports) module.exports = api; else root.PCashFill = api;
})(typeof window !== 'undefined' ? window : this);
