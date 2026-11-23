// Calculation engine — a line-for-line port of analyze() in the old Streamlit app.py.
(function (root) {
  const MONTHS = ["Jan","Feb","Mar","Apr","May","Jun","Jul","Aug","Sep","Oct","Nov","Dec"];

  // Python round(x, n) returns a float; toFixed rounds the exact binary value the same way.
  const pyRound = (x, n) => Number(x.toFixed(n));

  function findRow(rows, target) {
    let best = null, bestDiff = 9999;
    for (const r of rows) {
      const diff = Math.abs(r[0] - target);
      if (diff < bestDiff) { bestDiff = diff; best = r; }
    }
    return best;
  }
  function findLast(rows, test) {
    for (let i = rows.length - 1; i >= 0; i--) if (test(rows[i])) return rows[i];
    return null;
  }
  const findLastComplete = rows => findLast(rows, r => r[0] && r[1] && r[4] && r[9]);
  const findLastWithEarnings = rows => findLast(rows, r => r[0] && r[3] && r[10]);
  const findLastWithDividends = rows => findLast(rows, r => r[0] && r[1] && r[2] && r[4] && r[7] && r[8] && r[9]);

  function dateStr(d) {
    const year = Math.trunc(d);
    const mon = Math.max(1, Math.min(12, Math.round((d - year) * 100)));
    return `${MONTHS[mon - 1]} ${year}`;
  }

  function analyze(startDate, investment, rows) {
    const start = findRow(rows, startDate);
    const end = findLastComplete(rows);
    const divEnd = findLastWithDividends(rows) || end;

    const cpiS = start[4], cpiE = end[4];
    const rtrS = start[9], rtrE = end[9];
    const rpS = start[7], rpE = end[7];

    const scaled = investment * (cpiS / cpiE);
    const cpiFactor = cpiE / cpiS;
    const realTrF = rtrE / rtrS;
    const nomTrF = realTrF * cpiFactor;
    const realEnd = scaled * realTrF;
    const nomEnd = scaled * nomTrF;
    const priceRF = rpE / rpS;
    const priceOnly = scaled * priceRF;

    let divNomF = null, divRealF = null, divStartR = null, divEndR = null;
    if (start[2] && divEnd[2] && start[8] && divEnd[8]) {
      divNomF = divEnd[2] / start[2];
      divRealF = divEnd[8] / start[8];
      divStartR = start[8];
      divEndR = divEnd[8];
    }

    const earnEnd = findLastWithEarnings(rows) || end;
    let earnNomF = null, earnRealF = null, enS = null, enE = null, erS = null, erE = null;
    if (start[3] && earnEnd[3] && start[10] && earnEnd[10]) {
      enS = start[3]; enE = earnEnd[3]; erS = start[10]; erE = earnEnd[10];
      earnNomF = enE / enS; earnRealF = erE / erS;
    }

    const sy = Math.trunc(start[0]), sm = Math.round((start[0] - sy) * 100);
    const ey = Math.trunc(end[0]), em = Math.round((end[0] - ey) * 100);
    const years = (ey - sy) + (em - sm) / 12;
    const r = (v, n) => (v ? pyRound(v, n) : null);

    return {
      start_str: dateStr(start[0]),
      end_str: dateStr(end[0]),
      start_date: start[0],
      investment,
      scaled: pyRound(scaled, 0),
      cpi_start: pyRound(cpiS, 3),
      cpi_end: pyRound(cpiE, 3),
      cpi_factor: pyRound(cpiFactor, 2),
      real_tr_f: pyRound(realTrF, 2),
      nom_tr_f: pyRound(nomTrF, 1),
      real_end: pyRound(realEnd, 0),
      nom_end: pyRound(nomEnd, 0),
      price_r_f: pyRound(priceRF, 2),
      price_only: pyRound(priceOnly, 0),
      div_nom_f: r(divNomF, 1),
      div_real_f: r(divRealF, 1),
      div_start_r: r(divStartR, 2),
      div_end_r: r(divEndR, 2),
      earn_nom_f: r(earnNomF, 1),
      earn_real_f: r(earnRealF, 1),
      earn_nom_start: r(enS, 2),
      earn_nom_end: r(enE, 2),
      earn_real_start: r(erS, 2),
      earn_real_end: r(erE, 2),
      earn_end_str: dateStr(earnEnd[0]),
      years: Math.trunc(years),
    };
  }

  const comma = n => Math.round(n).toLocaleString("en-US");
  function fmt(n) {
    if (n === null || n === undefined) return "—";
    const a = Math.abs(n);
    if (a >= 1e9) return `$${(n / 1e9).toFixed(1)}B`;
    if (a >= 1e6) return `$${(n / 1e6).toFixed(1)}M`;
    if (a >= 1e3) return `$${comma(n)}`;
    return `$${n.toFixed(0)}`;
  }
  function fmtX(n) {
    if (n === null || n === undefined) return "—";
    if (n >= 100) return `${n.toFixed(0)}x`;
    if (n >= 10) return `${n.toFixed(1)}x`;
    return `${n.toFixed(2)}x`;
  }
  const fmtEps = n => (n === null || n === undefined ? "—" :
    `$${n.toLocaleString("en-US", { minimumFractionDigits: 2, maximumFractionDigits: 2 })}`);

  const api = { MONTHS, pyRound, findRow, findLastComplete, findLastWithEarnings, findLastWithDividends, dateStr, analyze, fmt, fmtX, fmtEps };
  if (typeof module !== "undefined") module.exports = api; else root.Engine = api;
})(this);
