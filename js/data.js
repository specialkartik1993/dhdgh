/* Workbook parsing: turns the Excel file into a plain data model.
 * Columns are matched by header name (case / dash / spacing insensitive), so
 * column order can change and extra columns are ignored. All totals and
 * averages are recomputed from the raw sheets, so editing any district or
 * AWC row flows through to every view. */
(function () {
  'use strict';

  const norm = (s) => String(s ?? '')
    .toLowerCase()
    .replace(/[‐-―−]/g, '-')
    .replace(/÷/g, '/')
    .replace(/[^a-z0-9%]+/g, ' ')
    .trim();

  const num = (v) => {
    if (v === null || v === undefined || v === '') return null;
    if (typeof v === 'number') return isFinite(v) ? v : null;
    const s = String(v).replace(/[,\s]/g, '');
    const isPct = s.endsWith('%');
    const n = parseFloat(s.replace('%', ''));
    if (!isFinite(n)) return null;
    return isPct ? n / 100 : n;
  };
  // Rates are stored 0–1; accept 0–100 too.
  const pct = (v) => {
    const n = num(v);
    if (n === null) return null;
    return n > 1.5 ? n / 100 : n;
  };
  const str = (v) => (v === null || v === undefined ? '' : String(v).trim());

  function findSheet(wb, ...names) {
    const wanted = names.map(norm);
    for (const w of wanted) {
      const hit = wb.SheetNames.find((n) => norm(n) === w);
      if (hit) return wb.Sheets[hit];
    }
    for (const w of wanted) {
      const hit = wb.SheetNames.find((n) => norm(n).includes(w));
      if (hit) return wb.Sheets[hit];
    }
    return null;
  }

  /* Read a table whose header row contains `anchor`. Returns array of
   * objects keyed by normalised header. Stops at the first fully blank row. */
  function readTable(ws, anchor) {
    if (!ws) return [];
    const rows = XLSX.utils.sheet_to_json(ws, { header: 1, defval: null, blankrows: true, raw: true });
    const a = norm(anchor);
    const hi = rows.findIndex((r) => r && r.some((c) => norm(c) === a));
    if (hi < 0) return [];
    const headers = rows[hi].map(norm);
    const out = [];
    for (let i = hi + 1; i < rows.length; i++) {
      const r = rows[i];
      if (!r || r.every((c) => c === null || c === '')) break;
      const o = {};
      headers.forEach((h, j) => { if (h) o[h] = r[j]; });
      out.push(o);
    }
    return out;
  }

  const get = (o, ...keys) => {
    for (const k of keys) {
      const v = o[norm(k)];
      if (v !== undefined && v !== null && v !== '') return v;
    }
    return null;
  };

  function readMeta(wb) {
    const ws = findSheet(wb, 'Read Me', 'Readme');
    const meta = { period: '', notes: {} };
    if (!ws) return meta;
    const rows = XLSX.utils.sheet_to_json(ws, { header: 1, defval: null });
    rows.forEach((r) => {
      if (r && r[0] && r[1]) meta.notes[norm(r[0])] = str(r[1]);
    });
    meta.title = str(rows[1] && rows[1][0]);
    const p = meta.notes['reporting period'] || '';
    const m = p.match(/(January|February|March|April|May|June|July|August|September|October|November|December)\s+\d{4}/i);
    meta.period = m ? m[0] : p.replace(/^[^:]*:\s*/, '').replace(/\.$/, '');
    meta.caution = meta.notes['important'] || '';
    return meta;
  }

  function parseWorkbook(buf) {
    const wb = XLSX.read(buf, { type: 'array', cellDates: true });

    const perf = readTable(findSheet(wb, 'District Performance'), 'District');
    if (!perf.length) throw new Error('Sheet "District Performance" was not found or has no rows.');

    const nutRows = readTable(findSheet(wb, 'Nutrition'), 'District');
    const matRows = readTable(findSheet(wb, 'Maternal Health'), 'District');
    const awcRows = readTable(findSheet(wb, 'AWC Details'), 'AWC Code');
    const alertRows = readTable(findSheet(wb, 'Alerts'), 'Alert Category');

    const byName = (rows) => {
      const m = new Map();
      rows.forEach((r) => { const d = str(get(r, 'District')); if (d) m.set(d.toLowerCase(), r); });
      return m;
    };
    const nutMap = byName(nutRows);
    const matMap = byName(matRows);

    const districts = perf
      .filter((r) => str(get(r, 'District')))
      .map((r) => {
        const name = str(get(r, 'District'));
        const n = nutMap.get(name.toLowerCase()) || {};
        const m = matMap.get(name.toLowerCase()) || {};
        const children = num(get(r, 'Active Children (0-6)', 'Active Children', 'Active Children 0 6')) || 0;
        const measured = num(get(r, 'Children Measured')) ?? num(get(n, 'Children Measured')) ?? 0;
        const samPct = pct(get(r, 'SAM % of Measured'));
        const mamPct = pct(get(r, 'MAM % of Measured'));
        const samCases = num(get(n, 'SAM Cases')) ?? (samPct !== null ? Math.round(samPct * measured) : null);
        const mamCases = num(get(n, 'MAM Cases')) ?? (mamPct !== null ? Math.round(mamPct * measured) : null);
        const pw = num(get(r, 'Pregnant Women')) ?? num(get(m, 'Active Pregnant Women')) ?? 0;
        const anaemiaCount = num(get(m, 'Pregnant Women with Anaemia (Count)', 'Pregnant Women with Anaemia'));
        const anaemiaPct = pct(get(r, 'Anaemia % of Pregnant Women')) ?? (anaemiaCount !== null && pw ? anaemiaCount / pw : null);
        return {
          division: str(get(r, 'Division')) || 'Unassigned',
          district: name,
          awcs: num(get(r, 'Total AWCs')) || 0,
          children,
          pw,
          lm: num(get(r, 'Lactating Mothers')) || 0,
          measured,
          me: pct(get(r, 'Measurement Efficiency')) ?? (children ? measured / children : null),
          samPct: samPct ?? (samCases !== null && measured ? samCases / measured : null),
          mamPct: mamPct ?? (mamCases !== null && measured ? mamCases / measured : null),
          thr: pct(get(r, 'THR Coverage')),
          hcm: pct(get(r, 'HCM Coverage')),
          hv: pct(get(r, 'Home Visit Completion')),
          op: pct(get(r, 'AWC Operational Rate')),
          awwS: num(get(r, 'AWW Sanctioned')) || 0,
          awwW: num(get(r, 'AWW Working')) || 0,
          awwV: num(get(r, 'AWW Vacant')) ?? Math.max(0, (num(get(r, 'AWW Sanctioned')) || 0) - (num(get(r, 'AWW Working')) || 0)),
          awhS: num(get(r, 'AWH Sanctioned')) || 0,
          awhW: num(get(r, 'AWH Working')) || 0,
          awhV: num(get(r, 'AWH Vacant')) ?? Math.max(0, (num(get(r, 'AWH Sanctioned')) || 0) - (num(get(r, 'AWH Working')) || 0)),
          infra: pct(get(r, 'Basic Infrastructure Coverage')),
          anc: pct(get(r, 'ANC Coverage')),
          anaemiaPct,
          // nutrition
          sam: samCases,
          mam: mamCases,
          sevStunt: pct(get(n, 'Severely Stunted %')),
          modStunt: pct(get(n, 'Moderately Stunted %')),
          newSam: num(get(n, 'New SAM Cases This Month')),
          samFollow: num(get(n, 'SAM Follow-up Completed')),
          samRef: num(get(n, 'SAM Referred')),
          // maternal
          anc1: pct(get(m, 'ANC 1 Coverage')),
          anc2: pct(get(m, 'ANC 2 Coverage')),
          anc3: pct(get(m, 'ANC 3 Coverage')),
          anc4: pct(get(m, 'ANC 4 Coverage')),
          anaemiaCount,
          pendingAnc: num(get(m, 'Pending Due ANC (Count)', 'Pending Due ANC')),
          gwg: pct(get(m, 'Optimum Gestational Weight Gain Coverage')),
        };
      });

    const awcs = awcRows
      .filter((r) => str(get(r, 'AWC Code')))
      .map((r) => {
        const active = num(get(r, 'Active Children')) || 0;
        const measured = num(get(r, 'Children Measured')) || 0;
        const tgt = num(get(r, 'Target Home Visits'));
        const done = num(get(r, 'Completed Home Visits'));
        const days = num(get(r, 'Days Open (Month)'));
        return {
          code: str(get(r, 'AWC Code')),
          division: str(get(r, 'Division')),
          district: str(get(r, 'District')),
          project: str(get(r, 'Project')),
          sector: str(get(r, 'Sector')),
          name: str(get(r, 'AWC Name')),
          active, measured,
          me: pct(get(r, 'Measurement Efficiency')) ?? (active ? measured / active : null),
          sam: num(get(r, 'SAM Children')) || 0,
          mam: num(get(r, 'MAM Children')) || 0,
          thr: num(get(r, 'Children Receiving THR')),
          hcm: num(get(r, 'Children Receiving HCM')),
          building: str(get(r, 'Building Condition')) || '—',
          toilet: str(get(r, 'Toilet Status')) || '—',
          water: str(get(r, 'Drinking Water Status')) || '—',
          days,
          op: pct(get(r, 'AWC Operational Rate')) ?? (days !== null ? days / 26 : null),
          hvTarget: tgt, hvDone: done,
          hvRate: pct(get(r, 'Home Visit Completion Rate')) ?? (tgt ? done / tgt : null),
          aww: str(get(r, 'AWW Status')) || '—',
          awh: str(get(r, 'AWH Status')) || '—',
        };
      });

    const divOf = new Map(districts.map((d) => [d.district.toLowerCase(), d.division]));
    const alerts = alertRows
      .filter((r) => str(get(r, 'Alert Category')))
      .map((r, i) => {
        const district = str(get(r, 'District'));
        return {
          id: i,
          district,
          division: divOf.get(district.toLowerCase()) || '',
          category: str(get(r, 'Alert Category')),
          issue: str(get(r, 'Trigger / Issue', 'Trigger', 'Issue')),
          priority: str(get(r, 'Priority')) || 'Medium',
          officer: str(get(r, 'Responsible Officer')),
          status: str(get(r, 'Action Status')) || 'Open',
        };
      });

    return { meta: readMeta(wb), districts, awcs, alerts, sheetNames: wb.SheetNames };
  }

  /* ---------- Aggregation ---------- */
  const sum = (arr, k) => arr.reduce((s, d) => s + (d[k] || 0), 0);
  const avg = (arr, k) => {
    const v = arr.map((d) => d[k]).filter((x) => x !== null && x !== undefined && isFinite(x));
    return v.length ? v.reduce((a, b) => a + b, 0) / v.length : null;
  };

  function aggregate(ds) {
    const children = sum(ds, 'children');
    const measured = sum(ds, 'measured');
    const awwS = sum(ds, 'awwS'), awhS = sum(ds, 'awhS');
    const pw = sum(ds, 'pw');
    return {
      n: ds.length,
      awcs: sum(ds, 'awcs'),
      children, measured,
      pw, lm: sum(ds, 'lm'),
      me: children ? measured / children : null,
      sam: sum(ds, 'sam'), mam: sum(ds, 'mam'),
      samPct: measured ? sum(ds, 'sam') / measured : null,
      mamPct: measured ? sum(ds, 'mam') / measured : null,
      thr: avg(ds, 'thr'), hcm: avg(ds, 'hcm'), hv: avg(ds, 'hv'), op: avg(ds, 'op'),
      infra: avg(ds, 'infra'), anc: avg(ds, 'anc'),
      awwV: sum(ds, 'awwV'), awhV: sum(ds, 'awhV'), awwS, awhS,
      awwVPct: awwS ? sum(ds, 'awwV') / awwS : null,
      awhVPct: awhS ? sum(ds, 'awhV') / awhS : null,
      sevStunt: avg(ds, 'sevStunt'), modStunt: avg(ds, 'modStunt'),
      newSam: sum(ds, 'newSam'), samFollow: sum(ds, 'samFollow'), samRef: sum(ds, 'samRef'),
      anc1: avg(ds, 'anc1'), anc2: avg(ds, 'anc2'), anc3: avg(ds, 'anc3'), anc4: avg(ds, 'anc4'),
      anaemiaCount: sum(ds, 'anaemiaCount'),
      anaemiaPct: pw ? sum(ds, 'anaemiaCount') / pw : avg(ds, 'anaemiaPct'),
      pendingAnc: sum(ds, 'pendingAnc'),
      gwg: avg(ds, 'gwg'),
    };
  }

  function groupBy(arr, k) {
    const m = new Map();
    arr.forEach((x) => { const key = x[k]; if (!m.has(key)) m.set(key, []); m.get(key).push(x); });
    return m;
  }

  window.WCD = { parseWorkbook, aggregate, groupBy, norm };
})();
