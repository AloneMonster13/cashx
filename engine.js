/* =====================================================================
   CashX -> Acumatica engine
   Pure logic: parse the CashX export, build AR + AP records, write xlsx.
   Works in the browser (global `CashXEngine`) and in Node (module.exports).
   ===================================================================== */
(function (root, factory) {
  if (typeof module === "object" && module.exports) module.exports = factory();
  else root.CashXEngine = factory();
})(typeof self !== "undefined" ? self : this, function () {
  "use strict";

  /* ---------- defaults (taken from your two template files) ---------- */
  const DEFAULTS = {
    currency: "LKR",
    location: "CTL",
    branch: "CTL",
    destBranch: "CTL",
    subAccount: 0,
    arDocType: "Invoice",
    apDocType: "Bill",
    arAccount: "2900009",
    apAccount: "5000091",
    arTax: "Non Taxable",
    apTax: "NON TAXABLE",
    arInvId: "AIR TICKETING",
    apInvId: "AIR_TICKETING_______",
    apDefaultVendor: "IATAHINGE_", // used for Airline rows
    apFallbackVendor: "", // used for non-airline suppliers that have no mapping
    descFields: ["invoiceNo", "clientNm", "paxName", "tktNo"],
    includeAirline: true,
    includeMisc: true,
    skipZero: true,
    postPeriod: null, // null = today minus 3 months
  };

  /* Selectable pieces of the DOC / LINE description */
  const DESC_FIELDS = {
    serNo: "Ser No",
    invoiceNo: "Invoice No",
    clientGrp: "Client Grp",
    clientNm: "Client Nm",
    paxName: "Pax Name",
    airline: "Airline Name",
    tktNo: "Tkt No",
    segment: "Segment",
    counselor: "Travel Counselor",
  };

  /* Known customers (from your AR template). Key = CashX Client Code. */
  const DEFAULT_CUSTOMERS = {
    M000020: "MM",
    S000001: "SCMINDIA__",
  };

  const AR_HEADERS = [
    "DOC TYPE", "INVOICE REF NBR", "Pre REF NBR", "LINE NBR", "LINE NBR (formula)",
    "CUSTOMER ID", "DATE", "POST PERIOD", "LOCATION", "ORIGINATING BRANCH",
    "CUSTOMER REF NBR", "Currency", "DOC DESCRIPTION (Formula)", "DOC DESCRIPTION",
    "LINE DESCRIPTION", "LINE DESCRIPTION (formula)", "INV ID", "EXT PRICE",
    "TOTAL AMOUNT", "ACCOUNT", "SUBACCOUNT", "DESTINATION BRANCH", "Tax Category",
  ];
  const AP_HEADERS = [
    "DOC TYPE", "BILL NO", "PRE BILL NUMBER", "LINE NBR", "CURRENCY", "Vendor id",
    "DATE", "POST PERIOD", "LOCATION", "ORIGINATING BRANCH", "Vendor Ref",
    "DOC DESCRIPTION", "LINE DESCRIPTION", "INV ID", "EXT PRICE", "TOTAL AMOUNT",
    "ACCOUNT", "SUBACCOUNT", "DESTINATION BRANCH", "Tax Category",
  ];

  /* ---------- small helpers ---------- */
  const round2 = (n) => Math.round((Number(n) + Number.EPSILON) * 100) / 100;
  const clean = (v) =>
    v === null || v === undefined ? "" : String(v).replace(/\u00a0/g, " ").replace(/\s+/g, " ").trim();
  const norm = (s) => clean(s).toLowerCase().replace(/[^a-z0-9]+/g, " ").trim();

  function cellValue(cell) {
    let v = cell.value;
    if (v && typeof v === "object" && !(v instanceof Date)) {
      if ("result" in v) v = v.result;
      else if (Array.isArray(v.richText)) v = v.richText.map((r) => r.text).join("");
      else if ("text" in v) v = v.text;
    }
    return v === undefined ? null : v;
  }
  const num = (v) => {
    if (v === null || v === undefined || v === "" || v === "-") return 0;
    const n = typeof v === "number" ? v : parseFloat(String(v).replace(/,/g, ""));
    return Number.isFinite(n) ? n : 0;
  };

  /* Default post period = current month minus 3 months, "MM-YYYY" */
  function defaultPostPeriod(today) {
    const t = today || new Date();
    const d = new Date(t.getFullYear(), t.getMonth() - 3, 1);
    return String(d.getMonth() + 1).padStart(2, "0") + "-" + d.getFullYear();
  }

  /* ---------------------------------------------------------------
     CashX dates are messy: the report is dd/mm/yyyy, but Excel turned
     every date whose day <= 12 into a real date with day/month SWAPPED
     (3 Nov 2026 became 11 Mar 2026), while dates with day > 12 stayed
     text. So: real dates -> swap back; text -> read as dd/mm/yyyy.
     --------------------------------------------------------------- */
  function parseCashXDate(v) {
    if (v instanceof Date && !isNaN(v)) {
      const y = v.getUTCFullYear(), m = v.getUTCMonth() + 1, d = v.getUTCDate();
      if (d <= 12) return { date: new Date(Date.UTC(y, d - 1, m)), swapped: true };
      return { date: new Date(Date.UTC(y, m - 1, d)), swapped: false };
    }
    const s = clean(v);
    const m = s.match(/^(\d{1,2})[\/.\-](\d{1,2})[\/.\-](\d{4})/);
    if (m) return { date: new Date(Date.UTC(+m[3], +m[2] - 1, +m[1])), swapped: false };
    return { date: null, swapped: false };
  }

  /* ---------------------------------------------------------------
     PARSE: ExcelJS workbook -> array of CashX line objects
     --------------------------------------------------------------- */
  function parseCashX(wb) {
    const ws = wb.worksheets[0];
    if (!ws) throw new Error("The workbook has no sheets.");

    // find the header row (contains "Ser No" and "Invoice No")
    let headerRow = 0;
    const maxScan = Math.min(ws.rowCount, 40);
    for (let r = 1; r <= maxScan; r++) {
      const names = [];
      ws.getRow(r).eachCell({ includeEmpty: false }, (c) => names.push(norm(cellValue(c))));
      if (names.includes("ser no") && names.includes("invoice no")) { headerRow = r; break; }
    }
    if (!headerRow) throw new Error("This doesn't look like a CashX export (couldn't find the 'Ser No' / 'Invoice No' header row).");

    const col = {};
    ws.getRow(headerRow).eachCell({ includeEmpty: false }, (c, i) => {
      const k = norm(cellValue(c));
      if (k && !(k in col)) col[k] = i;
    });
    const need = ["invoice no", "invoice date", "client nm", "gross fare", "supp amount"];
    const missing = need.filter((k) => !(k in col));
    if (missing.length) throw new Error("Missing CashX columns: " + missing.join(", "));

    const get = (row, k) => (k in col ? cellValue(row.getCell(col[k])) : null);
    const lines = [];
    let swappedDates = 0, badDates = 0;

    for (let r = headerRow + 1; r <= ws.rowCount; r++) {
      const row = ws.getRow(r);
      const ser = get(row, "ser no");
      const inv = clean(get(row, "invoice no"));
      if (!inv || typeof ser !== "number") continue; // skips totals / notes rows

      const dt = parseCashXDate(get(row, "invoice date"));
      if (dt.swapped) swappedDates++;
      if (!dt.date) badDates++;

      lines.push({
        excelRow: r,
        serNo: ser,
        type: clean(get(row, "type of svc")), // Airline | Misc.
        service: clean(get(row, "service")),
        invoiceNo: inv,
        date: dt.date,
        clientGrp: clean(get(row, "client grp")),
        clientCode: clean(get(row, "client code")),
        clientNm: clean(get(row, "client nm")),
        paxName: clean(get(row, "pax name")),
        supplier: clean(get(row, "supplier")),
        airline: clean(get(row, "airline name")),
        tktNo: clean(get(row, "tkt no")),
        segment: clean(get(row, "segment")),
        counselor: clean(get(row, "travel counselor")),
        curr: clean(get(row, "curr")) || "LKR",
        grossFare: num(get(row, "gross fare")),
        suppAmount: num(get(row, "supp amount")),
      });
    }
    if (!lines.length) throw new Error("No invoice lines were found in this file.");

    const period = ws.getRow(5).getCell(1).value;
    return { lines, meta: { headerRow, swappedDates, badDates, reportRange: clean(period && period.result ? period.result : period) } };
  }

  /* Parse an Acumatica customer export: needs a "Customer ID" and a name column */
  function parseCustomerList(wb) {
    const ws = wb.worksheets[0];
    if (!ws) return [];
    for (let r = 1; r <= Math.min(ws.rowCount, 15); r++) {
      let idCol = 0, nameCol = 0;
      ws.getRow(r).eachCell({ includeEmpty: false }, (c, i) => {
        const k = norm(cellValue(c));
        if (!idCol && /^(customer id|customer code|customer cd|id)$/.test(k)) idCol = i;
        else if (!nameCol && /^(customer name|customer|name|account name|company name)$/.test(k)) nameCol = i;
      });
      if (idCol && nameCol) {
        const out = [];
        for (let i = r + 1; i <= ws.rowCount; i++) {
          const id = clean(cellValue(ws.getRow(i).getCell(idCol)));
          const name = clean(cellValue(ws.getRow(i).getCell(nameCol)));
          if (id && name) out.push({ id, name });
        }
        return out;
      }
    }
    throw new Error("Couldn't find 'Customer ID' and 'Customer Name' columns in that file.");
  }

  /* ---------------------------------------------------------------
     Customer lookup:  1) manual/saved map by Client Code
                       2) Acumatica customer list (match by id or name)
     --------------------------------------------------------------- */
  function buildCustomerResolver(manualMap, acumaticaList) {
    const byName = (acumaticaList || []).map((c) => ({ id: c.id, n: norm(c.name) }));
    const ids = new Set((acumaticaList || []).map((c) => c.id));
    return function resolve(line) {
      const m = manualMap || {};
      if (m[line.clientCode]) return { id: m[line.clientCode], via: "mapping" };
      if (ids.has(line.clientCode)) return { id: line.clientCode, via: "acumatica" };
      // CashX truncates names at ~30 chars and appends "(CODE" - strip that for matching
      const nm = norm(line.clientNm.replace(/\(.*$/, ""));
      if (nm) {
        const hit = byName.find((c) => c.n === nm) ||
          byName.find((c) => c.n.startsWith(nm) || (nm.startsWith(c.n) && c.n.length > 6));
        if (hit) return { id: hit.id, via: "acumatica" };
      }
      return { id: "", via: "none" };
    };
  }

  /* ---------------------------------------------------------------
     BUILD: CashX lines -> AR + AP records
     --------------------------------------------------------------- */
  function build(parsed, userSettings, extras) {
    const S = Object.assign({}, DEFAULTS, userSettings || {});
    const X = Object.assign({ customers: {}, acumatica: [], vendors: {}, arItems: {}, apItems: {} }, extras || {});
    const acct = (v) => (/^\d+$/.test(String(v)) ? Number(v) : v);
    const postPeriod = S.postPeriod || defaultPostPeriod();
    const resolveCustomer = buildCustomerResolver(Object.assign({}, DEFAULT_CUSTOMERS, X.customers), X.acumatica);

    const lists = { customers: new Map(), vendors: new Map(), services: new Map() };
    const warnings = { unmappedCustomers: {}, unmappedVendors: {}, nonLkr: 0, skippedZeroAr: 0, skippedZeroAp: 0, noDate: 0 };
    const isAirline = (l) => /airline/i.test(l.type);

    let lines = parsed.lines.filter((l) => (isAirline(l) ? S.includeAirline : S.includeMisc));

    const description = (l) =>
      S.descFields
        .map((k) => (k === "serNo" ? String(l.serNo) : l[k === "airline" ? "airline" : k]))
        .map(clean)
        .filter((p) => p && p !== "-")
        .join(" / ")
        .slice(0, 250);

    lines.forEach((l) => {
      const s = lists.services.get(l.service) || { name: l.service || "(none)", rows: 0 };
      s.rows++; lists.services.set(l.service, s);
    });
    const dateOf = (l) => { if (!l.date) warnings.noDate++; return l.date; };

    /* ---------- AR ---------- */
    let arLines = lines.filter((l) => {
      if (S.skipZero && round2(l.grossFare) === 0) { warnings.skippedZeroAr++; return false; }
      return true;
    });
    // same invoice rows must sit together; number invoices by first appearance
    const order = new Map();
    arLines.forEach((l) => { if (!order.has(l.invoiceNo)) order.set(l.invoiceNo, order.size + 1); });
    arLines = arLines
      .map((l, i) => ({ l, i }))
      .sort((a, b) => order.get(a.l.invoiceNo) - order.get(b.l.invoiceNo) || a.i - b.i)
      .map((x) => x.l);

    const arTotals = new Map();
    arLines.forEach((l) => arTotals.set(l.invoiceNo, round2((arTotals.get(l.invoiceNo) || 0) + l.grossFare)));
    const lineCounter = new Map();

    const ar = arLines.map((l) => {
      const n = lineCounter.get(l.invoiceNo) || 0;
      lineCounter.set(l.invoiceNo, n + 1);
      const cust = resolveCustomer(l);
      const ck = l.clientCode || l.clientNm;
      const ce = lists.customers.get(ck) || { code: l.clientCode, name: l.clientNm, rows: 0, id: cust.id, via: cust.via };
      ce.rows++; lists.customers.set(ck, ce);
      if (cust.via === "none") {
        const k = l.clientCode || l.clientNm;
        (warnings.unmappedCustomers[k] = warnings.unmappedCustomers[k] || { code: l.clientCode, name: l.clientNm, rows: 0 }).rows++;
      }
      if (l.curr !== "LKR") warnings.nonLkr++;
      const desc = description(l);
      return [
        S.arDocType, order.get(l.invoiceNo), l.invoiceNo, n, n, cust.id, dateOf(l), postPeriod,
        S.location, S.branch, l.invoiceNo, S.currency, desc, desc, desc, desc,
        X.arItems[l.service] || S.arInvId, round2(l.grossFare), arTotals.get(l.invoiceNo),
        acct(S.arAccount), S.subAccount, S.destBranch, S.arTax,
      ];
    });

    /* ---------- AP ---------- */
    const apLines = lines.filter((l) => {
      if (S.skipZero && round2(l.suppAmount) === 0) { warnings.skippedZeroAp++; return false; }
      return true;
    });
    const ap = apLines.map((l, i) => {
      let vendor = X.vendors[l.supplier] || (isAirline(l) ? S.apDefaultVendor : S.apFallbackVendor);
      const vk = l.supplier || "(no supplier)";
      const ve = lists.vendors.get(vk) || { name: vk, rows: 0, id: vendor, airline: isAirline(l) };
      ve.rows++; lists.vendors.set(vk, ve);
      if (!vendor) {
        const k = l.supplier || "(no supplier)";
        (warnings.unmappedVendors[k] = warnings.unmappedVendors[k] || { name: k, rows: 0 }).rows++;
      }
      const desc = description(l);
      return [
        S.apDocType, i + 1, l.invoiceNo, 0, S.currency, vendor, dateOf(l), postPeriod,
        S.location, S.branch, l.invoiceNo, desc, desc,
        X.apItems[l.service] || S.apInvId, round2(l.suppAmount), round2(l.suppAmount),
        acct(S.apAccount), S.subAccount, S.destBranch, S.apTax,
      ];
    });

    const sum = (rows, idx) => round2(rows.reduce((a, r) => a + r[idx], 0));
    const stats = {
      postPeriod,
      arRows: ar.length,
      arInvoices: order.size,
      arTotal: sum(ar, 17),
      apRows: ap.length,
      apTotal: sum(ap, 14),
      swappedDates: parsed.meta.swappedDates,
    };
    stats.margin = round2(stats.arTotal - stats.apTotal);
    const arr = (m) => Array.from(m.values());
    return {
      ar, ap, stats, warnings, settings: S,
      lists: {
        customers: arr(lists.customers).sort((a, b) => (a.id ? 1 : 0) - (b.id ? 1 : 0) || b.rows - a.rows),
        vendors: arr(lists.vendors).sort((a, b) => (a.id ? 1 : 0) - (b.id ? 1 : 0) || b.rows - a.rows),
        services: arr(lists.services).sort((a, b) => b.rows - a.rows),
      },
    };
  }

  /* ---------------------------------------------------------------
     WRITE: records -> styled ExcelJS workbook
     `ExcelJS` is passed in so this file stays library-agnostic.
     --------------------------------------------------------------- */
  function makeWorkbook(ExcelJS, kind, rows) {
    const isAR = kind === "ar";
    const headers = isAR ? AR_HEADERS : AP_HEADERS;
    const dateCol = isAR ? 7 : 7;
    const moneyCols = isAR ? [18, 19] : [15, 16];
    const accent = isAR ? "FF2B4BF2" : "FF0F8B7D";

    const wb = new ExcelJS.Workbook();
    wb.creator = "CashX to Acumatica";
    wb.created = new Date();
    const ws = wb.addWorksheet(isAR ? "AR Invoices" : "Issued", { views: [{ state: "frozen", ySplit: 1 }] });

    ws.addRow(headers);
    rows.forEach((r) => ws.addRow(r));

    const head = ws.getRow(1);
    head.height = 22;
    head.eachCell((c) => {
      c.font = { bold: true, color: { argb: "FFFFFFFF" } };
      c.fill = { type: "pattern", pattern: "solid", fgColor: { argb: accent } };
      c.alignment = { vertical: "middle", horizontal: "center" };
    });

    ws.getColumn(dateCol).numFmt = "dd/mm/yyyy";
    moneyCols.forEach((i) => (ws.getColumn(i).numFmt = "#,##0.00"));
    headers.forEach((h, i) => {
      let w = Math.max(h.length + 3, 12);
      if (/DESCRIPTION/i.test(h)) w = 60;
      else if (/Vendor id|CUSTOMER ID|INV ID/i.test(h)) w = 22;
      ws.getColumn(i + 1).width = w;
    });
    ws.autoFilter = { from: { row: 1, column: 1 }, to: { row: 1, column: headers.length } };
    return wb;
  }

  return {
    DEFAULTS, DESC_FIELDS, DEFAULT_CUSTOMERS, AR_HEADERS, AP_HEADERS,
    defaultPostPeriod, parseCashX, parseCustomerList, build, makeWorkbook, parseCashXDate,
  };
});
