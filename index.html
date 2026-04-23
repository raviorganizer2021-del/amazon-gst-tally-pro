const state = {
  parsedReports: { b2b: [], b2c: [], str: [] },
  reportMeta: {
    b2b: { columns: [], mapping: {} },
    b2c: { columns: [], mapping: {} },
    str: { columns: [], mapping: {} }
  },
  companyProfiles: [],
  processed: null
};

const REPORT_SCHEMAS = {
  b2b: {
    sellerGstin: ["seller gstin"],
    invoiceNumber: ["invoice number"],
    voucherDate: ["invoice date"],
    transactionType: ["transaction type"],
    orderId: ["order id"],
    quantity: ["quantity"],
    itemDescription: ["item description"],
    hsn: ["hsn/sac"],
    partyName: ["buyer name"],
    gstin: ["customer bill to gstid", "customer ship to gstid"],
    state: ["ship to state", "bill to state"],
    totalAmount: ["invoice amount"],
    taxableValue: ["tax exclusive gross", "principal amount basis"],
    igst: ["igst tax"],
    cgst: ["cgst tax"],
    sgst: ["sgst tax", "utgst tax"],
    cess: ["compensatory cess tax"],
    tcsIgstAmount: ["tcs igst amount"],
    tcsCgstAmount: ["tcs cgst amount"],
    tcsSgstAmount: ["tcs sgst amount"],
    creditNoteNo: ["credit note no"],
    creditNoteDate: ["credit note date"],
    taxRate: ["igst rate", "cgst rate", "sgst rate", "utgst rate"]
  },
  b2c: {
    sellerGstin: ["seller gstin"],
    invoiceNumber: ["invoice number"],
    voucherDate: ["invoice date"],
    transactionType: ["transaction type"],
    orderId: ["order id"],
    quantity: ["quantity"],
    itemDescription: ["item description"],
    hsn: ["hsn/sac"],
    partyName: ["buyer name"],
    state: ["ship to state", "bill to state"],
    totalAmount: ["invoice amount"],
    taxableValue: ["tax exclusive gross", "principal amount basis"],
    igst: ["igst tax"],
    cgst: ["cgst tax"],
    sgst: ["sgst tax", "utgst tax"],
    cess: ["compensatory cess tax", "shipping cess tax amount"],
    tcsIgstAmount: ["tcs igst amount"],
    tcsCgstAmount: ["tcs cgst amount"],
    tcsSgstAmount: ["tcs sgst amount"],
    creditNoteNo: ["credit note no"],
    creditNoteDate: ["credit note date"],
    taxRate: ["igst rate", "cgst rate", "sgst rate", "utgst rate"]
  },
  str: {
    sellerGstin: ["gstin of supplier"],
    invoiceNumber: ["invoice number"],
    voucherDate: ["invoice date"],
    transactionType: ["transaction type"],
    orderId: ["order id", "transaction id"],
    quantity: ["quantity"],
    itemDescription: ["asin"],
    hsn: ["hsn code"],
    partyName: ["gstin of receiver", "ship to fc", "ship from fc"],
    gstin: ["gstin of receiver"],
    state: ["ship to state"],
    totalAmount: ["invoice value"],
    taxableValue: ["taxable value"],
    igst: ["igst amount"],
    cgst: ["cgst amount"],
    sgst: ["sgst amount", "utgst amount"],
    cess: ["compensatory cess amount"],
    taxRate: ["igst rate", "cgst rate", "sgst rate", "utgst rate"]
  }
};

const $ = (id) => document.getElementById(id);

document.addEventListener("DOMContentLoaded", () => {
  initializeDefaults();
  loadCompanyProfiles();
  bindEvents();
  renderSidebar();
  renderCompanyCards();
});

function initializeDefaults() {
  const today = new Date();
  $("defaultDate").value = today.toISOString().slice(0, 10);
  $("returnMonth").value = today.toISOString().slice(0, 7);
}

function bindEvents() {
  ["b2bFile", "b2cFile", "strFile"].forEach((id) => $(id).addEventListener("change", handleFileUpload));
  $("processBtn").addEventListener("click", processReports);
  $("demoFillBtn").addEventListener("click", loadDemoData);
  $("downloadSummaryBtn").addEventListener("click", () => downloadJson("gst-summary.json", state.processed?.summary || {}));
  $("downloadMappingsBtn").addEventListener("click", () => downloadJson("report-mappings.json", state.reportMeta));
  $("downloadGstr1Btn").addEventListener("click", () => downloadJson("gstr-1.json", state.processed?.gstReturns?.gstr1 || {}));
  $("downloadGstr3bBtn").addEventListener("click", () => downloadJson("gstr-3b.json", state.processed?.gstReturns?.gstr3b || []));
  $("downloadLedgersBtn").addEventListener("click", () => downloadCsv("ledgers.csv", state.processed?.ledgers || []));
  $("downloadVouchersBtn").addEventListener("click", () => downloadCsv("vouchers.csv", state.processed?.vouchers || []));
  $("saveRunBtn").addEventListener("click", saveRunHistory);
  $("addCompanyBtn").addEventListener("click", addCompanyProfile);
  $("downloadTallyBtn").addEventListener("click", () => {
    if (!state.processed?.xml) return alert("Process reports before downloading XML.");
    downloadText("amazon-tally-import.xml", state.processed.xml, "application/xml");
  });
  $("sellerFilter").addEventListener("change", () => {
    if (state.processed) processReports();
  });
}

function renderSidebar() {
  document.querySelectorAll(".side-link").forEach((button) => {
    button.addEventListener("click", () => {
      document.querySelectorAll(".side-link").forEach((node) => node.classList.remove("is-active"));
      button.classList.add("is-active");
      const target = $(button.dataset.target);
      if (target) target.scrollIntoView({ behavior: "smooth", block: "start" });
    });
  });
}

async function handleFileUpload(event) {
  const reportKey = event.target.id.replace("File", "").toLowerCase();
  const [file] = event.target.files || [];
  if (!file) return;
  const rows = await readSpreadsheetFile(file);
  const columns = extractColumns(rows);
  const mapping = detectSchemaMapping(columns, reportKey);
  state.reportMeta[reportKey] = { columns, mapping };
  state.parsedReports[reportKey] = normalizeRows(rows, reportKey, mapping);
  updateUploadStats();
  renderMappingReview();
}

async function readSpreadsheetFile(file) {
  const ext = file.name.split(".").pop().toLowerCase();
  if (ext === "csv") {
    return new Promise((resolve, reject) => {
      Papa.parse(file, {
        header: true,
        skipEmptyLines: true,
        complete: (result) => resolve(result.data || []),
        error: reject
      });
    });
  }
  const buffer = await file.arrayBuffer();
  const workbook = XLSX.read(buffer, { type: "array" });
  const firstSheet = workbook.SheetNames[0];
  return XLSX.utils.sheet_to_json(workbook.Sheets[firstSheet], { defval: "" });
}

function processReports() {
  const scope = $("processingScope").value;
  const sellerFilter = $("sellerFilter").value;
  let rows = scope === "all" ? Object.values(state.parsedReports).flat() : (state.parsedReports[scope] || []);
  rows = rows.filter((row) => sellerFilter === "ALL" || row.sellerGstin === sellerFilter);
  if (!rows.length) {
    alert("Selected scope or seller GSTIN me data available nahi hai.");
    return;
  }

  const summary = buildSummary(rows);
  const sellerSummaries = buildSellerSummaries(rows);
  const ledgers = buildLedgers(rows);
  const vouchers = buildVouchers(rows);
  const voucherTypeSummary = buildVoucherTypeSummary(vouchers);
  const gstReturns = buildGstReturns(rows);
  const sellerExports = buildSellerExports(rows, ledgers);
  const xml = buildTallyXml({
    companyName: getActiveCompanyName(),
    ledgers,
    vouchers
  });

  state.processed = { rows, summary, sellerSummaries, ledgers, vouchers, voucherTypeSummary, gstReturns, sellerExports, xml };
  renderProcessedData();
}

function normalizeRows(rows, reportType, mapping) {
  return rows.map((row) => normalizeRow(row, reportType, mapping)).filter((row) => row.sellerGstin && row.invoiceNumber);
}

function normalizeRow(rawRow, reportType, mapping) {
  const row = lowerCaseKeys(rawRow);
  const taxableValue = toSignedNumber(getMappedValue(row, mapping, "taxableValue"));
  const totalAmount = toSignedNumber(getMappedValue(row, mapping, "totalAmount"));
  const igst = toSignedNumber(getMappedValue(row, mapping, "igst"));
  const cgst = toSignedNumber(getMappedValue(row, mapping, "cgst"));
  const sgst = toSignedNumber(getMappedValue(row, mapping, "sgst"));
  const cess = toSignedNumber(getMappedValue(row, mapping, "cess"));
  const tcsIgstAmount = toSignedNumber(getMappedValue(row, mapping, "tcsIgstAmount"));
  const tcsCgstAmount = toSignedNumber(getMappedValue(row, mapping, "tcsCgstAmount"));
  const tcsSgstAmount = toSignedNumber(getMappedValue(row, mapping, "tcsSgstAmount"));
  const rate = detectTaxRate(getMappedValue(row, mapping, "taxRate"), taxableValue, igst, cgst, sgst);
  const transactionType = String(getMappedValue(row, mapping, "transactionType") || "").trim();
  const sellerGstin = sanitizeGstin(getMappedValue(row, mapping, "sellerGstin"));
  const gstin = sanitizeGstin(getMappedValue(row, mapping, "gstin"));
  const quantity = toNumber(getMappedValue(row, mapping, "quantity")) || 1;
  const invoiceNumber = String(getMappedValue(row, mapping, "invoiceNumber") || `AUTO-${Math.random().toString(36).slice(2, 7)}`).trim();
  const isReturn = /refund|cancel|return/i.test(transactionType) || totalAmount < 0;
  const voucherType = classifyVoucherType(reportType, transactionType, totalAmount, tcsIgstAmount + tcsCgstAmount + tcsSgstAmount);

  return {
    sourceType: reportType.toUpperCase(),
    sellerGstin,
    invoiceNumber,
    voucherDate: normalizeDate(getMappedValue(row, mapping, "voucherDate")) || $("defaultDate").value,
    transactionType,
    orderId: String(getMappedValue(row, mapping, "orderId") || "").trim(),
    quantity,
    itemDescription: String(getMappedValue(row, mapping, "itemDescription") || defaultItemDescription(reportType)).trim(),
    hsn: String(getMappedValue(row, mapping, "hsn") || "NA").trim(),
    partyName: String(getMappedValue(row, mapping, "partyName") || defaultPartyName(reportType)).trim(),
    gstin,
    state: normalizeState(getMappedValue(row, mapping, "state")) || "NA",
    taxableValue,
    totalAmount,
    igst,
    cgst,
    sgst,
    cess,
    taxRate: rate,
    voucherType,
    isReturn,
    creditNoteNo: String(getMappedValue(row, mapping, "creditNoteNo") || "").trim(),
    creditNoteDate: normalizeDate(getMappedValue(row, mapping, "creditNoteDate")),
    tcsIgstAmount,
    tcsCgstAmount,
    tcsSgstAmount,
    tcsAmount: round2(tcsIgstAmount + tcsCgstAmount + tcsSgstAmount)
  };
}

function classifyVoucherType(reportType, transactionType, totalAmount, tcsAmount) {
  if (reportType === "str") return $("stockVoucherName").value.trim() || "Stock Journal";
  if (tcsAmount !== 0) return $("tcsVoucherName").value.trim() || "Journal";
  if (/refund|cancel|return/i.test(transactionType) || totalAmount < 0) return $("returnVoucherName").value.trim() || "Credit Note";
  if (/debit/i.test(transactionType)) return $("debitVoucherName").value.trim() || "Debit Note";
  if (/credit/i.test(transactionType)) return $("creditVoucherName").value.trim() || "Credit Note";
  return $("salesVoucherName").value.trim() || "Sales";
}

function buildSummary(rows) {
  const buckets = groupBy(rows, (row) => row.sourceType);
  const data = Object.entries(buckets).map(([name, bucketRows]) => ({
    name,
    invoices: bucketRows.length,
    taxable: round2(sum(bucketRows, "taxableValue")),
    igst: round2(sum(bucketRows, "igst")),
    cgst: round2(sum(bucketRows, "cgst")),
    sgst: round2(sum(bucketRows, "sgst")),
    total: round2(sum(bucketRows, "totalAmount"))
  }));
  return {
    buckets: data,
    totals: {
      invoices: rows.length,
      taxable: round2(sum(rows, "taxableValue")),
      igst: round2(sum(rows, "igst")),
      cgst: round2(sum(rows, "cgst")),
      sgst: round2(sum(rows, "sgst")),
      total: round2(sum(rows, "totalAmount"))
    }
  };
}

function buildSellerSummaries(rows) {
  return Object.entries(groupBy(rows, (row) => row.sellerGstin)).map(([sellerGstin, sellerRows]) => ({
    sellerGstin,
    rows: sellerRows.length,
    b2bCount: sellerRows.filter((row) => row.sourceType === "B2B").length,
    b2cCount: sellerRows.filter((row) => row.sourceType === "B2C").length,
    strCount: sellerRows.filter((row) => row.sourceType === "STR").length,
    taxableValue: round2(sum(sellerRows, "taxableValue")),
    igst: round2(sum(sellerRows, "igst")),
    cgst: round2(sum(sellerRows, "cgst")),
    sgst: round2(sum(sellerRows, "sgst")),
    totalAmount: round2(sum(sellerRows, "totalAmount"))
  })).sort((a, b) => a.sellerGstin.localeCompare(b.sellerGstin));
}

function buildLedgers(rows) {
  const ledgerMap = new Map();
  const baseLedgers = [
    { ledgerName: $("salesLedgerName").value.trim(), group: "Sales Accounts" },
    { ledgerName: $("salesReturnLedgerName").value.trim(), group: "Sales Accounts" },
    { ledgerName: $("amazonLedgerName").value.trim(), group: $("ledgerGroup").value },
    { ledgerName: $("tcsLedgerName").value.trim(), group: "Duties & Taxes" },
    { ledgerName: "Output IGST", group: "Duties & Taxes" },
    { ledgerName: "Output CGST", group: "Duties & Taxes" },
    { ledgerName: "Output SGST", group: "Duties & Taxes" }
  ];
  baseLedgers.forEach((ledger) => ledgerMap.set(ledger.ledgerName, { ...ledger, gstin: "", state: "NA", openingBalance: 0 }));
  rows.forEach((row) => {
    const name = row.partyName || "Amazon Party";
    if (!ledgerMap.has(name)) {
      ledgerMap.set(name, {
        ledgerName: name,
        group: $("ledgerGroup").value,
        gstin: row.gstin || "",
        state: row.state || "NA",
        openingBalance: 0
      });
    }
  });
  return [...ledgerMap.values()];
}

function buildVouchers(rows) {
  const salesLedger = $("salesLedgerName").value.trim();
  const salesReturnLedger = $("salesReturnLedgerName").value.trim();
  const tcsLedger = $("tcsLedgerName").value.trim();
  return rows.flatMap((row) => {
    const entries = [{
      sellerGstin: row.sellerGstin,
      date: row.voucherDate,
      voucherType: row.voucherType,
      voucherNumber: row.isReturn && row.creditNoteNo ? row.creditNoteNo : row.invoiceNumber,
      partyLedger: row.partyName,
      salesLedger: row.isReturn ? salesReturnLedger : salesLedger,
      taxableValue: round2(Math.abs(row.taxableValue)),
      igst: round2(Math.abs(row.igst)),
      cgst: round2(Math.abs(row.cgst)),
      sgst: round2(Math.abs(row.sgst)),
      totalAmount: round2(Math.abs(row.totalAmount)),
      source: row.sourceType,
      orderId: row.orderId
    }];
    if (row.tcsAmount) {
      entries.push({
        sellerGstin: row.sellerGstin,
        date: row.voucherDate,
        voucherType: $("tcsVoucherName").value.trim() || "Journal",
        voucherNumber: `${row.invoiceNumber}-TCS`,
        partyLedger: row.partyName,
        salesLedger: tcsLedger,
        taxableValue: 0,
        igst: round2(Math.abs(row.tcsIgstAmount)),
        cgst: round2(Math.abs(row.tcsCgstAmount)),
        sgst: round2(Math.abs(row.tcsSgstAmount)),
        totalAmount: round2(Math.abs(row.tcsAmount)),
        source: `${row.sourceType}-TCS`,
        orderId: row.orderId
      });
    }
    return entries;
  });
}

function buildVoucherTypeSummary(vouchers) {
  return Object.entries(groupBy(vouchers, (voucher) => voucher.voucherType)).map(([voucherType, rows]) => ({
    voucherType,
    count: rows.length,
    totalAmount: round2(sum(rows, "totalAmount"))
  }));
}

function buildGstReturns(rows) {
  const b2b = rows.filter((row) => row.sourceType === "B2B" && row.gstin).map((row) => ({
    sellerGstin: row.sellerGstin,
    gstin: row.gstin,
    invoiceNumber: row.invoiceNumber,
    invoiceDate: row.voucherDate,
    placeOfSupply: row.state,
    taxableValue: round2(Math.abs(row.taxableValue)),
    taxAmount: round2(Math.abs(row.igst + row.cgst + row.sgst + row.cess)),
    totalInvoiceValue: round2(Math.abs(row.totalAmount))
  }));

  const b2cs = Object.entries(groupBy(rows.filter((row) => row.sourceType === "B2C"), (row) => `${row.sellerGstin}|${row.state}|${row.taxRate}`)).map(([key, groupRows]) => {
    const [sellerGstin, placeOfSupply, taxRate] = key.split("|");
    return {
      sellerGstin,
      placeOfSupply,
      taxRate: toNumber(taxRate),
      transactions: groupRows.length,
      taxableValue: round2(sumAbs(groupRows, "taxableValue")),
      taxAmount: round2(sumAbsTax(groupRows)),
      totalAmount: round2(sumAbs(groupRows, "totalAmount"))
    };
  });

  const hsn = Object.entries(groupBy(rows, (row) => `${row.sellerGstin}|${row.hsn}|${row.itemDescription}`)).map(([key, groupRows]) => {
    const [sellerGstin, hsnCode, description] = key.split("|");
    return {
      sellerGstin,
      hsn: hsnCode,
      description,
      quantity: round2(sumAbs(groupRows, "quantity")),
      taxableValue: round2(sumAbs(groupRows, "taxableValue")),
      taxAmount: round2(sumAbsTax(groupRows)),
      totalAmount: round2(sumAbs(groupRows, "totalAmount"))
    };
  });

  const gstr3b = [
    summarize3bBucket("Outward Taxable Supplies", rows.filter((row) => row.sourceType === "B2B" || row.sourceType === "B2C")),
    summarize3bBucket("Stock Transfers", rows.filter((row) => row.sourceType === "STR"))
  ];
  return { gstr1: { b2b, b2cs, hsn }, gstr3b };
}

function summarize3bBucket(bucket, rows) {
  return {
    bucket,
    taxableValue: round2(sumAbs(rows, "taxableValue")),
    igst: round2(sumAbs(rows, "igst")),
    cgst: round2(sumAbs(rows, "cgst")),
    sgst: round2(sumAbs(rows, "sgst")),
    totalTax: round2(sumAbsTax(rows))
  };
}

function buildSellerExports(rows, ledgers) {
  return Object.entries(groupBy(rows, (row) => row.sellerGstin)).map(([sellerGstin, sellerRows]) => {
    const sellerVouchers = buildVouchers(sellerRows);
    const companyProfile = getCompanyProfileForSellerGstin(sellerGstin);
    return {
      sellerGstin,
      companyName: companyProfile?.name || getActiveCompanyName(sellerGstin),
      b2b: stateSafeCsv(buildGstReturns(sellerRows).gstr1.b2b),
      b2cs: stateSafeCsv(buildGstReturns(sellerRows).gstr1.b2cs),
      hsn: stateSafeCsv(buildGstReturns(sellerRows).gstr1.hsn),
      gstr3b: buildGstReturns(sellerRows).gstr3b,
      xml: buildTallyXml({
        companyName: companyProfile?.name || getActiveCompanyName(sellerGstin),
        ledgers,
        vouchers: sellerVouchers
      })
    };
  }).sort((a, b) => a.sellerGstin.localeCompare(b.sellerGstin));
}

function buildTallyXml({ companyName, ledgers, vouchers }) {
  const ledgerXml = ledgers.map((ledger) => `
    <TALLYMESSAGE xmlns:UDF="TallyUDF">
      <LEDGER NAME="${escapeXml(ledger.ledgerName)}" RESERVEDNAME="">
        <PARENT>${escapeXml(ledger.group)}</PARENT>
        <OPENINGBALANCE>${round2(ledger.openingBalance)}</OPENINGBALANCE>
        <PARTYGSTIN>${escapeXml(ledger.gstin || "")}</PARTYGSTIN>
      </LEDGER>
    </TALLYMESSAGE>`).join("");

  const voucherXml = vouchers.map((voucher) => `
    <TALLYMESSAGE xmlns:UDF="TallyUDF">
      <VOUCHER VCHTYPE="${escapeXml(voucher.voucherType)}" ACTION="Create">
        <DATE>${formatTallyDate(voucher.date)}</DATE>
        <VOUCHERNUMBER>${escapeXml(voucher.voucherNumber)}</VOUCHERNUMBER>
        <PARTYLEDGERNAME>${escapeXml(voucher.partyLedger)}</PARTYLEDGERNAME>
        <ALLLEDGERENTRIES.LIST>
          <LEDGERNAME>${escapeXml(voucher.partyLedger)}</LEDGERNAME>
          <ISDEEMEDPOSITIVE>Yes</ISDEEMEDPOSITIVE>
          <AMOUNT>-${round2(voucher.totalAmount)}</AMOUNT>
        </ALLLEDGERENTRIES.LIST>
        <ALLLEDGERENTRIES.LIST>
          <LEDGERNAME>${escapeXml(voucher.salesLedger)}</LEDGERNAME>
          <ISDEEMEDPOSITIVE>No</ISDEEMEDPOSITIVE>
          <AMOUNT>${round2(voucher.taxableValue)}</AMOUNT>
        </ALLLEDGERENTRIES.LIST>
        ${voucher.igst ? taxEntry("Output IGST", voucher.igst) : ""}
        ${voucher.cgst ? taxEntry("Output CGST", voucher.cgst) : ""}
        ${voucher.sgst ? taxEntry("Output SGST", voucher.sgst) : ""}
      </VOUCHER>
    </TALLYMESSAGE>`).join("");

  return `<?xml version="1.0" encoding="UTF-8"?>
<ENVELOPE>
  <HEADER><TALLYREQUEST>Import Data</TALLYREQUEST></HEADER>
  <BODY>
    <IMPORTDATA>
      <REQUESTDESC>
        <REPORTNAME>All Masters</REPORTNAME>
        <STATICVARIABLES>
          <SVCURRENTCOMPANY>${escapeXml(companyName)}</SVCURRENTCOMPANY>
        </STATICVARIABLES>
      </REQUESTDESC>
      <REQUESTDATA>${ledgerXml}${voucherXml}</REQUESTDATA>
    </IMPORTDATA>
  </BODY>
</ENVELOPE>`;
}

function taxEntry(name, amount) {
  return `<ALLLEDGERENTRIES.LIST><LEDGERNAME>${escapeXml(name)}</LEDGERNAME><ISDEEMEDPOSITIVE>No</ISDEEMEDPOSITIVE><AMOUNT>${round2(amount)}</AMOUNT></ALLLEDGERENTRIES.LIST>`;
}

function renderProcessedData() {
  const { rows, summary, sellerSummaries, voucherTypeSummary, gstReturns, sellerExports, xml } = state.processed;
  $("rowsParsed").textContent = rows.length;
  $("gstLiability").textContent = formatMoney(summary.totals.igst + summary.totals.cgst + summary.totals.sgst);
  $("invoiceCountLabel").textContent = `${rows.length} processed invoice lines`;
  $("sellerSummaryCountLabel").textContent = `${sellerSummaries.length} seller GSTIN groups`;
  renderSellerFilter(rows);
  renderGstinCards(sellerSummaries);
  renderSellerSummaryTable(sellerSummaries);
  renderGstr1B2BTable(gstReturns.gstr1.b2b);
  renderGstr1B2CSTable(gstReturns.gstr1.b2cs);
  renderHsnTable(gstReturns.gstr1.hsn);
  renderGstr3bTable(gstReturns.gstr3b);
  renderVoucherTypeTable(voucherTypeSummary);
  renderSellerDownloadCards(sellerExports);
  $("saveStatus").textContent = "Run ready hai. Save Run History button se Supabase me store kar sakte hain.";
  renderCompanySellerOptions(rows);
}

function renderSellerFilter(rows) {
  const select = $("sellerFilter");
  const current = select.value || "ALL";
  const gstins = [...new Set(rows.map((row) => row.sellerGstin))].sort();
  select.innerHTML = `<option value="ALL">All Seller GSTINs</option>${gstins.map((gstin) => `<option value="${gstin}">${gstin}</option>`).join("")}`;
  select.value = gstins.includes(current) || current === "ALL" ? current : "ALL";
}

function renderGstinCards(sellerSummaries) {
  const root = $("gstinCards");
  root.innerHTML = sellerSummaries.map((item) => `
    <article class="gstin-card">
      <strong>${item.sellerGstin}</strong>
      <div>Rows: ${item.rows}</div>
      <div>Taxable: ${formatMoney(item.taxableValue)}</div>
      <div>GST: ${formatMoney(item.igst + item.cgst + item.sgst)}</div>
    </article>`).join("") || `<div class="empty-cell">No GSTIN cards yet.</div>`;
}

function renderSellerSummaryTable(rows) {
  const body = $("sellerSummaryTable").querySelector("tbody");
  body.innerHTML = rows.map((row) => `
    <tr>
      <td>${row.sellerGstin}</td>
      <td>${row.rows}</td>
      <td>${row.b2bCount}</td>
      <td>${row.b2cCount}</td>
      <td>${row.strCount}</td>
      <td>${formatMoney(row.taxableValue)}</td>
      <td>${formatMoney(row.igst)}</td>
      <td>${formatMoney(row.cgst)}</td>
      <td>${formatMoney(row.sgst)}</td>
      <td>${formatMoney(row.totalAmount)}</td>
    </tr>`).join("") || `<tr><td colspan="10" class="empty-cell">No seller GSTIN summary yet.</td></tr>`;
}

function renderGstr1B2BTable(rows) {
  const body = $("gstr1B2BTable").querySelector("tbody");
  body.innerHTML = rows.slice(0, 100).map((row) => `
    <tr>
      <td>${row.sellerGstin}</td>
      <td>${row.gstin}</td>
      <td>${row.invoiceNumber}</td>
      <td>${row.invoiceDate}</td>
      <td>${row.placeOfSupply}</td>
      <td>${formatMoney(row.taxableValue)}</td>
      <td>${formatMoney(row.taxAmount)}</td>
      <td>${formatMoney(row.totalInvoiceValue)}</td>
    </tr>`).join("") || `<tr><td colspan="8" class="empty-cell">No B2B rows yet.</td></tr>`;
}

function renderGstr1B2CSTable(rows) {
  const body = $("gstr1B2CSTable").querySelector("tbody");
  body.innerHTML = rows.slice(0, 100).map((row) => `
    <tr>
      <td>${row.sellerGstin}</td>
      <td>${row.placeOfSupply}</td>
      <td>${row.taxRate}%</td>
      <td>${row.transactions}</td>
      <td>${formatMoney(row.taxableValue)}</td>
      <td>${formatMoney(row.taxAmount)}</td>
      <td>${formatMoney(row.totalAmount)}</td>
    </tr>`).join("") || `<tr><td colspan="7" class="empty-cell">No B2CS rows yet.</td></tr>`;
}

function renderHsnTable(rows) {
  const body = $("hsnTable").querySelector("tbody");
  body.innerHTML = rows.slice(0, 100).map((row) => `
    <tr>
      <td>${row.sellerGstin}</td>
      <td>${row.hsn}</td>
      <td>${row.description}</td>
      <td>${row.quantity}</td>
      <td>${formatMoney(row.taxableValue)}</td>
      <td>${formatMoney(row.taxAmount)}</td>
      <td>${formatMoney(row.totalAmount)}</td>
    </tr>`).join("") || `<tr><td colspan="7" class="empty-cell">No HSN rows yet.</td></tr>`;
}

function renderGstr3bTable(rows) {
  const body = $("gstr3bTable").querySelector("tbody");
  body.innerHTML = rows.map((row) => `
    <tr>
      <td>${row.bucket}</td>
      <td>${formatMoney(row.taxableValue)}</td>
      <td>${formatMoney(row.igst)}</td>
      <td>${formatMoney(row.cgst)}</td>
      <td>${formatMoney(row.sgst)}</td>
      <td>${formatMoney(row.totalTax)}</td>
    </tr>`).join("") || `<tr><td colspan="6" class="empty-cell">No GSTR-3B summary yet.</td></tr>`;
}

function renderVoucherTypeTable(rows) {
  const table = $("voucherTypeTable").querySelector("tbody");
  table.innerHTML = rows.map((row) => `
    <tr>
      <td>${row.voucherType}</td>
      <td>${row.count}</td>
      <td>${formatMoney(row.totalAmount)}</td>
    </tr>`).join("") || `<tr><td colspan="3" class="empty-cell">No voucher summary yet.</td></tr>`;
}

function renderSellerDownloadCards(exportsBySeller) {
  const root = $("sellerDownloads");
  if (!exportsBySeller.length) {
    root.innerHTML = `<div class="empty-cell">No GSTIN-wise download cards yet.</div>`;
    return;
  }
  root.innerHTML = exportsBySeller.map((item, index) => `
    <article class="seller-card">
      <strong>${item.sellerGstin}</strong>
      <span>${item.companyName}</span>
      <div class="seller-actions">
        <button class="btn btn-ghost" data-export="${index}" data-kind="b2b">B2B CSV</button>
        <button class="btn btn-ghost" data-export="${index}" data-kind="b2cs">B2CS CSV</button>
        <button class="btn btn-ghost" data-export="${index}" data-kind="hsn">HSN CSV</button>
        <button class="btn btn-ghost" data-export="${index}" data-kind="gstr3b">GSTR3B JSON</button>
        <button class="btn btn-primary" data-export="${index}" data-kind="xml">Tally XML</button>
      </div>
    </article>`).join("");
  root.querySelectorAll("button[data-export]").forEach((button) => {
    button.addEventListener("click", () => downloadSellerFile(exportsBySeller[Number(button.dataset.export)], button.dataset.kind));
  });
}

function downloadSellerFile(item, kind) {
  const prefix = item.sellerGstin;
  if (kind === "gstr3b") return downloadJson(`${prefix}-GSTR3B.json`, item.gstr3b);
  if (kind === "xml") return downloadText(`${prefix}-Tally.xml`, item.xml, "application/xml");
  downloadText(`${prefix}-${kind.toUpperCase()}.csv`, item[kind], "text/csv");
}

function renderMappingReview() {
  renderMappingTable("b2b", "mappingTableB2B");
  renderMappingTable("b2c", "mappingTableB2C");
  renderMappingTable("str", "mappingTableSTR");
}

function renderMappingTable(reportKey, tableId) {
  const meta = state.reportMeta[reportKey];
  const body = $(tableId).querySelector("tbody");
  const schema = REPORT_SCHEMAS[reportKey];
  if (!meta.columns.length) {
    body.innerHTML = `<tr><td colspan="3" class="empty-cell">Upload ${reportKey.toUpperCase()} file to inspect mapping.</td></tr>`;
    return;
  }
  body.innerHTML = Object.keys(schema).map((field) => `
    <tr>
      <td>${field}</td>
      <td>${meta.mapping[field] || "-"}</td>
      <td>${meta.mapping[field] ? "Mapped" : "Missing"}</td>
    </tr>`).join("");
}

function extractColumns(rows) {
  const set = new Set();
  rows.slice(0, 25).forEach((row) => Object.keys(row || {}).forEach((key) => set.add(String(key).trim())));
  return [...set];
}

function detectSchemaMapping(columns, reportKey) {
  const schema = REPORT_SCHEMAS[reportKey];
  const lowered = columns.map((column) => column.toLowerCase());
  return Object.fromEntries(Object.entries(schema).map(([field, aliases]) => {
    const matched = aliases.find((alias) => lowered.includes(alias)) || aliases.find((alias) => lowered.some((column) => column.includes(alias)));
    const actual = columns.find((column) => column.toLowerCase() === matched) || columns.find((column) => column.toLowerCase().includes(matched || "__"));
    return [field, actual || ""];
  }));
}

function lowerCaseKeys(row) {
  return Object.fromEntries(Object.entries(row || {}).map(([key, value]) => [String(key).trim().toLowerCase(), value]));
}

function getMappedValue(row, mapping, field) {
  const key = String(mapping[field] || "").trim().toLowerCase();
  return key ? row[key] ?? "" : "";
}

function normalizeDate(value) {
  if (!value) return "";
  const date = new Date(value);
  if (!Number.isNaN(date.getTime())) return date.toISOString().slice(0, 10);
  return "";
}

function sanitizeGstin(value) {
  return String(value || "").trim().toUpperCase();
}

function normalizeState(value) {
  return String(value || "").trim().toUpperCase();
}

function toNumber(value) {
  return Number(String(value || "0").replace(/,/g, "").trim()) || 0;
}

function toSignedNumber(value) {
  return Number(String(value || "0").replace(/,/g, "").trim()) || 0;
}

function detectTaxRate(value, taxableValue, igst, cgst, sgst) {
  const explicit = toNumber(value) * 100;
  if (explicit) return round2(explicit);
  if (!taxableValue) return 0;
  return round2((Math.abs(igst) + Math.abs(cgst) + Math.abs(sgst)) / Math.abs(taxableValue) * 100);
}

function groupBy(rows, getKey) {
  return rows.reduce((acc, row) => {
    const key = getKey(row);
    if (!acc[key]) acc[key] = [];
    acc[key].push(row);
    return acc;
  }, {});
}

function sum(rows, key) {
  return rows.reduce((acc, row) => acc + (Number(row[key]) || 0), 0);
}

function sumAbs(rows, key) {
  return rows.reduce((acc, row) => acc + Math.abs(Number(row[key]) || 0), 0);
}

function sumAbsTax(rows) {
  return rows.reduce((acc, row) => acc + Math.abs((Number(row.igst) || 0) + (Number(row.cgst) || 0) + (Number(row.sgst) || 0) + (Number(row.cess) || 0)), 0);
}

function round2(value) {
  return Number((Number(value) || 0).toFixed(2));
}

function formatMoney(value) {
  return new Intl.NumberFormat("en-IN", { minimumFractionDigits: 2, maximumFractionDigits: 2 }).format(round2(value));
}

function formatTallyDate(dateString) {
  return String(dateString || "").replaceAll("-", "");
}

function escapeXml(value) {
  return String(value ?? "").replaceAll("&", "&amp;").replaceAll("<", "&lt;").replaceAll(">", "&gt;").replaceAll('"', "&quot;").replaceAll("'", "&apos;");
}

function csvCell(value) {
  return `"${String(value ?? "").replace(/"/g, '""')}"`;
}

function stateSafeCsv(rows) {
  if (!rows.length) return "";
  const headers = Object.keys(rows[0]);
  return [headers.join(","), ...rows.map((row) => headers.map((key) => csvCell(row[key])).join(","))].join("\n");
}

function downloadCsv(filename, rows) {
  if (!rows.length) return alert("Nothing to download yet.");
  downloadText(filename, stateSafeCsv(rows), "text/csv");
}

function downloadJson(filename, data) {
  if (!data || (typeof data === "object" && !Object.keys(data).length)) return alert("Nothing to download yet.");
  downloadText(filename, JSON.stringify(data, null, 2), "application/json");
}

function downloadText(filename, content, mimeType) {
  const blob = new Blob([content], { type: mimeType });
  const url = URL.createObjectURL(blob);
  const link = document.createElement("a");
  link.href = url;
  link.download = filename;
  link.click();
  URL.revokeObjectURL(url);
}

async function saveRunHistory() {
  if (!state.processed) {
    $("saveStatus").textContent = "Pehle reports process karein, phir save karein.";
    return;
  }

  $("saveStatus").textContent = "Saving run history to Supabase...";

  try {
    const uploadedReports = Object.entries(state.parsedReports)
      .filter(([, rows]) => rows.length)
      .map(([key]) => key.toUpperCase());

    const response = await fetch("/api/save-run", {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify({
        company_name: $("companyName").value.trim(),
        uploaded_reports: uploadedReports,
        summary_json: state.processed.summary,
        vouchers_json: state.processed.vouchers,
        seller_summary_json: state.processed.sellerSummaries,
        gstr1_json: state.processed.gstReturns?.gstr1 || {},
        gstr3b_json: state.processed.gstReturns?.gstr3b || []
      })
    });

    const result = await response.json();
    if (!response.ok) {
      const errorMessage = typeof result.error === "string" ? result.error : JSON.stringify(result.error);
      throw new Error(errorMessage);
    }

    $("saveStatus").textContent = "Run history Supabase me save ho gayi.";
  } catch (error) {
    $("saveStatus").textContent = `Save failed: ${error.message}`;
  }
}

function updateUploadStats() {
  $("uploadedCount").textContent = Object.values(state.parsedReports).filter((rows) => rows.length).length;
}

function defaultPartyName(reportType) {
  return reportType === "str" ? "Amazon Branch Transfer" : "Amazon Customer";
}

function defaultItemDescription(reportType) {
  return reportType === "str" ? "Amazon Stock Transfer Item" : "Amazon Sale Item";
}

function loadDemoData() {
  const date = $("defaultDate").value;
  state.parsedReports = {
    b2b: [
      { sourceType: "B2B", sellerGstin: "06AAHFO5288K1ZQ", invoiceNumber: "DED4-6284", voucherDate: date, transactionType: "Shipment", orderId: "408-9216832-9847560", quantity: 1, itemDescription: "Amazon B2B Product", hsn: "63059000", partyName: "SBG CONSULTANTS", gstin: "06AKQPR6345F1ZR", state: "HARYANA", taxableValue: 432.38, totalAmount: 454, igst: 0, cgst: 10.81, sgst: 10.81, cess: 0, taxRate: 5, voucherType: "Sales", isReturn: false, creditNoteNo: "", creditNoteDate: "", tcsIgstAmount: 0, tcsCgstAmount: 1.08, tcsSgstAmount: 1.08, tcsAmount: 2.16 }
    ],
    b2c: [
      { sourceType: "B2C", sellerGstin: "07AAHFO5288K1ZO", invoiceNumber: "B2C-1001", voucherDate: date, transactionType: "Shipment", orderId: "B2C-ORD", quantity: 1, itemDescription: "Amazon B2C Product", hsn: "42022990", partyName: "Amazon Customer", gstin: "", state: "MAHARASHTRA", taxableValue: 456.78, totalAmount: 539, igst: 82.22, cgst: 0, sgst: 0, cess: 0, taxRate: 18, voucherType: "Sales", isReturn: false, creditNoteNo: "", creditNoteDate: "", tcsIgstAmount: 2.28, tcsCgstAmount: 0, tcsSgstAmount: 0, tcsAmount: 2.28 }
    ],
    str: [
      { sourceType: "STR", sellerGstin: "27AAHFO5288K1ZM", invoiceNumber: "STR-101", voucherDate: date, transactionType: "FC_TRANSFER", orderId: "TR-1", quantity: 2, itemDescription: "B083ZVMJ5B", hsn: "48192090", partyName: "07AAHFO5288K1ZO", gstin: "07AAHFO5288K1ZO", state: "DELHI", taxableValue: 1000, totalAmount: 1000, igst: 50, cgst: 0, sgst: 0, cess: 0, taxRate: 5, voucherType: "Stock Journal", isReturn: false, creditNoteNo: "", creditNoteDate: "", tcsIgstAmount: 0, tcsCgstAmount: 0, tcsSgstAmount: 0, tcsAmount: 0 }
    ]
  };
  state.reportMeta = {
    b2b: { columns: Object.values(REPORT_SCHEMAS.b2b).flat(), mapping: detectSchemaMapping(Object.values(REPORT_SCHEMAS.b2b).flat(), "b2b") },
    b2c: { columns: Object.values(REPORT_SCHEMAS.b2c).flat(), mapping: detectSchemaMapping(Object.values(REPORT_SCHEMAS.b2c).flat(), "b2c") },
    str: { columns: Object.values(REPORT_SCHEMAS.str).flat(), mapping: detectSchemaMapping(Object.values(REPORT_SCHEMAS.str).flat(), "str") }
  };
  updateUploadStats();
  renderMappingReview();
  processReports();
}

function loadCompanyProfiles() {
  try {
    state.companyProfiles = JSON.parse(localStorage.getItem("amazon-gst-tally-companies") || "[]");
  } catch {
    state.companyProfiles = [];
  }
}

function saveCompanyProfiles() {
  localStorage.setItem("amazon-gst-tally-companies", JSON.stringify(state.companyProfiles));
}

function addCompanyProfile() {
  const name = $("tallyCompanyName").value.trim();
  const gstin = sanitizeGstin($("tallyCompanyGstin").value);
  const sellerGstin = sanitizeGstin($("companySellerGstin").value);
  if (!name) {
    $("saveStatus").textContent = "Company name required hai.";
    return;
  }
  const profile = { id: sellerGstin || `ALL-${Date.now()}`, name, gstin, sellerGstin };
  const index = state.companyProfiles.findIndex((item) => item.sellerGstin === sellerGstin && sellerGstin);
  if (index >= 0) state.companyProfiles[index] = profile;
  else state.companyProfiles.push(profile);
  saveCompanyProfiles();
  renderCompanyCards();
  $("saveStatus").textContent = sellerGstin
    ? `Company profile ${sellerGstin} ke liye save ho gayi.`
    : "Default company profile save ho gayi.";
}

function renderCompanyCards() {
  const root = $("companyCards");
  if (!root) return;
  if (!state.companyProfiles.length) {
    root.innerHTML = `<div class="empty-cell">Abhi koi company saved nahi hai.</div>`;
    return;
  }
  root.innerHTML = state.companyProfiles.map((item, index) => `
    <article class="seller-card">
      <strong>${item.name}</strong>
      <span>Seller GSTIN: ${item.sellerGstin || "All Seller GSTINs"}</span>
      <span>Company GSTIN: ${item.gstin || "Not set"}</span>
      <div class="seller-actions">
        <button class="btn btn-ghost" data-company-select="${index}">Use</button>
        <button class="btn btn-ghost" data-company-delete="${index}">Delete</button>
      </div>
    </article>`).join("");
  root.querySelectorAll("button[data-company-select]").forEach((button) => {
    button.addEventListener("click", () => selectCompanyProfile(Number(button.dataset.companySelect)));
  });
  root.querySelectorAll("button[data-company-delete]").forEach((button) => {
    button.addEventListener("click", () => deleteCompanyProfile(Number(button.dataset.companyDelete)));
  });
}

function selectCompanyProfile(index) {
  const item = state.companyProfiles[index];
  if (!item) return;
  $("tallyCompanyName").value = item.name || "";
  $("tallyCompanyGstin").value = item.gstin || "";
  $("companySellerGstin").value = item.sellerGstin || "";
  $("saveStatus").textContent = `${item.name} selected.`;
}

function deleteCompanyProfile(index) {
  state.companyProfiles.splice(index, 1);
  saveCompanyProfiles();
  renderCompanyCards();
  $("saveStatus").textContent = "Company profile delete ho gayi.";
}

function renderCompanySellerOptions(rows) {
  const select = $("companySellerGstin");
  if (!select) return;
  const current = select.value || "";
  const gstins = [...new Set(rows.map((row) => row.sellerGstin))].sort();
  select.innerHTML = `<option value="">Select Seller GSTIN</option>${gstins.map((gstin) => `<option value="${gstin}">${gstin}</option>`).join("")}`;
  select.value = gstins.includes(current) ? current : "";
}

function getCompanyProfileForSellerGstin(sellerGstin) {
  return state.companyProfiles.find((item) => item.sellerGstin === sellerGstin)
    || state.companyProfiles.find((item) => !item.sellerGstin)
    || null;
}

function getActiveCompanyName(sellerGstin = "") {
  return getCompanyProfileForSellerGstin(sellerGstin)?.name
    || $("tallyCompanyName").value.trim()
    || $("companyName").value.trim();
}
