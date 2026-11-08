// =========================================================
// DELIVERY NOTE PDF GENERATOR
// DESIGN FOLLOW WORD TEMPLATE
// =========================================================

const SOURCE_SHEET = "DN";
const FOLDER_NAME = "project";
const PDF_PREFIX = "Delivery Note ";
const MARGIN_INCH = 0.5;

const ACTION_HEADER = "GENERATE PDF";
const PDF_LINK_HEADER = "PDF";

const BRAND_COLOR = "#355E3B";


// =========================================================
// ON OPEN
// =========================================================

function onOpen() {

  try {

    const ui = SpreadsheetApp.getUi();

    ui.createMenu("📄 Delivery Note")
      .addItem(
        "Generate PDF dari baris aktif",
        "generateDeliveryNotePDF"
      )
      .addItem(
        "Setup tombol PDF",
        "setupPDFButton"
      )
      .addToUi();

  } catch (error) {

    console.log(
      "onOpen dijalankan di luar konteks Spreadsheet UI: " +
      error.message
    );

  }

}


// =========================================================
// SETUP CHECKBOX
// =========================================================

function setupPDFButton() {

  const ss =
    SpreadsheetApp.getActiveSpreadsheet();

  const sheet =
    ss.getSheetByName(SOURCE_SHEET);

  if (!sheet) {

    throw new Error(
      `Sheet "${SOURCE_SHEET}" tidak ditemukan.`
    );

  }


  let lastColumn =
    sheet.getLastColumn();


  let headers =
    sheet
      .getRange(
        1,
        1,
        1,
        lastColumn
      )
      .getDisplayValues()[0];


  // -------------------------------------------------------
  // ACTION COLUMN
  // -------------------------------------------------------

  let actionCol =
    headers.findIndex(
      h =>
        String(h)
          .trim()
          .toUpperCase() ===
        ACTION_HEADER
    ) + 1;


  if (actionCol === 0) {

    actionCol =
      lastColumn + 1;

    sheet
      .getRange(
        1,
        actionCol
      )
      .setValue(
        ACTION_HEADER
      );

  }


  // -------------------------------------------------------
  // PDF COLUMN
  // -------------------------------------------------------

  lastColumn =
    sheet.getLastColumn();

  headers =
    sheet
      .getRange(
        1,
        1,
        1,
        lastColumn
      )
      .getDisplayValues()[0];


  let pdfCol =
    headers.findIndex(
      h =>
        String(h)
          .trim()
          .toUpperCase() ===
        PDF_LINK_HEADER
    ) + 1;


  if (pdfCol === 0) {

    pdfCol =
      sheet.getLastColumn() + 1;

    sheet
      .getRange(
        1,
        pdfCol
      )
      .setValue(
        PDF_LINK_HEADER
      );

  }


  // -------------------------------------------------------
  // CHECKBOX
  // -------------------------------------------------------

  const lastRow =
    sheet.getLastRow();


  if (lastRow >= 2) {

    const actionRange =
      sheet.getRange(
        2,
        actionCol,
        lastRow - 1,
        1
      );


    actionRange.insertCheckboxes();

    actionRange.uncheck();

    actionRange
      .setHorizontalAlignment(
        "center"
      );

  }


  // -------------------------------------------------------
  // HEADER FORMAT
  // -------------------------------------------------------

  sheet
    .getRange(
      1,
      actionCol
    )
    .setFontWeight("bold")
    .setHorizontalAlignment("center");


  sheet
    .getRange(
      1,
      pdfCol
    )
    .setFontWeight("bold");


  // -------------------------------------------------------
  // SUCCESS MESSAGE
  // -------------------------------------------------------

  ss.toast(
    `Checkbox "${ACTION_HEADER}" berhasil dibuat.\n\n` +
    "Centang checkbox pada baris PO untuk membuat Delivery Note.",
    "Delivery Note",
    5
  );

}


// =========================================================
// ON EDIT
// =========================================================

function onEdit(e) {

  if (!e || !e.range) {
    return;
  }


  const sheet =
    e.range.getSheet();


  if (
    sheet.getName() !==
    SOURCE_SHEET
  ) {
    return;
  }


  if (
    e.range.getNumRows() !== 1 ||
    e.range.getNumColumns() !== 1
  ) {
    return;
  }


  const row =
    e.range.getRow();

  const col =
    e.range.getColumn();


  if (row <= 1) {
    return;
  }


  const headers =
    sheet
      .getRange(
        1,
        1,
        1,
        sheet.getLastColumn()
      )
      .getDisplayValues()[0];


  const actionCol =
    headers.findIndex(
      h =>
        String(h)
          .trim()
          .toUpperCase() ===
        ACTION_HEADER
    ) + 1;


  if (
    actionCol === 0 ||
    col !== actionCol
  ) {
    return;
  }


  if (e.value !== "TRUE") {
    return;
  }


  try {

    generateDeliveryNotePDF(row);

  } catch (error) {

    console.error(
      "Gagal membuat PDF:",
      error
    );

    e.source.toast(
      "Gagal membuat PDF:\n\n" +
      error.message,
      "Delivery Note",
      8
    );

  }

}


// =========================================================
// MAIN FUNCTION
// =========================================================

function generateDeliveryNotePDF(selectedRow) {

  const ss =
    SpreadsheetApp
      .getActiveSpreadsheet();


  const sourceSheet =
    ss.getSheetByName(
      SOURCE_SHEET
    );


  if (!sourceSheet) {

    throw new Error(
      `Sheet "${SOURCE_SHEET}" tidak ditemukan.`
    );

  }


  // -------------------------------------------------------
  // ACTIVE ROW
  // -------------------------------------------------------

  if (!selectedRow) {

    selectedRow =
      sourceSheet
        .getActiveRange()
        .getRow();

  }


  if (selectedRow <= 1) {

    throw new Error(
      "Pilih baris data terlebih dahulu."
    );

  }


  // =======================================================
  // HELPER
  // =======================================================

  function normalize(value) {

    return String(value || "")
      .trim()
      .toLowerCase()
      .replace(/\s+/g, " ");

  }


  function escapeHtml(value) {

    if (
      value === null ||
      value === undefined
    ) {
      return "";
    }


    return String(value)
      .replace(/&/g, "&amp;")
      .replace(/</g, "&lt;")
      .replace(/>/g, "&gt;")
      .replace(/"/g, "&quot;")
      .replace(/'/g, "&#039;");

  }


  function parseNumber(value) {

    if (
      value === null ||
      value === undefined ||
      value === ""
    ) {
      return 0;
    }


    let text =
      String(value)
        .trim();


    if (
      text.includes(".") &&
      text.includes(",")
    ) {

      text =
        text
          .replace(/\./g, "")
          .replace(",", ".");

    }

    else if (
      text.includes(".")
    ) {

      const parts =
        text.split(".");


      if (
        parts.length === 2 &&
        parts[1].length === 3
      ) {

        text =
          text.replace(/\./g, "");

      }

    }

    else if (
      text.includes(",")
    ) {

      const parts =
        text.split(",");


      if (
        parts.length === 2 &&
        parts[1].length === 3
      ) {

        text =
          text.replace(/,/g, "");

      } else {

        text =
          text.replace(",", ".");

      }

    }


    text =
      text.replace(
        /[^0-9.-]/g,
        ""
      );


    const number =
      Number(text);


    return isNaN(number)
      ? 0
      : number;

  }


  function formatNumber(value) {

    return parseNumber(value)
      .toLocaleString(
        "id-ID"
      );

  }


  function formatPrice(value) {

    return parseNumber(value)
      .toLocaleString(
        "id-ID",
        {
          minimumFractionDigits: 0,
          maximumFractionDigits: 0
        }
      );

  }


  // =======================================================
  // FORMAT BULAN
  // =======================================================

  function formatMonthYear(value) {

    if (!value) {
      return "";
    }


    const d =
      new Date(value);


    if (
      !isNaN(d.getTime())
    ) {

      return d
        .toLocaleDateString(
          "en-US",
          {
            month: "short",
            year: "2-digit"
          }
        )
        .replace(",", "");

    }


    return String(value);

  }


  // =======================================================
  // READ DATA
  // =======================================================

  const data =
    sourceSheet
      .getDataRange()
      .getDisplayValues();


  // =======================================================
  // FIND HEADER
  // =======================================================

  function findHeader(headerName) {

    const target =
      normalize(headerName);


    for (
      let r = 0;
      r < data.length;
      r++
    ) {

      for (
        let c = 0;
        c < data[r].length;
        c++
      ) {

        if (
          normalize(data[r][c]) ===
          target
        ) {

          return {
            row: r,
            col: c
          };

        }

      }

    }


    return null;

  }


  const poHeader =
    findHeader("NO PO");

  const skuHeader =
    findHeader("SKU");

  const qtyHeader =
    findHeader("qty");

  const priceHeader =
    findHeader("price");

  const clientHeader =
    findHeader("NAMA CLIENT");

  const addressHeader =
    findHeader("address");

  const picHeader =
    findHeader("PIC");


  if (!poHeader) {

    throw new Error(
      'Header "NO PO" tidak ditemukan.'
    );

  }


  if (!skuHeader) {

    throw new Error(
      'Header "SKU" tidak ditemukan.'
    );

  }


  if (!qtyHeader) {

    throw new Error(
      'Header "qty" tidak ditemukan.'
    );

  }


  if (!priceHeader) {

    throw new Error(
      'Header "price" tidak ditemukan.'
    );

  }


  // =======================================================
  // SELECTED PO
  // =======================================================

  const selectedDataRow =
    data[selectedRow - 1];


  const poNumber =
    selectedDataRow[
      poHeader.col
    ];


  if (
    !poNumber ||
    String(poNumber).trim() === ""
  ) {

    throw new Error(
      `NO PO pada baris ${selectedRow} kosong.`
    );

  }


  // =======================================================
  // CUSTOMER
  // =======================================================

  const clientName =
    clientHeader
      ? selectedDataRow[
          clientHeader.col
        ]
      : "";


  const clientAddress =
    addressHeader
      ? selectedDataRow[
          addressHeader.col
        ]
      : "";


  const clientPIC =
    picHeader
      ? selectedDataRow[
          picHeader.col
        ]
      : "";


  // =======================================================
  // DOCUMENT INFO
  // =======================================================

  const number =
    selectedDataRow[10] ||
    poNumber;


  const date =
    selectedDataRow[9] ||
    "";


  const terms =
    selectedDataRow[8] ||
    "";


  const dueDate =
    selectedDataRow[11] ||
    "";


  // =======================================================
  // ITEMS
  // =======================================================

  const items = [];

  let subtotal = 0;


  function splitItems(value) {

    if (
      value === null ||
      value === undefined ||
      String(value).trim() === ""
    ) {

      return [];

    }


    return String(value)
      .split(";")
      .map(function(item) {

        return item.trim();

      })
      .filter(function(item) {

        return item !== "";

      });

  }


  // =======================================================
  // LOOP ALL ROWS WITH SAME PO
  // =======================================================

  for (
    let r = 1;
    r < data.length;
    r++
  ) {

    const rowPO =
      data[r][poHeader.col];


    if (
      normalize(rowPO) !==
      normalize(poNumber)
    ) {

      continue;

    }


    const rawSKU =
      data[r][skuHeader.col];

    const rawQty =
      data[r][qtyHeader.col];

    const rawPrice =
      data[r][priceHeader.col];


    const skuList =
      splitItems(rawSKU);

    const qtyList =
      splitItems(rawQty);

    const priceList =
      splitItems(rawPrice);


    if (
      skuList.length === 0
    ) {

      continue;

    }


    if (
      skuList.length !== qtyList.length ||
      skuList.length !== priceList.length
    ) {

      throw new Error(
        "Jumlah SKU, Qty, dan Price tidak sama " +
        "pada baris " + (r + 1) +
        ".\n\n" +
        "SKU: " + skuList.join("; ") +
        "\nQty: " + qtyList.join("; ") +
        "\nPrice: " + priceList.join("; ")
      );

    }


    for (
      let i = 0;
      i < skuList.length;
      i++
    ) {

      const sku =
        skuList[i];

      const qty =
        parseNumber(
          qtyList[i]
        );

      const price =
        parseNumber(
          priceList[i]
        );


      const totalPrice =
        qty * price;


      subtotal +=
        totalPrice;


      items.push({

        sku:
          sku,

        description:
          `PO Distribusi Online B2B ${clientName} ${formatMonthYear(date)}`,

        department:
          "Managed Services",

        className:
          "Everpro MP",

        location:
          "NONE",

        qty:
          qty,

        units:
          "Pcs",

        price:
          price,

        total:
          totalPrice

      });

    }

  }


  // =======================================================
  // VALIDATE ITEMS
  // =======================================================

  if (
    items.length === 0
  ) {

    throw new Error(
      `Tidak ada item untuk NO PO "${poNumber}".`
    );

  }


  // =======================================================
  // SUMMARY
  // =======================================================

  const discount = 0;

  const afterDiscount =
    subtotal -
    discount;

  const vat =
    afterDiscount *
    0.11;

  const wht = 0;

  const prepayment = 0;

  const grandTotal =
    afterDiscount +
    vat -
    wht -
    prepayment;


  // =======================================================
  // TOTAL QTY
  // =======================================================

  const totalQty =
    items.reduce(
      function(sum, item) {

        return sum + item.qty;

      },
      0
    );


  // =======================================================
  // ITEM ROWS
  // =======================================================

  let itemRows = "";


  items.forEach(function(item) {

    itemRows += `

      <tr>

        <td class="item-code">
          ${escapeHtml(item.sku)}
        </td>

        <td class="item-desc">
          ${escapeHtml(item.description)}
        </td>

        <td>
          ${escapeHtml(item.department)}
        </td>

        <td>
          ${escapeHtml(item.className)}
        </td>

        <td>
          ${escapeHtml(item.location)}
        </td>

        <td class="center">
          ${formatNumber(item.qty)}
        </td>

        <td class="center">
          ${escapeHtml(item.units)}
        </td>

        <td class="money">
          ${formatPrice(item.price)}
        </td>

        <td class="money">
          ${formatPrice(item.total)}
        </td>

      </tr>

    `;

  });


  // =======================================================
  // HTML
  // =======================================================

  const html = `

<!DOCTYPE html>

<html>

<head>

<meta charset="UTF-8">

<style>

@page {

  size: A4;

  margin: ${MARGIN_INCH}in;

}


* {

  box-sizing: border-box;

}


body {

  margin: 0;

  padding: 0;

  font-family:
    Arial,
    Helvetica,
    sans-serif;

  font-size: 9pt;

  color: #000;

  line-height: 1.15;

}


/* =======================================================
   HEADER
   ======================================================= */

.header-table {

  width: 100%;

  border-collapse: collapse;

  margin-bottom: 12px;

}


.header-left {

  width: 48%;

  vertical-align: top;

  padding: 0;

}


.header-right {

  width: 52%;

  vertical-align: top;

  padding: 0;

}


.company-name {

  font-size: 10pt;

  font-weight: bold;

  margin-bottom: 3px;

}


.company-address {

  font-size: 10pt;

  line-height: 1.15;

}


.document-title {

  color: ${BRAND_COLOR};

  font-size: 20pt;

  font-weight: bold;

  text-align: left;

  padding-bottom: 5px;

}


/* =======================================================
   DOCUMENT INFO
   ======================================================= */

.doc-info {

  width: 100%;

  border-collapse: collapse;

  font-size: 9pt;

}


.doc-info td {

  padding: 1px 2px;

  vertical-align: top;

}


.doc-label {

  width: 25%;

  font-weight: normal;

}


.doc-colon {

  width: 3%;

  text-align: center;

}


.doc-value {

  width: 72%;

}


/* =======================================================
   CUSTOMER
   ======================================================= */

.customer-table {

  width: 100%;

  border-collapse: collapse;

  margin-top: 5px;

  margin-bottom: 8px;

}


.customer-table td {

  vertical-align: top;

}


.customer-box {

  width: 50%;

  padding-right: 14px;

}


.customer-title {

  font-weight: bold;

  margin-bottom: 3px;

}


.customer-name {

  font-weight: bold;

  margin-bottom: 2px;

}


.customer-detail {

  line-height: 1.2;

}


/* =======================================================
   SHIPPING
   ======================================================= */

.shipping {

  font-weight: bold;

  margin-top: 5px;

  margin-bottom: 10px;

}


/* =======================================================
   ITEMS TABLE
   ======================================================= */

.items-table {

  width: 100%;

  border-collapse: collapse;

  table-layout: fixed;

  font-size: 8pt;

}

/* =======================================================
   HEADER ROW — FULL GREEN BLOCK
   ======================================================= */

.items-table {
  border-collapse: collapse;
  -webkit-print-color-adjust: exact !important;
  print-color-adjust: exact !important;
}

.items-table thead {
  background-color: #355E3B !important;
  -webkit-print-color-adjust: exact !important;
  print-color-adjust: exact !important;
}

.items-table thead tr {
  background-color: #355E3B !important;
  -webkit-print-color-adjust: exact !important;
  print-color-adjust: exact !important;
}

.items-table thead th {
  background-color: #355E3B !important;
  background: #355E3B !important;
  color: #FFFFFF !important;
  font-weight: bold !important;
  text-align: center;
  padding: 5px 3px;
  vertical-align: middle;
  border: 1px solid #355E3B !important;

  -webkit-print-color-adjust: exact !important;
  print-color-adjust: exact !important;
}


/* =======================================================
   BODY TABLE
   ======================================================= */

.items-table tbody td {

  background-color: #FFFFFF;

  color: #000000;

  padding: 5px 3px;

  border: 1px solid #B7B7B7;

  vertical-align: top;

  line-height: 1.15;

}


/* =======================================================
   COLUMN WIDTHS
   ======================================================= */

.items-table th:nth-child(1),
.items-table td:nth-child(1) {

  width: 10%;

}


.items-table th:nth-child(2),
.items-table td:nth-child(2) {

  width: 28%;

}


.items-table th:nth-child(3),
.items-table td:nth-child(3) {

  width: 12%;

}


.items-table th:nth-child(4),
.items-table td:nth-child(4) {

  width: 11%;

}


.items-table th:nth-child(5),
.items-table td:nth-child(5) {

  width: 8%;

}


.items-table th:nth-child(6),
.items-table td:nth-child(6) {

  width: 6%;

}


.items-table th:nth-child(7),
.items-table td:nth-child(7) {

  width: 6%;

}


.items-table th:nth-child(8),
.items-table td:nth-child(8) {

  width: 10%;

}


.items-table th:nth-child(9),
.items-table td:nth-child(9) {

  width: 11%;

}


.item-code {

  font-weight: bold;

}


.item-desc {

  line-height: 1.15;

}


.center {

  text-align: center;

}


.money {

  text-align: right;

  white-space: nowrap;

}


/* =======================================================
   SUBTOTAL
   ======================================================= */

.subtotal-table {

  width: 100%;

  border-collapse: collapse;

  table-layout: fixed;

  font-size: 8.5pt;

}


.subtotal-table td {

  padding: 5px 3px;

  border: 1px solid #B7B7B7;

  vertical-align: middle;

  font-weight: bold;

}


.subtotal-label {

  text-align: left;

}


.subtotal-qty {

  text-align: center;

}


.subtotal-value {

  text-align: right;

  white-space: nowrap;

}


/* =======================================================
   SUMMARY
   ======================================================= */

.summary-table {

  width: 45%;

  margin-left: auto;

  border-collapse: collapse;

  margin-top: 3px;

  margin-bottom: 10px;

  font-size: 8.5pt;

}


.summary-table td {

  padding: 3px 4px;

}


.summary-label {

  width: 65%;

}


.summary-colon {

  width: 5%;

  text-align: center;

}


.summary-value {

  width: 30%;

  text-align: right;

}


.total-row {

  color: ${BRAND_COLOR};

  font-weight: bold;

  font-size: 10pt;

}


/* =======================================================
   DISCLAIMER
   ======================================================= */

.disclaimer {

  margin-top: 10px;

  font-size: 9pt;

  line-height: 1.25;

}


.disclaimer p {

  margin: 0 0 8px 0;

}


.disclaimer strong {

  font-weight: bold;

}


/* =======================================================
   SIGNATURE
   ======================================================= */

.signature-table {

  width: 100%;

  border-collapse: collapse;

  margin-top: 12px;

  font-size: 9pt;

}


.signature-table td {

  width: 33.33%;

  vertical-align: top;

  text-align: center;

  padding: 0 10px;

}


.signature-space {

  height: 65px;

}


.signature-company {

  font-weight: bold;

  text-align: right;

}


.approved-signature {

  text-align: right !important;

}


.approved-signature .signature-company {

  text-align: right;

}


/* =======================================================
   PRINT
   ======================================================= */

tr {

  page-break-inside: avoid;

}

</style>

</head>


<body>


<!-- =====================================================
     HEADER
     ===================================================== -->

<table class="header-table">

<tr>

<td class="header-left">

  <div class="company-name">
    PT Setiap Hari Dipakai
  </div>

  <div class="company-address">
    Jl. Ir H Juanda No.95, Lb. Siliwangi,
    Kecamatan Coblong, Kota Bandung,
    Jawa Barat 40132
  </div>

</td>


<td class="header-right">

  <div class="document-title">
    DELIVERY NOTE
  </div>


  <table class="doc-info">

    <tr>
      <td class="doc-label">Number</td>
      <td class="doc-colon">:</td>
      <td class="doc-value">${escapeHtml(number)}</td>
    </tr>

    <tr>
      <td class="doc-label">Date</td>
      <td class="doc-colon">:</td>
      <td class="doc-value">${escapeHtml(date)}</td>
    </tr>

    <tr>
      <td class="doc-label">PO Type</td>
      <td class="doc-colon">:</td>
      <td class="doc-value">PO Inventory Everpro</td>
    </tr>

    <tr>
      <td class="doc-label">Currency</td>
      <td class="doc-colon">:</td>
      <td class="doc-value">IDR - 1.00</td>
    </tr>

    <tr>
      <td class="doc-label">Department</td>
      <td class="doc-colon">:</td>
      <td class="doc-value">Managed Services</td>
    </tr>

    <tr>
      <td class="doc-label">Class</td>
      <td class="doc-colon">:</td>
      <td class="doc-value">Everpro MP</td>
    </tr>

    <tr>
      <td class="doc-label">Prepayment #</td>
      <td class="doc-colon">:</td>
      <td class="doc-value"></td>
    </tr>

    <tr>
      <td class="doc-label">Terms</td>
      <td class="doc-colon">:</td>
      <td class="doc-value">${escapeHtml(terms)}</td>
    </tr>

    <tr>
      <td class="doc-label">Due Date</td>
      <td class="doc-colon">:</td>
      <td class="doc-value">${escapeHtml(dueDate)}</td>
    </tr>

  </table>

</td>

</tr>

</table>


<!-- =====================================================
     SOLD TO / SHIP TO
     ===================================================== -->

<table class="customer-table">

<tr>

<td class="customer-box">

  <div class="customer-title">
    Sold to:
  </div>

  <div class="customer-name">
    ${escapeHtml(clientName)}
  </div>

  <div class="customer-detail">
    ${escapeHtml(clientAddress)}
  </div>

  <div class="customer-detail">
    ${escapeHtml(clientPIC)}
  </div>

</td>


<td class="customer-box">

  <div class="customer-title">
    Ship to:
  </div>

  <div class="customer-name">
    ${escapeHtml(clientName)}
  </div>

  <div class="customer-detail">
    ${escapeHtml(clientAddress)}
  </div>

  <div class="customer-detail">
    ${escapeHtml(clientPIC)}
  </div>

</td>

</tr>

</table>


<!-- =====================================================
     SHIPPING
     ===================================================== -->

<div class="shipping">

  SHIPPING TYPE DIAMBIL CUSTOMER

</div>


<!-- =====================================================
     ITEMS
     ===================================================== -->

<table class="items-table">

<thead>

<tr style="
  background-color: #355E3B !important;
  background: #355E3B !important;
  -webkit-print-color-adjust: exact !important;
  print-color-adjust: exact !important;
">
  <th style="
    background-color: #355E3B !important;
    background: #355E3B !important;
    color: #FFFFFF !important;
    font-weight: bold !important;
  ">Item Code</th>

  <th style="
    background-color: #355E3B !important;
    background: #355E3B !important;
    color: #FFFFFF !important;
    font-weight: bold !important;
  ">Item Descriptions</th>

  <th style="
    background-color: #355E3B !important;
    background: #355E3B !important;
    color: #FFFFFF !important;
    font-weight: bold !important;
  ">Department</th>

  <th style="
    background-color: #355E3B !important;
    background: #355E3B !important;
    color: #FFFFFF !important;
    font-weight: bold !important;
  ">Class</th>

  <th style="
    background-color: #355E3B !important;
    background: #355E3B !important;
    color: #FFFFFF !important;
    font-weight: bold !important;
  ">Location</th>

  <th style="
    background-color: #355E3B !important;
    background: #355E3B !important;
    color: #FFFFFF !important;
    font-weight: bold !important;
  ">Qty</th>

  <th style="
    background-color: #355E3B !important;
    background: #355E3B !important;
    color: #FFFFFF !important;
    font-weight: bold !important;
  ">Units</th>

  <th style="
    background-color: #355E3B !important;
    background: #355E3B !important;
    color: #FFFFFF !important;
    font-weight: bold !important;
  ">Unit Price</th>

  <th style="
    background-color: #355E3B !important;
    background: #355E3B !important;
    color: #FFFFFF !important;
    font-weight: bold !important;
  ">Total Price</th>
</tr>

</thead>


<tbody>

${itemRows}

</tbody>

</table>


<!-- =====================================================
     SUBTOTAL
     ===================================================== -->

<table class="subtotal-table">

<colgroup>

  <col style="width:10%;">
  <col style="width:28%;">
  <col style="width:12%;">
  <col style="width:11%;">
  <col style="width:8%;">
  <col style="width:6%;">
  <col style="width:6%;">
  <col style="width:10%;">
  <col style="width:11%;">

</colgroup>


<tr>

<td class="subtotal-label" colspan="5">
  Subtotal
</td>

<td class="subtotal-qty">
  ${formatNumber(totalQty)}
</td>

<td></td>

<td></td>

<td class="subtotal-value">
  ${formatPrice(subtotal)}
</td>

</tr>

</table>


<!-- =====================================================
     SUMMARY
     ===================================================== -->

<table class="summary-table">

<tr>
<td class="summary-label">Discount</td>
<td class="summary-colon">:</td>
<td class="summary-value">${formatPrice(discount)}</td>
</tr>

<tr>
<td class="summary-label">After Discount</td>
<td class="summary-colon">:</td>
<td class="summary-value">${formatPrice(afterDiscount)}</td>
</tr>

<tr>
<td class="summary-label">VAT</td>
<td class="summary-colon">:</td>
<td class="summary-value">${formatPrice(vat)}</td>
</tr>

<tr>
<td class="summary-label">WHT</td>
<td class="summary-colon">:</td>
<td class="summary-value">${formatPrice(wht)}</td>
</tr>

<tr>
<td class="summary-label">Prepayment</td>
<td class="summary-colon">:</td>
<td class="summary-value">${formatPrice(prepayment)}</td>
</tr>

<tr class="total-row">
<td class="summary-label">TOTAL</td>
<td class="summary-colon">:</td>
<td class="summary-value">${formatPrice(grandTotal)}</td>
</tr>

</table>


<!-- =====================================================
     DISCLAIMER
     ===================================================== -->

<div class="disclaimer">

<p>

<strong>
"Delivery Note / Faktur Penyerahan ini adalah bukti yang sah dan sempurna mengenai telah diterimanya Barang secara utuh dan lengkap serta sesuai dengan dokumen pemesanan pembelian Barang yang telah kami terima."
</strong>

</p>


<p>

<strong>
"This Delivery Note is legal and valid evidence for the entireness and completeness of the delivery of the Good pursuant to the purchase order document which has been provided to us."
</strong>

</p>


<p>

<strong>
"HARAP DIBUBUHKAN TANDA TANGAN, NAMA JELAS, TANGGAL TERIMA, DAN CAP PERUSAHAAN"
</strong>

</p>

</div>


<!-- =====================================================
     SIGNATURE
     ===================================================== -->

<table class="signature-table">

<tr>

<td>

<strong>
  Received,
</strong>

<div class="signature-space"></div>

</td>


<td>

<strong>
  Transported,
</strong>

<div class="signature-space"></div>

</td>


<td class="approved-signature">

<strong>
  Approved,
</strong>

<div class="signature-space"></div>

<div class="signature-company">
  PT SETIAP HARI DIPAKAI
</div>

</td>

</tr>

</table>


</body>

</html>

`;


  // =======================================================
  // CREATE PDF
  // =======================================================

  const htmlBlob =
    Utilities.newBlob(
      html,
      MimeType.HTML,
      `${PDF_PREFIX}.html`
    );


  const pdfBlob =
    htmlBlob.getAs(
      MimeType.PDF
    );


  const fileName =
    `${PDF_PREFIX}${poNumber}.pdf`;


  pdfBlob.setName(
    fileName
  );


  // =======================================================
  // GET / CREATE FOLDER
  // =======================================================

  const folders =
    DriveApp.getFoldersByName(
      FOLDER_NAME
    );


  let folder;


  if (folders.hasNext()) {

    folder =
      folders.next();

  } else {

    folder =
      DriveApp.createFolder(
        FOLDER_NAME
      );

  }


  // =======================================================
  // CREATE FILE
  // =======================================================

  const pdfFile =
    folder.createFile(
      pdfBlob
    );


  // =======================================================
  // WRITE PDF LINK
  // =======================================================

  const finalHeaders =
    sourceSheet
      .getRange(
        1,
        1,
        1,
        sourceSheet.getLastColumn()
      )
      .getDisplayValues()[0];


  const pdfCol =
    finalHeaders.findIndex(
      h =>
        String(h)
          .trim()
          .toUpperCase() ===
        PDF_LINK_HEADER
    ) + 1;


  if (pdfCol > 0) {

    sourceSheet
      .getRange(
        selectedRow,
        pdfCol
      )
      .setFormula(
        `=HYPERLINK("${pdfFile.getUrl()}","📄 Open PDF")`
      );

  }


  // =======================================================
  // RESET CHECKBOX
  // =======================================================

  const actionCol =
    finalHeaders.findIndex(
      h =>
        String(h)
          .trim()
          .toUpperCase() ===
        ACTION_HEADER
    ) + 1;


  if (actionCol > 0) {

    sourceSheet
      .getRange(
        selectedRow,
        actionCol
      )
      .setValue(false);

  }


  // =======================================================
  // RESULT
  // =======================================================

  ss.toast(
    `PDF ${poNumber} berhasil dibuat!`,
    "Delivery Note",
    5
  );


  return pdfFile.getUrl();

}
