import ExcelJS from "exceljs";

const COLORS = {
  navy: "17324D",
  blue: "2F75B5",
  paleBlue: "DCE6F1",
  stripeBlue: "D9EAF3",
  darkText: "1F2937",
  grayBorder: "E2E8F0",
  warningFill: "FFF2CC",
  warningText: "7F6000",
  white: "FFFFFF",
};

function cellText(value) {
  if (value === null || value === undefined) return "";
  if (value instanceof Date) return value.toISOString();
  if (typeof value === "object") {
    if (value.text !== undefined) return String(value.text).trim();
    if (value.hyperlink !== undefined) return String(value.hyperlink).trim();
    if (Array.isArray(value.richText)) {
      return value.richText.map((part) => part.text || "").join("").trim();
    }
    if (value.result !== undefined) return String(value.result).trim();
  }
  return String(value).trim();
}

function hostnameFromUrl(value) {
  try {
    return new URL(value).hostname.toLowerCase().replace(/^www\./, "");
  } catch {
    return "";
  }
}

export async function readPublicationSourcesStandalone(workbookPath) {
  const workbook = new ExcelJS.Workbook();
  await workbook.xlsx.readFile(workbookPath);
  const sheet = workbook.worksheets[0];
  if (!sheet || sheet.rowCount < 2) {
    throw new Error("The input workbook does not contain publication rows.");
  }

  const sources = [];
  for (let rowNumber = 2; rowNumber <= sheet.rowCount; rowNumber += 1) {
    const name = cellText(sheet.getCell(rowNumber, 1).value);
    const sourceUrl = cellText(sheet.getCell(rowNumber, 2).value);
    if (!name || !/^https?:\/\//i.test(sourceUrl)) continue;
    const hostname = hostnameFromUrl(sourceUrl);
    if (!hostname) continue;
    sources.push({
      sourceRow: rowNumber,
      name,
      sourceUrl,
      hostname,
      region: cellText(sheet.getCell(rowNumber, 5).value),
      county: cellText(sheet.getCell(rowNumber, 6).value),
      city: cellText(sheet.getCell(rowNumber, 7).value),
    });
  }

  const grouped = new Map();
  for (const source of sources) {
    if (!grouped.has(source.hostname)) {
      grouped.set(source.hostname, {
        ...source,
        aliases: [],
        regions: [],
        counties: [],
        cities: [],
        sourceRows: [],
      });
    }
    const record = grouped.get(source.hostname);
    record.aliases.push(source.name);
    if (source.region) record.regions.push(source.region);
    if (source.county) record.counties.push(source.county);
    if (source.city) record.cities.push(source.city);
    record.sourceRows.push(source.sourceRow);
  }

  return [...grouped.values()].map((source) => ({
    ...source,
    aliases: [...new Set(source.aliases)],
    regions: [...new Set(source.regions)],
    counties: [...new Set(source.counties)],
    cities: [...new Set(source.cities)],
  }));
}

function solidFill(argb) {
  return { type: "pattern", pattern: "solid", fgColor: { argb } };
}

function thinBottomBorder() {
  return { bottom: { style: "thin", color: { argb: COLORS.grayBorder } } };
}

function setTitle(sheet, endColumn, title, subtitle) {
  sheet.mergeCells(`A1:${endColumn}1`);
  sheet.mergeCells(`A2:${endColumn}2`);
  const titleCell = sheet.getCell("A1");
  titleCell.value = title;
  titleCell.fill = solidFill(COLORS.navy);
  titleCell.font = { bold: true, color: { argb: COLORS.white }, size: 18 };
  titleCell.alignment = { vertical: "middle", horizontal: "left" };
  sheet.getRow(1).height = 32;

  const subtitleCell = sheet.getCell("A2");
  subtitleCell.value = subtitle;
  subtitleCell.fill = solidFill(COLORS.paleBlue);
  subtitleCell.font = { italic: true, color: { argb: "334155" }, size: 10 };
  subtitleCell.alignment = { vertical: "middle", horizontal: "left", wrapText: true };
  sheet.getRow(2).height = 26;
}

function styleHeaderRow(row) {
  row.height = 32;
  row.eachCell({ includeEmpty: true }, (cell) => {
    cell.fill = solidFill(COLORS.blue);
    cell.font = { bold: true, color: { argb: COLORS.white }, size: 10 };
    cell.alignment = { vertical: "middle", horizontal: "left", wrapText: true };
    cell.border = {
      bottom: { style: "medium", color: { argb: COLORS.navy } },
    };
  });
}

function styleDataRows(sheet, startRow, endRow, columnCount) {
  if (endRow < startRow) return;
  for (let rowNumber = startRow; rowNumber <= endRow; rowNumber += 1) {
    const row = sheet.getRow(rowNumber);
    row.height = 42;
    for (let column = 1; column <= columnCount; column += 1) {
      const cell = row.getCell(column);
      if ((rowNumber - startRow) % 2 === 1) {
        cell.fill = solidFill(COLORS.stripeBlue);
      }
      cell.font = { color: { argb: COLORS.darkText }, size: 10 };
      cell.alignment = { vertical: "top", horizontal: "left", wrapText: true };
      cell.border = thinBottomBorder();
    }
  }
}

function setColumns(sheet, widths) {
  widths.forEach((width, index) => {
    sheet.getColumn(index + 1).width = width;
  });
}

function setHyperlink(cell, url) {
  if (!/^https?:\/\//i.test(String(url || ""))) return;
  cell.value = { text: String(url), hyperlink: String(url) };
  cell.font = { color: { argb: "0563C1" }, underline: true, size: 10 };
}

function excelDate(value) {
  if (!/^\d{4}-\d{2}-\d{2}$/.test(value || "")) return null;
  return new Date(`${value}T12:00:00Z`);
}

function configureSheet(sheet) {
  sheet.views = [{ state: "frozen", ySplit: 3, showGridLines: false }];
  sheet.pageSetup = {
    orientation: "landscape",
    fitToPage: true,
    fitToWidth: 1,
    fitToHeight: 0,
    margins: {
      left: 0.25,
      right: 0.25,
      top: 0.5,
      bottom: 0.5,
      header: 0.2,
      footer: 0.2,
    },
  };
}

export async function buildWorkbookStandalone({
  window,
  config,
  sources,
  developments,
  audits,
  pdfIndex,
  runStarted,
  dryRun,
  researchMode,
}) {
  const workbook = new ExcelJS.Workbook();
  workbook.creator = "Publication Development Research Automation";
  workbook.created = runStarted;
  workbook.modified = new Date();
  workbook.calcProperties.fullCalcOnLoad = true;

  const summary = workbook.addWorksheet("Summary", {
    views: [{ showGridLines: false }],
  });
  const developmentSheet = workbook.addWorksheet("Developments");
  const auditSheet = workbook.addWorksheet("Source Audit");
  const pdfSheet = workbook.addWorksheet("PDF Index");
  [developmentSheet, auditSheet, pdfSheet].forEach(configureSheet);

  setTitle(
    summary,
    "H",
    "Development Research Report",
    `Generated ${runStarted.toISOString()} | Model ${config.model} | ${researchMode}${dryRun ? " DRY RUN" : ""}`,
  );
  const metrics = [
    ["Metric", "Value"],
    ["Distinct websites reviewed", sources.length],
    ["Qualifying developments", developments.length],
    ["Plan PDF links found", pdfIndex.length],
    ["Plan PDFs downloaded", pdfIndex.filter((item) => item.status === "Downloaded").length],
    ["Sites with no qualifying results", audits.filter((item) => item.status === "no_results").length],
    ["Sites blocked or errored", audits.filter((item) => ["blocked", "error"].includes(item.status)).length],
  ];
  metrics.forEach((values, index) => {
    const row = 4 + index;
    summary.getCell(row, 1).value = values[0];
    summary.getCell(row, 2).value = values[1];
  });
  styleHeaderRow(summary.getRow(4));
  for (let row = 5; row <= 10; row += 1) {
    for (let col = 1; col <= 2; col += 1) {
      summary.getCell(row, col).border = thinBottomBorder();
      summary.getCell(row, col).alignment = { vertical: "top", wrapText: true };
    }
  }
  summary.mergeCells("D4:H4");
  summary.getCell("D4").value = "Scope and rules";
  summary.getCell("D4").fill = solidFill(COLORS.blue);
  summary.getCell("D4").font = { bold: true, color: { argb: COLORS.white } };
  summary.getCell("D4").alignment = { vertical: "middle" };
  const rules = [
    ["Geography", config.researchScope],
    ["Included", "Demolitions; ground-up retail/restaurant buildings; retail centers; residential and mixed-use development."],
    ["Excluded", "Ordinary tenant openings, leases, remodels, office-only, industrial-only, hotel-only, medical-only, and civic-only work."],
    ["Address rule", "Only source-supported addresses and cross streets are reported; unknown values remain blank."],
    ["Security", "Publication usernames and passwords are not used or copied into this report."],
  ];
  rules.forEach((rule, index) => {
    const row = 5 + index;
    summary.getCell(row, 4).value = rule[0];
    summary.mergeCells(row, 5, row, 8);
    summary.getCell(row, 5).value = rule[1];
    for (let col = 4; col <= 8; col += 1) {
      summary.getCell(row, col).alignment = { vertical: "top", wrapText: true };
      summary.getCell(row, col).border = thinBottomBorder();
    }
    summary.getRow(row).height = 40;
  });
  summary.mergeCells("A12:H12");
  summary.getCell("A12").value =
    "Review Developments for results, Source Audit for complete site coverage, and PDF Index for downloaded plan files.";
  summary.getCell("A12").fill = solidFill(COLORS.warningFill);
  summary.getCell("A12").font = { bold: true, color: { argb: COLORS.warningText } };
  summary.getCell("A12").alignment = { wrapText: true, vertical: "middle" };
  summary.getRow(12).height = 30;
  setColumns(summary, [30, 18, 4, 20, 22, 22, 22, 22]);

  const developmentHeaders = [
    "Publication Date", "Publication", "Article Title", "Development Type",
    "Project", "Status", "Description", "City", "State", "County",
    "Street Address", "Cross Street", "Units", "Retail Sq Ft",
    "Restaurant Sq Ft", "Demolition Details", "Developer", "Plan PDFs Found",
    "Confidence", "Evidence Notes", "Article URL", "Publication URL",
  ];
  setTitle(
    developmentSheet,
    "V",
    "Qualifying Developments",
    `${developments.length} distinct articles found`,
  );
  developmentSheet.addRow([]);
  developmentSheet.getRow(3).values = developmentHeaders;
  styleHeaderRow(developmentSheet.getRow(3));
  for (const item of developments) {
    const row = developmentSheet.addRow([
      excelDate(item.publication_date),
      item.publication_name,
      item.article_title,
      item.development_type,
      item.project_name,
      item.project_status,
      item.description,
      item.city,
      item.state,
      item.county,
      item.address,
      item.cross_street,
      item.units,
      item.retail_sq_ft,
      item.restaurant_sq_ft,
      item.demolition_details,
      item.developer,
      (item.plan_pdf_urls || []).length,
      item.confidence,
      item.evidence_notes,
      item.article_url,
      item.publication_url,
    ]);
    row.getCell(1).numFmt = "yyyy-mm-dd";
    for (let column = 13; column <= 18; column += 1) {
      row.getCell(column).numFmt = "#,##0";
    }
    setHyperlink(row.getCell(21), item.article_url);
    setHyperlink(row.getCell(22), item.publication_url);
  }
  const developmentLastRow = Math.max(3, developmentSheet.rowCount);
  styleDataRows(developmentSheet, 4, developmentLastRow, 22);
  developmentSheet.autoFilter = `A3:V${developmentLastRow}`;
  setColumns(developmentSheet, [
    13, 22, 42, 22, 25, 18, 55, 18, 10, 18, 28,
    24, 10, 14, 16, 42, 24, 12, 12, 42, 55, 35,
  ]);

  const auditHeaders = [
    "Source Row(s)", "Publication Name(s)", "Website", "Domain", "Region",
    "County", "City", "Review Status", "Qualifying Items", "Notes",
    "API Response ID",
  ];
  setTitle(
    auditSheet,
    "K",
    "Source Audit",
    "Every distinct website from the input workbook is represented.",
  );
  auditSheet.addRow([]);
  auditSheet.getRow(3).values = auditHeaders;
  styleHeaderRow(auditSheet.getRow(3));
  for (const item of audits) {
    const row = auditSheet.addRow([
      item.sourceRows.join(", "),
      item.aliases.join("; "),
      item.sourceUrl,
      item.hostname,
      item.regions.join("; "),
      item.counties.join("; "),
      item.cities.join("; ") || null,
      item.status,
      item.qualifyingItems,
      item.notes,
      item.responseId || null,
    ]);
    setHyperlink(row.getCell(3), item.sourceUrl);
    const status = String(item.status || "").toLowerCase();
    if (status === "blocked" || status === "error") {
      row.getCell(8).fill = solidFill("F4CCCC");
      row.getCell(8).font = { bold: true, color: { argb: "9C0006" } };
    } else if (status === "reviewed") {
      row.getCell(8).fill = solidFill("D9EAD3");
      row.getCell(8).font = { bold: true, color: { argb: "274E13" } };
    }
  }
  const auditLastRow = Math.max(3, auditSheet.rowCount);
  styleDataRows(auditSheet, 4, auditLastRow, 11);
  auditSheet.autoFilter = `A3:K${auditLastRow}`;
  setColumns(auditSheet, [14, 30, 45, 25, 20, 18, 18, 16, 14, 55, 30]);

  const pdfHeaders = [
    "Publication Date", "City", "Address", "Project", "Filename",
    "Download Status", "Size Bytes", "Notes", "PDF URL", "Source Article",
  ];
  setTitle(
    pdfSheet,
    "J",
    "Plan PDF Index",
    "Files are saved in the PDF subfolder using City - Address - YYYY-MM-DD.",
  );
  pdfSheet.addRow([]);
  pdfSheet.getRow(3).values = pdfHeaders;
  styleHeaderRow(pdfSheet.getRow(3));
  for (const item of pdfIndex) {
    const row = pdfSheet.addRow([
      excelDate(item.publication_date),
      item.city,
      item.address,
      item.project_name,
      item.filename,
      item.status,
      item.size_bytes,
      item.notes,
      item.pdf_url,
      item.source_article,
    ]);
    row.getCell(1).numFmt = "yyyy-mm-dd";
    row.getCell(7).numFmt = "#,##0";
    setHyperlink(row.getCell(9), item.pdf_url);
    setHyperlink(row.getCell(10), item.source_article);
  }
  const pdfLastRow = Math.max(3, pdfSheet.rowCount);
  styleDataRows(pdfSheet, 4, pdfLastRow, 10);
  pdfSheet.autoFilter = `A3:J${pdfLastRow}`;
  setColumns(pdfSheet, [13, 18, 30, 28, 55, 18, 14, 40, 55, 55]);

  return workbook;
}
