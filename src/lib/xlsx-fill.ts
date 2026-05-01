import ExcelJS from "exceljs";
import { saveAs } from "file-saver";
import { CZECH_MONTHS, deburr } from "./dates";

export interface DayRow {
  day: number;
  date: string;
  weekday: number;
  arrivalMin: number;
  departureMin: number;
  lunchStartMin: number;
  lunchEndMin: number;
  code: string;
  isWeekend: boolean;
  /** Only relevant for code === "S": user manually filled in times. */
  userEdited?: boolean;
}

export interface FillInput {
  jmeno: string;
  uvazek: string;
  praciste: string;
  month: number; // 1..12
  year: number;
  rows: DayRow[];
}

// ─── Colours ────────────────────────────────────────────────────────────
const CLR_WEEKEND = { type: "pattern", pattern: "solid", fgColor: { argb: "FFD9D9D9" } } as const;
const CLR_HOLIDAY = { type: "pattern", pattern: "solid", fgColor: { argb: "FFFFCCCC" } } as const;
const CLR_VACATION = { type: "pattern", pattern: "solid", fgColor: { argb: "FFD6E4F7" } } as const;

// ─── Fonts ──────────────────────────────────────────────────────────────
const FONT_NORMAL = { name: "Calibri", size: 10 };
const FONT_BOLD = { name: "Calibri", size: 10, bold: true };

// ─── Border helpers ─────────────────────────────────────────────────────
const BORDER_THIN = { style: "thin" as const, color: { argb: "FF000000" } };
const ALL_BORDERS = { top: BORDER_THIN, bottom: BORDER_THIN, left: BORDER_THIN, right: BORDER_THIN };

// ─── Number formats ─────────────────────────────────────────────────────
const FMT_TIME = "hh:mm";
const FMT_HOURS = "[h]:mm";

export async function generateXlsx(input: FillInput): Promise<void> {
  const wb = new ExcelJS.Workbook();
  wb.creator = "Docházka generátor";
  wb.created = new Date();

  const ws = wb.addWorksheet(`${CZECH_MONTHS[input.month - 1]} ${input.year}`, {
    pageSetup: { fitToPage: true, fitToWidth: 1, orientation: "portrait" },
  });

  ws.columns = [
    { key: "A", width: 0 },
    { key: "B", width: 5 },
    { key: "C", width: 5 },
    { key: "D", width: 8 },
    { key: "E", width: 8 },
    { key: "F", width: 8 },
    { key: "G", width: 8 },
    { key: "H", width: 8 },
    { key: "I", width: 8 },
    { key: "J", width: 8 },
    { key: "K", width: 8 },
    { key: "L", width: 8 },
    { key: "M", width: 10 },
  ];
  ws.getColumn("A").hidden = true;

  // Row 1
  ws.mergeCells("C1:M1");
  ws.getCell("C1").value = "JINTEK, z.ú., Křížkovského 554/12, 603 00 Brno";
  ws.getCell("C1").font = FONT_BOLD;
  ws.getRow(1).height = 16;

  // Row 2
  ws.mergeCells("C2:M2");
  ws.getCell("C2").value = "Evidence docházky";
  ws.getCell("C2").font = FONT_BOLD;
  ws.getRow(2).height = 14;

  // Row 3
  ws.getRow(3).height = 14;
  ws.getCell("C3").value = "Jméno:";
  ws.getCell("C3").font = FONT_BOLD;
  ws.getCell("D3").value = input.jmeno;
  ws.mergeCells("D3:E3");
  ws.getCell("F3").value = `úvazek ${input.uvazek}`;
  ws.mergeCells("F3:G3");
  ws.getCell("H3").value = "Měsíc:";
  ws.getCell("H3").font = FONT_BOLD;
  ws.getCell("I3").value = CZECH_MONTHS[input.month - 1];
  ws.getCell("J3").value = "Rok:";
  ws.getCell("J3").font = FONT_BOLD;
  ws.getCell("K3").value = input.year;
  ws.getCell("M3").value = input.praciste;

  // Row 4 — column headers
  const HEADER_ROW = 4;
  ws.getRow(HEADER_ROW).height = 16;
  const headers: [string, string][] = [
    ["C", "Datum"],
    ["D", "Příchod"],
    ["E", "Odchod"],
    ["F", "Jedn. od"],
    ["G", "Jedn. do"],
    ["H", "Oběd od"],
    ["I", "Oběd do"],
    ["J", "Důvod"],
    ["K", "Ost. od"],
    ["L", "Ost. do"],
    ["M", "Hod."],
  ];
  for (const [col, label] of headers) {
    const cell = ws.getCell(`${col}${HEADER_ROW}`);
    cell.value = label;
    cell.font = FONT_BOLD;
    cell.border = ALL_BORDERS;
    cell.alignment = { horizontal: "center", wrapText: true };
  }

  // Day rows
  const FIRST_DATA_ROW = 5;
  const WEEKDAYS_SHORT = ["ne", "po", "út", "st", "čt", "pá", "so"];

  for (const d of input.rows) {
    const rowNum = FIRST_DATA_ROW + d.day - 1;
    const row = ws.getRow(rowNum);
    row.height = 15;

    let fill: typeof CLR_WEEKEND | typeof CLR_HOLIDAY | typeof CLR_VACATION | undefined;
    if (d.isWeekend) fill = CLR_WEEKEND;
    else if (d.code === "S") fill = CLR_HOLIDAY;
    else if (d.code === "D") fill = CLR_VACATION;

    function applyCell(col: string, value: ExcelJS.CellValue, numFmt?: string) {
      const cell = ws.getCell(`${col}${rowNum}`);
      cell.value = value;
      cell.font = FONT_NORMAL;
      cell.border = ALL_BORDERS;
      if (fill) cell.fill = fill;
      if (numFmt) cell.numFmt = numFmt;
      cell.alignment = { horizontal: "center" };
    }

    applyCell("B", WEEKDAYS_SHORT[d.weekday]);
    applyCell("C", d.day);

    if (d.isWeekend) {
      for (const c of ["D", "E", "F", "G", "H", "I", "J", "K", "L", "M"]) applyCell(c, null);
      continue;
    }

    if (d.code === "S") {
      applyCell("D", null, FMT_TIME);
      applyCell("E", null, FMT_TIME);
      applyCell("H", null, FMT_TIME);
      applyCell("I", null, FMT_TIME);
      applyCell("F", null);
      applyCell("G", null);
      applyCell("K", null);
      applyCell("L", null);
      applyCell("J", d.code);
      applyCell("M", null, FMT_HOURS);
      continue;
    }

    if (d.code) {
      applyCell("D", d.arrivalMin / 1440, FMT_TIME);
      applyCell("E", d.departureMin / 1440, FMT_TIME);
      applyCell("H", d.lunchStartMin / 1440, FMT_TIME);
      applyCell("I", d.lunchEndMin / 1440, FMT_TIME);
      applyCell("F", null);
      applyCell("G", null);
      applyCell("K", null);
      applyCell("L", null);
      applyCell("J", d.code);
      applyCell("M", null, FMT_HOURS);
      continue;
    }

    applyCell("D", d.arrivalMin / 1440, FMT_TIME);
    applyCell("E", d.departureMin / 1440, FMT_TIME);
    applyCell("H", d.lunchStartMin / 1440, FMT_TIME);
    applyCell("I", d.lunchEndMin / 1440, FMT_TIME);
    applyCell("F", null);
    applyCell("G", null);
    applyCell("K", null);
    applyCell("L", null);
    applyCell("J", null);
    applyCell(
      "M",
      { formula: `(E${rowNum}-D${rowNum})-(I${rowNum}-H${rowNum})+(L${rowNum}-K${rowNum})` },
      FMT_HOURS,
    );
  }

  const lastDataRow = FIRST_DATA_ROW + input.rows.length - 1;

  // Summary block
  const summaryData: [string, string][] = [
    ["odpracovaná doba:", `SUM(M${FIRST_DATA_ROW}:M${lastDataRow})`],
    ["svátek:", "0"],
    ["dovolená:", "0"],
    ["sick day:", "0"],
    ["DPN:", "0"],
    ["OČR:", "0"],
    ["Placené volno:", "0"],
    ["Celkem:", `SUM(M${lastDataRow + 1}:M${lastDataRow + 7})`],
  ];

  let sRow = lastDataRow + 2;
  for (const [label, formula] of summaryData) {
    ws.getCell(`K${sRow}`).value = label;
    ws.getCell(`K${sRow}`).font = FONT_BOLD;
    const mCell = ws.getCell(`M${sRow}`);
    mCell.value = { formula };
    mCell.numFmt = FMT_HOURS;
    mCell.border = ALL_BORDERS;
    sRow++;
  }

  // Legend
  const legendRow = sRow + 1;
  ws.getCell(`C${legendRow}`).value =
    "D = dovolená  |  S = svátek  |  SD = sick day  |  DPN = nemoc  |  OČR = ošetřovné  |  PV = placené volno  |  HO = home office  |  PC = pracovní cesta  |  ŠK = školení  |  SO = soukromé";
  ws.getCell(`C${legendRow}`).font = { name: "Calibri", size: 8, italic: true };
  ws.mergeCells(`C${legendRow}:M${legendRow}`);

  const buf = await wb.xlsx.writeBuffer();
  const safeName = deburr(input.jmeno).replace(/\s+/g, "");
  const mm = String(input.month).padStart(2, "0");
  saveAs(
    new Blob([buf], {
      type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    }),
    `Dochazka_${safeName || "uzivatel"}_${input.year}-${mm}.xlsx`,
  );
}
