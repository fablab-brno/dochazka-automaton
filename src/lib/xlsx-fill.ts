import ExcelJS from "exceljs";
import { CZECH_MONTHS, deburr, lastDayOfMonth, minutesToExcelFraction } from "./dates";

export interface DayRow {
  day: number;
  date: string; // YYYY-MM-DD
  weekday: number; // 0=Sun..6=Sat
  arrivalMin: number; // minutes
  departureMin: number;
  lunchStartMin: number;
  lunchEndMin: number;
  code: string; // "" or D/S/SD/...
  isWeekend: boolean;
  /** True when the user has manually edited times on this row.
   * Relevant for code === "S": only edited holiday rows export times. */
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

export async function loadTemplate(): Promise<ExcelJS.Workbook> {
  const res = await fetch("/template/vzor_dochzka.xlsx");
  if (!res.ok) throw new Error("Šablona nenalezena.");
  const buf = await res.arrayBuffer();
  const wb = new ExcelJS.Workbook();
  await wb.xlsx.load(buf);
  return wb;
}

export function fillWorkbook(wb: ExcelJS.Workbook, input: FillInput): void {
  // Visible monthly sheet — third sheet, do NOT rename it.
  const ws = wb.worksheets[2];
  if (!ws) throw new Error("V šabloně chybí třetí list.");

  // Header
  ws.getCell("D4").value = input.jmeno;
  ws.getCell("F4").value = `úvazek ${input.uvazek}`;
  ws.getCell("I4").value = CZECH_MONTHS[input.month - 1];
  ws.getCell("K4").value = input.year;
  ws.getCell("M4").value = input.praciste;

  const last = lastDayOfMonth(input.year, input.month);

  for (const r of input.rows) {
    if (r.day < 1 || r.day > last) continue;
    const row = 7 + r.day;
    const D = `D${row}`, E = `E${row}`, H = `H${row}`, I = `I${row}`, J = `J${row}`, M = `M${row}`;

    // Weekend: leave everything empty.
    if (r.isWeekend && !r.code) {
      for (const a of [D, E, H, I, J, M]) ws.getCell(a).value = null;
      continue;
    }

    // State holiday "S": code only. If user edited times, write them; otherwise blank.
    if (r.code === "S") {
      if (r.userEdited) {
        ws.getCell(D).value = minutesToExcelFraction(r.arrivalMin);
        ws.getCell(E).value = minutesToExcelFraction(r.departureMin);
        ws.getCell(H).value = minutesToExcelFraction(r.lunchStartMin);
        ws.getCell(I).value = minutesToExcelFraction(r.lunchEndMin);
      } else {
        for (const a of [D, E, H, I]) ws.getCell(a).value = null;
      }
      ws.getCell(J).value = "S";
      ws.getCell(M).value = null;
      continue;
    }

    // Vacation/sick/etc (D, SD, DPN, OČR, PV, HO, PC, ŠK, SO):
    // keep the standard times defaults, blank M.
    if (r.code) {
      ws.getCell(D).value = minutesToExcelFraction(r.arrivalMin);
      ws.getCell(E).value = minutesToExcelFraction(r.departureMin);
      ws.getCell(H).value = minutesToExcelFraction(r.lunchStartMin);
      ws.getCell(I).value = minutesToExcelFraction(r.lunchEndMin);
      ws.getCell(J).value = r.code;
      ws.getCell(M).value = null;
      continue;
    }

    // Plain workday — leave M (formula) intact.
    ws.getCell(D).value = minutesToExcelFraction(r.arrivalMin);
    ws.getCell(E).value = minutesToExcelFraction(r.departureMin);
    ws.getCell(H).value = minutesToExcelFraction(r.lunchStartMin);
    ws.getCell(I).value = minutesToExcelFraction(r.lunchEndMin);
    ws.getCell(J).value = null;
  }
}

export async function downloadFilled(
  wb: ExcelJS.Workbook,
  jmeno: string,
  year: number,
  month: number,
): Promise<void> {
  const out = await wb.xlsx.writeBuffer();
  const blob = new Blob([out], {
    type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
  });
  const safeName = deburr(jmeno).replace(/\s+/g, "");
  const mm = String(month).padStart(2, "0");
  const fname = `Dochazka_${safeName || "uzivatel"}_${year}-${mm}.xlsx`;
  const url = URL.createObjectURL(blob);
  const a = document.createElement("a");
  a.href = url;
  a.download = fname;
  document.body.appendChild(a);
  a.click();
  document.body.removeChild(a);
  setTimeout(() => URL.revokeObjectURL(url), 1000);
}
