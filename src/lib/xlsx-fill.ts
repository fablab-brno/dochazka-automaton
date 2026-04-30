import XLSX from "xlsx-js-style";
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
}

export interface FillInput {
  jmeno: string;
  uvazek: string;
  month: number; // 1..12
  year: number;
  rows: DayRow[];
}

function setCell(ws: XLSX.WorkSheet, addr: string, value: any) {
  const existing = ws[addr] || {};
  if (typeof value === "string") {
    ws[addr] = { ...existing, t: "s", v: value, w: undefined, f: undefined };
    delete (ws[addr] as any).f;
  } else if (typeof value === "number") {
    ws[addr] = { ...existing, t: "n", v: value, w: undefined, f: undefined };
    delete (ws[addr] as any).f;
  }
}

function clearCell(ws: XLSX.WorkSheet, addr: string) {
  if (ws[addr]) {
    // Preserve style, remove value/formula
    const cell = ws[addr] as any;
    delete cell.v;
    delete cell.f;
    delete cell.w;
    delete cell.t;
  }
}

export async function loadTemplate(): Promise<XLSX.WorkBook> {
  const res = await fetch("/template/vzor_dochzka.xlsx");
  if (!res.ok) throw new Error("Šablona nenalezena.");
  const buf = await res.arrayBuffer();
  return XLSX.read(buf, { type: "array", cellStyles: true, cellFormula: true });
}

export function fillWorkbook(wb: XLSX.WorkBook, input: FillInput): void {
  // Use third sheet
  const sheetName = wb.SheetNames[2];
  if (!sheetName) throw new Error("V šabloně chybí třetí list.");
  const ws = wb.Sheets[sheetName];

  // Header
  setCell(ws, "D4", input.jmeno);
  setCell(ws, "I4", CZECH_MONTHS[input.month - 1]);
  setCell(ws, "K4", input.year);
  setCell(ws, "M4", input.uvazek);

  const last = lastDayOfMonth(input.year, input.month);

  for (const r of input.rows) {
    if (r.day < 1 || r.day > last) continue;
    const row = 7 + r.day;
    const D = `D${row}`, E = `E${row}`, H = `H${row}`, I = `I${row}`, J = `J${row}`, M = `M${row}`;

    if (r.isWeekend && !r.code) {
      // Leave untouched (template handles greying)
      continue;
    }

    if (r.code) {
      setCell(ws, D, minutesToExcelFraction(r.arrivalMin));
      setCell(ws, E, minutesToExcelFraction(r.departureMin));
      setCell(ws, H, minutesToExcelFraction(r.lunchStartMin));
      setCell(ws, I, minutesToExcelFraction(r.lunchEndMin));
      setCell(ws, J, r.code);
      clearCell(ws, M);
    } else {
      // Plain workday
      setCell(ws, D, minutesToExcelFraction(r.arrivalMin));
      setCell(ws, E, minutesToExcelFraction(r.departureMin));
      setCell(ws, H, minutesToExcelFraction(r.lunchStartMin));
      setCell(ws, I, minutesToExcelFraction(r.lunchEndMin));
      // Clear J (důvod). Leave M (formula) intact.
      clearCell(ws, J);
    }
  }
}

export function downloadFilled(wb: XLSX.WorkBook, jmeno: string, year: number, month: number): void {
  const out = XLSX.write(wb, { bookType: "xlsx", type: "array", cellStyles: true });
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
