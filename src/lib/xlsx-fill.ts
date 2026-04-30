import JSZip from "jszip";
import { saveAs } from "file-saver";
import { CZECH_MONTHS, deburr, lastDayOfMonth } from "./dates";

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

const TEMPLATE_URL = "/template/vzor_dochzka.xlsx";

function escXml(s: string): string {
  return s.replace(/&/g, "&amp;").replace(/</g, "&lt;").replace(/>/g, "&gt;");
}

function escRef(ref: string): string {
  return ref.replace(/[$]/g, "\\$");
}

/**
 * Patch helpers — they only mutate the requested <c r="..."/> element and
 * preserve its `s="..."` style attribute. Anything not explicitly written
 * (cell xfs, merges, conditional formatting, tables, shared formulas in M)
 * stays byte-identical to the template.
 */
function makePatcher(initialXml: string) {
  let xml = initialXml;

  function cellRegex(ref: string): RegExp {
    // Match either a self-closing <c .../> or a full <c ...>...</c>.
    return new RegExp(
      `<c r="${escRef(ref)}"([^>]*?)(?:\\s*/>|>[\\s\\S]*?</c>)`,
    );
  }

  function stripTypeAttr(attrs: string): string {
    return attrs.replace(/\s+t="[^"]*"/g, "");
  }

  function setNum(ref: string, val: number) {
    const re = cellRegex(ref);
    xml = xml.replace(re, (_, attrs) => {
      const a = stripTypeAttr(attrs);
      return `<c r="${ref}"${a}><v>${val}</v></c>`;
    });
  }

  function setStr(ref: string, text: string) {
    const re = cellRegex(ref);
    xml = xml.replace(re, (_, attrs) => {
      const a = stripTypeAttr(attrs);
      return `<c r="${ref}"${a} t="inlineStr"><is><t xml:space="preserve">${escXml(text)}</t></is></c>`;
    });
  }

  function clr(ref: string) {
    const re = cellRegex(ref);
    xml = xml.replace(re, (_, attrs) => {
      const a = stripTypeAttr(attrs);
      return `<c r="${ref}"${a}/>`;
    });
  }

  return {
    setNum,
    setStr,
    clr,
    get xml() {
      return xml;
    },
  };
}

async function loadTemplateZip(): Promise<{ zip: JSZip; sheetPath: string; xml: string }> {
  const res = await fetch(TEMPLATE_URL);
  if (!res.ok) throw new Error("Šablona nenalezena.");
  const buf = await res.arrayBuffer();
  const zip = await JSZip.loadAsync(buf);

  const wbFile = zip.file("xl/workbook.xml");
  const relsFile = zip.file("xl/_rels/workbook.xml.rels");
  if (!wbFile || !relsFile) throw new Error("Šablona je poškozená.");
  const wbXml = await wbFile.async("string");
  const relsXml = await relsFile.async("string");

  const sheetMatches = [...wbXml.matchAll(/<sheet [^>]*r:id="([^"]+)"/g)];
  if (sheetMatches.length < 3) throw new Error("V šabloně chybí třetí list.");
  const targetRId = sheetMatches[2][1];

  const relMatch = relsXml.match(
    new RegExp(`Id="${targetRId}"[^>]*Target="([^"]+)"`),
  );
  if (!relMatch) throw new Error("Nelze najít cestu k třetímu listu.");
  const target = relMatch[1].replace(/^\/?xl\//, "").replace(/^\//, "");
  const sheetPath = `xl/${target}`;

  const sheetFile = zip.file(sheetPath);
  if (!sheetFile) throw new Error(`Sheet ${sheetPath} nenalezen v šabloně.`);
  const xml = await sheetFile.async("string");
  return { zip, sheetPath, xml };
}

export async function generateXlsx(input: FillInput): Promise<void> {
  const { zip, sheetPath, xml } = await loadTemplateZip();
  const p = makePatcher(xml);

  // Header row 4. NOTE: do NOT touch J4 — it has a VLOOKUP that derives the
  // month number from I4. Excel recalculates on open.
  p.setStr("D4", input.jmeno);
  p.setStr("F4", `úvazek ${input.uvazek}`);
  p.setStr("I4", CZECH_MONTHS[input.month - 1]);
  p.setNum("J4", input.month);
  p.setNum("K4", input.year);
  p.setStr("M4", input.praciste);

  const last = lastDayOfMonth(input.year, input.month);

  for (const d of input.rows) {
    if (d.day < 1 || d.day > last) continue;
    const r = 7 + d.day;

    // Weekend: clear D, E, H, I, J, M. Leave F, G, K, L alone.
    if (d.isWeekend && !d.code) {
      for (const c of ["D", "E", "H", "I", "J", "M"]) p.clr(`${c}${r}`);
      continue;
    }

    // State holiday: code only, no prefilled times, M cleared.
    // Exception: if the user manually filled times, write them through.
    if (d.code === "S") {
      if (d.userEdited) {
        p.setNum(`D${r}`, d.arrivalMin / 1440);
        p.setNum(`E${r}`, d.departureMin / 1440);
        p.setNum(`H${r}`, d.lunchStartMin / 1440);
        p.setNum(`I${r}`, d.lunchEndMin / 1440);
      } else {
        for (const c of ["D", "E", "H", "I"]) p.clr(`${c}${r}`);
      }
      p.setStr(`J${r}`, "S");
      p.clr(`M${r}`);
      continue;
    }

    // Vacation / sick / etc: standard times, code, M cleared.
    if (d.code) {
      p.setNum(`D${r}`, d.arrivalMin / 1440);
      p.setNum(`E${r}`, d.departureMin / 1440);
      p.setNum(`H${r}`, d.lunchStartMin / 1440);
      p.setNum(`I${r}`, d.lunchEndMin / 1440);
      p.setStr(`J${r}`, d.code);
      p.clr(`M${r}`);
      continue;
    }

    // Plain workday: fill standard times, leave M untouched
    // (its shared formula will recalc on open).
    p.setNum(`D${r}`, d.arrivalMin / 1440);
    p.setNum(`E${r}`, d.departureMin / 1440);
    p.setNum(`H${r}`, d.lunchStartMin / 1440);
    p.setNum(`I${r}`, d.lunchEndMin / 1440);
    p.clr(`J${r}`);
  }

  zip.file(sheetPath, p.xml);
  zip.remove("xl/calcChain.xml");

  const blob = await zip.generateAsync({
    type: "blob",
    mimeType:
      "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    compression: "DEFLATE",
  });

  const safeName = deburr(input.jmeno).replace(/\s+/g, "");
  const mm = String(input.month).padStart(2, "0");
  const fname = `Dochazka_${safeName || "uzivatel"}_${input.year}-${mm}.xlsx`;
  saveAs(blob, fname);
}
