import ICAL from "ical.js";
import { ymd } from "./dates";

const VACATION_KEYWORDS = ["dovolená", "dovolena", "vacation", "ooo", "pto", "volno"];

function isVacationSummary(summary: string): boolean {
  const s = summary.toLowerCase();
  return VACATION_KEYWORDS.some(k => s.includes(k));
}

function dateToYMD(d: Date): string {
  return ymd(d.getFullYear(), d.getMonth() + 1, d.getDate());
}

/**
 * Parse ICS text and return a Set of YYYY-MM-DD strings that are vacation days
 * within [yearStart, yearEnd] (exclusive end-year filter not strict — caller
 * filters by month later).
 */
export function parseVacationDays(icsText: string, windowStart: Date, windowEnd: Date): Set<string> {
  const result = new Set<string>();
  let jcal: any;
  try {
    jcal = ICAL.parse(icsText);
  } catch (e) {
    throw new Error("ICS soubor nelze přečíst (neplatný formát).");
  }
  const comp = new ICAL.Component(jcal);
  const vevents = comp.getAllSubcomponents("vevent");

  // Workday window for "covers full day" check (in minutes from midnight)
  const WORK_START = 8 * 60;
  const WORK_END = 16 * 60 + 30;

  for (const ve of vevents) {
    const event = new ICAL.Event(ve);
    const summary: string = event.summary || "";
    if (!isVacationSummary(summary)) continue;

    const isAllDay = event.startDate?.isDate === true;

    const expand = (occStart: any, occEnd: any) => {
      const startJs: Date = occStart.toJSDate();
      const endJs: Date = occEnd ? occEnd.toJSDate() : new Date(startJs.getTime() + 60 * 60 * 1000);
      if (isAllDay) {
        // All-day: end is exclusive in ICS
        const cur = new Date(startJs);
        while (cur < endJs) {
          if (cur >= windowStart && cur <= windowEnd) {
            result.add(dateToYMD(cur));
          }
          cur.setDate(cur.getDate() + 1);
        }
      } else {
        // Timed: only count if it covers the workday on a single calendar date
        const sameDay = startJs.toDateString() === new Date(endJs.getTime() - 1).toDateString();
        if (!sameDay) {
          // Multi-day timed event: include each fully-covered day
          const cur = new Date(startJs);
          cur.setHours(0, 0, 0, 0);
          while (cur < endJs) {
            const dayStart = new Date(cur); dayStart.setHours(8, 0, 0, 0);
            const dayEnd = new Date(cur); dayEnd.setHours(16, 30, 0, 0);
            if (startJs <= dayStart && endJs >= dayEnd) {
              if (cur >= windowStart && cur <= windowEnd) result.add(dateToYMD(cur));
            }
            cur.setDate(cur.getDate() + 1);
          }
        } else {
          const startMin = startJs.getHours() * 60 + startJs.getMinutes();
          const endMin = endJs.getHours() * 60 + endJs.getMinutes();
          if (startMin <= WORK_START && endMin >= WORK_END) {
            if (startJs >= windowStart && startJs <= windowEnd) {
              result.add(dateToYMD(startJs));
            }
          }
        }
      }
    };

    if (event.isRecurring()) {
      const it = event.iterator();
      let next: any;
      let safety = 0;
      while ((next = it.next()) && safety < 2000) {
        safety++;
        const occ = event.getOccurrenceDetails(next);
        if (occ.startDate.toJSDate() > windowEnd) break;
        if (occ.endDate.toJSDate() < windowStart) continue;
        expand(occ.startDate, occ.endDate);
      }
    } else {
      expand(event.startDate, event.endDate);
    }
  }
  return result;
}

export async function fetchIcs(url: string): Promise<string> {
  // Try direct first
  try {
    const res = await fetch(url);
    if (res.ok) {
      const text = await res.text();
      if (text.includes("BEGIN:VCALENDAR")) return text;
    }
  } catch {
    // CORS or network — fall through to proxy
  }
  const proxy = `/api/ics?url=${encodeURIComponent(url)}`;
  const res = await fetch(proxy);
  if (!res.ok) throw new Error(`Nepodařilo se načíst kalendář (HTTP ${res.status}).`);
  const text = await res.text();
  if (!text.includes("BEGIN:VCALENDAR")) throw new Error("Odpověď není platný ICS soubor.");
  return text;
}
