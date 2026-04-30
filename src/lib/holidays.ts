import { lsGet, lsSet } from "./storage";

export interface Holiday { date: string; localName: string; name: string }

export async function fetchCzechHolidays(year: number): Promise<Set<string>> {
  const cacheKey = `holidays_cz_${year}`;
  const cached = lsGet(cacheKey);
  if (cached) {
    try {
      const arr = JSON.parse(cached) as Holiday[];
      return new Set(arr.map(h => h.date));
    } catch {}
  }
  try {
    const res = await fetch(`https://date.nager.at/api/v3/PublicHolidays/${year}/CZ`);
    if (!res.ok) return new Set();
    const data = (await res.json()) as Holiday[];
    lsSet(cacheKey, JSON.stringify(data));
    return new Set(data.map(h => h.date));
  } catch {
    return new Set();
  }
}
