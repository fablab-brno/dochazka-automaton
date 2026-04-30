## Docházka generátor — Plan

A 3-step wizard that fills the existing JINTEK attendance template (.xlsx) for a chosen month, pre-populating vacation days from a shared M365 calendar (.ics) and Czech state holidays.

### Template inspection (done — one deviation flagged)

I inspected `vzor_dochzka.xlsx` and the example `Docházka_Vejtasa.xlsx`. The 3rd sheet is `únor 2026` with these header cells:

- `D4` = Jméno (string) ✓ matches spec
- `I4` = month name in Czech (e.g. "Březen") ✓ matches spec
- `J4` = VLOOKUP for month number — leave alone ✓
- `K4` = year (number) ✓ matches spec
- **`F4`** currently holds the literal `"OČ: "` label — it is NOT the úvazek cell.
- **`L4`** holds the label `"úvazek:"`, and **`M4`** holds the úvazek value (e.g. `"X,X"` / `"0,8"`).
- There is **no "Pracoviště" cell** anywhere in the template.

The spec says: `F4 = "úvazek " + value` and `M4 = Pracoviště`. The actual template structure is the opposite: `M4 = úvazek` and there is no place for Pracoviště. **I will follow the template's actual layout** (write úvazek to `M4`, skip Pracoviště) and call this out in the UI as a small note. If you want Pracoviště written somewhere specific, tell me a cell address. The "Pracoviště" field will still be collected and persisted, just not written into the XLSX in v1.

Day rows: row `7 + n` for day `n`. Columns `D` (příchod), `E` (odchod), `H` (oběd od), `I` (oběd do), `J` (kód/důvod), `M` (shared formula `=(E-D)-(I-H)+(L-K)`). Rows beyond month length stay untouched — the template's conditional formatting greys them via `S7=DAY(EOMONTH)`.

### Step 1 — ICS link

- Single text input + "Zapamatovat" checkbox + "Načíst" button.
- Try `fetch(url)` directly. On CORS/network failure, retry via `/api/ics?url=...` server route (TanStack server route, not Vercel Edge).
- Parse with `ical.js`. Filter VEVENTs whose SUMMARY (case-insensitive) contains any of: dovolená, vacation, ooo, pto, volno.
- Expand RRULEs and multi-day events into a Set of `YYYY-MM-DD` strings.
- All-day events → vacation. Timed events covering ≥ the workday window (08:00–16:30) → vacation. Partial-day timed events → ignored.
- Show preview: count of matched events + first 5 dates.

### Step 2 — Month + identity

- Měsíc dropdown (Leden…Prosinec). Default: if `today.day >= 20` use current month, else previous month.
- Rok number input. Default matches the chosen month/year.
- Jméno text (persisted: `jmeno`).
- Úvazek text, default `"1,0"` (persisted: `uvazek`).
- Pracoviště text, default `"Fablab"` (persisted: `praciste`) — collected but not written to XLSX (see deviation above).
- Collapsible "Pokročilé": default times Příchod 08:00, Odchod 16:30, Oběd 12:00–12:30 (persisted: `default_times`).

### Step 3 — Preview table + download

Row per day with columns: Datum | Den | Příchod | Odchod | Oběd od | Oběd do | Kód | Hodiny.

Row pre-fill priority:
1. Sat/Sun → empty times, empty code, grey row.
2. Czech state holiday (cached from `date.nager.at` per-year in localStorage) → code `S`, defaults stay, light-red row, hours `—`.
3. Date in ICS vacation set → code `D`, defaults stay, light-blue row, hours `—`.
4. Else workday → defaults filled, hours = 8.

- Kód cell editable (select: ``, D, S, SD, DPN, OČR, PV, HO, PC, ŠK, SO).
- Time cells editable (HH:MM). Hodiny recomputed live; blank to `—` when any non-empty code is set.
- Footer totals (display-only): odpracovaná doba (sum of workday hours) + counts·8 for S/D/SD/DPN/OČR/PV.
- "Stáhnout XLSX" → fills template, downloads as `Dochazka_<jmeno>_<YYYY>-<MM>.xlsx` (jméno transliterated, spaces removed).

### XLSX writing rules

Library: `xlsx-js-style` (preserves styles; SheetJS community fork). Load `/template/vzor_dochzka.xlsx` as ArrayBuffer, modify cells on the 3rd sheet in place, write back as ArrayBuffer, trigger download.

Header writes:
- `D4` = jméno (string)
- `I4` = month name capitalized in Czech
- `K4` = rok (number)
- `M4` = úvazek string (the template's úvazek cell — see deviation note)
- Do NOT touch `J4`, `L4`, `F4`.

For each day `n = 1..lastDay`, row `r = 7 + n`:
- Weekend → leave `D/E/H/I/J/M` untouched; the template handles greying.
- Holiday (`S`), vacation (`D`), or any non-empty code:
  - `D{r}` = 0.3333333333333333 (or user's default Příchod / 24)
  - `E{r}` = 0.6875 (default Odchod / 24)
  - `H{r}` = 0.5
  - `I{r}` = 0.5208333333333333
  - `J{r}` = code letter (string)
  - `M{r}` = delete cell (so the existing summary block stops counting it as worked time — Excel re-evaluates on open)
- Plain workday (no code):
  - `D/E/H/I` set to user's times / 24
  - `J{r}` empty
  - `M{r}` LEAVE the existing shared formula intact
- Days beyond `lastDay` (rows 37/38 if not present) → untouched.
- Never write to `F`, `G`, `K`, `L` columns in day rows.

### Server route

`src/routes/api/ics.ts` — TanStack server route with GET handler. Validates that `url` query param starts with `https://outlook.office365.com/` or another trusted ICS host (basic SSRF guard), fetches it, returns the body as `text/calendar` with permissive CORS.

### File map

```text
public/template/vzor_dochzka.xlsx          (copy of uploaded template)
src/routes/index.tsx                       (wizard host: step state, persistence)
src/routes/api/ics.ts                      (ICS proxy server route)
src/components/wizard/StepIcs.tsx
src/components/wizard/StepIdentity.tsx
src/components/wizard/StepPreview.tsx
src/lib/ics.ts                             (ical.js parsing + vacation detection)
src/lib/holidays.ts                        (nager.at fetch + cache)
src/lib/xlsx-fill.ts                       (template load + cell writes + download)
src/lib/dates.ts                           (Czech month names, weekday short, helpers)
src/lib/storage.ts                         (localStorage helpers)
```

Dependencies to add: `xlsx-js-style`, `ical.js`.

### Open question

**F4 / M4 / Pracoviště mismatch (see top of plan).** I'll proceed by writing úvazek to `M4` and skipping Pracoviště. If you want Pracoviště written into a specific free cell (or want me to overwrite the `OČ: ` label in `F4` with `úvazek <value>` despite the existing layout), say so and I'll adjust.

### Non-goals (v1)

No reading of past XLSX files. No auto-detect of Jednání/PC/ŠK from calendar. No writing to summary block M39:M46 (Excel recomputes).
