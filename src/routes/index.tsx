import { createFileRoute } from "@tanstack/react-router";
import { useEffect, useMemo, useState } from "react";
import {
  CZECH_MONTHS,
  CZECH_WEEKDAYS_SHORT,
  defaultMonthYear,
  formatHHMM,
  lastDayOfMonth,
  parseHHMM,
  ymd,
} from "@/lib/dates";
import { lsGet, lsGetJSON, lsSet, lsSetJSON } from "@/lib/storage";
import { fetchCzechHolidays } from "@/lib/holidays";
import { fetchIcs, parseVacationDays } from "@/lib/ics";
import {
  downloadFilled,
  fillWorkbook,
  loadTemplate,
  type DayRow,
} from "@/lib/xlsx-fill";

export const Route = createFileRoute("/")({
  component: Index,
});

const CODES = ["", "D", "S", "SD", "DPN", "OČR", "PV", "HO", "PC", "ŠK", "SO"] as const;

interface DefaultTimes {
  arrival: string;
  departure: string;
  lunchStart: string;
  lunchEnd: string;
}

const DEFAULT_TIMES: DefaultTimes = {
  arrival: "08:00",
  departure: "16:30",
  lunchStart: "12:00",
  lunchEnd: "12:30",
};

function Index() {
  const [step, setStep] = useState(1);
  const [icsUrl, setIcsUrl] = useState("");
  const [remember, setRemember] = useState(true);
  const [icsLoading, setIcsLoading] = useState(false);
  const [icsError, setIcsError] = useState<string | null>(null);
  const [vacationDays, setVacationDays] = useState<Set<string>>(new Set());
  const [icsLoaded, setIcsLoaded] = useState(false);

  const initial = defaultMonthYear();
  const [month, setMonth] = useState<number>(initial.month);
  const [year, setYear] = useState<number>(initial.year);
  const [jmeno, setJmeno] = useState("");
  const [uvazek, setUvazek] = useState("1,0");
  const [praciste, setPraciste] = useState("Fablab");
  const [times, setTimes] = useState<DefaultTimes>(DEFAULT_TIMES);

  const [holidays, setHolidays] = useState<Set<string>>(new Set());
  const [rows, setRows] = useState<DayRow[]>([]);

  // Load persisted values
  useEffect(() => {
    const savedUrl = lsGet("ics_url");
    if (savedUrl) {
      setIcsUrl(savedUrl);
      setRemember(true);
    }
    setJmeno(lsGet("jmeno") ?? "");
    setUvazek(lsGet("uvazek") ?? "1,0");
    setPraciste(lsGet("praciste") ?? "Fablab");
    setTimes(lsGetJSON<DefaultTimes>("default_times", DEFAULT_TIMES));
  }, []);

  // Persist identity
  useEffect(() => { if (jmeno) lsSet("jmeno", jmeno); }, [jmeno]);
  useEffect(() => { lsSet("uvazek", uvazek); }, [uvazek]);
  useEffect(() => { lsSet("praciste", praciste); }, [praciste]);
  useEffect(() => { lsSetJSON("default_times", times); }, [times]);

  async function handleLoadIcs() {
    setIcsLoading(true);
    setIcsError(null);
    try {
      const text = await fetchIcs(icsUrl.trim());
      // Parse a wide window so user can change month afterwards
      const start = new Date(year - 1, 0, 1);
      const end = new Date(year + 1, 11, 31);
      const days = parseVacationDays(text, start, end);
      setVacationDays(days);
      setIcsLoaded(true);
      if (remember) lsSet("ics_url", icsUrl.trim());
      else lsSet("ics_url", "");
    } catch (e: any) {
      setIcsError(e?.message || "Načtení selhalo.");
      setIcsLoaded(false);
    } finally {
      setIcsLoading(false);
    }
  }

  function skipIcs() {
    setVacationDays(new Set());
    setIcsLoaded(true);
    setStep(2);
  }

  // When entering step 3, build rows
  useEffect(() => {
    if (step !== 3) return;
    let cancelled = false;
    (async () => {
      const hol = await fetchCzechHolidays(year);
      if (cancelled) return;
      setHolidays(hol);
      const last = lastDayOfMonth(year, month);
      const arrM = parseHHMM(times.arrival) ?? 8 * 60;
      const depM = parseHHMM(times.departure) ?? 16 * 60 + 30;
      const lsM = parseHHMM(times.lunchStart) ?? 12 * 60;
      const leM = parseHHMM(times.lunchEnd) ?? 12 * 60 + 30;
      const newRows: DayRow[] = [];
      for (let d = 1; d <= last; d++) {
        const date = new Date(year, month - 1, d);
        const wd = date.getDay();
        const isWeekend = wd === 0 || wd === 6;
        const dateStr = ymd(year, month, d);
        let code = "";
        if (!isWeekend) {
          if (hol.has(dateStr)) code = "S";
          else if (vacationDays.has(dateStr)) code = "D";
        }
        newRows.push({
          day: d,
          date: dateStr,
          weekday: wd,
          arrivalMin: arrM,
          departureMin: depM,
          lunchStartMin: lsM,
          lunchEndMin: leM,
          code,
          isWeekend,
        });
      }
      setRows(newRows);
    })();
    return () => { cancelled = true; };
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [step, month, year]);

  function updateRow(idx: number, patch: Partial<DayRow>) {
    setRows(prev => prev.map((r, i) => (i === idx ? { ...r, ...patch } : r)));
  }

  function rowHours(r: DayRow): number | null {
    if (r.isWeekend && !r.code) return null;
    // Holiday rows only contribute hours when the user has actually filled them in.
    if (r.code === "S" && !r.userEdited) return null;
    // Other codes (D, SD, ...) don't contribute to worked/holiday hours.
    if (r.code && r.code !== "S") return null;
    const work = r.departureMin - r.arrivalMin;
    const lunch = r.lunchEndMin - r.lunchStartMin;
    const mins = work - lunch;
    return mins / 60;
  }

  const totals = useMemo(() => {
    let worked = 0;
    let svatek = 0;
    const counts: Record<string, number> = { D: 0, SD: 0, DPN: 0, "OČR": 0, PV: 0 };
    for (const r of rows) {
      if (r.isWeekend && !r.code) continue;
      if (r.code === "S") {
        // Only edited holiday rows roll into the Svátek total.
        const h = rowHours(r);
        if (h !== null) svatek += h;
        continue;
      }
      if (r.code) {
        if (counts[r.code] !== undefined) counts[r.code]++;
        continue;
      }
      const h = rowHours(r);
      if (h !== null) worked += h;
    }
    return {
      worked,
      svatek,
      dovolena: counts.D * 8,
      sick: counts.SD * 8,
      dpn: counts.DPN * 8,
      ocr: counts["OČR"] * 8,
      pv: counts.PV * 8,
    };
  }, [rows]);

  async function handleDownload() {
    try {
      const wb = await loadTemplate();
      fillWorkbook(wb, { jmeno, uvazek, praciste, month, year, rows });
      await downloadFilled(wb, jmeno, year, month);
    } catch (e: any) {
      alert(e?.message || "Generování selhalo.");
    }
  }

  return (
    <div className="min-h-screen bg-background">
      <div className="mx-auto max-w-3xl px-4 py-8">
        <header className="mb-8">
          <h1 className="text-3xl font-bold tracking-tight">Docházka generátor</h1>
          <p className="mt-1 text-sm text-muted-foreground">
            Vyplní měsíční výkaz docházky podle sdíleného kalendáře.
          </p>
          <Stepper step={step} />
        </header>

        {step === 1 && (
          <StepIcs
            icsUrl={icsUrl}
            setIcsUrl={setIcsUrl}
            remember={remember}
            setRemember={setRemember}
            loading={icsLoading}
            error={icsError}
            loaded={icsLoaded}
            vacationCount={vacationDays.size}
            sampleDates={Array.from(vacationDays).sort().slice(0, 5)}
            onLoad={handleLoadIcs}
            onNext={() => setStep(2)}
            onSkip={skipIcs}
          />
        )}

        {step === 2 && (
          <StepIdentity
            month={month} setMonth={setMonth}
            year={year} setYear={setYear}
            jmeno={jmeno} setJmeno={setJmeno}
            uvazek={uvazek} setUvazek={setUvazek}
            praciste={praciste} setPraciste={setPraciste}
            times={times} setTimes={setTimes}
            onBack={() => setStep(1)}
            onNext={() => setStep(3)}
          />
        )}

        {step === 3 && (
          <StepPreview
            month={month} year={year}
            rows={rows} updateRow={updateRow}
            holidays={holidays}
            totals={totals}
            onBack={() => setStep(2)}
            onDownload={handleDownload}
          />
        )}
      </div>
    </div>
  );
}

function Stepper({ step }: { step: number }) {
  const labels = ["Kalendář", "Identita", "Náhled"];
  return (
    <ol className="mt-6 flex items-center gap-2 text-sm">
      {labels.map((l, i) => {
        const n = i + 1;
        const active = n === step;
        const done = n < step;
        return (
          <li key={l} className="flex items-center gap-2">
            <span
              className={
                "flex h-7 w-7 items-center justify-center rounded-full text-xs font-semibold " +
                (active
                  ? "bg-primary text-primary-foreground"
                  : done
                    ? "bg-primary/30 text-primary-foreground"
                    : "bg-muted text-muted-foreground")
              }
            >
              {n}
            </span>
            <span className={active ? "font-medium" : "text-muted-foreground"}>{l}</span>
            {n < labels.length && <span className="mx-1 text-muted-foreground">›</span>}
          </li>
        );
      })}
    </ol>
  );
}

function Field({ label, children }: { label: string; children: React.ReactNode }) {
  return (
    <label className="block">
      <span className="mb-1 block text-sm font-medium">{label}</span>
      {children}
    </label>
  );
}

const inputCls =
  "w-full rounded-md border border-input bg-background px-3 py-2 text-sm shadow-xs outline-none transition focus:border-ring focus:ring-2 focus:ring-ring/30";

function StepIcs(props: {
  icsUrl: string; setIcsUrl: (v: string) => void;
  remember: boolean; setRemember: (v: boolean) => void;
  loading: boolean; error: string | null; loaded: boolean;
  vacationCount: number; sampleDates: string[];
  onLoad: () => void; onNext: () => void; onSkip: () => void;
}) {
  return (
    <section className="rounded-lg border bg-card p-6 shadow-sm">
      <h2 className="text-xl font-semibold">1. Sdílený kalendář</h2>
      <p className="mt-1 text-sm text-muted-foreground">
        Vlož odkaz na sdílený M365 kalendář (.ics). Z kalendáře načteme dny dovolené.
      </p>
      <div className="mt-4 space-y-3">
        <Field label="ICS odkaz">
          <input
            type="url"
            className={inputCls}
            placeholder="https://outlook.office365.com/owa/calendar/<id>@<tenant>/<token>/calendar.ics"
            value={props.icsUrl}
            onChange={(e) => props.setIcsUrl(e.target.value)}
          />
        </Field>
        <label className="flex items-center gap-2 text-sm">
          <input
            type="checkbox"
            checked={props.remember}
            onChange={(e) => props.setRemember(e.target.checked)}
          />
          Zapamatovat odkaz v tomto prohlížeči
        </label>
        <div className="flex flex-wrap gap-2">
          <button
            type="button"
            disabled={!props.icsUrl || props.loading}
            onClick={props.onLoad}
            className="rounded-md bg-primary px-4 py-2 text-sm font-medium text-primary-foreground hover:bg-primary/90 disabled:opacity-50"
          >
            {props.loading ? "Načítám…" : "Načíst"}
          </button>
          <button
            type="button"
            onClick={props.onSkip}
            className="rounded-md border border-input bg-background px-4 py-2 text-sm font-medium hover:bg-accent"
          >
            Přeskočit
          </button>
        </div>
        {props.error && (
          <p className="rounded-md bg-destructive/10 p-3 text-sm text-destructive">{props.error}</p>
        )}
        {props.loaded && (
          <div className="rounded-md bg-muted p-3 text-sm">
            Načteno {props.vacationCount} dní dovolené.
            {props.sampleDates.length > 0 && (
              <span className="text-muted-foreground"> (např. {props.sampleDates.join(", ")})</span>
            )}
          </div>
        )}
      </div>
      <div className="mt-6 flex justify-end gap-2">
        <button
          type="button"
          disabled={!props.loaded}
          onClick={props.onNext}
          className="rounded-md bg-primary px-4 py-2 text-sm font-medium text-primary-foreground hover:bg-primary/90 disabled:opacity-50"
        >
          Další →
        </button>
      </div>
    </section>
  );
}

function StepIdentity(props: {
  month: number; setMonth: (v: number) => void;
  year: number; setYear: (v: number) => void;
  jmeno: string; setJmeno: (v: string) => void;
  uvazek: string; setUvazek: (v: string) => void;
  praciste: string; setPraciste: (v: string) => void;
  times: DefaultTimes; setTimes: (v: DefaultTimes) => void;
  onBack: () => void; onNext: () => void;
}) {
  const [advOpen, setAdvOpen] = useState(false);
  return (
    <section className="rounded-lg border bg-card p-6 shadow-sm">
      <h2 className="text-xl font-semibold">2. Měsíc a údaje</h2>
      <div className="mt-4 grid gap-4 sm:grid-cols-2">
        <Field label="Měsíc">
          <select
            className={inputCls}
            value={props.month}
            onChange={(e) => props.setMonth(Number(e.target.value))}
          >
            {CZECH_MONTHS.map((m, i) => (
              <option key={m} value={i + 1}>{m}</option>
            ))}
          </select>
        </Field>
        <Field label="Rok">
          <input
            type="number"
            className={inputCls}
            value={props.year}
            onChange={(e) => props.setYear(Number(e.target.value))}
          />
        </Field>
        <Field label="Jméno">
          <input
            type="text"
            className={inputCls}
            value={props.jmeno}
            onChange={(e) => props.setJmeno(e.target.value)}
            placeholder="Jan Novák"
          />
        </Field>
        <Field label="Úvazek">
          <input
            type="text"
            className={inputCls}
            value={props.uvazek}
            onChange={(e) => props.setUvazek(e.target.value)}
          />
        </Field>
        <Field label="Pracoviště">
          <input
            type="text"
            className={inputCls}
            value={props.praciste}
            onChange={(e) => props.setPraciste(e.target.value)}
          />
        </Field>
      </div>

      <div className="mt-4">
        <button
          type="button"
          onClick={() => setAdvOpen(o => !o)}
          className="text-sm text-muted-foreground underline-offset-4 hover:underline"
        >
          {advOpen ? "▾" : "▸"} Pokročilé — výchozí časy pracovního dne
        </button>
        {advOpen && (
          <div className="mt-3 grid gap-3 sm:grid-cols-4">
            {(
              [
                ["arrival", "Příchod"],
                ["departure", "Odchod"],
                ["lunchStart", "Oběd od"],
                ["lunchEnd", "Oběd do"],
              ] as const
            ).map(([k, lbl]) => (
              <Field key={k} label={lbl}>
                <input
                  type="time"
                  className={inputCls}
                  value={props.times[k]}
                  onChange={(e) =>
                    props.setTimes({ ...props.times, [k]: e.target.value })
                  }
                />
              </Field>
            ))}
          </div>
        )}
      </div>

      <div className="mt-6 flex justify-between">
        <button
          type="button"
          onClick={props.onBack}
          className="rounded-md border border-input bg-background px-4 py-2 text-sm font-medium hover:bg-accent"
        >
          ← Zpět
        </button>
        <button
          type="button"
          disabled={!props.jmeno.trim()}
          onClick={props.onNext}
          className="rounded-md bg-primary px-4 py-2 text-sm font-medium text-primary-foreground hover:bg-primary/90 disabled:opacity-50"
        >
          Další →
        </button>
      </div>
    </section>
  );
}

function StepPreview(props: {
  month: number; year: number;
  rows: DayRow[]; updateRow: (i: number, p: Partial<DayRow>) => void;
  holidays: Set<string>;
  totals: { worked: number; svatek: number; dovolena: number; sick: number; dpn: number; ocr: number; pv: number };
  onBack: () => void; onDownload: () => void;
}) {
  return (
    <section className="rounded-lg border bg-card p-6 shadow-sm">
      <h2 className="text-xl font-semibold">
        3. Náhled — {CZECH_MONTHS[props.month - 1]} {props.year}
      </h2>
      <p className="mt-1 text-sm text-muted-foreground">
        Hodnoty můžeš upravit. Stažený XLSX obsahuje stejné pole, vzorce a formátování jako šablona.
      </p>

      <div className="mt-4 overflow-x-auto">
        <table className="w-full min-w-[720px] border-collapse text-sm">
          <thead>
            <tr className="border-b text-left text-xs uppercase text-muted-foreground">
              <th className="px-2 py-2">Datum</th>
              <th className="px-2 py-2">Den</th>
              <th className="px-2 py-2">Příchod</th>
              <th className="px-2 py-2">Odchod</th>
              <th className="px-2 py-2">Oběd od</th>
              <th className="px-2 py-2">Oběd do</th>
              <th className="px-2 py-2">Kód</th>
              <th className="px-2 py-2 text-right">Hodiny</th>
            </tr>
          </thead>
          <tbody>
            {props.rows.map((r, idx) => (
              <DayRowEditor
                key={r.day}
                row={r}
                isHoliday={props.holidays.has(r.date)}
                onChange={(p) => props.updateRow(idx, p)}
              />
            ))}
          </tbody>
          <tfoot>
            <tr className="border-t font-medium">
              <td colSpan={7} className="px-2 py-2 text-right">Odpracováno:</td>
              <td className="px-2 py-2 text-right">{props.totals.worked.toFixed(1)} h</td>
            </tr>
          </tfoot>
        </table>
      </div>

      <div className="mt-4 grid gap-2 rounded-md bg-muted p-4 text-sm sm:grid-cols-2">
        <div>Svátek (S): <strong>{props.totals.svatek} h</strong></div>
        <div>Dovolená (D): <strong>{props.totals.dovolena} h</strong></div>
        <div>Sick day (SD): <strong>{props.totals.sick} h</strong></div>
        <div>DPN: <strong>{props.totals.dpn} h</strong></div>
        <div>OČR: <strong>{props.totals.ocr} h</strong></div>
        <div>Placené volno (PV): <strong>{props.totals.pv} h</strong></div>
      </div>

      <div className="mt-6 flex justify-between">
        <button
          type="button"
          onClick={props.onBack}
          className="rounded-md border border-input bg-background px-4 py-2 text-sm font-medium hover:bg-accent"
        >
          ← Zpět
        </button>
        <button
          type="button"
          onClick={props.onDownload}
          className="rounded-md bg-primary px-4 py-2 text-sm font-medium text-primary-foreground hover:bg-primary/90"
        >
          Stáhnout XLSX
        </button>
      </div>
    </section>
  );
}

function DayRowEditor({
  row, isHoliday, onChange,
}: {
  row: DayRow;
  isHoliday: boolean;
  onChange: (p: Partial<DayRow>) => void;
}) {
  const wdLabel = CZECH_WEEKDAYS_SHORT[row.weekday];
  let rowCls = "border-b";
  if (row.isWeekend) rowCls += " bg-muted/60 text-muted-foreground";
  else if (isHoliday) rowCls += " bg-red-50";
  else if (row.code === "D") rowCls += " bg-blue-50";

  const editable = !row.isWeekend || !!row.code;
  const hours = (() => {
    if (row.isWeekend && !row.code) return null;
    if (row.code) return null;
    const w = row.departureMin - row.arrivalMin;
    const l = row.lunchEndMin - row.lunchStartMin;
    return (w - l) / 60;
  })();

  function setTime(field: keyof DayRow, value: string) {
    const m = parseHHMM(value);
    if (m === null) return;
    onChange({ [field]: m } as Partial<DayRow>);
  }

  const cellInputCls =
    "w-24 rounded border border-input bg-background px-2 py-1 text-xs disabled:opacity-50";

  return (
    <tr className={rowCls}>
      <td className="px-2 py-1.5 whitespace-nowrap font-mono text-xs">{row.date}</td>
      <td className="px-2 py-1.5">{wdLabel}</td>
      <td className="px-2 py-1.5">
        <input
          type="time"
          className={cellInputCls}
          disabled={!editable}
          value={formatHHMM(row.arrivalMin)}
          onChange={(e) => setTime("arrivalMin", e.target.value)}
        />
      </td>
      <td className="px-2 py-1.5">
        <input
          type="time"
          className={cellInputCls}
          disabled={!editable}
          value={formatHHMM(row.departureMin)}
          onChange={(e) => setTime("departureMin", e.target.value)}
        />
      </td>
      <td className="px-2 py-1.5">
        <input
          type="time"
          className={cellInputCls}
          disabled={!editable}
          value={formatHHMM(row.lunchStartMin)}
          onChange={(e) => setTime("lunchStartMin", e.target.value)}
        />
      </td>
      <td className="px-2 py-1.5">
        <input
          type="time"
          className={cellInputCls}
          disabled={!editable}
          value={formatHHMM(row.lunchEndMin)}
          onChange={(e) => setTime("lunchEndMin", e.target.value)}
        />
      </td>
      <td className="px-2 py-1.5">
        <select
          className="rounded border border-input bg-background px-2 py-1 text-xs"
          value={row.code}
          onChange={(e) => onChange({ code: e.target.value })}
        >
          {CODES.map((c) => (
            <option key={c} value={c}>{c || "—"}</option>
          ))}
        </select>
      </td>
      <td className="px-2 py-1.5 text-right tabular-nums">
        {hours === null ? "—" : hours.toFixed(1)}
      </td>
    </tr>
  );
}
