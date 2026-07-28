# Live data sync — setup

Pushes the **Copeland First Pass Q Report** workbook into Supabase so the dashboard
loads live and refreshes itself. Uploading a file still works as a fallback.

```
Workbook changes  →  Power Automate  →  Office Script reads sheets to JSON
                  →  HTTP POST to Supabase RPC  →  dashboard auto-refreshes
```

Nobody logs in. No password is stored anywhere.

---

## Before you start

Check one thing: in Power Automate, create a flow and search actions for **HTTP**.
If the plain `HTTP` action is greyed out or blocked, that tenant lacks the premium
connector and this route stops here — tell me and I'll switch to a different approach.

**Your ingest secret** (treat like a password — it's the only thing guarding writes):

```
2769cca6ec7318717d562a43548d1428a3babe3d85ced4cf
```

Rotate any time with:
```sql
update app_secrets
set value = encode(gen_random_bytes(24),'hex')
where key = 'fp_ingest_secret';
```

---

## Step 1 — Create the Office Script

Open the workbook in **Excel on the web** → **Automate** → **New Script**. Paste this,
name it `Export First Pass Report`, save.

It converts Excel date serials to `yyyy-mm-dd` so days can't shift by a timezone.

```ts
function main(workbook: ExcelScript.Workbook) {
  const out: { rows: object[]; names: object[] } = { rows: [], names: [] };

  const serialToISO = (v: string | number | boolean): string => {
    if (typeof v === "number" && v > 0) {
      const d = new Date(Math.round((v - 25569) * 86400 * 1000));
      return d.toISOString().slice(0, 10);
    }
    const s = String(v ?? "").trim();
    const m = s.match(/^(\d{1,2})[\/\-](\d{1,2})[\/\-](\d{4})$/); // M/D/YYYY
    if (m) {
      return `${m[3]}-${m[1].padStart(2, "0")}-${m[2].padStart(2, "0")}`;
    }
    return /^\d{4}-\d{2}-\d{2}/.test(s) ? s.slice(0, 10) : "";
  };

  const txt = (v: string | number | boolean): string =>
    v === null || v === undefined ? "" : String(v).trim();

  // ---- Daily Results ----
  const sheet =
    workbook.getWorksheet("Daily Results") ??
    workbook.getWorksheet("old daily results");
  if (!sheet) throw new Error('Sheet "Daily Results" not found');

  const values = sheet.getUsedRange().getValues();
  // Column order must match the workbook:
  // Panel, Date, Pass_Fail, Reason, Line, Emp, Routing_Hours, Tech, Layout, Wire, FA, Sales_Order
  for (let i = 1; i < values.length; i++) {
    const r = values[i];
    const date = serialToISO(r[1]);
    if (!date) continue; // skip blank/!invalid dates
    out.rows.push({
      panel: txt(r[0]),
      date: date,
      pass_fail: txt(r[2]),
      reason: txt(r[3]),
      line: txt(r[4]),
      emp: txt(r[5]),
      routing_hours: txt(r[6]),
      tech: txt(r[7]),
      layout: txt(r[8]),
      wire: txt(r[9]),
      fa: txt(r[10]),
      sales_order: txt(r[11]),
    });
  }

  // ---- Names (initials -> full name) ----
  const ns = workbook.getWorksheet("Names");
  if (ns) {
    const nv = ns.getUsedRange().getValues();
    const hdr = nv[0].map((h) => txt(h).toLowerCase());
    const ci = hdr.indexOf("initials");
    const cn = hdr.indexOf("name");
    if (ci >= 0 && cn >= 0) {
      for (let i = 1; i < nv.length; i++) {
        const ini = txt(nv[i][ci]);
        const nm = txt(nv[i][cn]);
        if (ini && nm) out.names.push({ initials: ini, name: nm });
      }
    }
  }

  console.log(`rows=${out.rows.length} names=${out.names.length}`);
  return out;
}
```

---

## Step 2 — Build the flow

**Power Automate** → **Create** → **Automated cloud flow**.

**Trigger:** *When a file is modified (properties only)* — SharePoint
- Site Address: the **PanelServices** site
- Library Name: the library holding the report

> Prefer a fixed cadence instead? Use **Recurrence** (e.g. every 15 minutes). Simpler,
> and avoids retriggering on unrelated files in the library.

**Action 2:** *Run script* — Excel Online (Business)
- Location / Document Library / File: the report workbook
- Script: `Export First Pass Report`

**Action 3:** *HTTP*
- **Method:** `POST`
- **URI:**
  ```
  https://ptbhguthosenkffjhbry.supabase.co/rest/v1/rpc/ingest_fp_report
  ```
- **Headers:**
  | Key | Value |
  |---|---|
  | `Content-Type` | `application/json` |
  | `apikey` | *(anon key — Step 3)* |
  | `Authorization` | `Bearer` + *(same anon key)* |

- **Body** — the two `result` values come from the Run script action's dynamic content:
  ```json
  {
    "p_secret": "2769cca6ec7318717d562a43548d1428a3babe3d85ced4cf",
    "p_rows":   @{outputs('Run_script')?['body/result']?['rows']},
    "p_names":  @{outputs('Run_script')?['body/result']?['names']}
  }
  ```

Save and **Test** → **Manually**.

---

## Step 3 — The anon key

Already embedded in the dashboard (it's a public read key, safe to use here). Copy the
value of `SB_ANON` from near the top of `first-pass/index.html`, or grab it from
Supabase → Project Settings → API → `anon` `public`.

---

## Step 4 — Confirm

A successful run returns:

```json
{ "ok": true, "rows": 1234, "names": 12 }
```

Then open the dashboard — it should show a green **● Live · updated …** next to the
period. It re-checks every 60 seconds and redraws only when the workbook actually
changed. **⟳ Refresh** forces an immediate re-read.

---

## Security notes

- The write path is a `SECURITY DEFINER` function gated by the secret above. RLS stays
  **on** and the public anon key **cannot** write — verified by test.
- No `service_role` key is used anywhere, so nothing bypasses RLS.
- The secret lives only inside the flow (server-side, protected by M365).
- An empty payload is **rejected**, so a misfiring flow can't wipe the table. The whole
  ingest is one transaction — a bad run rolls back and the dashboard keeps the last
  good data.
- Each sync is a full refresh, so edits and deletions in the workbook are reflected exactly.

## Troubleshooting

| Result | Cause |
|---|---|
| `401` / `"unauthorized"` | `p_secret` doesn't match `app_secrets` |
| `404` on the RPC URL | Check the path is `/rest/v1/rpc/ingest_fp_report` |
| `"refusing empty payload"` | Script found no dated rows — check the sheet name and that column B is the date |
| Flow OK but dashboard empty | Hard-refresh the browser; confirm rows exist in `fp_daily_results` |
