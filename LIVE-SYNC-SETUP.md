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

## Step 1 — Choose how the flow reads the workbook

Two ways. **The flow needs one of them** — this step isn't optional, it's how the data
gets read at all.

| | **1A · Office Script** ✅ **recommended** | **1B · Connector only** |
|---|---|---|
| Row limit | **None** — reads the whole sheet | 256 unless pagination is configured |
| Changes your workbook | No | **Yes** — sheet must become an Excel Table |
| Column headers renamed | Still works (reads by position) | **Breaks** — mapping is by header text |
| Date conversion | Done in the script | Depends on a connector setting |
| Extra flow actions | 0 | 2 (`Select`) |

**Use 1A.** For handling every row it's the safer choice: no cap to configure, and it
reads by column position exactly like the dashboard's own parser — so a renamed header
can't silently break the sync. It also leaves the workbook untouched, which matters
since other people use that file.

1B's 256-row cap is *configurable*, not removable, and getting it wrong fails quietly
rather than loudly. That's the main reason to avoid it here.

---

### Option 1A — Office Script

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

### Option 1B — No script, connector only

**First:** in Excel, select the Daily Results data → **Insert → Table** (tick *My table
has headers*). Do the same for the Names sheet. The connector can only read Tables.

Then in the flow use these instead of *Run script*:

**Action:** *List rows present in a table* — Excel Online (Business)
- File: the report workbook · Table: your Daily Results table
- **⚙️ Settings → Pagination → On**, Threshold `100000`
- **Advanced → DateTime Format → `ISO 8601`**

> ⚠️ **The 256-row trap.** Without pagination on, this action silently returns only the
> first 256 rows — and since a partial payload isn't an *empty* one, it would replace
> your good data with a truncated set. Turn pagination on. The ingest returns
> `prev_rows` so you can sanity-check the count after each run.

**Action:** *Select* — Data Operation. From: `value` of the list action. Map:

| To | From |
|---|---|
| `panel` | `item()?['Panel']` |
| `date` | `item()?['Date']` |
| `pass_fail` | `item()?['Pass_Fail']` |
| `reason` | `item()?['Reason']` |
| `line` | `item()?['Line']` |
| `emp` | `item()?['Emp']` |
| `routing_hours` | `item()?['Routing_Hours']` |
| `tech` | `item()?['Tech']` |
| `layout` | `item()?['Layout']` |
| `wire` | `item()?['Wire']` |
| `fa` | `item()?['FA']` |
| `sales_order` | `item()?['Sales_Order']` |

The left column is fixed — those are the keys the database expects. The right column
must match your **actual header text**; correct it if your headers differ.

Repeat both actions for the **Names** table, mapping only `initials` and `name`.

In Step 2's HTTP body, use the two `Select` outputs in place of the script results:

```json
{
  "p_secret": "2769cca6ec7318717d562a43548d1428a3babe3d85ced4cf",
  "p_rows":   @{body('Select')},
  "p_names":  @{body('Select_2')}
}
```

Dates are handled either way — the database accepts ISO dates, ISO timestamps, Excel
serial numbers and `M/D/YYYY`, all verified.

---

## Step 2 — Build the flow

**Power Automate** → **Create** → **Automated cloud flow**.

**Trigger:** *When a file is modified (properties only)* — SharePoint
- Site Address: the **PanelServices** site
- Library Name: the library holding the report

> Prefer a fixed cadence instead? Use **Recurrence** (e.g. every 15 minutes). Simpler,
> and avoids retriggering on unrelated files in the library.

**Action 2:** reading the workbook — whichever you chose in Step 1
- **1A:** *Run script* — Excel Online (Business). File: the report workbook,
  Script: `Export First Pass Report`
- **1B:** the two *List rows present in a table* + *Select* pairs

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

- **Body** *(Option 1A — from the Run script action's dynamic content)*:
  ```json
  {
    "p_secret": "2769cca6ec7318717d562a43548d1428a3babe3d85ced4cf",
    "p_rows":   @{outputs('Run_script')?['body/result']?['rows']},
    "p_names":  @{outputs('Run_script')?['body/result']?['names']}
  }
  ```
  *(Option 1B uses `@{body('Select')}` / `@{body('Select_2')}` — see above.)*

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
{ "ok": true, "rows": 1234, "names": 12, "skipped_no_date": 0, "prev_rows": 1230 }
```

- `rows` — what's now live. Compare against `prev_rows`; a sudden large drop on
  Option 1B almost always means pagination is off and the read truncated at 256.
- `skipped_no_date` — rows whose Date cell couldn't be read. A few blank trailing
  rows are normal; a large number means the Date column isn't where it's expected.

Then open the dashboard — it should show a green **● Live · updated …** next to the
period. It re-checks every 60 seconds and redraws only when the workbook actually
changed. **⟳ Refresh** forces an immediate re-read.

---

## Handling very large sheets (optional)

The single POST in Step 2 is fine for typical volumes. If the sheet grows large enough
that one request gets unwieldy, switch to the chunked path — it has **no row limit**,
and live data is never partially replaced.

Three calls, same URL pattern and headers as Step 2:

**1. Begin** → `POST /rest/v1/rpc/fp_ingest_begin`
```json
{ "p_secret": "…" }
```
Returns `{ "ok": true, "token": "…" }`. Save `token` in a variable.

**2. Send batches** → `POST /rest/v1/rpc/fp_ingest_chunk`

Initialize an integer variable `idx = 0`, then **Do until** `idx >= length(<your rows>)`:
```json
{ "p_secret": "…", "p_token": "<token>",
  "p_rows": @{take(skip(<your rows>, variables('idx')), 5000)} }
```
Increment `idx` by 5000 at the end of each loop. Batch size is yours to tune.

**3. Commit** → `POST /rest/v1/rpc/fp_ingest_commit`
```json
{ "p_secret": "…", "p_token": "<token>", "p_names": <names array> }
```

Only this last call touches live data, and it does so in one transaction. Until it
runs, the dashboard keeps serving the previous sync — so a flow that dies halfway
leaves the last good data in place rather than a half-loaded table. Tokens are
single-use, so a retried or duplicated run can't double-apply.

Verified: 1200 rows over 3 batches, live data intact mid-load, wrong and reused tokens
both rejected.

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
| `"refusing empty payload"` | The read returned nothing — check the sheet/table name |
| `"no rows had a readable date"` | Date column not mapped or in an unexpected format |
| Only 256 rows synced | Option 1B with pagination off — enable it |
| Flow OK but dashboard empty | Hard-refresh the browser; confirm rows exist in `fp_daily_results` |
