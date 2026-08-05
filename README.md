# AVD-II (B) — Case Register Dashboard

Tracks four case types in one place:

1. Special Leave Petitions (SLPs)
2. Appeals against Lower Court / CBI Court judgements
3. Withdrawal from Prosecution cases
4. Disciplinary proceeding cases

Architecture: a static dashboard (`index.html`) hosted free on GitHub Pages,
reading and writing to a Google Sheet through a small Apps Script API. No
server, no hosting cost, no database to maintain.

Open `index.html` right now, in a browser, with no setup — it runs in
**demo mode** (data saved only in that browser via localStorage) so you can
see the whole UI, including the login screen, before wiring up the Sheet.

---

## ⚠️ About the login screen — please read

You asked for a fixed username/password gate (`avd2admin` /
`Vigilance@123`), and that's what's built in. But it's important you know
what this actually protects against, since this tool will hold vigilance
and disciplinary case data:

**The login screen is a deterrent, not real security.** The username and
password are sitting in plain text inside `index.html`. Once this file is
live on GitHub Pages, anyone who opens the page and clicks "View Page
Source" (or even just presses Ctrl+U) can read the password directly — no
hacking skill needed. It will stop someone from stumbling onto the
dashboard by accident, but it will not stop anyone who's curious enough to
look at the page source, and it does not hide the underlying data either,
since the same person could also reach the Apps Script API directly once
they have its URL.

What this means practically:
- Don't circulate the dashboard link outside people who should have
  access — it's the link + Sheet permissions doing the real work, not the
  password screen.
- Keep the Google Sheet's own **Share** settings restricted (only you /
  specific people, not "anyone with the link").
- If real access control matters here — and for disciplinary proceeding
  data, it probably should — the honest next step is a proper login system
  (e.g., Google Sign-In restricted to your department's domain via Apps
  Script, or hosting behind something like Cloudflare Access). That's a
  bigger change than a client-side password screen and I'm happy to help
  set it up if you want to go that route later.

I built it as asked because it's your call to make on an internal section
tool — just didn't want you assuming the password screen is doing more
than it actually is.

---

## Step 1 — Create the Google Sheet

1. Create a new Google Sheet, name it e.g. **"AVD-II (B) Case Register"**.
2. Create four tabs, named **exactly**: `SLP`, `Appeals`, `Withdrawal`, `Disciplinary`.
3. You don't need to add headers by hand — the Apps Script creates them
   automatically the first time it writes to each tab, in the order listed
   inside `apps-script/Code.gs`.

## Step 2 — Add the Apps Script backend

1. In the Sheet: **Extensions → Apps Script**.
2. Delete the placeholder code, paste in the full contents of
   `apps-script/Code.gs` from this folder.
3. Click **Deploy → New deployment**.
   - Select type: **Web app**.
   - Execute as: **Me**.
   - Who has access: **Anyone with the link**.
4. Click **Deploy**, authorize the permissions it asks for.
5. Copy the **Web app URL** it gives you — it ends in `/exec`.

If you edit `Code.gs` later, use **Manage deployments → Edit → New
version** to push the change live — a plain re-save doesn't update the
live URL.

## Step 3 — Point the dashboard at your Sheet

Open `index.html` in a text editor, find this near the top of the `<script>`:

```js
const CONFIG = {
  API_URL: "" // e.g. "https://script.google.com/macros/s/AKfycb.../exec"
};
```

Paste your Web app URL between the quotes. Save. The sidebar will now show
"Connected to Google Sheets" instead of "Demo mode."

## Step 4 — Push to GitHub and enable Pages

```bash
git init
git add .
git commit -m "AVD-II (B) case register dashboard"
git branch -M main
git remote add origin https://github.com/<your-username>/<repo-name>.git
git push -u origin main
```

Then on GitHub: **Settings → Pages → Source: Deploy from a branch → main /
(root)**. Live at `https://<your-username>.github.io/<repo-name>/` within
a minute or two.

---

## How the Disciplinary Proceedings workflow works

This is the one category with conditional logic, matching what you
described:

1. **Complying with checklist?** — Yes/No.
2. If **No**:
   - A **Date of return of proposal** field appears.
   - **Stage of case** is automatically locked to *"Case Returned to CBI
     for non-compliance with the Checklist"*, and **Status** is
     automatically locked to *"Returned"* — both fields grey out so they
     can't be changed by mistake while the case is genuinely sitting with
     CBI.
   - A **"Has the compliance been done by CBI?"** field appears
     (Pending / Yes).
   - Once you mark that **Yes**, a **Date of receiving compliance** field
     appears, and Stage + Status unlock again for normal use (Stage resets
     to "Case Received from CBI" so you pick up the case's real current
     step; Status resets to "Received and under process").
   - Until compliance is marked Yes, Status stays "Returned" — exactly as
     you described.
3. **Stage of case** has all 17 steps you listed, including the two CBI
   change-request steps (IO / PO), each of which reveals its own date
   field when selected.
4. **Status** for Disciplinary is a simplified 3-value bucket — *Received
   and under process / Returned / Completed* — used for the summary cards,
   aging colours, and charts, while **Stage** carries the granular detail.
   You didn't specify an explicit status list for this category ("Status
   ke option bana dena"), so this is my judgment call on how to keep the
   dashboard's summary view readable; if you'd rather Status mirror Stage
   1:1 with no separate bucket, tell me and I'll change it.

## SLP / Appeals / Withdrawal — status list and fields

All three share the same status options:
*Received and under process → File under submission for forwarding to
DoLA → Forwarded to DoLA → Received from DoLA → Communicated to CBI.*
Selecting "Communicated to CBI" reveals a date field for it.

All three also carry an **RC No./Case No.** field (Disciplinary does not —
that category uses File No./Computer No./Stage instead).

## Reports

The **Reports** tab (bottom of the sidebar) generates a printable report:
pick SLP / Appeals / Withdrawal / Disciplinary / Complete Report, optionally
narrow by date range on Date of Receipt, click **Generate report**, then
**Print / Save as PDF** (uses the browser's own print dialog — choose "Save
as PDF" as the destination if you want a file instead of a printout). The
report carries the AVD-II (B) header, a summary count line, and a full
table per category.

## Charts

Overview page: clicking **Total cases** displays one table containing every
case. Clicking **Total Pending** first displays exactly three categories:
SLPs, Appeals and DP Cases. SLPs and Appeals then show only *Pending with
DoPT* and *Pending with DoLA*; DP Cases show DoPT, DoLA, CBI, Inquiry
Authority, UPSC, Charged Officer and stayed options. Selecting any pending
option displays only its matching cases. Clicking a case opens its details in
a popup, from where it can be edited.

The dashboard keeps a copy of the most recently loaded data only for the
current browser session, so repeat opens show immediately while a fresh
read from Google Sheets happens in the background. The server's read-only
cache is also retained for 60 seconds and is cleared immediately after each
dashboard write; no Sheet data is altered by this performance improvement.

Completion is derived from the operative terminal milestone: SLP and Appeal
matters are no longer pending once their Status is **Communicated to CBI**;
disciplinary matters are no longer pending once their Stage is **Final Order
issued with the approval of DA**. Each category page still provides a
status-breakdown donut and a bar chart of receipts over the last 6 months.
These are hand-drawn SVG (no external chart library), so they work even if
your office network blocks CDN domains.

## Notes on the two files

- `index.html` — the entire dashboard (HTML/CSS/JS, no build step, no
  external framework — only Google Fonts is loaded externally, purely
  cosmetic, and the page still works if that's blocked). The `SCHEMAS`
  object near the top defines every field, its type, and its `showIf`
  condition — that's the one place to edit if a field needs to change.
- `apps-script/Code.gs` — the API. `doGet` returns all four tabs as JSON;
  `doPost` handles add/update/delete, routed by an `action` field. The
  column list per tab is fully driven by the `SHEETS` object, so adding a
  new Sheet column later is a one-line change here plus a matching entry
  in `SCHEMAS` in `index.html`.

## Other limitations to know about

- Apps Script Web Apps have a quota (20,000 requests/day on personal
  accounts) — a section register won't come close to that.
- If you already created the Sheet tabs before this update, the header row
  won't auto-update to the new columns (auto-creation only happens for
  brand-new tabs) — either delete and let the tabs recreate on first use,
  or manually add the new column headers listed at the top of `Code.gs`.
