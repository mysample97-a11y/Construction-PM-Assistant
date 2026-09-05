# BIM Multi-Site Delivery Tracker

Track BIM production across several similar sites from a spreadsheet you already keep. The app reads your weekly status grid, works out where each site actually stands, and — optionally — has an AI model of your choosing write the commentary.

It is a static site. No server, no account, no database. Everything happens in your browser.

---

## The two ideas it is built on

**1. The app calculates; the AI interprets.**

Every percentage, rate, forecast, blocker age and risk is computed in JavaScript from your spreadsheet. The model receives those results as fixed facts and is explicitly instructed not to recalculate them.

This is not caution for its own sake. Ask a language model to work out a forecast from a status grid and it will return numbers that look right and are not, with no signal that anything is wrong — and a coordinator will act on them. So the boundary is enforced in code and tested: the suite asserts that the payload sent to the model contains no raw weekly grid.

**2. Sensitive information never enters the tool.**

Sites are `A-01`, people are `R1`. The mapping from those codes to real clients, addresses and names lives in a separate file that stays with you and is never loaded here. That is stronger than redaction: the data cannot leak because it was never present.

The one gap is that **task names are free text and are sent to the AI**. The app scans them for anything that looks like a real place or person and warns you before sending, but that scan is a safety net, not a guarantee. Write `Draft model - substructure`, not `Substructure at <a real building>`.

---

## What it does

- **Reads the register workbook** — sheet 1 lists your sites, sheet 2 is the standard task framework, then one sheet per site with a task list, a prerequisites table and a waiting-on log.
- **Shows you what it read** before anything else happens: which sheet was matched to which site and how, how many tasks and weeks, what could not be read, and the result of the privacy scan.
- **Computes, per site**: progress by count and by weight, throughput per week, a forecast finish from the actual completion rate, tasks past their target week, tasks stalled in WIP, blocked and waiting work with how long it has been stuck, outstanding prerequisites, scope added after kickoff, work running ahead of its dependencies, resource load, and data-quality problems.
- **Computes across sites**: which sites are slipping, who is spread across too many at once, and — the one a single-site view cannot see — **blockers that appear on more than one site**, which are process problems rather than site problems.
- **Runs site by site**, so you control token spend and can stop after any one.
- **Saves every report the moment it lands**, so a crash, a cancel or a rate limit never loses finished work.
- **Exports** to Excel, Word and PDF, with or without the AI narrative.

---

## Getting it running

### GitHub Pages

1. Push these files to `main`.
2. **Settings → Pages → Source → GitHub Actions**.

The workflow runs both test suites and only deploys if they pass. It is pure static files, so **Deploy from a branch** works too.

### Locally

Browsers refuse to load JavaScript modules over `file://`, so double-clicking `index.html` gives a blank page. Serve the folder:

```bash
python3 -m http.server 8000     # then open http://localhost:8000
```

---

## The workbook

Start from `templates/BIM-Multi-Site-Tracker-TEMPLATE.xlsx`.

```
Sheet 1  "Sites"           one row per site code - the index for the workbook
Sheet 2  "Task Template"   the standard framework, plus the prerequisite list
Sheet 3+ "A-01", "A-02"    one per site: tasks x weekly status, prereqs, blocker log
```

### The ID system

Every cross-reference uses an ID, never a name, so renaming a task breaks nothing.

| ID | Meaning |
|---|---|
| `C030` | Category |
| `T030-01` | Task **under C030** — the ID says where it belongs. Reusable as a 4D WBS code. |
| `P01` / `PT01` | Prerequisite, site-level / template-level |
| `R1` | Resource |
| `L01` | Waiting-on log entry |

### Status values

| Status | Means |
|---|---|
| `Not started` | Nobody is on it |
| `WIP` | Actively worked this week |
| `Blocked` | Stopped by something **inside** your team's control — rework, resource pulled away |
| `Waiting on` | Stopped by something **outside** it — client, another discipline, supplier |
| `Finished` | Meets the Definition of done for its category |
| `N/A` | Does not apply here — excluded from percentages rather than counted as incomplete |
| *(blank)* | The task did not exist that week. This is how scope growth is measured — do not backfill blanks. |

**Blocked and Waiting on must stay distinct**, or the bottleneck report is meaningless: one you fix yourself, the other you chase someone for. Either requires a row in the log saying why.

### Week columns

Headed with the **date the week ends**, with the week number above. The app finds them by checking whether the header is a date, so you can add columns to the right and it just works.

---

## How the numbers are produced

Read this before acting on them.

**Progress** is reported two ways because they diverge and neither is the whole truth: by task count, and weighted by the optional Weight column. The app always says which.

**Blank weeks carry forward.** A task's status in any week is its last recorded status — people fill in what moved, not every cell.

**Throughput** counts tasks that became Finished in each week. Two rates are shown: the all-time average, which is stable, and the last three weeks, which reflects what is happening now. **Where they disagree, the disagreement is the finding.**

**Forecast** is remaining tasks ÷ completion rate. Nothing more sophisticated, and that is deliberate — it uses your actual observed rate rather than a plan nobody is meeting. It is an indicator, not a re-planned date. The calculation uses the unrounded rate; the displayed rate is rounded for reading only.

**A site with zero completions gets no forecast at all**, rather than an infinite one.

**The plan line** on the curve comes from the Target week column. If you have not filled it in, there is no plan line and the app says so instead of inventing one.

**Category rows carry your own status**, which may differ from the arithmetic of the tasks under them. Neither is automatically right, so both are shown and a large gap is flagged.

---

## Before you use a key

- **Your API key stays in this browser.** Held in memory, or in `sessionStorage` if you opt in. It is never written into a session file, an export, or the autosave. Tested.
- **But a key in a browser is not secret from whoever is at that browser.** Use a key created for this tool, restricted to one model, with a spend cap.
- **Rate limits, not tokens, are what stop you.** Free tiers publish generous token allowances alongside a low requests-per-minute cap, so a token-only meter shows headroom right up to the 429. The usage rail tracks both, and retries count — each is a real request the provider counts.
- **The limit figures are defaults you can edit**, not read from your account. Providers change them without notice.

---

## Sessions and crash safety

**Autosave** writes to IndexedDB after every change and after every completed report. Reload the tab and your work is there.

**Save session** writes a `.json` file you keep. That file survives a cleared browser, moves between machines, and can be sent to a colleague. It contains the parsed workbook, every report and your settings — and deliberately no API key.

Browser storage is not durable. Save a session file at the end of each reporting cycle.

---

## Exports

| Format | How |
|---|---|
| **Excel** `.xlsx` | Full data export: summary, per-site figures, blockers, risks, scope growth, resource load, weekly throughput, and any AI reports |
| **Word** `.rtf` | A formatted report Word opens natively. `.rtf` rather than `.docx` because it needs no zip writer, so the app keeps zero runtime dependencies |
| **PDF** | The browser's own print-to-PDF from a purpose-built print view — better typography than a bundled PDF library, and always matches what you see |

Exports include the computed figures whether or not an AI report exists.

---

## Running the tests

```bash
node tools/verify-engine.mjs                          # 147 checks
npm install                                            # jsdom, for the UI test only
node --experimental-vm-modules tools/smoke-dom.mjs     # 75 checks
```

The engine suite checks the arithmetic against a fixture whose answers are known by hand, then runs the **actual shipped template** to prove the parser and the template have not drifted apart. The UI suite boots the real app under jsdom and walks the whole workflow — load, confirm, select, generate against a stubbed provider, hit a rate limit, export, save and reload a session — failing on any thrown error, any empty render, or any key appearing where it should not.

The app itself has no runtime dependencies.

---

## Project layout

```
index.html                  App shell
assets/css/                 tokens - base - components - app
assets/vendor/              SheetJS (Apache-2.0), vendored so there is no CDN call
templates/                  The tracker template users download
js/core/
  parser.js                 Workbook to model, plus the privacy screen
  engine.js                 Every computed figure
  ai.js                     Providers, retry/cancel, payloads, prompts
  tokens.js                 Token and request-window accounting
  session.js                State, autosave, session files, key handling
  exports.js                Excel, RTF, print
  charts.js                 SVG curve, throughput bars, meters
  util.js                   DOM helpers, dates, modals
js/views/app.js             The four-step workflow
tools/                      Test suites
```

---

## Limitations, stated plainly

- **The forecast is a straight-line extrapolation.** It does not know that QA is faster than modelling, or that a site is about to get another modeller. It tells you where the current rate leads, which is usually the useful question, but it is not a re-plan.
- **Dependencies are checked, not scheduled.** The app flags work running ahead of what it depends on; it does not compute a critical path. With weekly status and no durations there is nothing to compute one from.
- **Weight is optional and defaults to 1.** Leave it blank and a one-hour task counts the same as a week of modelling.
- **The privacy scan is heuristic.** It catches obvious names. You remain the check.
- **No resource levelling.** Resource load is reported; nothing is scheduled around it.

---

## Licence

MIT. SheetJS is bundled under Apache-2.0 — see `assets/vendor/XLSX-LICENSE.txt`.
