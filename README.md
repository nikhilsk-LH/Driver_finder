# Driver Finder

A Netflix-themed, glassmorphic dashboard for finding **which drivers run at a given facility**. Drop in a roster (CSV, Excel, or a shared Google Sheet), pick a facility from the dropdown, and every driver who operates there appears as a card — no more chasing messages to work out who covers where.

It runs entirely in the browser. Your file is parsed client-side and never uploaded to a server.

## Features

- **Facility → driver lookup** — select a facility and see its drivers instantly.
- **Roster upload** — `.csv`, `.xlsx`, `.xls`, or a shared Google Sheet link.
- **Smart column detection** — auto-finds the facility, first-name, last-name, and shift columns.
- **Full names** — first and last name shown together on each card.
- **Multi-facility drivers** — a driver whose facility cell lists more than one site (e.g. `LUX, GSD`) appears under each of those facilities, tagged with a "+N more sites" badge.
- **Preferred facility** — each card shows the facility where that driver has taken the most blocks across the roster.
- **Smooth, on-brand motion** — card lift and a shine sweep across the name on hover; respects `prefers-reduced-motion`.

## Getting started

Requires [Node.js](https://nodejs.org/) 18+.

```bash
npm install
npm run dev
```

Then open the local URL Vite prints (usually http://localhost:5173).

To build a production bundle:

```bash
npm run build     # outputs to dist/
npm run preview   # serve the built bundle locally
```

## Expected data shape

One row per driver-block, with a header row. The app looks for columns named like:

- **Facility / Site / Location / Hub** — the facility. May contain multiple sites separated by `,` `;` `/` or `|`.
- **First Name** (+ optional **Last Name**) — or a single **Name** / **Driver** column.
- **Shift / Slot / Block** — optional; shown as a badge.

Any other columns (Driver ID, Phone, etc.) are shown as details on each card. Blank-facility rows are ignored.

> If the auto-detection picks the wrong column for a particular sheet, adjust the patterns in `detectColumns()` inside `src/DriverNetwork.jsx`.

## A note on data privacy

The repository ships with **synthetic sample data only** (generated in `buildDemo()`).

Real rosters contain personal data (names, phone numbers, IDs) and belong to your organisation — **do not commit them**. The included `.gitignore` already excludes `.csv` / `.xlsx` / `.xls` files and a `/data/` folder so you don't push a spreadsheet by accident. If this repo will ever hold anything work-specific, keep it **private**.

## Deploying

The build is a static site, so it works on GitHub Pages, Vercel, Netlify, or any static host. `vite.config.js` sets `base: "./"` so relative asset paths work from a subpath. For GitHub Pages, push the contents of `dist/` (or use a Pages action) and enable Pages in the repo settings.

## Tech

React 18 + Vite. Roster parsing via [SheetJS (`xlsx`)](https://sheetjs.com/) and [PapaParse](https://www.papaparse.com/). All styling is self-contained in an injected stylesheet — no CSS framework to configure.
