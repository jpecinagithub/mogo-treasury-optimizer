# MOGO Uganda — Treasury Optimizer

> **Eleving Group | Vite + React + Tailwind | 100% local, no server upload**

A clean, MOGO-inspired treasury web app to **upload your Excel, edit restrictions, and calculate an optimized payment calendar** — built to impress. Designed after https://www.mogo.co.ke/ (navy `#0B1E3A` + orange `#FF6B00`).

**Live features:**
- 📤 **Upload Excel** (`.xlsx` Task sheet) — auto-detects payments, rails and restrictions
- 🎛️ **Restrictions panel** — 20+ extracted constraints (liquidity, windows, limits, deadlines, rules) — fully editable, deletable, add new, Save/Load
- 🧮 **CALCULATE OPTIMIZATION** — respects all edited restrictions, never allows negative balances, protects Central Bank `13M` minimum
- 📅 **Payments in chronological order** — ranked by deadline + breach cost, with timeline `14:00 → 00:00` and RTGS/FX cutoff `16:30`
- 💱 **FX trades** — EUR→KES / USD→KES sizing & timing before cutoff
- ⏸️ **Deferred / Skipped** — grace-aware (Tax 2 days, no penalty)
- 📒 **Ledger by account** — live balances Bank A / Bank B / Wallet C (total equiv. KES)
- ✅ **Restriction Compliance Check** — 10 live checks (Vendors→B, Loans→A, no overdraft, minimum, windows, limits, deadlines)
- 📥 **Export Excel** — 4 sheets identical to `Task 1 Completed`: `Task1 Answer` + `Bank A/B/Wallet Activity`

**Stack:** Vite 5 • React 19 • TailwindCSS 3 • SheetJS `xlsx` • FontAwesome

---

## Quick start

```bash
npm install
npm run dev      # http://localhost:5173
npm run build    # production → dist/
npm run preview  # preview build
```

Upload `Financial Controller Homework - Eleving Group - Task 1.xlsx` (or your own variant with `Payment | Amount | Currency` columns) and click **CALCULATE OPTIMIZATION**.

## How it works

1. **Parse Excel** (`src/App.jsx:90`): extracts opening balances (`Bank A 170k EUR / 10M KES / 100k USD`, `Wallet C 10M`), inflows `10M/h`, FX `150/130`, windows (`RTGS 06:00–16:30`, `Wallet 22:00`), limits (`PesaLink ≤1M`, `M-Pesa ≤250k`), and 8+ business deadlines from payment notes.
2. **Edit** any restriction in the left panel (value inline, toggle, duplicate/delete, Add). `Pending → Recalculate` banner appears — old results dim until you recalculate.
3. **Engine** (`src/App.jsx:408` `calculateInternal`): FX sizing, guarded funding `min(balA, fundB)` (never negative), chronological simulation with wallet sweeps `:05`, auto-defer to protect `13M` minimum, split-aware for limits.
4. **Verify** in `Restriction Compliance Check` — all green means every rule satisfied.

## Project structure

```
src/
  App.jsx      # all UI + treasury engine (upload, restrictions, KPIs, calendar, ledgers, export)
  main.jsx
  index.css    # Tailwind + MOGO theme (uganda-flag, mogo-card)
index.html
vite.config.js
tailwind.config.js
```

## Excel format

- Sheet `Task` with spec text (optional) + table header `Payment | Amount | Currency | Notes | Banking information...`
- Supports `KES / EUR / USD`, `RTGS / PesaLink / M-Pesa / Within-bank / Tax API / Wallet API`
- If spec text missing, minimal editable defaults are created

## License

MIT — for portfolio / internal use.
