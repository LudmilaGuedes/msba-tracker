# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

## Repository layout

This repo is not a single app — it's a small research-to-dashboard pipeline for an international student's MSBA (Master of Science in Business Analytics) program search:

- `data/msba_programs_database.xlsx` — the source-of-truth dataset (138 programs × 55 columns, sheet `MSBA Programs`), produced by manual research rounds described in `resumo_contexto_tecnico.md`.
- `resumo_contexto_tecnico.md` — project brief: schema definition, data-collection methodology, and phase history. Read this before changing the data schema.
- `msba-dashboard/` — the actual React app (Phase 3). **All commands below run from inside this directory.**

## Commands (run inside `msba-dashboard/`)

```
npm run dev       # start Vite dev server on port 5173 (fixed via strictPort — see note below)
npm run build     # tsc -b (typecheck) && vite build — this is the closest thing to a test suite; no unit tests exist
npm run preview   # preview the production build
npm run lint      # oxlint
```

**Port 5173 is pinned on purpose** (`server.port` + `strictPort` in `vite.config.ts`). All user data (ratings, comments, folders, custom programs) lives in localStorage, which is per-origin *including the port* — if the app ever runs on another port, the user's data appears to vanish. If 5173 is busy, find and stop the stale process on it instead of changing the port.

There is no test framework configured — validate changes via `npm run build` (typecheck) and manual verification in the browser.

**Node/npm PATH note (this machine only):** Node is installed at `C:\Program Files\nodejs\` but that directory is not on the system/user `PATH`. If `node`/`npm` aren't found, either prefix the session PATH (`export PATH="/c/Program Files/nodejs:$PATH"` in bash) or call the full path (`& "C:\Program Files\nodejs\npm.cmd" ...` in PowerShell).

## Architecture

### The schema is the single source of truth

`src/types/columns.ts` defines all 55 program fields as a `COLUMNS` array: each entry maps the original Portuguese spreadsheet header to a camelCase `key`, a `type` (`text` | `longtext` | `number` | `percent` | `currency` | `enum`), a `group` (one of the 7 categories from `resumo_contexto_tecnico.md`), and — for `enum` fields — the allowed `options`. Every other part of the app derives from this array instead of hardcoding field lists:

- `utils/dataParser.ts` maps spreadsheet headers → keys using `COLUMNS`.
- `components/table/DataTable.tsx`, `components/ProgramDetailModal.tsx`, and `components/AddProgramModal.tsx` all render their fields by iterating `COLUMNS`/`COLUMN_GROUPS`.
- `utils/exportData.ts` writes spreadsheet columns in `COLUMNS` order.

**When adding/renaming/removing a data field, edit `columns.ts` only** — do not hardcode new field lists elsewhere.

Three fields are schema-constrained enums rendered as `<select>` everywhere they're edited (inline table cells, the detail modal, and the "add program" form): `estado` (US state codes from `types/usStates.ts`), `tipo` (Pública/Privada), and `stemDesignated` (Sim/Não). This is deliberate — free-text entry in these fields would silently break the state/type filters.

### Data flow: three sources merged at runtime, nothing mutates the original file

1. **Original rows** — `utils/dataParser.ts` fetches `public/data/msba_programs_database.xlsx` (a copy of `data/msba_programs_database.xlsx`) via SheetJS at runtime (not at build time) and parses it into `RawProgram[]`. Program IDs come from `utils/programId.ts`: `prog:<nome>|<estado>|<cidade>` (normalized), so deleting or reordering rows no longer affects user data. **Renaming** a row's `Nome da Universidade`/`Estado`/`Cidade` changes its id and orphans its ratings/comments — avoid that, or migrate the user's localStorage keys along with it. (IDs used to be the row position `p{index}`; deleting Michigan State and Virginia Tech on 2026-10-01 shifted every later comment onto the wrong college. `utils/legacyIdMap.ts` is the one-time migration from those old ids and must not be edited by hand.)
2. **Manually-added rows** — programs created via "Adicionar Faculdade" (`AddProgramModal.tsx`) are stored as full `RawProgram` objects (id prefix `custom_...`) in their own localStorage bucket, kept entirely separate from the original file.
3. **Per-program user overlay** — ratings, comments, folder assignments, and inline-edited cell values are *not* stored on the program itself; they live in a separate `userData` map keyed by program id, in localStorage.

`context/DataContext.tsx` combines all three into `MergedProgram[]` (`mergeProgram()`) — this is the only shape UI components consume. Never read `RawProgram` or the localStorage overlay directly from a component; go through `useData()`.

localStorage keys (all versioned with a `:v1` suffix so a future breaking schema change can migrate): `msba-tracker:user-data:v1`, `msba-tracker:folders:v1`, `msba-tracker:custom-programs:v1`.

### The "Verificar" sentinel

Missing data is never blank or null — it's the literal string `"Verificar"` (Portuguese for "verify/check"), inherited from the original research spreadsheet's convention for unconfirmed data. `utils/format.ts` (`isVerificar()`, `parseNumeric()`, `formatCellValue()`) centralizes handling of this convention for display, sorting, and filtering. When a form field is left blank (including an unselected `<select>`), it's coerced to `"Verificar"` on save (see `addCustomProgram` in `DataContext.tsx`). Any new numeric/display logic must treat `"Verificar"` as "unknown", not as `0` or `NaN`.

### Export

`utils/exportData.ts` exports whatever `MergedProgram[]` it's given — in practice, the *currently filtered* view passed down from `App.tsx`, not unconditionally the full dataset. It writes the 55 schema columns in `COLUMNS` order, then appends three extra columns (`Minha Nota (0-5)`, `Meus Comentários`, `Minhas Pastas`) that only exist in the export, not in the schema.

### Data enrichment scripts (`scripts/`)

Per-university "Data Enrichment" requests (audit one row against fresh web research, correct/fill fields, leave the rest untouched) each get their own one-off `scripts/etl_enrich_<school>.py`, matched by `Nome da Universidade` + `Estado` + `Cidade` (never by row number, and never a plain name match when two rows could share a name — e.g. `UCF` vs `UCF (Professional)`, `Miami U Ohio` vs `University of Miami`). Broader `scripts/etl_add_<state>_programs.py` scripts append new rows for a state-wide sweep. Both kinds only ever touch `data/msba_programs_database.xlsx` (run them from the repo root), then the caller copies the result into `msba-dashboard/public/data/msba_programs_database.xlsx` so the dashboard picks it up.

**Backup retention policy:** `scripts/db_retention.py` provides `rotate_backup()` — call it immediately before opening/overwriting `msba_programs_database.xlsx`. It keeps exactly one backup, at the fixed path `data/msba_programs_database_backup.xlsx` (gitignored) (Option A: fixed names, not timestamped — the dashboard reads a static path, so the live filename must never change). Do not hand-roll a new `shutil.copy2(..., f"...backup-{timestamp}...")` per script; that was the old pattern and it left 20+ stray timestamped backups in the repo root before this policy existed. Any new script that saves the database should do:

```python
from db_retention import rotate_backup, DB_PATH
rotate_backup()
wb = openpyxl.load_workbook(DB_PATH)
# ...
wb.save(DB_PATH)
```

### "Deadline Aplicação": earliest date first

The "Estratégia de Aplicação" tab and the Round 1/2/3 badge (`utils/strategyParsing.ts` → `classifyLeadTimeBucket`) classify each program by the **first month that appears** in `Deadline Aplicação`, using fixed calendar windows that are the same for every program, whatever its start month: Round 1 = Sep–Dec, Round 2 = Jan–May, Round 3 = Jun–Aug. The rounds are a planning tool for *when to apply*, not a lead-time measure. The user's rule: always classify by the **earliest possible deadline** of the main intake (Round 1, Early Action/Decision, scholarship/assistantship priority date), since applying first gives the best admission and funding odds. So when writing this field in any enrichment script:

- Start the text with that earliest deadline. List international before domestic when the international one is earlier, and the main intake (usually Fall) before Spring/Summer.
- Don't put any other month before it: no start month like "Fall 2027 (agosto)", and no "aplicações abrem 1 set" (an opening date is not a deadline; put it at the end).
- Use `Round 1: <data>; Round 2: <data>` with arabic numerals, since `parseRounds` doesn't understand "Round I".

### Dependency note

`xlsx` is intentionally installed from SheetJS's own CDN (`https://cdn.sheetjs.com/xlsx-0.20.3/xlsx-0.20.3.tgz`) in `package.json`, not the npm registry package — the registry version has two unpatched CVEs (prototype pollution, ReDoS). Don't `npm install xlsx` from the registry to "fix" or update it.
