# Programme Dashboard (Phase 3)

The `DASHBOARD` sheet in each programme workbook is built from **PivotTables
and PivotCharts sourced from the tables `powerquery/queries.pq` loads** —
not from more Power Query. This is deliberate: reshaping and joining data
is Power Query's job (Phase 1–2); slicing and viewing it is a PivotTable's
job, and it's much faster to try five different views of the same table
with PivotTables than to write five different M queries.

This doc is the recipe for adding a new dashboard item without needing
another round of M code — plus what's already built as worked examples.

## The building blocks (source tables)

Everything below is a query already loaded from `powerquery/queries.pq`.
Swap `FS_` for another programme's code once that programme's queries
exist (`PE_`, `SES_`, `COACH_`, `NUTR_`, `PSYC_`).

| Table | One row per | Good for |
|---|---|---|
| `FS_ModuleOverview` | module | module facts, sorting/filtering, Hours per Credit outliers |
| `FS_Assess` | assessment | assessment load, weighting, due-week clustering |
| `FS_LO` | learning outcome | LO lists, LO coverage |
| `FS_Mapping` | (assessment, LO, skill) link | cross-referencing assessments ↔ LOs ↔ graduate skills |
| `FS_SkillsCoverage` | catalogue skill | gap analysis — `Covered` TRUE/FALSE, which modules cover it |
| `FS_Meta_Wide` / `FS_Hours_Wide` | module | raw wide detail behind `ModuleOverview`, if you need a Meta or Hours field it doesn't carry |

If the question you want to answer needs a field that isn't in any of
these tables (a different item type, e.g. `Weekly`), that's a Phase 2
gap, not something a PivotTable can work around — see `powerquery/README.md`
("Item scope") for adding an item back in.

## The general recipe

1. **Name the question first**, not the mechanics — "which modules are
   lecture-heavy?", not "I need a bar chart." The question tells you
   which source table has the right granularity (a module-level question
   → `ModuleOverview`; a per-assessment question → `Assess`).
2. Click any cell inside the source table (or select it by name in the
   Name Box) → **Insert → PivotTable** (or **PivotChart** if you want the
   visual directly).
3. **Existing Worksheet** → `DASHBOARD` → pick an empty cell with room to
   grow below/right of it. Pivots resize themselves as data changes, so
   leave a few blank rows/columns as a buffer against overlap.
4. Drag fields into **Rows** / **Columns** / **Values** / **Filters** in
   the PivotTable Fields pane. Values defaults to **Count** for text
   fields and **Sum** for numbers — check it picked the one you want
   (right-click the Values field → **Value Field Settings**).
5. Format: percentages for weightings (`Value Field Settings → Number
   Format → Percentage`), conditional formatting for anything you want to
   spot at a glance (see below).
6. **Data → Refresh All** re-runs the Power Query queries *and* updates
   every PivotTable from the new results, in that order. Right-clicking a
   single PivotTable → **Refresh** only re-reads the already-loaded table
   — use that for a quick re-layout, use Refresh All after editing a
   query or when the source workbooks have changed.

## Worked examples already on the sheet

**Assessment Map** (from `FS_Assess`):
- Rows: `ModuleTitle`
- Columns: `Due Week`
- Values: Sum of `Weighting`, formatted `%`
- Reads as a module × week grid — a filled column is a clustered week.

**Module Overview** (from `FS_ModuleOverview`, loaded directly, not a
pivot — it's already one row per module, nothing to aggregate):
- Placed at the top of `DASHBOARD` as the sheet's spine; everything else
  either drills into it or sits below it.

**Graduate Skills Coverage** (from `FS_SkillsCoverage`, also loaded
directly):
- Conditional formatting rule on the `Covered` column: **Home →
  Conditional Formatting → New Rule → Format only cells that contain →
  Cell Value → equal to → FALSE** → red fill. Gaps become visible without
  reading the column.

**Suggested, not yet built** (no new queries needed — same recipe):
- Assessment Type balance: PivotChart from `FS_Assess`, Rows =
  `Assessment Type`, Values = Count of `Assessment No`.
- Hours by component: bar chart straight from `FS_Hours_Wide`'s columns
  (`Lecture`, `Seminar`, `Practical`, …) per module.

## Tips & gotchas

- **Text fields won't Sum.** If a Values field defaults to Count and you
  expected Sum, the underlying column probably isn't numerically typed in
  the query — fix the type in Power Query (`Table.TransformColumnTypes`),
  don't work around it with a pivot calculation.
- **Week/number columns can sort like text.** If `Due Week` sorts
  1, 10, 11, 2, 3… instead of 1, 2, 3…, right-click the column header row
  → **Sort → Sort by value** rather than trusting the default.
- **Keep the raw query tables hidden**, not deleted. A PivotTable reads a
  hidden sheet's table exactly as if it were visible — hiding is purely
  cosmetic and doesn't need re-pointing anything.
- **A PivotTable is a snapshot**, not a live filter — it only reflects a
  query's output as of the last refresh. If a number looks stale, refresh
  before assuming the data's wrong.
- **One tab per item is fine.** There's no requirement to cram everything
  onto one physical view — a second sheet (`DASHBOARD 2`, or split by
  topic) is a reasonable answer to "this is getting crowded," not a sign
  something went wrong.

## Extending to another programme

Once `<ProgCode>_ModuleOverview` / `_Assess` / `_SkillsCoverage` etc. exist
for another programme's workbook (same queries, different `ProgCode`
prefix, per `powerquery/README.md`), build its `DASHBOARD` sheet the same
way in that workbook — each workbook keeps its own copy, since Power
Query doesn't share queries or PivotTables across separate files. Once a
couple of programmes' dashboards look right, copying the layout (not the
data) between them is largely mechanical.
