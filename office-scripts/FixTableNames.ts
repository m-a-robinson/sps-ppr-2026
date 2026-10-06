/**
 * Office Script: make every table on a module sheet match its sheet name.
 *
 * Tables on sheet 4NX1SPS should be called _4NX1SPS_Meta, _4NX1SPS_LO,
 * _4NX1SPS_Assess and so on. This script finds tables whose names don't
 * match their sheet (after a module code change, a copied sheet, or a
 * template that was never renamed) and renames them.
 *
 * How to use (Excel on the web or Microsoft 365 desktop: Automate > New Script):
 *   1. Paste this script in and run it with DRY_RUN = true. Nothing is
 *      changed; the output lists every rename it would make.
 *   2. If the list looks right, set DRY_RUN = false and run it again.
 *   3. Repeat for each programme workbook.
 *
 * Safeguards:
 *   - Only sheets named like 4001SPS or 4NX1SPS are checked.
 *   - Tables that already have the right name are left alone.
 *   - A table is not renamed if its new name is already used anywhere in the
 *     workbook (no duplicates); the output tells you which ones to check.
 *   - Protected sheets are unprotected only for the rename and then protected
 *     again with their original settings. If a sheet can't be unprotected,
 *     nothing at all is renamed.
 */

const DRY_RUN = true;
const PASSWORD = ""; // sheet protection password, if the sheets have one
const MODULE_SHEET = /^[0-9][0-9A-Z]{3}SPS$/i;
const ITEMS = ["Meta", "LO", "Assess", "Weekly", "Hours", "Mapping", "Aims", "Syllabus", "Overview", "Notes"];

interface Rename {
  sheet: ExcelScript.Worksheet;
  sheetName: string;
  table: ExcelScript.Table;
  id: string;
  from: string;
  to: string;
  temp: string;
}

interface Unprotected {
  sheet: ExcelScript.Worksheet;
  options: ExcelScript.WorksheetProtectionOptions;
}

function main(workbook: ExcelScript.Workbook) {
  const notes: string[] = [];
  const plan: Rename[] = [];
  let correct = 0;

  // Table names and defined names share one namespace, so collect both.
  const taken = new Set<string>();
  workbook.getTables().forEach((t) => taken.add(t.getName().toLowerCase()));
  workbook.getNames().forEach((n) => taken.add(n.getName().toLowerCase()));

  for (const sheet of workbook.getWorksheets()) {
    const sheetName = sheet.getName();
    const tables = sheet.getTables();
    if (!MODULE_SHEET.test(sheetName)) {
      if (/SPS/i.test(sheetName) && tables.length > 0) {
        notes.push(`Sheet "${sheetName}" isn't in the 4001SPS format, so its tables were not checked.`);
      }
      continue;
    }
    for (const table of tables) {
      const name = table.getName();
      const body = name.startsWith("_") ? name.substring(1) : name;
      const cut = body.lastIndexOf("_");
      // "Meta2" or "LO4" come from copied tables: drop the trailing number.
      const suffix = cut >= 0 ? body.substring(cut + 1).replace(/\d+$/, "") : "";
      const item = ITEMS.find((i) => i.toLowerCase() === suffix.toLowerCase());
      if (!item) {
        notes.push(`${sheetName}: "${name}" isn't a proforma table, left alone.`);
        continue;
      }
      if (item === "Meta") checkMetaCode(table, sheetName, notes);
      const target = `_${sheetName}_${item}`;
      const id = table.getId();
      // Already correct if Excel resolves the target name to this same table
      // (these workbooks also hold an internal name without the underscore).
      const holder = name === target ? table : workbook.getTable(target);
      if (holder && holder.getId() === id) {
        correct++;
        continue;
      }
      plan.push({ sheet, sheetName, table, id, from: name, to: target, temp: "" });
    }
  }

  // Drop any rename whose new name is already in use, or already claimed by
  // another rename. Repeat until stable, because a dropped rename keeps its
  // old name, which another rename may have been counting on being freed.
  let changed = true;
  while (changed) {
    changed = false;
    const leaving = new Set<string>(plan.map((p) => p.from.toLowerCase()));
    const claimed = new Set<string>();
    for (let i = 0; i < plan.length; i++) {
      const key = plan[i].to.toLowerCase();
      if ((taken.has(key) && !leaving.has(key)) || claimed.has(key)) {
        notes.push(`${plan[i].sheetName}: "${plan[i].from}" NOT renamed, because "${plan[i].to}" already exists. Check which table holds the right data.`);
        plan.splice(i, 1);
        changed = true;
        break;
      }
      claimed.add(key);
    }
  }

  notes.forEach((n) => console.log(n));
  plan.forEach((p) => console.log(`${p.sheetName}: "${p.from}" -> "${p.to}"`));

  if (plan.length === 0) {
    console.log(`Nothing to rename. ${correct} tables already match their sheet.`);
    return;
  }
  if (DRY_RUN) {
    console.log(`DRY RUN: ${plan.length} tables would be renamed, ${correct} already match. Set DRY_RUN = false and run again to apply.`);
    return;
  }

  // Unprotect only the sheets that need a change, remembering their settings.
  const unprotected: Unprotected[] = [];
  const failed: string[] = [];
  const sheetsToChange = new Map<string, ExcelScript.Worksheet>();
  plan.forEach((p) => sheetsToChange.set(p.sheetName, p.sheet));
  sheetsToChange.forEach((sheet, sheetName) => {
    const protection = sheet.getProtection();
    if (!protection.getProtected()) return;
    const options = protection.getOptions();
    try {
      protection.unprotect(PASSWORD);
      if (protection.getProtected()) throw new Error("still protected");
      unprotected.push({ sheet, options });
    } catch (error) {
      failed.push(sheetName);
    }
  });

  const reprotect = () =>
    unprotected.forEach((u) => {
      if (PASSWORD) u.sheet.getProtection().protect(u.options, PASSWORD);
      else u.sheet.getProtection().protect(u.options);
    });

  if (failed.length > 0) {
    reprotect();
    console.log(`STOPPED: couldn't unprotect ${failed.join(", ")}. Put the sheet password in PASSWORD at the top and run again. Nothing was renamed.`);
    return;
  }

  try {
    // Step 1: move each table to a temporary name, so codes swapped between
    // two sheets can't collide halfway through.
    plan.forEach((p, i) => {
      let temp = `_RENAME${i}_${p.to.substring(1)}`;
      while (taken.has(temp.toLowerCase())) temp += "X";
      p.table.setName(temp);
      p.temp = temp;
    });
    // Step 2: give each table its final name.
    plan.forEach((p) => {
      const table = workbook.getTable(p.temp);
      if (table) table.setName(p.to);
    });
  } finally {
    reprotect();
  }

  plan.forEach((p) => {
    const table = workbook.getTable(p.to);
    if (!table || table.getId() !== p.id) console.log(`CHECK: "${p.to}" on ${p.sheetName} didn't take. Rename it by hand.`);
  });

  console.log(`Done: ${plan.length} tables renamed, ${correct} already matched. ${unprotected.length} sheet(s) were unprotected and protected again.`);
}

/** Report (but don't change) a Module Code in the Meta table that differs from the sheet name. */
function checkMetaCode(table: ExcelScript.Table, sheetName: string, notes: string[]) {
  const body = table.getRangeBetweenHeaderAndTotal();
  if (!body) return;
  const rows = body.getValues();
  for (const row of rows) {
    if (String(row[0]).trim().toLowerCase() === "module code") {
      const code = String(row[1]).trim();
      if (code !== "" && code.toUpperCase() !== sheetName.toUpperCase()) {
        notes.push(`${sheetName}: the Meta table still says Module Code "${code}". Update it by hand.`);
      }
      return;
    }
  }
}
