/**
 * excel.js
 * ------------------------------------------------------------------
 * Workbook intelligence layer for TourPlan AI.
 *
 * Responsibilities:
 *   1. Parse the uploaded .xlsx into an ExcelJS workbook object.
 *   2. AUTO-DETECT the tour-plan structure (header row, column order,
 *      data block boundaries, total-formula row, year/month labels)
 *      by reading the workbook itself — nothing about row/column
 *      position is hardcoded, so this works for any tour-plan
 *      workbook that follows the "Date / Day / Town Name / Work
 *      With / DB NAME / TC / PC / VALUE" header convention.
 *   3. Extract the existing rows as a reusable "template entry" pool
 *      (Town Name, Work With, DB NAME, TC, PC, VALUE).
 *   4. Rewrite the workbook in place for a target (year, month):
 *      dates, weekdays, Sunday-clearing, and the 250000 VALUE rule.
 *   5. Serialize the mutated workbook back to a downloadable file.
 *
 * Uses the global `ExcelJS` object (loaded from the ExcelJS browser
 * bundle in index.html). Unlike SheetJS-family libraries, ExcelJS
 * keeps the full OOXML document model (styles, merges, borders,
 * fonts, fills, row heights, column widths, hidden rows/columns,
 * print settings, ...) in memory and only changes what you actually
 * touch. Because every write in this file sets ONLY `cell.value`
 * (never `cell.style`), every other property of every cell — and
 * every workbook-level setting we never reference — survives the
 * round trip untouched. This is what fixes the formatting loss that
 * happened with xlsx-js-style.
 *
 * NOTE: ExcelJS is 1-indexed (row 1 / column 1 is the first row/
 * column), unlike SheetJS's 0-indexed { r, c } pairs. All row/column
 * numbers in this file are 1-based.
 * ------------------------------------------------------------------
 */

import { MONTH_ABBR, MONTH_FULL, generateMonthDates } from './planner.js';
import { applyMonthlyValueRule } from './value.js';

/** Normalizes a header label for fuzzy matching: collapses whitespace, lowercases. */
function norm(s) {
  return String(s == null ? '' : s).replace(/\s+/g, ' ').trim().toLowerCase();
}

/**
 * Resolves the "plain" value of an ExcelJS cell, unwrapping the rich
 * object shapes ExcelJS uses for formulas, hyperlinks and rich text.
 * Returns `undefined` for a genuinely empty cell.
 */
function cellValue(cell) {
  if (!cell) return undefined;
  const v = cell.value;
  if (v === null || v === undefined) return undefined;
  if (v instanceof Date) return v;
  if (typeof v === 'object') {
    if (Array.isArray(v.richText)) return v.richText.map(part => part.text).join('');
    if (v.result !== undefined) return v.result; // formula cell: { formula, result }
    if (v.text !== undefined) return v.text;       // hyperlink cell: { text, hyperlink }
    return undefined;
  }
  return v;
}

/** Reads a raw cell value from a worksheet at 1-based (row, col), or undefined if empty. */
function readCell(ws, r, c) {
  return cellValue(ws.getCell(r, c));
}

/**
 * Parses an ArrayBuffer into an ExcelJS workbook.
 * Throws a friendly Error if the file is not a readable workbook.
 */
async function parseWorkbook(arrayBuffer) {
  const workbook = new ExcelJS.Workbook();
  try {
    await workbook.xlsx.load(arrayBuffer);
  } catch (err) {
    throw new Error('This file could not be read as an Excel workbook. Please upload a valid .xlsx file.');
  }
  if (!workbook.worksheets || workbook.worksheets.length === 0) {
    throw new Error('The uploaded workbook does not contain any sheets.');
  }
  return workbook;
}

/**
 * Scans a single sheet for the "Date | Day | ..." header row.
 * Returns null if this sheet doesn't look like a tour-plan sheet.
 */
function detectHeaderInSheet(ws) {
  const rowCount = ws.rowCount;
  const colCount = ws.columnCount;
  if (!rowCount || !colCount) return null;

  for (let r = 1; r <= rowCount; r++) {
    for (let c = 1; c <= colCount; c++) {
      const v = readCell(ws, r, c);
      if (norm(v) === 'date') {
        // Look a few columns ahead (allowing for the occasional gap) for "day"
        for (let c2 = c + 1; c2 <= Math.min(c + 3, colCount); c2++) {
          if (norm(readCell(ws, r, c2)) === 'day') {
            return { headerRow: r, dateCol: c, dayCol: c2 };
          }
        }
      }
    }
  }
  return null;
}

/** Known header label -> internal field name, for fuzzy column mapping. */
const HEADER_ALIASES = [
  { field: 'date', test: v => norm(v) === 'date' },
  { field: 'day', test: v => norm(v) === 'day' },
  { field: 'town', test: v => norm(v).includes('town') },
  { field: 'workWith', test: v => norm(v).includes('work with') },
  { field: 'db', test: v => norm(v).includes('db name') || norm(v) === 'db' },
  { field: 'tc', test: v => norm(v) === 'tc' },
  { field: 'pc', test: v => norm(v) === 'pc' },
  { field: 'value', test: v => norm(v) === 'value' }
];

/**
 * Builds the column map for a sheet given its detected header row,
 * falling back to the canonical A..H offsets relative to the Date
 * column whenever a label can't be confidently matched.
 */
function buildColumnMap(ws, headerRow, dateCol, colCount) {
  const fallbackOffsets = { date: 0, day: 1, town: 2, workWith: 3, db: 4, tc: 5, pc: 6, value: 7 };
  const map = {};

  for (let c = dateCol; c <= Math.min(dateCol + 7, colCount); c++) {
    const v = readCell(ws, headerRow, c);
    const hit = HEADER_ALIASES.find(a => a.test(v));
    if (hit && map[hit.field] === undefined) {
      map[hit.field] = c;
    }
  }

  Object.keys(fallbackOffsets).forEach(field => {
    if (map[field] === undefined) {
      map[field] = dateCol + fallbackOffsets[field];
    }
  });

  return map;
}

/**
 * Locates the row holding the SUM(...) total formula in the VALUE column,
 * searching downward from the first data row.
 */
function findTotalRow(ws, valueCol, dataStartRow, rowCount) {
  for (let r = dataStartRow; r <= rowCount; r++) {
    const cell = ws.getCell(r, valueCol);
    const formula = cell && cell.formula;
    if (typeof formula === 'string' && /sum/i.test(formula)) {
      return r;
    }
  }
  return -1;
}

/**
 * Searches the rows above the header for a 4-digit year cell and a
 * month-label cell (e.g. " JUN", "June"). Returns address + parsing
 * metadata needed to reconstruct the label later in the same style.
 */
function findYearMonthCells(ws, headerRow, colCount) {
  let yearInfo = null;
  let monthInfo = null;

  for (let r = 1; r < headerRow; r++) {
    for (let c = 1; c <= colCount; c++) {
      const v = readCell(ws, r, c);
      if (v === undefined || v === null || v === '') continue;

      if (!yearInfo && typeof v === 'number' && Number.isInteger(v) && v >= 1900 && v <= 2100) {
        yearInfo = { r, c, value: v };
      }

      if (!monthInfo && typeof v === 'string') {
        const upper = v.toUpperCase();
        const found = MONTH_ABBR.find(abbr => upper.includes(abbr));
        if (found) {
          const idx = upper.indexOf(found);
          monthInfo = {
            r, c,
            prefix: v.slice(0, idx),
            token: v.slice(idx, idx + 3),
            suffix: v.slice(idx + 3),
            monthIndex: MONTH_ABBR.indexOf(found)
          };
        }
      }
    }
  }

  return { yearInfo, monthInfo };
}

/**
 * Best-effort search for a "Name-" style label above the header row,
 * returning the adjacent value purely for display in the UI (it is
 * never written to). Returns null if nothing matches.
 */
function findPersonName(ws, headerRow, colCount) {
  for (let r = 1; r < headerRow; r++) {
    for (let c = 1; c <= colCount; c++) {
      const v = readCell(ws, r, c);
      const label = norm(v).replace(/[-:]+$/, '');
      if (label === 'name') {
        for (let c2 = c + 1; c2 <= Math.min(c + 4, colCount); c2++) {
          const candidate = readCell(ws, r, c2);
          if (candidate !== undefined && String(candidate).trim() !== '') {
            return String(candidate).trim();
          }
        }
      }
    }
  }
  return null;
}

/**
 * Full structure detection for an uploaded workbook.
 * Tries every sheet until one matches the tour-plan header convention.
 */
function detectStructure(workbook) {
  for (const ws of workbook.worksheets) {
    const headerHit = detectHeaderInSheet(ws);
    if (!headerHit) continue;

    const rowCount = ws.rowCount;
    const colCount = ws.columnCount;
    const colMap = buildColumnMap(ws, headerHit.headerRow, headerHit.dateCol, colCount);
    const dataStartRow = headerHit.headerRow + 1;
    const totalRow = findTotalRow(ws, colMap.value, dataStartRow, rowCount);

    // Convention observed in real-world templates: one blank separator
    // row sits between the last data row and the total-formula row.
    let maxDataRows;
    if (totalRow > dataStartRow) {
      const rawGap = totalRow - dataStartRow;
      maxDataRows = rawGap - 1;
    } else {
      // No SUM formula found — fall back to a generous 31-row block.
      maxDataRows = Math.max(31, rowCount - dataStartRow + 1);
    }
    if (maxDataRows < 28) maxDataRows = 28;

    const { yearInfo, monthInfo } = findYearMonthCells(ws, headerHit.headerRow, colCount);
    const personName = findPersonName(ws, headerHit.headerRow, colCount);

    return {
      sheetName: ws.name,
      headerRow: headerHit.headerRow,
      dataStartRow,
      maxDataRows,
      totalRow: totalRow > 0 ? totalRow : null,
      colMap,
      yearInfo,
      monthInfo,
      personName
    };
  }

  throw new Error(
    'Could not find a recognizable Tour Plan layout (a "Date / Day / Town Name ... VALUE" header row) in this workbook.'
  );
}

/**
 * Reads the existing data rows and extracts the reusable "template
 * entries" — the Town Name / Work With / DB NAME / TC / PC / VALUE
 * combination for every row that represents an actual working day
 * (i.e. not an already-blank Sunday row). Order is preserved.
 */
function extractTemplateEntries(workbook, structure) {
  const ws = workbook.getWorksheet(structure.sheetName);
  const { colMap, dataStartRow, maxDataRows } = structure;
  const entries = [];

  for (let i = 0; i < maxDataRows; i++) {
    const r = dataStartRow + i;
    const town = readCell(ws, r, colMap.town);
    const value = readCell(ws, r, colMap.value);
    const isBlankRow = (town === undefined || String(town).trim() === '') && (value === undefined || value === '');

    if (isBlankRow) continue;

    entries.push({
      town: readCell(ws, r, colMap.town),
      workWith: readCell(ws, r, colMap.workWith),
      db: readCell(ws, r, colMap.db),
      tc: readCell(ws, r, colMap.tc),
      pc: readCell(ws, r, colMap.pc),
      value: Number(readCell(ws, r, colMap.value)) || 0,
      sourceRow: r
    });
  }

  if (entries.length === 0) {
    throw new Error('No working-day rows were found in the uploaded workbook to use as a template.');
  }

  return entries;
}

/**
 * Writes a value into a cell. Only `cell.value` is touched — font,
 * fill, border, alignment, number format and every other style
 * property already on the cell (inherited from the template) is left
 * completely alone, so formatting is preserved automatically.
 */
function setCell(ws, r, c, value) {
  const cell = ws.getCell(r, c);
  cell.value = (value === undefined || value === null || value === '') ? null : value;
}

/** Fully blanks a cell's value/formula while keeping its style intact. */
function clearCell(ws, r, c) {
  ws.getCell(r, c).value = null;
}

/**
 * Rewrites the detected sheet in place for the target (year, monthIndex).
 * Returns a summary object describing what was done, for the UI to display.
 */
function applyTourPlan(workbook, structure, templateEntries, year, monthIndex) {
  const ws = workbook.getWorksheet(structure.sheetName);
  const { colMap, dataStartRow, maxDataRows, yearInfo, monthInfo } = structure;

  const dates = generateMonthDates(year, monthIndex);
  const truncatedWarning = dates.length > maxDataRows;
  const usableDates = dates.slice(0, maxDataRows);

  // --- Update the year / month header labels, preserving their original style ---
  if (yearInfo) {
    setCell(ws, yearInfo.r, yearInfo.c, year);
  }
  if (monthInfo) {
    const sameCase = monthInfo.token === monthInfo.token.toUpperCase();
    const newToken = sameCase ? MONTH_ABBR[monthIndex] : (MONTH_ABBR[monthIndex][0] + MONTH_ABBR[monthIndex].slice(1).toLowerCase());
    setCell(ws, monthInfo.r, monthInfo.c, `${monthInfo.prefix}${newToken}${monthInfo.suffix}`);
  }

  // --- Walk every row slot in the data block ---
  let entryCursor = 0;
  const workingRows = []; // { rowIndex, value } in calendar order, for the VALUE rule

  for (let i = 0; i < maxDataRows; i++) {
    const r = dataStartRow + i;

    if (i >= usableDates.length) {
      // Template has more row slots than this month has days (e.g. 31-row
      // template, 30-day month) — clear the whole row, mirroring how the
      // source workbook itself leaves trailing rows blank.
      [colMap.date, colMap.day, colMap.town, colMap.workWith, colMap.db, colMap.tc, colMap.pc, colMap.value]
        .forEach(c => clearCell(ws, r, c));
      continue;
    }

    const d = usableDates[i];
    setCell(ws, r, colMap.date, d.jsDate);
    setCell(ws, r, colMap.day, d.dayName);

    if (d.isSunday) {
      setCell(ws, r, colMap.town, ' '); // matches the source convention for off-days
      clearCell(ws, r, colMap.workWith);
      clearCell(ws, r, colMap.db);
      clearCell(ws, r, colMap.tc);
      clearCell(ws, r, colMap.pc);
      clearCell(ws, r, colMap.value);
    } else {
      const entry = templateEntries[entryCursor % templateEntries.length];
      entryCursor++;

      setCell(ws, r, colMap.town, entry.town);
      setCell(ws, r, colMap.workWith, entry.workWith);
      setCell(ws, r, colMap.db, entry.db);
      setCell(ws, r, colMap.tc, entry.tc);
      setCell(ws, r, colMap.pc, entry.pc);
      setCell(ws, r, colMap.value, entry.value);

      workingRows.push({ rowIndex: r, value: entry.value });
    }
  }

  // --- Enforce the 250000 monthly VALUE rule on the last working day only ---
  const valueResult = applyMonthlyValueRule(workingRows);
  if (valueResult.lastWorkingRowIndex !== null) {
    setCell(ws, valueResult.lastWorkingRowIndex, colMap.value,
      valueResult.adjustedEntries[valueResult.adjustedEntries.length - 1].value);
  }

  // Ask Excel to recompute the SUM(...) total formula when the file is opened
  // (ExcelJS does not evaluate formulas itself, so the cached result on the
  // total cell is left as-is and Excel recalculates it on open).
  workbook.calcProperties = workbook.calcProperties || {};
  workbook.calcProperties.fullCalcOnLoad = true;

  return {
    daysInTargetMonth: dates.length,
    rowsWritten: usableDates.length,
    workingDays: workingRows.length,
    sundays: usableDates.filter(d => d.isSunday).length,
    originalTotal: valueResult.originalTotal,
    finalTotal: valueResult.finalTotal,
    adjustment: valueResult.adjustment,
    truncatedWarning,
    monthLabel: `${MONTH_FULL[monthIndex]} ${year}`
  };
}

/** Serializes the (mutated) workbook back into a downloadable Blob. */
async function exportWorkbook(workbook) {
  const buffer = await workbook.xlsx.writeBuffer();
  return new Blob([buffer], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' });
}

export {
  parseWorkbook,
  detectStructure,
  extractTemplateEntries,
  applyTourPlan,
  exportWorkbook
};
