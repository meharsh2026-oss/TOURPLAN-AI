/**
 * value.js
 * ------------------------------------------------------------------
 * Pure VALUE-column business logic for TourPlan AI.
 *
 * Rule (per spec):
 *   - The sum of VALUE across all *working* days (Sundays ignored,
 *     since Sundays carry no VALUE at all) must equal MONTHLY_TARGET.
 *   - Every working day keeps the VALUE carried over from the source
 *     template EXCEPT the last working day of the month, whose VALUE
 *     is adjusted (up or down) so the monthly total lands exactly on
 *     the target.
 * ------------------------------------------------------------------
 */

const MONTHLY_TARGET = 250000;

/**
 * @param {Array<{rowIndex:number, value:number}>} workingEntries
 *        Ordered list of working-day rows (Sundays already excluded),
 *        in calendar order, each carrying the VALUE inherited from
 *        the source template.
 * @returns {{
 *   adjustedEntries: Array<{rowIndex:number, value:number}>,
 *   originalTotal: number,
 *   finalTotal: number,
 *   adjustment: number,
 *   lastWorkingRowIndex: number|null
 * }}
 */
function applyMonthlyValueRule(workingEntries) {
  if (!Array.isArray(workingEntries) || workingEntries.length === 0) {
    return {
      adjustedEntries: [],
      originalTotal: 0,
      finalTotal: 0,
      adjustment: 0,
      lastWorkingRowIndex: null
    };
  }

  const originalTotal = workingEntries.reduce((sum, e) => sum + (Number(e.value) || 0), 0);

  // Clone so we never mutate the caller's array in place.
  const adjustedEntries = workingEntries.map(e => ({ ...e }));

  const lastIdx = adjustedEntries.length - 1;
  const lastEntry = adjustedEntries[lastIdx];

  const sumExcludingLast = originalTotal - (Number(lastEntry.value) || 0);
  const newLastValue = MONTHLY_TARGET - sumExcludingLast;

  const adjustment = newLastValue - (Number(lastEntry.value) || 0);
  lastEntry.value = newLastValue;

  const finalTotal = adjustedEntries.reduce((sum, e) => sum + (Number(e.value) || 0), 0);

  return {
    adjustedEntries,
    originalTotal,
    finalTotal,
    adjustment,
    lastWorkingRowIndex: lastEntry.rowIndex
  };
}

export { MONTHLY_TARGET, applyMonthlyValueRule };
