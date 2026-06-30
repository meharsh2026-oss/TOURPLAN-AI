/**
 * planner.js
 * ------------------------------------------------------------------
 * Pure calendar logic for TourPlan AI.
 * No DOM access, no Excel access — just dates, weekdays and months.
 * Everything here is deterministic and unit-testable in isolation.
 * ------------------------------------------------------------------
 */

const MONTH_FULL = [
  'January', 'February', 'March', 'April', 'May', 'June',
  'July', 'August', 'September', 'October', 'November', 'December'
];

const MONTH_ABBR = [
  'JAN', 'FEB', 'MAR', 'APR', 'MAY', 'JUN',
  'JUL', 'AUG', 'SEP', 'OCT', 'NOV', 'DEC'
];

const DAY_NAME = [
  'SUNDAY', 'MONDAY', 'TUESDAY', 'WEDNESDAY', 'THURSDAY', 'FRIDAY', 'SATURDAY'
];

/**
 * Returns the number of days in a given month/year.
 * Relies on the JS Date engine's native leap-year arithmetic
 * (day 0 of the *next* month == last day of *this* month).
 * @param {number} year  full 4-digit year, e.g. 2026
 * @param {number} monthIndex 0-based month index (0 = January)
 */
function daysInMonth(year, monthIndex) {
  return new Date(year, monthIndex + 1, 0).getDate();
}

/**
 * True if `year` is a leap year (Gregorian rule).
 */
function isLeapYear(year) {
  return (year % 4 === 0 && year % 100 !== 0) || (year % 400 === 0);
}

/**
 * Builds the full list of calendar days for a month.
 * @param {number} year
 * @param {number} monthIndex 0-based (0 = Jan ... 11 = Dec)
 * @returns {Array<{dayNumber:number, jsDate:Date, weekdayIndex:number, dayName:string, isSunday:boolean}>}
 */
function generateMonthDates(year, monthIndex) {
  const total = daysInMonth(year, monthIndex);
  const out = [];
  for (let d = 1; d <= total; d++) {
    // Noon avoids any DST/timezone edge cases shifting the calendar day.
    const jsDate = new Date(year, monthIndex, d, 12, 0, 0);
    const weekdayIndex = jsDate.getDay(); // 0 = Sunday
    out.push({
      dayNumber: d,
      jsDate,
      weekdayIndex,
      dayName: DAY_NAME[weekdayIndex],
      isSunday: weekdayIndex === 0
    });
  }
  return out;
}

/**
 * Given a (year, monthIndex) pair, returns the next calendar month,
 * rolling the year over at December -> January.
 */
function nextMonth(year, monthIndex) {
  if (monthIndex === 11) return { year: year + 1, monthIndex: 0 };
  return { year, monthIndex: monthIndex + 1 };
}

/**
 * Best-effort parse of a free-text month label found inside a workbook
 * (e.g. " JUN", "June", "Jul-26") into a 0-based month index.
 * Returns -1 if nothing recognisable is found.
 */
function parseMonthLabel(label) {
  if (!label) return -1;
  const clean = String(label).trim().toUpperCase().slice(0, 3);
  const idx = MONTH_ABBR.indexOf(clean);
  return idx;
}

export {
  MONTH_FULL,
  MONTH_ABBR,
  DAY_NAME,
  daysInMonth,
  isLeapYear,
  generateMonthDates,
  nextMonth,
  parseMonthLabel
};
