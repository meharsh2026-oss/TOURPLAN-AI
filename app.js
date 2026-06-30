/**
 * app.js
 * ------------------------------------------------------------------
 * UI orchestration for TourPlan AI. Owns DOM wiring and the
 * upload -> detect -> generate -> download flow. All Excel logic
 * lives in excel.js; all date logic lives in planner.js; all VALUE
 * logic lives in value.js — this file just coordinates them and
 * drives the on-screen "AI status panel" animation.
 * ------------------------------------------------------------------
 */

import { MONTH_FULL, nextMonth } from './planner.js';
import { parseWorkbook, detectStructure, extractTemplateEntries, applyTourPlan, exportWorkbook } from './excel.js';

/* ----------------------------- DOM refs ----------------------------- */

const dropzone = document.getElementById('dropzone');
const fileInput = document.getElementById('fileInput');
const fileChip = document.getElementById('fileChip');
const fileChipName = document.getElementById('fileChipName');
const fileChipRemove = document.getElementById('fileChipRemove');
const uploadMessage = document.getElementById('uploadMessage');
const detectedGrid = document.getElementById('detectedGrid');
const detName = document.getElementById('detName');
const detMonth = document.getElementById('detMonth');
const detRows = document.getElementById('detRows');
const detEntries = document.getElementById('detEntries');

const monthSelect = document.getElementById('monthSelect');
const yearSelect = document.getElementById('yearSelect');
const generateBtn = document.getElementById('generateBtn');

const panelStatus = document.getElementById('panel-status');
const statusPill = document.getElementById('statusPill');
const terminal = document.getElementById('terminal');
const scanGrid = document.getElementById('scanGrid');

const panelResult = document.getElementById('panel-result');
const resMonth = document.getElementById('resMonth');
const resWorking = document.getElementById('resWorking');
const resSundays = document.getElementById('resSundays');
const resTotal = document.getElementById('resTotal');
const resWarning = document.getElementById('resWarning');
const downloadBtn = document.getElementById('downloadBtn');
const resetBtn = document.getElementById('resetBtn');

/* ------------------------------ State -------------------------------- */

const state = {
  arrayBuffer: null,
  fileName: null,
  structurePreview: null,   // structure detected at upload time, for display only
  generatedBlob: null,
  generatedFileName: null
};

/* --------------------------- Small helpers ---------------------------- */

const sleep = ms => new Promise(resolve => setTimeout(resolve, ms));

function showMessage(el, text, isWarning) {
  el.textContent = text;
  el.hidden = false;
  el.classList.toggle('inline-message--warn', !!isWarning);
}

function hideMessage(el) {
  el.hidden = true;
  el.textContent = '';
}

function formatINR(n) {
  return '₹' + Math.round(n).toLocaleString('en-IN');
}

/* ----------------------- Populate period selects ----------------------- */

MONTH_FULL.forEach((name, idx) => {
  const opt = document.createElement('option');
  opt.value = String(idx);
  opt.textContent = name;
  monthSelect.appendChild(opt);
});

function populateYearOptions(centerYear) {
  yearSelect.innerHTML = '';
  for (let y = centerYear - 1; y <= centerYear + 3; y++) {
    const opt = document.createElement('option');
    opt.value = String(y);
    opt.textContent = String(y);
    yearSelect.appendChild(opt);
  }
  yearSelect.value = String(centerYear);
}

function setDefaultPeriod(detected) {
  const today = new Date();
  let baseYear = today.getFullYear();
  let baseMonth = today.getMonth();

  if (detected && detected.monthInfo && detected.monthInfo.monthIndex >= 0) {
    const sourceYear = detected.yearInfo ? Number(detected.yearInfo.value) : baseYear;
    const rolled = nextMonth(sourceYear, detected.monthInfo.monthIndex);
    baseYear = rolled.year;
    baseMonth = rolled.monthIndex;
  } else {
    const rolled = nextMonth(baseYear, baseMonth);
    baseYear = rolled.year;
    baseMonth = rolled.monthIndex;
  }

  populateYearOptions(baseYear);
  monthSelect.value = String(baseMonth);
  yearSelect.value = String(baseYear);
}

populateYearOptions(new Date().getFullYear());

/* ------------------------------ Upload flow ----------------------------- */

function resetUploadUI() {
  fileChip.hidden = true;
  detectedGrid.hidden = true;
  hideMessage(uploadMessage);
  generateBtn.disabled = true;
  state.arrayBuffer = null;
  state.fileName = null;
  state.structurePreview = null;
  fileInput.value = '';
}

async function handleFile(file) {
  hideMessage(uploadMessage);

  if (!file) return;

  const isXlsx = /\.xlsx$/i.test(file.name);
  if (!isXlsx) {
    showMessage(uploadMessage, 'Please upload a .xlsx file — other formats are not supported.', false);
    return;
  }

  try {
    const buffer = await file.arrayBuffer();
    const workbook = await parseWorkbook(buffer);
    const structure = detectStructure(workbook);
    const entries = extractTemplateEntries(workbook, structure);

    state.arrayBuffer = buffer;
    state.fileName = file.name;

    fileChipName.textContent = file.name;
    fileChip.hidden = false;

    // structure.yearInfo.value is already resolved by excel.js during detection
    const resolvedYear = structure.yearInfo ? structure.yearInfo.value : null;

    detName.textContent = structure.personName || '—';
    detMonth.textContent = structure.monthInfo
      ? `${MONTH_FULL[structure.monthInfo.monthIndex]}${resolvedYear ? ' ' + resolvedYear : ''}`
      : '—';
    detRows.textContent = `${structure.maxDataRows} rows`;
    detEntries.textContent = `${entries.length} entries`;
    detectedGrid.hidden = false;

    state.structurePreview = structure;
    setDefaultPeriod(structure);

    generateBtn.disabled = false;
  } catch (err) {
    resetUploadUI();
    showMessage(uploadMessage, err.message || 'This workbook could not be processed.', false);
  }
}

dropzone.addEventListener('click', () => fileInput.click());
dropzone.addEventListener('keydown', e => {
  if (e.key === 'Enter' || e.key === ' ') {
    e.preventDefault();
    fileInput.click();
  }
});

fileInput.addEventListener('change', () => {
  if (fileInput.files && fileInput.files[0]) handleFile(fileInput.files[0]);
});

['dragenter', 'dragover'].forEach(evt => {
  dropzone.addEventListener(evt, e => {
    e.preventDefault();
    dropzone.classList.add('is-dragover');
  });
});

['dragleave', 'drop'].forEach(evt => {
  dropzone.addEventListener(evt, e => {
    e.preventDefault();
    dropzone.classList.remove('is-dragover');
  });
});

dropzone.addEventListener('drop', e => {
  const file = e.dataTransfer.files && e.dataTransfer.files[0];
  if (file) handleFile(file);
});

fileChipRemove.addEventListener('click', e => {
  e.stopPropagation();
  resetUploadUI();
});

/* --------------------------- AI status panel ---------------------------- */

const STEPS = [
  'Reading Workbook...',
  'Detecting Sheet...',
  'Updating Dates...',
  'Updating Days...',
  'Applying Sunday Rules...',
  'Calculating VALUE...',
  'Generating Workbook...',
  'Download Ready.'
];

function buildScanGrid(cellCount) {
  scanGrid.innerHTML = '';
  const cells = [];
  for (let i = 0; i < cellCount; i++) {
    const div = document.createElement('div');
    div.className = 'scan-cell';
    scanGrid.appendChild(div);
    cells.push(div);
  }
  return cells;
}

function addTerminalLine(text, state = 'active') {
  const line = document.createElement('div');
  line.className = 'terminal-line';
  line.dataset.state = state;
  line.innerHTML = `<span class="terminal-line__marker">${state === 'done' ? '✓' : state === 'warn' ? '!' : '›'}</span><span class="terminal-line__text">${text}</span>`;
  terminal.appendChild(line);
  terminal.scrollTop = terminal.scrollHeight;
  return line;
}

async function runStatusAnimation(scanCellCount) {
  terminal.innerHTML = '';
  statusPill.textContent = 'Working';
  statusPill.dataset.state = 'working';
  const cells = buildScanGrid(scanCellCount);

  let cellCursor = 0;
  const litPerStep = Math.ceil(scanCellCount / (STEPS.length - 1));

  for (let i = 0; i < STEPS.length; i++) {
    const isLast = i === STEPS.length - 1;
    const line = addTerminalLine(STEPS[i], isLast ? 'done' : 'active');
    await sleep(isLast ? 120 : 260);

    if (!isLast) {
      const target = Math.min(scanCellCount, cellCursor + litPerStep);
      for (; cellCursor < target; cellCursor++) {
        cells[cellCursor].classList.add('is-lit');
      }
      line.dataset.state = 'done';
      line.querySelector('.terminal-line__marker').textContent = '✓';
    } else {
      cells.forEach(c => c.classList.add('is-done'));
    }
  }

  statusPill.textContent = 'Done';
  statusPill.dataset.state = 'done';
}

/* ------------------------------ Generation ------------------------------ */

generateBtn.addEventListener('click', async () => {
  if (!state.arrayBuffer) {
    showMessage(uploadMessage, 'Please upload a workbook before generating a tour plan.', false);
    return;
  }

  const monthIndex = Number(monthSelect.value);
  const year = Number(yearSelect.value);

  generateBtn.disabled = true;
  panelStatus.hidden = false;
  panelResult.hidden = true;
  panelStatus.scrollIntoView({ behavior: 'smooth', block: 'nearest' });

  try {
    // Re-parse a fresh workbook from the original bytes every run, so
    // repeated generations never compound on top of a previous mutation.
    const workbook = await parseWorkbook(state.arrayBuffer);
    const structure = detectStructure(workbook);
    const templateEntries = extractTemplateEntries(workbook, structure);

    const animation = runStatusAnimation(structure.maxDataRows);
    const summary = applyTourPlan(workbook, structure, templateEntries, year, monthIndex);
    const blob = await exportWorkbook(workbook);

    await animation;

    state.generatedBlob = blob;
    state.generatedFileName = `TourPlan_${MONTH_FULL[monthIndex]}_${year}.xlsx`;

    resMonth.textContent = summary.monthLabel;
    resWorking.textContent = String(summary.workingDays);
    resSundays.textContent = String(summary.sundays);
    resTotal.textContent = formatINR(summary.finalTotal);

    if (summary.truncatedWarning) {
      showMessage(
        resWarning,
        `This template has ${structure.maxDataRows} row slots, but ${summary.daysInTargetMonth} days were requested — the last ${summary.daysInTargetMonth - summary.rowsWritten} day(s) could not be added.`,
        true
      );
    } else {
      hideMessage(resWarning);
    }

    panelResult.hidden = false;
    panelResult.scrollIntoView({ behavior: 'smooth', block: 'nearest' });
  } catch (err) {
    addTerminalLine(err.message || 'Something went wrong while generating the workbook.', 'warn');
    statusPill.textContent = 'Error';
    statusPill.dataset.state = 'warn';
  } finally {
    generateBtn.disabled = false;
  }
});

/* ------------------------------- Download -------------------------------- */

downloadBtn.addEventListener('click', () => {
  if (!state.generatedBlob) return;
  const url = URL.createObjectURL(state.generatedBlob);
  const a = document.createElement('a');
  a.href = url;
  a.download = state.generatedFileName;
  document.body.appendChild(a);
  a.click();
  document.body.removeChild(a);
  setTimeout(() => URL.revokeObjectURL(url), 4000);
});

resetBtn.addEventListener('click', () => {
  resetUploadUI();
  panelStatus.hidden = true;
  panelResult.hidden = true;
  state.generatedBlob = null;
  state.generatedFileName = null;
  window.scrollTo({ top: 0, behavior: 'smooth' });
});
