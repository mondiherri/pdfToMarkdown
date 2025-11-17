import * as pdfjsLib from 'https://cdnjs.cloudflare.com/ajax/libs/pdf.js/4.4.168/pdf.mjs';

pdfjsLib.GlobalWorkerOptions.workerSrc =
  'https://cdnjs.cloudflare.com/ajax/libs/pdf.js/4.4.168/pdf.worker.min.mjs';

const fileInput = document.getElementById('pdfInput');
const loadingEl = document.getElementById('loading');
const summaryEl = document.getElementById('summary');
const pageGroupsEl = document.getElementById('pageGroups');
const pageDetailsEl = document.getElementById('pageDetails');
const pageSelectEl = document.getElementById('pageSelect');
const prevButton = document.getElementById('prevPage');
const nextButton = document.getElementById('nextPage');
const pageIndicator = document.getElementById('pageIndicator');
const canvas = document.getElementById('pageCanvas');
const ctx = canvas.getContext('2d');

let pdfDoc = null;
let analysis = null;
let currentPage = 1;

fileInput.addEventListener('change', handleFileSelection);
prevButton.addEventListener('click', () => goToPage(currentPage - 1));
nextButton.addEventListener('click', () => goToPage(currentPage + 1));
pageSelectEl.addEventListener('change', (event) => goToPage(Number(event.target.value)));

function setLoading(message = '') {
  loadingEl.textContent = message;
}

async function handleFileSelection(event) {
  const file = event.target.files?.[0];
  if (!file) return;

  resetUi();
  setLoading('Loading PDF…');

  const typedArray = new Uint8Array(await file.arrayBuffer());
  pdfDoc = await pdfjsLib.getDocument({ data: typedArray }).promise;
  setLoading('Analyzing document structure…');
  analysis = await analyzeDocument(pdfDoc);
  setLoading('');

  renderSummary(analysis);
  renderPageGroups(analysis.groups);
  populatePageSelect(pdfDoc.numPages);
  await renderPage(1);
  renderPageDetails(1);
}

function resetUi() {
  analysis = null;
  pdfDoc = null;
  currentPage = 1;
  summaryEl.textContent = 'Load a document to view its metadata.';
  pageGroupsEl.innerHTML = '';
  pageDetailsEl.textContent = 'Select a page to inspect its structure.';
  pageSelectEl.innerHTML = '';
  pageSelectEl.disabled = true;
  prevButton.disabled = true;
  nextButton.disabled = true;
  pageIndicator.textContent = '';
  ctx.clearRect(0, 0, canvas.width, canvas.height);
}

function renderSummary(docAnalysis) {
  if (!docAnalysis) return;
  const tocSummary = docAnalysis.tableOfContentsPages.length
    ? `Detected on page(s): ${docAnalysis.tableOfContentsPages.join(', ')}`
    : 'Not detected';

  const groupCounts = docAnalysis.groups.reduce((acc, group) => {
    acc[group.label] = (acc[group.label] || 0) + 1;
    return acc;
  }, {});

  const groupList = Object.entries(groupCounts)
    .map(([label, count]) => `<li><strong>${label}</strong>: ${count} page(s)</li>`)
    .join('');

  summaryEl.innerHTML = `
    <p><strong>Pages:</strong> ${docAnalysis.pageCount}</p>
    <p><strong>Table of contents:</strong> ${tocSummary}</p>
    <p><strong>Index pages detected:</strong> ${docAnalysis.indexPages.length ? docAnalysis.indexPages.join(', ') : 'Not detected'}</p>
    <div>
      <p><strong>Page distribution:</strong></p>
      <ul>${groupList || '<li>No pages analyzed</li>'}</ul>
    </div>
  `;
}

function renderPageGroups(groups) {
  pageGroupsEl.innerHTML = groups
    .map((group) => `<li>Page ${group.page}: ${group.label}</li>`)
    .join('');
}

function populatePageSelect(pageCount) {
  pageSelectEl.innerHTML = Array.from({ length: pageCount })
    .map((_, index) => `<option value="${index + 1}">Page ${index + 1}</option>`)
    .join('');
  pageSelectEl.disabled = false;
  prevButton.disabled = false;
  nextButton.disabled = false;
  pageSelectEl.value = '1';
  pageIndicator.textContent = `Page 1 of ${pageCount}`;
}

async function goToPage(pageNumber) {
  if (!pdfDoc || pageNumber < 1 || pageNumber > pdfDoc.numPages) return;
  currentPage = pageNumber;
  pageSelectEl.value = String(pageNumber);
  pageIndicator.textContent = `Page ${pageNumber} of ${pdfDoc.numPages}`;
  renderPageDetails(pageNumber);
  await renderPage(pageNumber);
  prevButton.disabled = pageNumber === 1;
  nextButton.disabled = pageNumber === pdfDoc.numPages;
}

async function renderPage(pageNumber) {
  const page = await pdfDoc.getPage(pageNumber);
  const viewport = page.getViewport({ scale: 1.25 });
  canvas.height = viewport.height;
  canvas.width = viewport.width;

  const renderContext = {
    canvasContext: ctx,
    viewport,
  };

  await page.render(renderContext).promise;
}

function renderPageDetails(pageNumber) {
  if (!analysis) return;
  const details = analysis.pageDetails[pageNumber - 1];
  if (!details) return;

  pageDetailsEl.innerHTML = `
    <p><strong>Classification:</strong> ${details.groupLabel}</p>
    <p><strong>Detected header:</strong> ${details.headerLine || 'Not detected'}</p>
    <p><strong>Detected footer:</strong> ${details.footerLine || 'Not detected'}</p>
    <p><strong>Estimated columns:</strong> ${details.columnCount || 0}</p>
    <p><strong>Invisible text items:</strong> ${details.invisibleItems}</p>
    <p><strong>Table of contents cues:</strong> ${(details.tocEntryRatio * 100).toFixed(0)}% of lines</p>
    <p><strong>Index cues:</strong> ${details.indexEntryCount} line(s)</p>
  `;
}

async function analyzeDocument(pdf) {
  const summary = {
    pageCount: pdf.numPages,
    tableOfContentsPages: [],
    indexPages: [],
    groups: [],
    pageDetails: [],
  };

  for (let pageNumber = 1; pageNumber <= pdf.numPages; pageNumber += 1) {
    setLoading(`Analyzing page ${pageNumber} of ${pdf.numPages}…`);
    const page = await pdf.getPage(pageNumber);
    const viewport = page.getViewport({ scale: 1 });
    const textContent = await page.getTextContent({ normalizeWhitespace: true });
    const items = textContent.items
      .filter((item) => item.str && item.str.trim())
      .map((item) => ({
        text: item.str.trim(),
        x: item.transform[4],
        y: item.transform[5],
        width: item.width ?? Math.abs(item.transform[0]),
      }));

    const details = evaluatePage(items, viewport, pageNumber);
    summary.groups.push({ page: pageNumber, label: details.groupLabel });
    summary.pageDetails.push(details);

    if (details.hasTableOfContentsHeading || details.tocEntryRatio > 0.3) {
      summary.tableOfContentsPages.push(pageNumber);
    }
    if (details.hasIndexHeading || details.indexEntryCount > 3) {
      summary.indexPages.push(pageNumber);
    }
  }

  return summary;
}

function evaluatePage(items, viewport, pageNumber) {
  const lines = buildLines(items);
  const pageWidth = viewport.width || 1;
  const headerLine = lines.find((line) => line.relativeY > 0.85)?.text ?? '';
  const footerLine = [...lines].reverse().find((line) => line.relativeY < 0.15)?.text ?? '';

  let tocEntryCount = 0;
  let indexEntryCount = 0;
  let hasTocHeading = false;
  let hasIndexHeading = false;

  lines.forEach((line) => {
    const lower = line.text.toLowerCase();
    if (lower.includes('table of contents') || lower === 'contents') {
      hasTocHeading = true;
    }
    if (lower.startsWith('index') || lower.includes('alphabetical index')) {
      hasIndexHeading = true;
    }

    if (looksLikeTocEntry(line.text)) tocEntryCount += 1;
    if (looksLikeIndexEntry(line.text)) indexEntryCount += 1;
  });

  const tocEntryRatio = lines.length ? tocEntryCount / lines.length : 0;
  const groupLabel = classifyPage({ tocEntryRatio, indexEntryCount, hasTocHeading, hasIndexHeading });
  const columnCount = estimateColumnCount(lines, pageWidth);
  const invisibleItems = items.filter((item) => item.width < 0.1).length;

  return {
    pageNumber,
    headerLine,
    footerLine,
    columnCount,
    invisibleItems,
    tocEntryRatio,
    indexEntryCount,
    hasTableOfContentsHeading: hasTocHeading,
    hasIndexHeading,
    groupLabel,
  };
}

function buildLines(items) {
  const lineMap = new Map();
  items.forEach((item) => {
    const key = Math.round(item.y / 2);
    const row = lineMap.get(key) || [];
    row.push(item);
    lineMap.set(key, row);
  });

  const lines = [...lineMap.entries()]
    .map(([key, row]) => {
      const sorted = row.sort((a, b) => a.x - b.x);
      const text = sorted.map((item) => item.text).join(' ').replace(/\s{2,}/g, ' ').trim();
      const y = key * 2;
      const minX = sorted.length ? sorted[0].x : 0;
      return { text, y, minX };
    })
    .filter((line) => line.text)
    .sort((a, b) => b.y - a.y);

  const maxY = lines[0]?.y ?? 1;
  const minY = lines[lines.length - 1]?.y ?? 0;
  const span = Math.max(maxY - minY, 1);

  return lines.map((line) => ({
    ...line,
    relativeY: (line.y - minY) / span,
  }));
}

function looksLikeTocEntry(line) {
  const dottedLeader = /\.{2,}/.test(line);
  const endingPageNumber = /\s\d{1,4}$/.test(line);
  const hasLetters = /[A-Za-z]/.test(line);
  return hasLetters && (dottedLeader || endingPageNumber);
}

function looksLikeIndexEntry(line) {
  const pattern = /^[A-Z](?:[A-Za-z\s,\-]+)?\s+\d{1,4}$/;
  return pattern.test(line.trim());
}

function classifyPage({ tocEntryRatio, indexEntryCount, hasTocHeading, hasIndexHeading }) {
  if (hasTocHeading || tocEntryRatio >= 0.35) {
    return 'mainly Table of contents';
  }
  if (hasIndexHeading || indexEntryCount >= 4) {
    return 'mainly index';
  }
  return 'mainly text';
}

function estimateColumnCount(lines, pageWidth) {
  const xPositions = lines.map((line) => line.minX).filter((x) => Number.isFinite(x));
  if (!xPositions.length || !pageWidth) return 0;
  xPositions.sort((a, b) => a - b);
  let clusters = 1;
  const threshold = pageWidth * 0.15;
  for (let i = 1; i < xPositions.length; i += 1) {
    if (Math.abs(xPositions[i] - xPositions[i - 1]) > threshold) {
      clusters += 1;
    }
  }
  return Math.min(clusters, 4);
}
