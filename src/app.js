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
      const analyzeButton = document.getElementById('analyzeBtn');
      const previewButton = document.getElementById('previewBtn');
      const extractButton = document.getElementById('extractBtn');
      const extractImagesButton = document.getElementById('extractImagesBtn');
      const imageToolButton = document.getElementById('imageTool');
      const shapesToolButton = document.getElementById('shapesTool');
      const markdownButton = document.getElementById('markdownBtn');
      const jsonButton = document.getElementById('jsonBtn');
      const csvButton = document.getElementById('csvBtn');
      const previewPanel = document.getElementById('previewPanel');
      const selectionLayer = document.getElementById('selectionLayer');
      const selectToolButton = document.getElementById('selectTool');
      const overrideSummaryEl = document.getElementById('overrideSummary');
      const tocThresholdInput = document.getElementById('tocThreshold');
      const indexMinimumInput = document.getElementById('indexMinimum');
      const invisibleThresholdInput = document.getElementById('invisibleThreshold');
      const applySettingsButton = document.getElementById('applySettings');
      const resetSettingsButton = document.getElementById('resetSettings');
      const searchInput = document.getElementById('searchInput');
      const searchResultsEl = document.getElementById('searchResults');
      const entityListEl = document.getElementById('entityList');
      const imageGridEl = document.getElementById('imageGrid');
      let currentRenderTask = null;

      let pdfDoc = null;
      let analysis = null;
      let currentPage = 1;
      let pdfData = null;
      let pdfFileName = 'pdf-analysis';
      let extractedImages = [];
      let docExtractionStruct = [];
      let activeTool = 'select';
      let showImageHighlights = false;
      let showCategoryHighlights = false;
      const selectionBox = document.createElement('div');
      selectionBox.className = 'selection-box';
      selectionBox.style.display = 'none';
      selectionLayer?.appendChild(selectionBox);
      const pageImageRects = new Map();
      let selectionMenu = null;
      let currentViewportState = {
        baseWidth: 0,
        baseHeight: 0,
        scale: 1,
        outputScale: 1,
      };
      let isDraggingSelection = false;
      let dragStart = null;
      const headerOverridePages = new Set();
      const footerOverridePages = new Set();
      const columnPreviewSelections = new Map();
      const textAreaOverlayPages = new Set();
      const headerFooterOverlayPages = new Set();
      const invisibleTextOverlayPages = new Set();
      let headerRememberScope = 'page';
      let footerRememberScope = 'page';
      let headerRememberEnabled = false;
      let footerRememberEnabled = false;
      const STORAGE_KEYS = {
        overrides: 'pdfAnalyzerOverrides',
        heuristics: 'pdfAnalyzerHeuristics',
      };
      const heuristicDefaults = {
        tocRatioThreshold: 0.35,
        indexEntryMinimum: 4,
        invisibleWidthThreshold: 0.1,
      };
      let heuristics = loadStoredHeuristics();

      fileInput.addEventListener('change', handleFileSelection);
      prevButton.addEventListener('click', () => goToPage(currentPage - 1));
      nextButton.addEventListener('click', () => goToPage(currentPage + 1));
      pageSelectEl.addEventListener('change', handlePageInput);
      analyzeButton.addEventListener('click', () => runAnalysis());
      previewButton.addEventListener('click', focusPreviewPanel);
      extractButton.addEventListener('click', exportAnalysisReport);
      extractImagesButton.addEventListener('click', handleExtractImages);
      markdownButton.addEventListener('click', exportMarkdownDocument);
      jsonButton.addEventListener('click', exportJsonAnalysis);
      csvButton.addEventListener('click', exportCsvSummary);
      window.addEventListener('resize', handleResize);
      applySettingsButton.addEventListener('click', applyHeuristicChanges);
      resetSettingsButton.addEventListener('click', resetHeuristicsToDefault);
      searchInput.addEventListener('input', handleSearchInput);
      syncHeuristicInputs();
      if (selectToolButton) {
        selectToolButton.addEventListener('click', () => setActiveTool('select'));
      }
      if (imageToolButton) {
        imageToolButton.addEventListener('click', () => {
          showImageHighlights = !showImageHighlights;
          imageToolButton.classList.toggle('active', showImageHighlights);
          renderImageHighlights(currentPage);
        });
      }
      if (shapesToolButton) {
        shapesToolButton.addEventListener('click', () => {
          showCategoryHighlights = !showCategoryHighlights;
          shapesToolButton.classList.toggle('active', showCategoryHighlights);
          renderCategoryHighlights(currentPage);
        });
      }
      setupSelectionHandlers();

      function setLoading(message = '') {
        loadingEl.textContent = message || 'Status: idle';
      }

      function storageAvailable(){
        try {
          const key = '__pdf_analyzer_test__';
          window.localStorage.setItem(key, '1');
          window.localStorage.removeItem(key);
          return true;
        } catch (error) {
          return false;
        }
      };

      function readFromStorage(key, fallback = null) {
        if (!storageAvailable()) return fallback;
        try {
          const raw = window.localStorage.getItem(key);
          return raw ? JSON.parse(raw) : fallback;
        } catch (error) {
          return fallback;
        }
      }

      function writeToStorage(key, value) {
        if (!storageAvailable) return;
        try {
          window.localStorage.setItem(key, JSON.stringify(value));
        } catch (error) {
          console.warn('Failed to persist settings', error);
        }
      }

      function loadStoredHeuristics() {
        const stored = readFromStorage(STORAGE_KEYS.heuristics, {});
        return {
          ...heuristicDefaults,
          ...stored,
        };
      }

      function persistHeuristics() {
        writeToStorage(STORAGE_KEYS.heuristics, heuristics);
      }

      function syncHeuristicInputs() {
        if (tocThresholdInput) {
          tocThresholdInput.value = Math.round((heuristics.tocRatioThreshold || 0.35) * 100);
        }
        if (indexMinimumInput) {
          indexMinimumInput.value = heuristics.indexEntryMinimum || heuristicDefaults.indexEntryMinimum;
        }
        if (invisibleThresholdInput) {
          invisibleThresholdInput.value = heuristics.invisibleWidthThreshold || heuristicDefaults.invisibleWidthThreshold;
        }
      }

      function applyHeuristicChanges() {
        const tocValue = Number(tocThresholdInput?.value) / 100 || heuristicDefaults.tocRatioThreshold;
        const indexValue = Number(indexMinimumInput?.value) || heuristicDefaults.indexEntryMinimum;
        const invisibleValue = Number(invisibleThresholdInput?.value) || heuristicDefaults.invisibleWidthThreshold;
        heuristics = {
          tocRatioThreshold: clampNumber(tocValue, 0.05, 0.95),
          indexEntryMinimum: clampNumber(indexValue, 1, 20),
          invisibleWidthThreshold: clampNumber(invisibleValue, 0.01, 1),
        };
        persistHeuristics();
        syncHeuristicInputs();
        if (pdfData) {
          setLoading('Re-running analysis with new settings…');
          runAnalysis();
        } else {
          setLoading('Settings saved. Upload a PDF to analyze.');
        }
      }

      function resetHeuristicsToDefault() {
        heuristics = { ...heuristicDefaults };
        persistHeuristics();
        syncHeuristicInputs();
        if (pdfData) {
          setLoading('Restoring defaults and re-running analysis…');
          runAnalysis();
        }
      }

      function clampNumber(value, min, max) {
        if (!Number.isFinite(value)) return min;
        return Math.min(Math.max(value, min), max);
      }

      function prepareSearchPanel() {
        if (!analysis) return;
        searchInput.disabled = false;
        if (searchInput.value.trim()) {
          handleSearchInput({ target: searchInput });
        } else {
          searchResultsEl.innerHTML = '<p>Type in the search box to jump to matches.</p>';
        }
        renderEntityList();
      }

      function handleSearchInput(event) {
        if (!analysis?.searchIndex?.length) {
          searchResultsEl.innerHTML = '<p>Run an analysis to enable search.</p>';
          return;
        }
        const query = event?.target?.value?.trim() || '';
        if (!query) {
          searchResultsEl.innerHTML = '<p>Type in the search box to jump to matches.</p>';
          return;
        }
        const results = runSearch(query);
        renderSearchResults(results, query);
      }

      function runSearch(query) {
        const lower = query.toLowerCase();
        const matches = [];
        analysis.searchIndex.forEach((entry) => {
          const index = entry.textLower.indexOf(lower);
          if (index === -1) return;
          const start = Math.max(index - 60, 0);
          const end = Math.min(index + query.length + 60, entry.textLower.length);
          const snippet = entry.text.slice(start, end).replace(/\s+/g, ' ');
          matches.push({ pageNumber: entry.pageNumber, snippet });
        });
        return matches.slice(0, 25);
      }

      function renderSearchResults(results = [], query = '') {
        const safeQuery = escapeHtml(query);
        if (!results.length) {
          searchResultsEl.innerHTML = `<p>No matches for <strong>${safeQuery}</strong>.</p>`;
          return;
        }
        const highlighted = results
          .map((result) => {
            const safeSnippet = escapeHtml(result.snippet);
            const regex = new RegExp(`(${escapeRegExp(query)})`, 'ig');
            const snippet = safeSnippet.replace(regex, '<mark>$1</mark>');
            return `
              <div class="search-result" data-jump-page="${result.pageNumber}">
                <strong>Page ${result.pageNumber}</strong>
                <div class="search-snippet">${snippet}</div>
              </div>
            `;
          })
          .join('');
        searchResultsEl.innerHTML = highlighted;
      }

      searchResultsEl.addEventListener('click', (event) => {
        const target = event.target.closest('[data-jump-page]');
        if (!target) return;
        const pageNumber = Number(target.dataset.jumpPage);
        if (Number.isFinite(pageNumber)) {
          goToPage(pageNumber);
        }
      });

      entityListEl.addEventListener('click', (event) => {
        const chip = event.target.closest('[data-entity-page]');
        if (!chip) return;
        const pageNumber = Number(chip.dataset.entityPage);
        const text = chip.dataset.entityText || chip.textContent;
        if (text) {
          searchInput.value = text;
          handleSearchInput({ target: searchInput });
        }
        if (Number.isFinite(pageNumber)) {
          goToPage(pageNumber);
        }
      });

      function renderEntityList() {
        if (!analysis?.entities?.length) {
          entityListEl.innerHTML = '<p class="hint">No recurring entities detected yet.</p>';
          return;
        }
        entityListEl.innerHTML = analysis.entities
          .map(
            (entity) =>
              `<span class="entity-chip" data-entity-page="${entity.firstPage}" data-entity-text="${escapeHtml(
                entity.text
              )}">${escapeHtml(entity.text)}</span>`
          )
          .join('');
      }

      const EXTRACTION_CATEGORIES = ['Title', 'Paragraph', 'Header', 'Footer', 'Column', 'Descriptor'];
      const CATEGORY_COLORS = {
        Title: 'category-title',
        Paragraph: 'category-paragraph',
        Header: 'category-header',
        Footer: 'category-footer',
        Column: 'category-column',
        Image: 'category-image',
        Descriptor: 'category-descriptor',
      };

      function setActiveTool(tool) {
        activeTool = tool;
        if (selectToolButton) {
          if (tool === 'select') {
            selectToolButton.classList.add('active');
          } else {
            selectToolButton.classList.remove('active');
          }
        }
      }

      function setupSelectionHandlers() {
        if (!canvas) return;
        canvas.addEventListener('mousedown', (event) => {
          if (activeTool !== 'select' || !pdfDoc) return;
          const point = getCanvasPoint(event);
          if (!point) return;
          hideSelectionMenu();
          isDraggingSelection = true;
          dragStart = point;
          updateSelectionBox({ x: point.x, y: point.y, width: 0, height: 0, displayX: point.displayX, displayY: point.displayY, displayWidth: 0, displayHeight: 0 });
        });

        canvas.addEventListener('mousemove', (event) => {
          if (!isDraggingSelection || !dragStart) return;
          const point = getCanvasPoint(event);
          if (!point) return;
          const rect = buildRectFromPoints(dragStart, point);
          updateSelectionBox(rect);
        });

        canvas.addEventListener('mouseup', (event) => {
          if (!isDraggingSelection || !dragStart) return;
          const point = getCanvasPoint(event);
          isDraggingSelection = false;
          if (!point) {
            hideSelectionMenu();
            return;
          }
          const rect = buildRectFromPoints(dragStart, point);
          dragStart = null;
          if (rect.width < 5 || rect.height < 5) {
            hideSelectionMenu();
            selectionBox.style.display = 'none';
            return;
          }
          openSelectionMenu(rect, event);
        });

        document.addEventListener('click', (event) => {
          if (!selectionMenu) return;
          if (event.target.closest('.selection-menu')) return;
          hideSelectionMenu();
        });
      }

      function getCanvasPoint(event) {
        const rect = canvas.getBoundingClientRect();
        const scaleX = canvas.width / rect.width;
        const scaleY = canvas.height / rect.height;
        const x = (event.clientX - rect.left) * scaleX;
        const y = (event.clientY - rect.top) * scaleY;
        if (Number.isFinite(x) && Number.isFinite(y)) {
          return {
            x,
            y,
            displayX: event.clientX - rect.left,
            displayY: event.clientY - rect.top,
            displayWidth: rect.width,
            displayHeight: rect.height,
          };
        }
        return null;
      }

      function buildRectFromPoints(a, b) {
        const x = Math.min(a.x, b.x);
        const y = Math.min(a.y, b.y);
        const width = Math.abs(a.x - b.x);
        const height = Math.abs(a.y - b.y);
        const displayX = Math.min(a.displayX, b.displayX);
        const displayY = Math.min(a.displayY, b.displayY);
        const displayWidth = Math.abs(a.displayX - b.displayX);
        const displayHeight = Math.abs(a.displayY - b.displayY);
        return { x, y, width, height, displayX, displayY, displayWidth, displayHeight };
      }

      function updateSelectionBox(rect) {
        if (!selectionLayer || !selectionBox) return;
        selectionBox.style.display = 'block';
        const canvasRect = canvas.getBoundingClientRect();
        const layerRect = selectionLayer.getBoundingClientRect();
        const left = (canvasRect.left - layerRect.left) + (rect.displayX || 0);
        const top = (canvasRect.top - layerRect.top) + (rect.displayY || 0);
        selectionBox.style.left = `${left}px`;
        selectionBox.style.top = `${top}px`;
        selectionBox.style.width = `${rect.displayWidth || 0}px`;
        selectionBox.style.height = `${rect.displayHeight || 0}px`;
      }

      function hideSelectionMenu() {
        if (selectionMenu?.parentNode) {
          selectionMenu.parentNode.removeChild(selectionMenu);
        }
        selectionMenu = null;
      }

      function openSelectionMenu(rect, event) {
        if (!previewPanel) return;
        hideSelectionMenu();
        selectionMenu = document.createElement('div');
        selectionMenu.className = 'selection-menu';
        selectionMenu.innerHTML = `
          <p>Label selected area as:</p>
          ${EXTRACTION_CATEGORIES.map((label) => `<button type="button" data-extract-label="${label}">${label}</button>`).join('')}
        `;
        selectionMenu.addEventListener('click', (clickEvent) => {
          const button = clickEvent.target.closest('[data-extract-label]');
          if (!button) return;
          const label = button.dataset.extractLabel;
          recordExtraction(label, rect);
          hideSelectionMenu();
        });
        const previewRect = previewPanel.getBoundingClientRect();
        const offsetLeft = (event.clientX - previewRect.left) + 8;
        const offsetTop = (event.clientY - previewRect.top) + 8;
        selectionMenu.style.left = `${offsetLeft}px`;
        selectionMenu.style.top = `${offsetTop}px`;
        selectionMenu.style.position = 'absolute';
        previewPanel.appendChild(selectionMenu);
      }

      function recordExtraction(category, rect) {
        if (!pdfDoc || !rect || !currentViewportState.baseWidth || !currentViewportState.baseHeight) return;
        const scale = currentViewportState.scale || 1;
        const outputScale = currentViewportState.outputScale || 1;
        const baseWidth = currentViewportState.baseWidth || 1;
        const baseHeight = currentViewportState.baseHeight || 1;
        const pdfX = rect.x / (scale * outputScale);
        const pdfYFromTop = rect.y / (scale * outputScale);
        const pdfWidth = rect.width / (scale * outputScale);
        const pdfHeight = rect.height / (scale * outputScale);
        const pdfY = baseHeight - (pdfYFromTop + pdfHeight);
        const normalizedRect = {
          x: clampNumber(pdfX / baseWidth, 0, 1),
          y: clampNumber(pdfY / baseHeight, 0, 1),
          width: clampNumber(pdfWidth / baseWidth, 0, 1),
          height: clampNumber(pdfHeight / baseHeight, 0, 1),
        };
        docExtractionStruct.push({
          pageNumber: currentPage,
          category,
          pdfRect: {
            x: pdfX,
            y: pdfY,
            width: pdfWidth,
            height: pdfHeight,
            pageWidth: baseWidth,
            pageHeight: baseHeight,
          },
          normalizedRect,
          canvasRect: rect,
          recordedAt: new Date().toISOString(),
        });
        setLoading(`Saved ${category} region on page ${currentPage}.`);
        renderCategoryHighlights(currentPage);
      }
      function deriveDrawImageRect(args = []) {
        const [, x = 0, y = 0, w = 0, h = 0] = args;
        const width = args.length >= 5 ? w : 0;
        const height = args.length >= 5 ? h : 0;
        return {
          x: Number(x) || 0,
          y: Number(y) || 0,
          width: Number(width) || 0,
          height: Number(height) || 0,
        };
      }

      function buildTransformedQuad(transform, rect) {
        const { x, y, width, height } = rect;
        const points = [
          { x, y },
          { x: x + width, y },
          { x: x + width, y: y + height },
          { x, y: y + height },
        ];
        return points.map((point) => ({
          x: transform.a * point.x + transform.c * point.y + transform.e,
          y: transform.b * point.x + transform.d * point.y + transform.f,
        }));
      }

      function quadToBounds(quad = []) {
        const xs = quad.map((p) => p.x);
        const ys = quad.map((p) => p.y);
        const minX = Math.min(...xs);
        const minY = Math.min(...ys);
        const maxX = Math.max(...xs);
        const maxY = Math.max(...ys);
        return { x: minX, y: minY, width: maxX - minX, height: maxY - minY };
      }

      function dedupeRects(rects = []) {
        const seen = new Set();
        const results = [];
        rects.forEach((rect) => {
          const key = [rect.x, rect.y, rect.width, rect.height].map((v) => Math.round(v * 1000)).join(':');
          if (seen.has(key)) return;
          seen.add(key);
          results.push(rect);
        });
        return results;
      }

      function renderImageHighlights(pageNumber) {
        if (!selectionLayer) return;
        selectionLayer.querySelectorAll('.image-highlight').forEach((node) => node.remove());
        if (!showImageHighlights) return;
        const rects = pageImageRects.get(pageNumber) || [];
        if (!rects.length) return;
        const canvasRect = canvas.getBoundingClientRect();
        const layerRect = selectionLayer.getBoundingClientRect();
        const offsetLeft = canvasRect.left - layerRect.left;
        const offsetTop = canvasRect.top - layerRect.top;
        rects.forEach((rect) => {
          const highlight = document.createElement('div');
          highlight.className = 'image-highlight';
          const left = (rect.x || 0) * canvasRect.width + offsetLeft;
          const top = (rect.y || 0) * canvasRect.height + offsetTop;
          const width = (rect.width || 0) * canvasRect.width;
          const height = (rect.height || 0) * canvasRect.height;
          highlight.style.left = `${left}px`;
          highlight.style.top = `${top}px`;
          highlight.style.width = `${width}px`;
          highlight.style.height = `${height}px`;
          selectionLayer.appendChild(highlight);
        });
      }

      function normalizeRect(entry) {
        if (!entry) return null;
        if (entry.normalizedRect) return entry.normalizedRect;
        const pdfRect = entry.pdfRect;
        if (!pdfRect?.pageWidth || !pdfRect?.pageHeight) return null;
        return {
          x: clampNumber((pdfRect.x || 0) / pdfRect.pageWidth, 0, 1),
          y: clampNumber((pdfRect.y || 0) / pdfRect.pageHeight, 0, 1),
          width: clampNumber((pdfRect.width || 0) / pdfRect.pageWidth, 0, 1),
          height: clampNumber((pdfRect.height || 0) / pdfRect.pageHeight, 0, 1),
        };
      }

      function renderCategoryHighlights(pageNumber) {
        if (!selectionLayer) return;
        selectionLayer.querySelectorAll('.category-highlight').forEach((node) => node.remove());
        if (!showCategoryHighlights) return;
        const canvasRect = canvas.getBoundingClientRect();
        const layerRect = selectionLayer.getBoundingClientRect();
        const offsetLeft = canvasRect.left - layerRect.left;
        const offsetTop = canvasRect.top - layerRect.top;
        const entries = (docExtractionStruct || []).filter((entry) => entry.pageNumber === pageNumber);
        const rects = [];
        entries.forEach((entry) => {
          const normalized = normalizeRect(entry);
          if (!normalized) return;
          rects.push({ category: entry.category || 'Descriptor', rect: normalized });
        });
        const imageRects = pageImageRects.get(pageNumber) || [];
        imageRects.forEach((rect) => rects.push({ category: 'Image', rect }));
        rects.forEach((item) => {
          const colorClass = CATEGORY_COLORS[item.category] || 'category-paragraph';
          const overlay = document.createElement('div');
          overlay.className = `category-highlight ${colorClass}`;
          const left = (item.rect.x || 0) * canvasRect.width + offsetLeft;
          const top = (item.rect.y || 0) * canvasRect.height + offsetTop;
          const width = (item.rect.width || 0) * canvasRect.width;
          const height = (item.rect.height || 0) * canvasRect.height;
          overlay.style.left = `${left}px`;
          overlay.style.top = `${top}px`;
          overlay.style.width = `${width}px`;
          overlay.style.height = `${height}px`;
          selectionLayer.appendChild(overlay);
        });
      }


      function clearImageGrid() {
        if (!imageGridEl) return;
        imageGridEl.innerHTML = '<p class="hint">Run "Extract Images" to see any embedded artwork.</p>';
      }

      function renderImageGrid(images = []) {
        if (!imageGridEl) return;
        if (!images.length) {
          imageGridEl.innerHTML = '<p class="hint">No images detected in this document.</p>';
          return;
        }
        imageGridEl.innerHTML = images
          .map(
            (image, index) => `
              <div class="image-card">
                <img src="${image.dataUrl}" alt="Extracted image ${index + 1} from page ${image.pageNumber || '?'}" loading="lazy" />
                <div class="image-meta">
                  <span class="pill muted">Page ${image.pageNumber || '?'}</span>
                  <span>${image.width || '?'} × ${image.height || '?'}</span>
                </div>
              </div>
            `
          )
          .join('');
      }

      async function handleExtractImages() {
        if (!pdfData) {
          setLoading('Select a PDF to extract images.');
          return;
        }
        if (extractImagesButton) extractImagesButton.disabled = true;
        try {
          if (!pdfDoc) {
            setLoading('Loading PDF…');
            pdfDoc = await pdfjsLib.getDocument({ data: pdfData }).promise;
          }
          setLoading('Extracting images from PDF…');
          extractedImages = await collectImagesFromDocument(pdfDoc);
          renderImageGrid(extractedImages);
          const message = extractedImages.length
            ? `Extracted ${extractedImages.length} image${extractedImages.length === 1 ? '' : 's'}.`
            : 'No embedded images were detected.';
          setLoading(message);
        } catch (error) {
          console.error('Failed to extract images', error);
          setLoading('Unable to extract images from this PDF.');
        } finally {
          if (extractImagesButton) extractImagesButton.disabled = !pdfDoc;
        }
      }

      async function collectImagesFromDocument(pdf) {
        if (!pdf) return [];
        const images = [];
        for (let pageNumber = 1; pageNumber <= pdf.numPages; pageNumber += 1) {
          setLoading(`Extracting images from page ${pageNumber} of ${pdf.numPages}…`);
          const page = await pdf.getPage(pageNumber);
          const pageImages = await extractImagesFromPage(page);
          images.push(...pageImages);
        }
        setLoading('');
        return images;
      }

      async function extractImagesFromPage(page) {
        if (!page) return [];
        try {
          const viewport = page.getViewport({ scale: 0.5 });
          const width = Math.max(Math.ceil(viewport.width), 1);
          const height = Math.max(Math.ceil(viewport.height), 1);
          const { context } = createCanvasWithContext(width, height);
          if (!context) return [];
          await page.render({ canvasContext: context, viewport }).promise;
          const operatorList = await page.getOperatorList();
          const results = [];
          const seen = new Set();
          for (let i = 0; i < operatorList.fnArray.length; i += 1) {
            const fnId = operatorList.fnArray[i];
            if (!isImageOperation(fnId)) continue;
            const args = operatorList.argsArray[i];
            // eslint-disable-next-line no-await-in-loop
            const image = await resolveImageFromArgs(args, page);
            if (!image?.dataUrl) continue;
            if (seen.has(image.dataUrl)) continue;
            seen.add(image.dataUrl);
            results.push({ ...image, pageNumber: page.pageNumber });
          }
          return results;
        } catch (error) {
          console.warn('Image extraction failed for a page', error);
          return [];
        }
      }

      async function resolveImageFromArgs(args = [], page) {
        const [firstArg] = args || [];
        if (firstArg?.data && firstArg?.width && firstArg?.height) {
          return convertInlineImage(firstArg);
        }
        if (typeof firstArg === 'string') {
          return readImageFromStore(page?.objs, firstArg);
        }
        return null;
      }

      async function convertInlineImage(imageData) {
        const width = Math.max(Math.round(imageData.width || 0), 1);
        const height = Math.max(Math.round(imageData.height || 0), 1);
        if (!width || !height || !imageData.data) return null;
        const { canvas, context } = createCanvasWithContext(width, height);
        if (!context) return null;
        const dataArray = imageData.data instanceof Uint8ClampedArray ? imageData.data : new Uint8ClampedArray(imageData.data);
        const inline = new ImageData(dataArray, width, height);
        context.putImageData(inline, 0, 0);
        return { dataUrl: canvas.toDataURL('image/png'), width, height };
      }

      async function readImageFromStore(store, name) {
        if (!store || !name) return null;
        try {
          const direct = store.get(name);
          const converted = await convertImageLike(direct);
          if (converted) return converted;
        } catch (error) {
          // Object may not be ready yet; fall back to async getter.
        }
        return new Promise((resolve) => {
          try {
            store.get(name, async (image) => {
              resolve(await convertImageLike(image));
            });
          } catch (error) {
            resolve(null);
          }
        });
      }

      async function convertImageLike(image) {
        if (!image) return null;
        const width = Math.max(Math.round(image.width || image.bitmapWidth || image._width || 0), 1);
        const height = Math.max(Math.round(image.height || image.bitmapHeight || image._height || 0), 1);
        if (!width || !height) return null;
        const { canvas, context } = createCanvasWithContext(width, height);
        if (!context) return null;
        try {
          if (image instanceof ImageData) {
            context.putImageData(image, 0, 0);
          } else if (image?.data) {
            const data = image.data instanceof Uint8ClampedArray ? image.data : new Uint8ClampedArray(image.data);
            const imageData = new ImageData(data, width, height);
            context.putImageData(imageData, 0, 0);
          } else if (typeof ImageBitmap !== 'undefined' && image instanceof ImageBitmap) {
            context.drawImage(image, 0, 0, width, height);
          } else if (typeof OffscreenCanvas !== 'undefined' && image instanceof OffscreenCanvas) {
            context.drawImage(image, 0, 0, width, height);
          } else if (image instanceof HTMLCanvasElement || image instanceof HTMLImageElement) {
            context.drawImage(image, 0, 0, width, height);
          } else {
            return null;
          }
          return { dataUrl: canvas.toDataURL('image/png'), width, height };
        } catch (error) {
          console.warn('Could not convert image', error);
          return null;
        }
      }

      function createCanvasWithContext(width, height) {
        const canvasEl = document.createElement('canvas');
        canvasEl.width = Math.max(width, 1);
        canvasEl.height = Math.max(height, 1);
        const context = canvasEl.getContext('2d');
        return { canvas: canvasEl, context };
      }

      function isImageOperation(fnId) {
        const ops = pdfjsLib.OPS || {};
        return (
          fnId === ops.paintImageXObject ||
          fnId === ops.paintImageXObjectRepeat ||
          fnId === ops.paintInlineImageXObject ||
          fnId === ops.paintInlineImageXObjectGroup ||
          fnId === ops.paintImageMaskXObject ||
          fnId === ops.paintJpegXObject
        );
      }

      function escapeRegExp(text = '') {
        return text.replace(/[.*+?^${}()|[\]\\]/g, '\\$&');
      }

      function escapeHtml(text = '') {
        const div = document.createElement('div');
        div.textContent = text;
        return div.innerHTML;
      }

      async function handleFileSelection(event) {
        const file = event.target.files?.[0];
        resetUi();
        if (!file) {
          setLoading('');
          return;
        }

        setLoading('Reading PDF…');
        pdfData = new Uint8Array(await file.arrayBuffer());
        const baseName = file.name?.replace(/\.pdf$/i, '') || 'pdf-analysis';
        pdfFileName = baseName.replace(/[^a-z0-9-_]+/gi, '-').toLowerCase() || 'pdf-analysis';
        analyzeButton.disabled = false;
        setLoading('Ready. Click Analyze to inspect the document.');
      }

      function resetUi() {
        analysis = null;
        pdfDoc = null;
        currentPage = 1;
        pdfData = null;
        pdfFileName = 'pdf-analysis';
        headerOverridePages.clear();
        footerOverridePages.clear();
        textAreaOverlayPages.clear();
        columnPreviewSelections.clear();
        headerFooterOverlayPages.clear();
        invisibleTextOverlayPages.clear();
        headerRememberScope = 'page';
        footerRememberScope = 'page';
        headerRememberEnabled = false;
        footerRememberEnabled = false;
        summaryEl.textContent = 'Load a document to view its metadata.';
        pageGroupsEl.innerHTML = '';
        pageDetailsEl.textContent = 'Select a page to inspect its structure.';
        pageSelectEl.value = '';
        pageSelectEl.disabled = true;
        pageSelectEl.removeAttribute('max');
        pageSelectEl.removeAttribute('min');
        pageSelectEl.removeAttribute('step');
        prevButton.disabled = true;
        nextButton.disabled = true;
        pageIndicator.textContent = 'Page —';
        ctx.clearRect(0, 0, canvas.width, canvas.height);
        analyzeButton.disabled = true;
        previewButton.disabled = true;
        extractButton.disabled = true;
        if (extractImagesButton) extractImagesButton.disabled = true;
        markdownButton.disabled = true;
        jsonButton.disabled = true;
        csvButton.disabled = true;
        if (overrideSummaryEl) {
          overrideSummaryEl.innerHTML = '<p>No header/footer overrides applied yet.</p>';
        }
        searchInput.value = '';
        searchInput.disabled = true;
        searchResultsEl.innerHTML = '<p>Load a document to enable search.</p>';
        entityListEl.textContent = '';
        extractedImages = [];
        docExtractionStruct = [];
        pageImageRects.clear();
        showImageHighlights = false;
        showCategoryHighlights = false;
        imageToolButton?.classList.remove('active');
        shapesToolButton?.classList.remove('active');
        clearImageGrid();
        selectionBox.style.display = 'none';
        hideSelectionMenu();
        renderCategoryHighlights(currentPage);
        renderImageHighlights(currentPage);
        setLoading('');
      }

      async function runAnalysis() {
        if (!pdfData) {
          setLoading('Select a PDF to analyze.');
          return;
        }

        analyzeButton.disabled = true;
        previewButton.disabled = true;
        extractButton.disabled = true;
        if (extractImagesButton) extractImagesButton.disabled = true;
        markdownButton.disabled = true;
        headerOverridePages.clear();
        footerOverridePages.clear();
        columnPreviewSelections.clear();
        textAreaOverlayPages.clear();
        headerFooterOverlayPages.clear();
        invisibleTextOverlayPages.clear();
        docExtractionStruct = [];
        pageImageRects.clear();
        showImageHighlights = false;
        showCategoryHighlights = false;
        imageToolButton?.classList.remove('active');
        shapesToolButton?.classList.remove('active');
        selectionBox.style.display = 'none';
        hideSelectionMenu();

        setLoading('Loading PDF…');
        pdfDoc = await pdfjsLib.getDocument({ data: pdfData }).promise;
        setLoading('Analyzing document structure…');
        analysis = await analyzeDocument(pdfDoc);
        restoreOverridesForDocument();
        setLoading('');

        renderSummary(analysis);
        renderPageGroups(analysis.groups);
        renderOverrideSummary();
        populatePageSelect(pdfDoc.numPages);
        await renderPage(1);
        renderPageDetails(1);
        previewButton.disabled = false;
        extractButton.disabled = false;
        extractedImages = [];
        clearImageGrid();
        if (extractImagesButton) extractImagesButton.disabled = false;
        markdownButton.disabled = false;
        jsonButton.disabled = false;
        csvButton.disabled = false;
        prepareSearchPanel();
        analyzeButton.disabled = false;
      }

      function renderSummary(docAnalysis) {
        if (!docAnalysis) return;
        const tocSummary = docAnalysis.tableOfContentsPages.length
          ? `Detected on page(s): ${docAnalysis.tableOfContentsPages.join(', ')}`
          : 'Not detected';

        const groupCounts = docAnalysis.groups.reduce((acc, group) => {
          const formatted = formatGroupLabel(group.label);
          acc[formatted] = (acc[formatted] || 0) + 1;
          return acc;
        }, {});

        const groupList = Object.entries(groupCounts)
          .map(([label, count]) => `<li><strong>${label}</strong>: ${count} page(s)</li>`)
          .join('');

        const heuristicsUsed = docAnalysis.settingsUsed || heuristics;
        const heuristicsMarkup = `
          <p class="hint">
            <strong>Current heuristics</strong>: TOC ≥ ${Math.round(
              (heuristicsUsed.tocRatioThreshold || heuristicDefaults.tocRatioThreshold) * 100
            )}% of lines,
            Index cues ≥ ${heuristicsUsed.indexEntryMinimum || heuristicDefaults.indexEntryMinimum},
            Invisible width &lt; ${(heuristicsUsed.invisibleWidthThreshold || heuristicDefaults.invisibleWidthThreshold).toFixed(2)}
          </p>
        `;

        summaryEl.innerHTML = `
          <p><strong>Pages:</strong> ${docAnalysis.pageCount}</p>
          <p><strong>Table of contents:</strong> ${tocSummary}</p>
          <p><strong>Index pages detected:</strong> ${docAnalysis.indexPages.length ? docAnalysis.indexPages.join(', ') : 'Not detected'}</p>
          <div>
            <p><strong>Page distribution:</strong></p>
            <ul>${groupList || '<li>No pages analyzed</li>'}</ul>
          </div>
          ${heuristicsMarkup}
        `;
      }

      function renderPageGroups(groups) {
        if (!groups?.length) {
          pageGroupsEl.innerHTML = '<li>No pages analyzed</li>';
          return;
        }

        const ranges = [];
        groups.forEach((entry) => {
          const last = ranges[ranges.length - 1];
          if (last && last.label === entry.label && last.end === entry.page - 1) {
            last.end = entry.page;
          } else {
            ranges.push({ start: entry.page, end: entry.page, label: entry.label });
          }
        });

        pageGroupsEl.innerHTML = ranges
          .map((range) => {
            const label = formatGroupLabel(range.label);
            const rangeText = range.start === range.end ? `Page ${range.start}` : `Pages ${range.start}-${range.end}`;
            return `<li data-range-start="${range.start}" data-range-end="${range.end}">${rangeText}: ${label}</li>`;
          })
          .join('');
      }

      function populatePageSelect(pageCount) {
        pageSelectEl.disabled = false;
        pageSelectEl.min = '1';
        pageSelectEl.max = String(pageCount);
        pageSelectEl.step = '1';
        pageSelectEl.value = '1';
        prevButton.disabled = false;
        nextButton.disabled = false;
        pageIndicator.textContent = `Page 1 of ${pageCount}`;
      }

      function handlePageInput(event) {
        if (!pdfDoc) {
          event.target.value = '';
          return;
        }
        const value = Number(event.target.value);
        if (!Number.isFinite(value)) {
          event.target.value = String(currentPage);
          return;
        }
        const clamped = Math.min(Math.max(Math.round(value), 1), pdfDoc.numPages);
        event.target.value = String(clamped);
        goToPage(clamped);
      }

      async function goToPage(pageNumber) {
        if (!pdfDoc || pageNumber < 1 || pageNumber > pdfDoc.numPages) return;
        currentPage = pageNumber;
        pageSelectEl.value = String(pageNumber);
        pageIndicator.textContent = `Page ${pageNumber} of ${pdfDoc.numPages}`;
        selectionBox.style.display = 'none';
        hideSelectionMenu();
        renderPageDetails(pageNumber);
        await renderPage(pageNumber);
        prevButton.disabled = pageNumber === 1;
        nextButton.disabled = pageNumber === pdfDoc.numPages;
        renderCategoryHighlights(pageNumber);
        renderImageHighlights(pageNumber);
      }

      function focusPreviewPanel() {
        if (!pdfDoc) return;
        previewPanel?.scrollIntoView({ behavior: 'smooth', block: 'start' });
      }

      async function renderPage(pageNumber) {
        if (!pdfDoc) return;
        const page = await pdfDoc.getPage(pageNumber);
        const baseViewport = page.getViewport({ scale: 1 });
        const desiredHeight = Math.max(window.innerHeight || 0, 100);
        const scale = desiredHeight / (baseViewport.height || 1);
        const viewport = page.getViewport({ scale });
        const outputScale = window.devicePixelRatio || 1;
        currentViewportState = {
          baseWidth: baseViewport.width || 1,
          baseHeight: baseViewport.height || 1,
          scale,
          outputScale,
        };

        if (currentRenderTask) {
          try {
            await currentRenderTask.promise;
          } catch (err) {
            console.error('Previous render aborted', err);
          }
          currentRenderTask = null;
        }

        const displayWidth = Math.floor(viewport.width);
        const displayHeight = Math.floor(viewport.height);
        canvas.style.height = `${displayHeight}px`;
        canvas.style.width = `${displayWidth}px`;
        canvas.height = Math.floor(viewport.height * outputScale);
        canvas.width = Math.floor(viewport.width * outputScale);

        const transform = outputScale !== 1 ? [outputScale, 0, 0, outputScale, 0, 0] : null;

        const renderContext = {
          canvasContext: ctx,
          viewport,
          transform,
        };

        const detectedImageRects = [];
        const originalDrawImage = ctx.drawImage.bind(ctx);
        ctx.drawImage = function instrumentedDrawImage(image, ...args) {
          try {
            const rect = deriveDrawImageRect(args);
            const transformMatrix = ctx.getTransform();
            const quad = buildTransformedQuad(transformMatrix, rect);
            const bounds = quadToBounds(quad);
            if (bounds.width > 0 && bounds.height > 0) {
              const normalized = {
                x: clampNumber(bounds.x / canvas.width, 0, 1),
                y: clampNumber(bounds.y / canvas.height, 0, 1),
                width: clampNumber(bounds.width / canvas.width, 0, 1),
                height: clampNumber(bounds.height / canvas.height, 0, 1),
              };
              detectedImageRects.push(normalized);
            }
          } catch (error) {
            console.warn('Failed to record image bounds', error);
          }
          return originalDrawImage(image, ...args);
        };

        currentRenderTask = page.render(renderContext);
        try {
          await currentRenderTask.promise;
        } finally {
          ctx.drawImage = originalDrawImage;
          currentRenderTask = null;
        }
        pageImageRects.set(pageNumber, dedupeRects(detectedImageRects));
        drawColumnOverlay(pageNumber);
        drawTextAreasOverlay(pageNumber);
        drawHeaderFooterOverlay(pageNumber);
        drawInvisibleOverlay(pageNumber);
        renderImageHighlights(pageNumber);
        renderCategoryHighlights(pageNumber);
      }

      function handleResize() {
        if (!pdfDoc) return;
        renderPage(currentPage).catch((err) => console.error('Resize render failed', err));
        renderImageHighlights(currentPage);
        renderCategoryHighlights(currentPage);
      }

      function renderPageDetails(pageNumber) {
        if (!analysis) return;
        const details = analysis.pageDetails[pageNumber - 1];
        if (!details) return;

        const normalizedGroupLabel = formatGroupLabel(details.groupLabel);
        const headerFlagged = headerOverridePages.has(pageNumber);
        const footerFlagged = footerOverridePages.has(pageNumber);
        const headerDetected = Boolean(details.headerLine);
        const footerDetected = Boolean(details.footerLine);
        const columnCount = details.columnCount || 0;
        const selectedColumn = columnPreviewSelections.get(pageNumber) || 1;
        const previewActive = columnPreviewSelections.has(pageNumber);
        const hasTextAreas = Boolean(details.textAreas?.length);
        const textOverlayActive = textAreaOverlayPages.has(pageNumber);
        const headerFooterOverlayPossible = Boolean(details.headerBounds || details.footerBounds);
        const headerFooterOverlayActive = headerFooterOverlayPages.has(pageNumber);
        const invisibleOverlayPossible = Boolean(details.invisibleBoxes?.length);
        const invisibleOverlayActive = invisibleTextOverlayPages.has(pageNumber);
        let overlayStateChanged = false;
        if (!hasTextAreas && textAreaOverlayPages.delete(pageNumber)) {
          overlayStateChanged = true;
        }
        if (!headerFooterOverlayPossible && headerFooterOverlayPages.delete(pageNumber)) {
          overlayStateChanged = true;
        }
        if (!invisibleOverlayPossible && invisibleTextOverlayPages.delete(pageNumber)) {
          overlayStateChanged = true;
        }
        if (overlayStateChanged) {
          persistOverrides();
        }

        const columnControlsMarkup = columnCount
          ? `
            <div class="column-line">
              <p><strong>Estimated columns:</strong> ${columnCount}</p>
              <div class="column-controls">
                <label>
                  Column
                  <input type="number" min="1" max="${columnCount}" value="${Math.min(selectedColumn, columnCount)}" data-column-input />
                </label>
                <button type="button" data-column-preview>Show raster</button>
                <button type="button" data-column-clear ${previewActive ? '' : 'disabled'}>Clear</button>
              </div>
            </div>
            <p class="column-preview-status">
              ${
                previewActive
                  ? `Previewing column ${Math.min(selectedColumn, columnCount)} of ${columnCount}.`
                  : 'Enter a column number to overlay its raster on the preview.'
              }
            </p>
          `
          : `<p><strong>Estimated columns:</strong> ${columnCount}</p>`;

        const textOverlayControls = hasTextAreas
          ? `
            <div class="overlay-toggle-group">
              <label class="override-option">
                <input type="checkbox" data-text-overlay ${textOverlayActive ? 'checked' : ''} />
                Highlight text areas
              </label>
            </div>
          `
          : '';

        const overlayControlsMarkup = headerFooterOverlayPossible || invisibleOverlayPossible
          ? `
            <div class="overlay-toggle-group">
              ${
                headerFooterOverlayPossible
                  ? `<label class="override-option">
                      <input type="checkbox" data-header-footer-overlay ${headerFooterOverlayActive ? 'checked' : ''} />
                      Highlight header/footer cues
                    </label>`
                  : ''
              }
              ${
                invisibleOverlayPossible
                  ? `<label class="override-option">
                      <input type="checkbox" data-invisible-overlay ${invisibleOverlayActive ? 'checked' : ''} />
                      Highlight invisible text
                    </label>`
                  : ''
              }
            </div>
          `
          : '';

        pageDetailsEl.innerHTML = `
          <p><strong>Classification:</strong> ${normalizedGroupLabel}</p>
          <div class="detail-row">
            <p><strong>Detected header:</strong> ${details.headerLine || 'Not detected'}${
              headerFlagged ? ' <em>(flagged as normal text)</em>' : ''
            }</p>
            <div class="override-set">
              <label class="override-option">
                <input type="checkbox" data-header-override ${headerDetected ? '' : 'disabled'} ${
                  headerFlagged ? 'checked' : ''
                } />
                Treat as normal text
              </label>
              <label class="override-option">
                <input type="checkbox" data-header-remember ${headerDetected ? '' : 'disabled'} ${
                  headerRememberEnabled ? 'checked' : ''
                } />
                Remember for
                <select data-header-scope ${
                  headerDetected && headerRememberEnabled ? '' : 'disabled'
                }>
                  ${buildScopeOptions(headerRememberScope)}
                </select>
              </label>
            </div>
          </div>
          <div class="detail-row">
            <p><strong>Detected footer:</strong> ${details.footerLine || 'Not detected'}${
              footerFlagged ? ' <em>(flagged as normal text)</em>' : ''
            }</p>
            <div class="override-set">
              <label class="override-option">
                <input type="checkbox" data-footer-override ${footerDetected ? '' : 'disabled'} ${
                  footerFlagged ? 'checked' : ''
                } />
                Treat as normal text
              </label>
              <label class="override-option">
                <input type="checkbox" data-footer-remember ${footerDetected ? '' : 'disabled'} ${
                  footerRememberEnabled ? 'checked' : ''
                } />
                Remember for
                <select data-footer-scope ${
                  footerDetected && footerRememberEnabled ? '' : 'disabled'
                }>
                  ${buildScopeOptions(footerRememberScope)}
                </select>
              </label>
            </div>
          </div>
          ${columnControlsMarkup}
          ${textOverlayControls}
          ${overlayControlsMarkup}
          <p><strong>Invisible text items:</strong> ${details.invisibleItems}</p>
          <p><strong>Table of contents cues:</strong> ${(details.tocEntryRatio * 100).toFixed(0)}% of lines</p>
          <p><strong>Index cues:</strong> ${details.indexEntryCount} line(s)</p>
        `;

        attachDetailInteractions(pageNumber, details, normalizedGroupLabel);
      }

      function attachDetailInteractions(pageNumber, details, normalizedGroupLabel) {
        const headerCheckbox = pageDetailsEl.querySelector('[data-header-override]');
        const footerCheckbox = pageDetailsEl.querySelector('[data-footer-override]');
        const headerRememberToggle = pageDetailsEl.querySelector('[data-header-remember]');
        const footerRememberToggle = pageDetailsEl.querySelector('[data-footer-remember]');
        const headerScopeSelect = pageDetailsEl.querySelector('[data-header-scope]');
        const footerScopeSelect = pageDetailsEl.querySelector('[data-footer-scope]');
        const columnInput = pageDetailsEl.querySelector('[data-column-input]');
        const columnPreviewBtn = pageDetailsEl.querySelector('[data-column-preview]');
        const columnClearBtn = pageDetailsEl.querySelector('[data-column-clear]');
        const textOverlayToggle = pageDetailsEl.querySelector('[data-text-overlay]');
        const headerFooterOverlayToggle = pageDetailsEl.querySelector('[data-header-footer-overlay]');
        const invisibleOverlayToggle = pageDetailsEl.querySelector('[data-invisible-overlay]');

        headerCheckbox?.addEventListener('change', () => {
          const scope = headerRememberToggle?.checked ? headerScopeSelect?.value || 'page' : 'page';
          applyOverrideChange({
            type: 'header',
            pageNumber,
            isChecked: headerCheckbox.checked,
            scope,
            groupLabel: normalizedGroupLabel,
          });
        });

        footerCheckbox?.addEventListener('change', () => {
          const scope = footerRememberToggle?.checked ? footerScopeSelect?.value || 'page' : 'page';
          applyOverrideChange({
            type: 'footer',
            pageNumber,
            isChecked: footerCheckbox.checked,
            scope,
            groupLabel: normalizedGroupLabel,
          });
        });

        headerRememberToggle?.addEventListener('change', () => {
          headerRememberEnabled = headerRememberToggle.checked;
          if (headerScopeSelect) headerScopeSelect.disabled = !headerRememberToggle.checked;
        });

        footerRememberToggle?.addEventListener('change', () => {
          footerRememberEnabled = footerRememberToggle.checked;
          if (footerScopeSelect) footerScopeSelect.disabled = !footerRememberToggle.checked;
        });

        headerScopeSelect?.addEventListener('change', () => {
          headerRememberScope = headerScopeSelect.value;
        });

        footerScopeSelect?.addEventListener('change', () => {
          footerRememberScope = footerScopeSelect.value;
        });

        columnPreviewBtn?.addEventListener('click', async () => {
          if (!columnInput || !details.columnCount) return;
          let column = Number(columnInput.value);
          if (!Number.isFinite(column)) column = 1;
          column = Math.min(Math.max(Math.round(column), 1), details.columnCount);
          columnPreviewSelections.set(pageNumber, column);
          await renderPage(currentPage);
          renderPageDetails(pageNumber);
          persistOverrides();
        });

        columnClearBtn?.addEventListener('click', async () => {
          columnPreviewSelections.delete(pageNumber);
          await renderPage(currentPage);
          renderPageDetails(pageNumber);
          persistOverrides();
        });

        textOverlayToggle?.addEventListener('change', async () => {
          if (textOverlayToggle.checked) {
            textAreaOverlayPages.add(pageNumber);
          } else {
            textAreaOverlayPages.delete(pageNumber);
          }
          await renderPage(currentPage);
          persistOverrides();
        });

        headerFooterOverlayToggle?.addEventListener('change', async () => {
          if (headerFooterOverlayToggle.checked) {
            headerFooterOverlayPages.add(pageNumber);
          } else {
            headerFooterOverlayPages.delete(pageNumber);
          }
          await renderPage(currentPage);
          persistOverrides();
        });

        invisibleOverlayToggle?.addEventListener('change', async () => {
          if (invisibleOverlayToggle.checked) {
            invisibleTextOverlayPages.add(pageNumber);
          } else {
            invisibleTextOverlayPages.delete(pageNumber);
          }
          await renderPage(currentPage);
          persistOverrides();
        });
      }

      function applyOverrideChange({ type, pageNumber, isChecked, scope, groupLabel }) {
        const targetSet = type === 'header' ? headerOverridePages : footerOverridePages;
        const targetPages = getPagesForScope(scope, pageNumber, groupLabel);
        targetPages.forEach((page) => {
          if (isChecked) {
            targetSet.add(page);
          } else {
            targetSet.delete(page);
          }
        });

        renderPageDetails(pageNumber);
        renderOverrideSummary();
        persistOverrides();
      }

      function getPagesForScope(scope, pageNumber, groupLabel) {
        if (!analysis) return [pageNumber];
        if (scope === 'document') {
          return analysis.pageDetails.map((detail) => detail.pageNumber);
        }
        if (scope === 'group') {
          return analysis.pageDetails
            .filter((detail) => formatGroupLabel(detail.groupLabel) === groupLabel)
            .map((detail) => detail.pageNumber);
        }
        return [pageNumber];
      }

      function renderOverrideSummary() {
        if (!overrideSummaryEl) return;
        if (!analysis) {
          overrideSummaryEl.innerHTML = '<p>No header/footer overrides applied yet.</p>';
          return;
        }

        const headerRanges = buildOverrideRanges([...headerOverridePages]);
        const footerRanges = buildOverrideRanges([...footerOverridePages]);

        if (!headerRanges.length && !footerRanges.length) {
          overrideSummaryEl.innerHTML = '<p>No header/footer overrides applied yet.</p>';
          return;
        }

        const headerMarkup = headerRanges.length
          ? `<div><strong>Header overrides</strong>${buildRangeList(headerRanges)}</div>`
          : '';
        const footerMarkup = footerRanges.length
          ? `<div><strong>Footer overrides</strong>${buildRangeList(footerRanges)}</div>`
          : '';

        overrideSummaryEl.innerHTML = `${headerMarkup}${footerMarkup}` || '<p>No header/footer overrides applied yet.</p>';
      }

      function buildOverrideRanges(pages = []) {
        if (!analysis || !pages.length) return [];
        const sorted = [...pages].sort((a, b) => a - b);
        const ranges = [];
        sorted.forEach((page) => {
          const detail = analysis.pageDetails[page - 1];
          const label = formatGroupLabel(detail?.groupLabel || '');
          const last = ranges[ranges.length - 1];
          if (last && last.end === page - 1 && last.label === label) {
            last.end = page;
          } else {
            ranges.push({ start: page, end: page, label });
          }
        });
        return ranges;
      }

      function buildRangeList(ranges = []) {
        if (!ranges.length) return '';
        const items = ranges
          .map((range) => {
            const pageLabel = range.start === range.end ? `Page ${range.start}` : `Pages ${range.start}-${range.end}`;
            return `<li>${pageLabel}: ${range.label}</li>`;
          })
          .join('');
        return `<ul>${items}</ul>`;
      }

      function buildScopeOptions(selectedValue = 'page') {
        return `
          <option value="page" ${selectedValue === 'page' ? 'selected' : ''}>Only this page</option>
          <option value="group" ${selectedValue === 'group' ? 'selected' : ''}>All pages in this group</option>
          <option value="document" ${selectedValue === 'document' ? 'selected' : ''}>All pages in the document</option>
        `;
      }

      function persistOverrides() {
        if (!pdfFileName) return;
        const data = readFromStorage(STORAGE_KEYS.overrides, {});
        data[pdfFileName] = {
          header: [...headerOverridePages],
          footer: [...footerOverridePages],
          columnPreview: Object.fromEntries(columnPreviewSelections),
          textAreas: [...textAreaOverlayPages],
          headerFooter: [...headerFooterOverlayPages],
          invisible: [...invisibleTextOverlayPages],
        };
        writeToStorage(STORAGE_KEYS.overrides, data);
      }

      function restoreOverridesForDocument() {
        headerOverridePages.clear();
        footerOverridePages.clear();
        columnPreviewSelections.clear();
        textAreaOverlayPages.clear();
        headerFooterOverlayPages.clear();
        invisibleTextOverlayPages.clear();

        const data = readFromStorage(STORAGE_KEYS.overrides, {});
        const record = data?.[pdfFileName];
        if (!record) return;
        (record.header || []).forEach((page) => headerOverridePages.add(page));
        (record.footer || []).forEach((page) => footerOverridePages.add(page));
        if (record.columnPreview) {
          Object.entries(record.columnPreview).forEach(([page, column]) => {
            const pageNumber = Number(page);
            const columnNumber = Number(column);
            if (Number.isFinite(pageNumber) && Number.isFinite(columnNumber)) {
              columnPreviewSelections.set(pageNumber, columnNumber);
            }
          });
        }
        (record.textAreas || []).forEach((page) => textAreaOverlayPages.add(page));
        (record.headerFooter || []).forEach((page) => headerFooterOverlayPages.add(page));
        (record.invisible || []).forEach((page) => invisibleTextOverlayPages.add(page));
      }

      function drawColumnOverlay(pageNumber) {
        const selection = columnPreviewSelections.get(pageNumber);
        const details = analysis?.pageDetails?.[pageNumber - 1];
        if (!selection || !details?.columnCount) return;

        const columnCount = details.columnCount;
        const columnIndex = Math.min(Math.max(selection, 1), columnCount);
        const columnWidth = canvas.width / columnCount;
        const x = columnWidth * (columnIndex - 1);

        ctx.save();
        ctx.fillStyle = 'rgba(18, 89, 255, 0.2)';
        ctx.fillRect(x, 0, columnWidth, canvas.height);
        ctx.strokeStyle = 'rgba(18, 89, 255, 0.8)';
        ctx.lineWidth = 2;
        ctx.strokeRect(x, 0, columnWidth, canvas.height);
        ctx.restore();
      }

      function drawTextAreasOverlay(pageNumber) {
        if (!analysis || !textAreaOverlayPages.has(pageNumber)) return;
        const details = analysis.pageDetails[pageNumber - 1];
        const areas = details?.textAreas;
        if (!areas?.length) return;

        ctx.save();
        ctx.fillStyle = 'rgba(255, 165, 0, 0.18)';
        ctx.strokeStyle = 'rgba(255, 165, 0, 0.9)';
        ctx.lineWidth = 1;

        areas.forEach((area) => {
          const x = (area.x || 0) * canvas.width;
          const y = (area.y || 0) * canvas.height;
          const width = (area.width || 0) * canvas.width;
          const height = (area.height || 0) * canvas.height;
          if (!width || !height) return;
          ctx.fillRect(x, y, width, height);
          ctx.strokeRect(x, y, width, height);
        });

        ctx.restore();
      }

      function drawHeaderFooterOverlay(pageNumber) {
        if (!analysis || !headerFooterOverlayPages.has(pageNumber)) return;
        const details = analysis.pageDetails[pageNumber - 1];
        const boxes = [];
        if (details?.headerBounds) boxes.push(details.headerBounds);
        if (details?.footerBounds) boxes.push(details.footerBounds);
        if (!boxes.length) return;

        ctx.save();
        ctx.strokeStyle = 'rgba(76, 175, 80, 0.8)';
        ctx.fillStyle = 'rgba(76, 175, 80, 0.15)';
        ctx.lineWidth = 2;

        boxes.forEach((box) => {
          const x = (box.x || 0) * canvas.width;
          const y = (box.y || 0) * canvas.height;
          const width = (box.width || 0) * canvas.width;
          const height = (box.height || 0.02) * canvas.height;
          if (!width || !height) return;
          ctx.fillRect(x, y, width, height);
          ctx.strokeRect(x, y, width, height);
        });

        ctx.restore();
      }

      function drawInvisibleOverlay(pageNumber) {
        if (!analysis || !invisibleTextOverlayPages.has(pageNumber)) return;
        const details = analysis.pageDetails[pageNumber - 1];
        const boxes = details?.invisibleBoxes;
        if (!boxes?.length) return;

        ctx.save();
        ctx.strokeStyle = 'rgba(244, 67, 54, 0.85)';
        ctx.fillStyle = 'rgba(244, 67, 54, 0.2)';
        ctx.lineWidth = 1;

        boxes.forEach((box) => {
          const x = (box.x || 0) * canvas.width;
          const y = (box.y || 0) * canvas.height;
          const width = (box.width || 0) * canvas.width;
          const height = (box.height || 0) * canvas.height || 1;
          if (!width || !height) return;
          ctx.fillRect(x, y, width, height);
          ctx.strokeRect(x, y, width, height);
        });

        ctx.restore();
      }

      function formatGroupLabel(label = '') {
        const lower = label.toLowerCase();
        if (lower.includes('table')) {
          return 'Mainly TOC';
        }
        if (lower.includes('index')) {
          return 'Mainly Index';
        }
        return 'Mainly Text';
      }
      // Analyzes the pdf Structure to extract text and other statistics about the document.
      async function analyzeDocument(pdf) {
        const summary = {
          pageCount: pdf.numPages,
          tableOfContentsPages: [],
          indexPages: [],
          groups: [],
          pageDetails: [],
          textContent: [],
        };

        const tocCutoff = heuristics.tocRatioThreshold || heuristicDefaults.tocRatioThreshold;
        const indexMinimum = heuristics.indexEntryMinimum || heuristicDefaults.indexEntryMinimum;

        for (let pageNumber = 1; pageNumber <= pdf.numPages; pageNumber += 1) {
          setLoading(`Analyzing page ${pageNumber} of ${pdf.numPages}…`);
          // Reads the page from the PDF
          const page = await pdf.getPage(pageNumber);
          const viewport = page.getViewport({ scale: 1 });
          // Reads the text content from the PDF
          const textContent = await page.getTextContent({ normalizeWhitespace: true });
          const items = textContent.items
            .filter((item) => item.str && item.str.trim())
            .map((item) => {
              const width = item.width ?? Math.abs(item.transform[0]) ?? 0;
              const height = Math.abs(item.transform[3]) || Math.abs(item.transform[1]) || item.height || 0;
              const fontSize = height;
              // objTransformed = PDFJS.Util.transform(viewPort.transform, objText.transform);  // Translate objText transform to correct ViewPort x, y coordinates.
              const rect = viewport.convertToViewportRectangle([
                item.transform[4],
                item.transform[5],
                item.transform[4] + width,
                item.transform[5] - height,
              ]);
              const [rx1 = 0, ry1 = 0, rx2 = 0, ry2 = 0] = rect;
              const rectX = Math.min(rx1, rx2);
              const rectY = Math.min(ry1, ry2);
              const rectWidth = Math.abs(rx1 - rx2);
              const rectHeight = Math.abs(ry1 - ry2);
              return {
                text: item.str.trim(),
                x: item.transform[4],
                y: item.transform[5],
                width,
                height,
                fontSize,
                normalizedX: viewport.width ? rectX / viewport.width : 0,
                normalizedY: viewport.height ? rectY / viewport.height : 0,
                normalizedWidth: viewport.width ? rectWidth / viewport.width : 0,
                normalizedHeight: viewport.height ? rectHeight / viewport.height : 0,
              };
            });

          const details = evaluatePage(items, viewport, pageNumber);
          const { readingOrder, ...pageMetrics } = details;
          const safeReadingOrder = readingOrder.map((line) => ({
            text: line.text,
            headingLevel: line.headingLevel,
            blockType: line.blockType,
            pageNumber: line.pageNumber,
          }));
          summary.groups.push({ page: pageNumber, label: details.groupLabel });
          summary.pageDetails.push(pageMetrics);
          summary.textContent.push({ pageNumber, lines: safeReadingOrder });

          if (details.hasTableOfContentsHeading || details.tocEntryRatio >= tocCutoff) {
            summary.tableOfContentsPages.push(pageNumber);
          }
          if (details.hasIndexHeading || details.indexEntryCount >= indexMinimum) {
            summary.indexPages.push(pageNumber);
          }
        }

        summary.settingsUsed = { ...heuristics };
        summary.searchIndex = buildSearchIndex(summary.textContent);
        summary.entities = extractEntities(summary.textContent);
        return summary;
      }
      
      // Evaluates one page of the document. EHE
      function evaluatePage(items, viewport, pageNumber) {
        const lines = buildLines(items);
        const pageWidth = viewport.width || 1;
        const columnCount = annotateColumnAssignments(lines, pageWidth);
        const averageFontSize = lines.length
          ? lines.reduce((sum, line) => sum + (line.fontSize || 0), 0) / lines.length
          : 0;

        annotateHeadingLevels(lines, averageFontSize);

        const headerCandidate = lines.find((line) => line.relativeY > 0.85);
        if (headerCandidate) headerCandidate.blockType = 'header';
        const footerCandidate = [...lines].reverse().find((line) => line.relativeY < 0.15);
        if (footerCandidate && footerCandidate !== headerCandidate) {
          footerCandidate.blockType = 'footer';
        }

        const headerLine = headerCandidate?.text ?? '';
        const footerLine = footerCandidate?.text ?? '';
        const headerBounds = buildLineBounds(headerCandidate);
        const footerBounds = footerCandidate && footerCandidate !== headerCandidate ? buildLineBounds(footerCandidate) : null;

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
        const invisibleThreshold = heuristics.invisibleWidthThreshold || heuristicDefaults.invisibleWidthThreshold;
        const invisibleBoxes = items
          .filter((item) => (item.width || 0) < invisibleThreshold)
          .map((item) => ({
            x: item.normalizedX || 0,
            y: item.normalizedY || 0,
            width: item.normalizedWidth || 0,
            height: item.normalizedHeight || 0,
          }))
          .filter((box) => box.width > 0 && box.height > 0);
        const invisibleItems = invisibleBoxes.length;
        const textAreas = items
          .map((item) => ({
            x: item.normalizedX || 0,
            y: item.normalizedY || 0,
            width: item.normalizedWidth || 0,
            height: item.normalizedHeight || 0,
          }))
          .filter((area) => area.width > 0 && area.height > 0);
        const readingOrder = buildReadingOrder(lines, pageNumber);

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
          textAreas,
          headerBounds,
          footerBounds,
          invisibleBoxes,
          readingOrder,
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
            const rowItems = row.slice();
            const sorted = rowItems.slice().sort((a, b) => a.x - b.x);
            const text = sorted.map((item) => item.text).join(' ').replace(/\s{2,}/g, ' ').trim();
            const y = key * 2;
            const minX = sorted.length ? sorted[0].x : 0;
            const averageFontSize = sorted.length
              ? sorted.reduce((sum, item) => sum + (item.fontSize || item.height || 0), 0) / sorted.length
              : 0;
            return {
              text,
              y,
              minX,
              fontSize: averageFontSize,
              columnIndex: 0,
              headingLevel: 0,
              blockType: 'body',
              items: rowItems,
            };
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

      function buildLineBounds(line) {
        if (!line?.items?.length) return null;
        let minX = 1;
        let minY = 1;
        let maxX = 0;
        let maxY = 0;
        line.items.forEach((item) => {
          const startX = item.normalizedX ?? 0;
          const endX = startX + (item.normalizedWidth || 0);
          const startY = item.normalizedY ?? 0;
          const endY = startY + (item.normalizedHeight || 0);
          minX = Math.min(minX, startX);
          minY = Math.min(minY, startY);
          maxX = Math.max(maxX, endX);
          maxY = Math.max(maxY, endY);
        });
        const width = Math.max(maxX - minX, 0.01);
        const height = Math.max(maxY - minY, 0.01);
        return {
          x: clampNumber(minX, 0, 1),
          y: clampNumber(minY, 0, 1),
          width: clampNumber(width, 0.01, 1),
          height: clampNumber(height, 0.01, 1),
        };
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
        const tocCutoff = heuristics.tocRatioThreshold || heuristicDefaults.tocRatioThreshold;
        const indexMinimum = heuristics.indexEntryMinimum || heuristicDefaults.indexEntryMinimum;
        if (hasTocHeading || tocEntryRatio >= tocCutoff) {
          return 'mainly Table of contents';
        }
        if (hasIndexHeading || indexEntryCount >= indexMinimum) {
          return 'mainly index';
        }
        return 'mainly text';
      }

      function annotateColumnAssignments(lines, pageWidth) {
        if (!lines.length || !pageWidth) return 0;
        const threshold = pageWidth * 0.15;
        const centers = [];

        lines.forEach((line) => {
          const existing = centers.find((center) => Math.abs(center.x - line.minX) <= threshold);
          if (existing) {
            existing.lines.push(line);
            const total = existing.lines.length;
            existing.x = ((existing.x * (total - 1)) + line.minX) / total;
          } else {
            centers.push({ x: line.minX, lines: [line] });
          }
        });

        centers.sort((a, b) => a.x - b.x);
        centers.forEach((center, index) => {
          center.lines.forEach((line) => {
            line.columnIndex = index;
          });
        });

        return centers.length || 0;
      }

      function buildReadingOrder(lines, pageNumber) {
        if (!lines.length) return [];
        const columnGroups = new Map();
        lines.forEach((line) => {
          const columnIndex = Number.isFinite(line.columnIndex) ? line.columnIndex : 0;
          const group = columnGroups.get(columnIndex) || [];
          group.push(line);
          columnGroups.set(columnIndex, group);
        });

        const orderedLines = [];
        [...columnGroups.keys()]
          .sort((a, b) => a - b)
          .forEach((columnIndex) => {
            const group = columnGroups.get(columnIndex) || [];
            group
              .slice()
              .sort((a, b) => b.y - a.y)
              .forEach((line) => {
                orderedLines.push({ ...line, pageNumber });
              });
          });

        return orderedLines;
      }

      function buildSearchIndex(pages = []) {
        return pages.map((page) => {
          const text = (page.lines || [])
            .map((line) => line.text || '')
            .join(' ')
            .replace(/\s+/g, ' ')
            .trim();
          return {
            pageNumber: page.pageNumber,
            text,
            textLower: text.toLowerCase(),
          };
        });
      }

      const entityStopWords = new Set([
        'section',
        'article',
        'page',
        'exhibit',
        'schedule',
        'figure',
      ]);

      function extractEntities(pages = []) {
        const counts = new Map();
        pages.forEach((page) => {
          const body = (page.lines || []).map((line) => line.text || '').join(' ');
          const namePattern = /\b([A-Z][a-z]+(?:\s+[A-Z][a-z]+){0,3})\b/g;
          const acronymPattern = /\b([A-Z]{3,})\b/g;
          let match;
          while ((match = namePattern.exec(body))) {
            registerEntity(counts, match[1], page.pageNumber);
          }
          while ((match = acronymPattern.exec(body))) {
            registerEntity(counts, match[1], page.pageNumber);
          }
        });
        return [...counts.values()].sort((a, b) => b.count - a.count).slice(0, 25);
      }

      function registerEntity(map, rawText, pageNumber) {
        const normalized = rawText.replace(/\s+/g, ' ').trim();
        if (normalized.length < 3) return;
        const lower = normalized.toLowerCase();
        if (entityStopWords.has(lower)) return;
        const entry = map.get(lower) || { text: normalized, count: 0, firstPage: pageNumber };
        entry.count += 1;
        entry.firstPage = Math.min(entry.firstPage, pageNumber);
        map.set(lower, entry);
      }

      function annotateHeadingLevels(lines, averageFontSize) {
        lines.forEach((line) => {
          line.headingLevel = detectHeadingLevel(line, averageFontSize);
        });
      }

      function detectHeadingLevel(line, averageFontSize) {
        const text = line.text?.trim() || '';
        if (!text) return 0;

        const numberedMatch = text.match(/^(\d+(?:\.\d+){0,5})[\s\-\)]/);
        if (numberedMatch) {
          const depth = numberedMatch[1].split('.').length;
          return Math.min(Math.max(depth, 1), 6);
        }

        const romanMatch = text.match(/^(?:[IVXLCDM]+\.|[A-Z]\))\s+/);
        if (romanMatch) {
          return 2;
        }

        const relativeFont = averageFontSize ? (line.fontSize || averageFontSize) / averageFontSize : 1;
        const letterMatches = text.match(/[A-Za-z]/g) || [];
        const uppercaseMatches = text.match(/[A-Z]/g) || [];
        const uppercaseRatio = letterMatches.length ? uppercaseMatches.length / letterMatches.length : 0;
        const isShort = text.length <= 80;

        if ((relativeFont > 1.25 && isShort) || (uppercaseRatio > 0.6 && isShort)) {
          return 2;
        }

        if (text.endsWith(':') && isShort) {
          return 3;
        }

        return 0;
      }

      function exportAnalysisReport() {
        if (!analysis) return;
        const markdown = buildMarkdownReport(analysis);
        const blob = new Blob([markdown], { type: 'text/markdown' });
        const url = URL.createObjectURL(blob);
        const link = document.createElement('a');
        link.href = url;
        link.download = `${pdfFileName || 'pdf-analysis'}-structure.md`;
        document.body.appendChild(link);
        link.click();
        document.body.removeChild(link);
        URL.revokeObjectURL(url);
      }

      function exportMarkdownDocument() {
        if (!analysis?.textContent?.length) return;
        const markdown = buildReadingOrderMarkdown({
          pages: analysis.textContent,
          title: pdfFileName || 'PDF Markdown Export',
          headerOverrides: headerOverridePages,
          footerOverrides: footerOverridePages,
        });
        const blob = new Blob([markdown], { type: 'text/markdown' });
        const url = URL.createObjectURL(blob);
        const link = document.createElement('a');
        link.href = url;
        link.download = `${pdfFileName || 'pdf-analysis'}-text.md`;
        document.body.appendChild(link);
        link.click();
        document.body.removeChild(link);
        URL.revokeObjectURL(url);
      }

      function exportJsonAnalysis() {
        if (!analysis) return;
        const payload = {
          fileName: pdfFileName,
          generatedAt: new Date().toISOString(),
          heuristics,
          analysis,
        };
        const blob = new Blob([JSON.stringify(payload, null, 2)], { type: 'application/json' });
        const url = URL.createObjectURL(blob);
        const link = document.createElement('a');
        link.href = url;
        link.download = `${pdfFileName || 'pdf-analysis'}-analysis.json`;
        document.body.appendChild(link);
        link.click();
        document.body.removeChild(link);
        URL.revokeObjectURL(url);
      }

      function exportCsvSummary() {
        if (!analysis?.pageDetails?.length) return;
        const headers = [
          'page',
          'classification',
          'header',
          'footer',
          'columns',
          'invisible_text_items',
          'toc_ratio',
          'index_cues',
        ];
        const rows = analysis.pageDetails.map((detail) => [
          detail.pageNumber,
          formatGroupLabel(detail.groupLabel),
          detail.headerLine || '',
          detail.footerLine || '',
          detail.columnCount || 0,
          detail.invisibleItems || 0,
          (detail.tocEntryRatio * 100).toFixed(1),
          detail.indexEntryCount || 0,
        ]);
        const csv = [headers, ...rows]
          .map((row) => row.map((value) => `"${String(value).replace(/"/g, '""')}"`).join(','))
          .join('\n');
        const blob = new Blob([csv], { type: 'text/csv' });
        const url = URL.createObjectURL(blob);
        const link = document.createElement('a');
        link.href = url;
        link.download = `${pdfFileName || 'pdf-analysis'}-summary.csv`;
        document.body.appendChild(link);
        link.click();
        document.body.removeChild(link);
        URL.revokeObjectURL(url);
      }

      function buildMarkdownReport(docAnalysis) {
        const tocPages = docAnalysis.tableOfContentsPages.length
          ? docAnalysis.tableOfContentsPages.join(', ')
          : 'Not detected';
        const indexPages = docAnalysis.indexPages.length ? docAnalysis.indexPages.join(', ') : 'Not detected';

        let markdown = `# PDF Analysis Report\n\n`;
        markdown += `- **Pages**: ${docAnalysis.pageCount}\n`;
        markdown += `- **Table of contents**: ${tocPages}\n`;
        markdown += `- **Index pages**: ${indexPages}\n\n`;

        markdown += `## Page Groups\n`;
        docAnalysis.groups.forEach((group) => {
          markdown += `- Page ${group.page}: ${formatGroupLabel(group.label)}\n`;
        });

        markdown += `\n## Page Details\n`;
        docAnalysis.pageDetails.forEach((detail) => {
          markdown += `\n### Page ${detail.pageNumber}\n`;
          markdown += `- Classification: ${formatGroupLabel(detail.groupLabel)}\n`;
          const headerText = detail.headerLine || 'Not detected';
          const footerText = detail.footerLine || 'Not detected';
          const headerNote =
            detail.headerLine && headerOverridePages.has(detail.pageNumber)
              ? `${detail.headerLine} (flagged as normal text)`
              : headerText;
          const footerNote =
            detail.footerLine && footerOverridePages.has(detail.pageNumber)
              ? `${detail.footerLine} (flagged as normal text)`
              : footerText;
          markdown += `- Header: ${headerNote}\n`;
          markdown += `- Footer: ${footerNote}\n`;
          markdown += `- Estimated columns: ${detail.columnCount || 0}\n`;
          markdown += `- Invisible text items: ${detail.invisibleItems}\n`;
          markdown += `- TOC cues: ${(detail.tocEntryRatio * 100).toFixed(0)}% of lines\n`;
          markdown += `- Index cues: ${detail.indexEntryCount}\n`;
        });

        return markdown;
      }

      function buildReadingOrderMarkdown({ pages = [], title = 'PDF Markdown Export', headerOverrides, footerOverrides } = {}) {
        const sections = [];
        const contentBlocks = [];
        const anchorCounts = new Map();
        let paragraphBuffer = '';

        const headerSet = headerOverrides instanceof Set ? headerOverrides : new Set(headerOverrides);
        const footerSet = footerOverrides instanceof Set ? footerOverrides : new Set(footerOverrides);

        function flushParagraph() {
          const cleaned = paragraphBuffer.trim();
          if (cleaned) {
            contentBlocks.push(cleaned);
          }
          paragraphBuffer = '';
        }

        function registerAnchor(text) {
          const base = slugifyHeading(text) || `section-${sections.length + 1}`;
          const count = anchorCounts.get(base) || 0;
          anchorCounts.set(base, count + 1);
          return count ? `${base}-${count}` : base;
        }

        pages.forEach((page) => {
          const allowHeader = headerSet.has(page.pageNumber);
          const allowFooter = footerSet.has(page.pageNumber);
          (page.lines || []).forEach((line) => {
            const blockType = line.blockType || 'body';
            if (!allowHeader && blockType === 'header') return;
            if (!allowFooter && blockType === 'footer') return;

            const text = cleanMarkdownText(line.text);
            if (!text) return;

            if (line.headingLevel > 0) {
              flushParagraph();
              const level = Math.min(line.headingLevel, 6);
              const anchor = registerAnchor(text);
              sections.push({ title: text, level, anchor });
              contentBlocks.push(`<a id="${anchor}"></a>\n${'#'.repeat(level)} ${text}`);
            } else {
              paragraphBuffer = appendParagraphLine(paragraphBuffer, text);
            }
          });
          flushParagraph();
        });

        const headingTitle = `# ${title || 'PDF Markdown Export'}`;
        const toc = buildTableOfContents(sections);
        const body = contentBlocks.join('\n\n');
        return `${headingTitle}\n\n${toc}${body}\n`;
      }

      function appendParagraphLine(buffer, text) {
        const cleaned = text.trim();
        if (!buffer) {
          return cleaned;
        }
        const trimmedBuffer = buffer.trim();
        if (!trimmedBuffer) {
          return cleaned;
        }
        if (trimmedBuffer.endsWith('-') && !trimmedBuffer.endsWith('--')) {
          return `${trimmedBuffer.slice(0, -1)}${cleaned}`;
        }
        return `${trimmedBuffer} ${cleaned}`.replace(/\s+/g, ' ').trim();
      }

      function cleanMarkdownText(text = '') {
        return text
          .replace(/[\u2022\u2023\u25E6\u2043]/g, '- ')
          .replace(/\s+/g, ' ')
          .replace(/\s+([,.;:])/g, '$1')
          .trim();
      }

      function buildTableOfContents(sections = []) {
        const heading = '## Table of Contents\n';
        if (!sections.length) {
          return `${heading}- _No headings detected_\n\n`;
        }
        const lines = sections
          .map((section) => {
            const indent = '  '.repeat(Math.max(section.level - 1, 0));
            return `${indent}- [${section.title}](#${section.anchor})`;
          })
          .join('\n');
        return `${heading}${lines}\n\n`;
      }

      function slugifyHeading(text = '') {
        return text
          .toLowerCase()
          .replace(/[^a-z0-9\s-]/g, '')
          .trim()
          .replace(/\s+/g, '-')
          .replace(/-+/g, '-');
      }
