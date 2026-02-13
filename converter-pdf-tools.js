/**
 * PDF 結合・分割 コンバーター
 * pdf-lib を使用。結合と分割（ページ抽出）のサブモード切替
 * ページレベルの並べ替え・選択機能付き
 */
const PdfToolsConverter = (() => {
  const panel = Utils.$('panel-pdf-tools');
  const { PDFDocument } = PDFLib;

  let _pageUidCounter = 0;
  function nextPageUid() { return ++_pageUidCounter; }

  // ── サブモード管理 ──
  const submodeToggle = panel.querySelector('.submode-toggle');
  const submodeButtons = panel.querySelectorAll('.submode-btn');

  function initSubmode() {
    submodeButtons.forEach((btn, idx) => {
      btn.addEventListener('click', () => {
        const mode = btn.dataset.submode;
        submodeButtons.forEach((b) => b.classList.remove('active'));
        btn.classList.add('active');
        submodeToggle.dataset.activeIndex = idx;

        panel.querySelectorAll('.subpanel').forEach((sp) => sp.classList.add('hidden'));
        Utils.$(`subpanel-${mode}`).classList.remove('hidden');
      });
    });
  }

  // ==========================================
  // PDF 結合
  // ==========================================
  const mergeState = {
    pdfs: [],    // { id, name, size, data (Uint8Array), pageCount, expanded }
    pages: [],   // { uid, pdfId, pageIndex (0-based), label }
    idCounter: 0,
    previewUrl: null,
  };

  let mergeOptions = null;

  function buildPagesForPdf(pdf) {
    const pages = [];
    for (let i = 0; i < pdf.pageCount; i++) {
      pages.push({
        uid: nextPageUid(),
        pdfId: pdf.id,
        pageIndex: i,
        label: `${pdf.name} - P${i + 1}`,
      });
    }
    return pages;
  }

  async function handleMergeFiles(files) {
    Loading.show('PDFを読み込み中...');
    try {
      for (const file of files) {
        const data = new Uint8Array(await file.arrayBuffer());
        const pdfDoc = await PDFDocument.load(data, { ignoreEncryption: true });
        const pdf = {
          id: mergeState.idCounter++,
          name: file.name,
          size: file.size,
          data,
          pageCount: pdfDoc.getPageCount(),
          expanded: false,
        };
        mergeState.pdfs.push(pdf);
        const newPages = buildPagesForPdf(pdf);
        mergeState.pages.push(...newPages);
      }
      showMergeUI();
    } catch (err) {
      Toast.show('PDFの読み込みに失敗しました', 'error');
      console.error(err);
    } finally {
      Loading.hide();
    }
  }

  function showMergeUI() {
    if (mergeState.pdfs.length === 0) return;
    Utils.$('dropZone-pdf-merge').classList.add('compact');
    Utils.$('pdfList-pdf-merge').classList.remove('hidden');
    if (mergeOptions) mergeOptions.show();
    Utils.$('actionsPanel-pdf-merge').classList.remove('hidden');
    renderMergeList();
    debouncedMergePreview();
  }

  function resetMerge() {
    mergeState.pdfs = [];
    mergeState.pages = [];
    mergeState.idCounter = 0;
    mergeState.previewUrl = Utils.revokeBlobUrl(mergeState.previewUrl);

    Utils.$('dropZone-pdf-merge').classList.remove('compact');
    Utils.$('pdfList-pdf-merge').classList.add('hidden');
    if (mergeOptions) mergeOptions.hide();
    Utils.$('previewSection-pdf-merge').classList.add('hidden');
    Utils.$('actionsPanel-pdf-merge').classList.add('hidden');
    Utils.$('pdfListContainer-pdf-merge').innerHTML = '';
  }

  // ── ドラッグハンドル用SVG ──
  const DRAG_HANDLE_SVG = `<svg width="14" height="14" viewBox="0 0 14 14" fill="none">
    <circle cx="4" cy="3" r="1.2" fill="currentColor"/>
    <circle cx="10" cy="3" r="1.2" fill="currentColor"/>
    <circle cx="4" cy="7" r="1.2" fill="currentColor"/>
    <circle cx="10" cy="7" r="1.2" fill="currentColor"/>
    <circle cx="4" cy="11" r="1.2" fill="currentColor"/>
    <circle cx="10" cy="11" r="1.2" fill="currentColor"/>
  </svg>`;
  const UP_SVG = `<svg width="14" height="14" viewBox="0 0 14 14" fill="none"><path d="M7 2v10M3 6l4-4 4 4" stroke="currentColor" stroke-width="1.5" stroke-linecap="round" stroke-linejoin="round"/></svg>`;
  const DOWN_SVG = `<svg width="14" height="14" viewBox="0 0 14 14" fill="none"><path d="M7 12V2M3 8l4 4 4-4" stroke="currentColor" stroke-width="1.5" stroke-linecap="round" stroke-linejoin="round"/></svg>`;
  const REMOVE_SVG = `<svg width="14" height="14" viewBox="0 0 14 14" fill="none"><path d="M3 3l8 8M11 3l-8 8" stroke="currentColor" stroke-width="1.5" stroke-linecap="round"/></svg>`;
  const EXPAND_SVG = `<svg width="14" height="14" viewBox="0 0 14 14" fill="none"><path d="M3 5l4 4 4-4" stroke="currentColor" stroke-width="1.5" stroke-linecap="round" stroke-linejoin="round"/></svg>`;
  const COLLAPSE_SVG = `<svg width="14" height="14" viewBox="0 0 14 14" fill="none"><path d="M3 9l4-4 4 4" stroke="currentColor" stroke-width="1.5" stroke-linecap="round" stroke-linejoin="round"/></svg>`;

  function renderMergeList() {
    const container = Utils.$('pdfListContainer-pdf-merge');
    container.innerHTML = '';

    mergeState.pdfs.forEach((pdf, idx) => {
      const pdfPages = mergeState.pages.filter((p) => p.pdfId === pdf.id);
      const card = document.createElement('div');
      card.className = 'sheet-card';
      card.dataset.pdfId = pdf.id;
      card.draggable = true;

      card.innerHTML = `
        <div class="sheet-card-main">
          <div class="sheet-drag-handle" title="ドラッグして並び替え">${DRAG_HANDLE_SVG}</div>
          <div class="sheet-info">
            <span class="sheet-name">${Utils.escapeHtml(pdf.name)}</span>
            <span class="sheet-file-name">${pdfPages.length} / ${pdf.pageCount} ページ ・ ${Utils.formatFileSize(pdf.size)}</span>
          </div>
          <div class="sheet-card-actions">
            <button class="sheet-btn expand-toggle-btn" title="${pdf.expanded ? '閉じる' : 'ページ展開'}" data-pdf-id="${pdf.id}">
              ${pdf.expanded ? COLLAPSE_SVG : EXPAND_SVG}
            </button>
            <button class="sheet-btn move-up-btn" title="上に移動" ${idx === 0 ? 'disabled' : ''} data-pdf-id="${pdf.id}">${UP_SVG}</button>
            <button class="sheet-btn move-down-btn" title="下に移動" ${idx === mergeState.pdfs.length - 1 ? 'disabled' : ''} data-pdf-id="${pdf.id}">${DOWN_SVG}</button>
            <button class="sheet-btn danger-remove-btn" title="削除" data-pdf-id="${pdf.id}">${REMOVE_SVG}</button>
          </div>
        </div>
        <div class="page-expand-area ${pdf.expanded ? 'open' : ''}" data-pdf-id="${pdf.id}">
          <div class="page-grid" data-pdf-id="${pdf.id}"></div>
        </div>
      `;
      container.appendChild(card);

      if (pdf.expanded) {
        renderMergePageGrid(card.querySelector(`.page-grid[data-pdf-id="${pdf.id}"]`), pdf.id);
      }
    });

    bindMergeEvents();
  }

  function renderMergePageGrid(gridEl, pdfId) {
    gridEl.innerHTML = '';
    const pdfPages = mergeState.pages.filter((p) => p.pdfId === pdfId);

    pdfPages.forEach((page, idx) => {
      const card = document.createElement('div');
      card.className = 'page-card';
      card.dataset.pageUid = page.uid;
      card.draggable = true;
      card.innerHTML = `
        <span class="page-card-number">${page.pageIndex + 1}</span>
        <span class="page-card-label">P${page.pageIndex + 1}</span>
        <div class="page-card-actions">
          <button class="sheet-btn page-move-up-btn" title="上に移動" ${idx === 0 ? 'disabled' : ''} data-page-uid="${page.uid}">
            ${UP_SVG}
          </button>
          <button class="sheet-btn page-move-down-btn" title="下に移動" ${idx === pdfPages.length - 1 ? 'disabled' : ''} data-page-uid="${page.uid}">
            ${DOWN_SVG}
          </button>
          <button class="sheet-btn danger-remove-btn page-remove-btn" title="削除" data-page-uid="${page.uid}">
            ${REMOVE_SVG}
          </button>
        </div>
      `;
      gridEl.appendChild(card);
    });

    bindMergePageEvents(gridEl, pdfId);
  }

  let draggedMergePageUid = null;

  function bindMergePageEvents(gridEl, pdfId) {
    gridEl.querySelectorAll('.page-move-up-btn').forEach((btn) => {
      btn.addEventListener('click', (e) => {
        e.stopPropagation();
        const uid = parseInt(btn.dataset.pageUid, 10);
        const idx = mergeState.pages.findIndex((p) => p.uid === uid);
        // Find previous page in same pdf
        const pdfPages = mergeState.pages.filter((p) => p.pdfId === pdfId);
        const localIdx = pdfPages.findIndex((p) => p.uid === uid);
        if (localIdx <= 0) return;
        const prevUid = pdfPages[localIdx - 1].uid;
        const prevGlobalIdx = mergeState.pages.findIndex((p) => p.uid === prevUid);
        [mergeState.pages[prevGlobalIdx], mergeState.pages[idx]] = [mergeState.pages[idx], mergeState.pages[prevGlobalIdx]];
        renderMergeList();
        debouncedMergePreview();
      });
    });

    gridEl.querySelectorAll('.page-move-down-btn').forEach((btn) => {
      btn.addEventListener('click', (e) => {
        e.stopPropagation();
        const uid = parseInt(btn.dataset.pageUid, 10);
        const idx = mergeState.pages.findIndex((p) => p.uid === uid);
        const pdfPages = mergeState.pages.filter((p) => p.pdfId === pdfId);
        const localIdx = pdfPages.findIndex((p) => p.uid === uid);
        if (localIdx >= pdfPages.length - 1) return;
        const nextUid = pdfPages[localIdx + 1].uid;
        const nextGlobalIdx = mergeState.pages.findIndex((p) => p.uid === nextUid);
        [mergeState.pages[idx], mergeState.pages[nextGlobalIdx]] = [mergeState.pages[nextGlobalIdx], mergeState.pages[idx]];
        renderMergeList();
        debouncedMergePreview();
      });
    });

    gridEl.querySelectorAll('.page-remove-btn').forEach((btn) => {
      btn.addEventListener('click', (e) => {
        e.stopPropagation();
        const uid = parseInt(btn.dataset.pageUid, 10);
        mergeState.pages = mergeState.pages.filter((p) => p.uid !== uid);
        renderMergeList();
        debouncedMergePreview();
      });
    });

    // Page-level drag & drop within grid
    gridEl.querySelectorAll('.page-card').forEach((card) => {
      card.addEventListener('dragstart', (e) => {
        e.stopPropagation();
        draggedMergePageUid = parseInt(card.dataset.pageUid, 10);
        card.classList.add('dragging');
        e.dataTransfer.effectAllowed = 'move';
        e.dataTransfer.setData('text/plain', 'page-' + card.dataset.pageUid);
      });
      card.addEventListener('dragend', () => {
        card.classList.remove('dragging');
        draggedMergePageUid = null;
        gridEl.querySelectorAll('.page-card').forEach((c) => c.classList.remove('drag-over-card'));
      });
      card.addEventListener('dragover', (e) => {
        e.preventDefault();
        e.stopPropagation();
        e.dataTransfer.dropEffect = 'move';
        if (draggedMergePageUid !== null && draggedMergePageUid !== parseInt(card.dataset.pageUid, 10)) {
          card.classList.add('drag-over-card');
        }
      });
      card.addEventListener('dragleave', () => card.classList.remove('drag-over-card'));
      card.addEventListener('drop', (e) => {
        e.preventDefault();
        e.stopPropagation();
        card.classList.remove('drag-over-card');
        const targetUid = parseInt(card.dataset.pageUid, 10);
        if (draggedMergePageUid === null || draggedMergePageUid === targetUid) return;
        const fromIdx = mergeState.pages.findIndex((p) => p.uid === draggedMergePageUid);
        const toIdx = mergeState.pages.findIndex((p) => p.uid === targetUid);
        if (fromIdx === -1 || toIdx === -1) return;
        const [moved] = mergeState.pages.splice(fromIdx, 1);
        mergeState.pages.splice(toIdx, 0, moved);
        renderMergeList();
        debouncedMergePreview();
      });
    });
  }

  let draggedPdfId = null;

  function bindMergeEvents() {
    const container = Utils.$('pdfListContainer-pdf-merge');

    container.querySelectorAll('.expand-toggle-btn').forEach((btn) => {
      btn.addEventListener('click', () => {
        const pdfId = parseInt(btn.dataset.pdfId, 10);
        const pdf = mergeState.pdfs.find((p) => p.id === pdfId);
        if (pdf) {
          pdf.expanded = !pdf.expanded;
          renderMergeList();
        }
      });
    });

    container.querySelectorAll('.move-up-btn').forEach((btn) => {
      btn.addEventListener('click', () => {
        const idx = mergeState.pdfs.findIndex((p) => p.id === parseInt(btn.dataset.pdfId, 10));
        if (idx > 0) {
          [mergeState.pdfs[idx - 1], mergeState.pdfs[idx]] = [mergeState.pdfs[idx], mergeState.pdfs[idx - 1]];
          reorderPagesForFiles();
          renderMergeList(); debouncedMergePreview();
        }
      });
    });

    container.querySelectorAll('.move-down-btn').forEach((btn) => {
      btn.addEventListener('click', () => {
        const idx = mergeState.pdfs.findIndex((p) => p.id === parseInt(btn.dataset.pdfId, 10));
        if (idx < mergeState.pdfs.length - 1) {
          [mergeState.pdfs[idx], mergeState.pdfs[idx + 1]] = [mergeState.pdfs[idx + 1], mergeState.pdfs[idx]];
          reorderPagesForFiles();
          renderMergeList(); debouncedMergePreview();
        }
      });
    });

    container.querySelectorAll(':scope > .sheet-card > .sheet-card-main > .sheet-card-actions > .danger-remove-btn').forEach((btn) => {
      btn.addEventListener('click', () => {
        const pdfId = parseInt(btn.dataset.pdfId, 10);
        mergeState.pdfs = mergeState.pdfs.filter((p) => p.id !== pdfId);
        mergeState.pages = mergeState.pages.filter((p) => p.pdfId !== pdfId);
        if (mergeState.pdfs.length === 0) { resetMerge(); } else { renderMergeList(); debouncedMergePreview(); }
      });
    });

    // File-level Drag & Drop
    container.querySelectorAll(':scope > .sheet-card').forEach((card) => {
      card.addEventListener('dragstart', (e) => {
        // Only handle file-level drag if started from handle
        if (e.target.closest('.page-card')) return;
        draggedPdfId = parseInt(card.dataset.pdfId, 10);
        card.classList.add('dragging');
        e.dataTransfer.effectAllowed = 'move';
        e.dataTransfer.setData('text/plain', 'file-' + card.dataset.pdfId);
      });
      card.addEventListener('dragend', () => {
        card.classList.remove('dragging');
        draggedPdfId = null;
        container.querySelectorAll('.sheet-card').forEach((c) => c.classList.remove('drag-over-card'));
      });
      card.addEventListener('dragover', (e) => {
        if (e.target.closest('.page-card') || e.target.closest('.page-grid')) return;
        e.preventDefault();
        e.dataTransfer.dropEffect = 'move';
        if (draggedPdfId !== null && draggedPdfId !== parseInt(card.dataset.pdfId, 10)) {
          card.classList.add('drag-over-card');
        }
      });
      card.addEventListener('dragleave', (e) => {
        if (!card.contains(e.relatedTarget)) {
          card.classList.remove('drag-over-card');
        }
      });
      card.addEventListener('drop', (e) => {
        if (e.target.closest('.page-card') || e.target.closest('.page-grid')) return;
        e.preventDefault();
        card.classList.remove('drag-over-card');
        const targetId = parseInt(card.dataset.pdfId, 10);
        if (draggedPdfId === null || draggedPdfId === targetId) return;
        const fromIdx = mergeState.pdfs.findIndex((p) => p.id === draggedPdfId);
        const toIdx = mergeState.pdfs.findIndex((p) => p.id === targetId);
        if (fromIdx === -1 || toIdx === -1) return;
        const [moved] = mergeState.pdfs.splice(fromIdx, 1);
        mergeState.pdfs.splice(toIdx, 0, moved);
        reorderPagesForFiles();
        renderMergeList(); debouncedMergePreview();
      });
    });
  }

  function reorderPagesForFiles() {
    const reordered = [];
    for (const pdf of mergeState.pdfs) {
      const pdfPages = mergeState.pages.filter((p) => p.pdfId === pdf.id);
      reordered.push(...pdfPages);
    }
    mergeState.pages = reordered;
  }

  function _applyMetadata(pdfDoc, opts) {
    if (opts.title) pdfDoc.setTitle(opts.title);
    if (opts.author) pdfDoc.setAuthor(opts.author);
    if (opts.subject) pdfDoc.setSubject(opts.subject);
    if (opts.keywords) pdfDoc.setKeywords(opts.keywords.split(',').map((k) => k.trim()));
  }

  async function mergePdfs() {
    if (mergeState.pages.length === 0) return null;
    const merged = await PDFDocument.create();

    // Cache loaded source PDFs
    const pdfCache = new Map();
    for (const pdf of mergeState.pdfs) {
      pdfCache.set(pdf.id, await PDFDocument.load(pdf.data, { ignoreEncryption: true }));
    }

    // Copy pages one by one in pages[] order
    for (const page of mergeState.pages) {
      const srcDoc = pdfCache.get(page.pdfId);
      if (!srcDoc) continue;
      const [copiedPage] = await merged.copyPages(srcDoc, [page.pageIndex]);
      merged.addPage(copiedPage);
    }

    if (mergeOptions) {
      _applyMetadata(merged, mergeOptions.getOptions());
    }
    return merged;
  }

  const debouncedMergePreview = Utils.debounce(async () => {
    if (mergeState.pages.length === 0) return;
    const previewInfo = Utils.$('previewInfo-pdf-merge');
    const pdfPreview = Utils.$('pdfPreview-pdf-merge');
    previewInfo.textContent = '結合中...';

    try {
      const merged = await mergePdfs();
      if (!merged) return;
      const bytes = await merged.save();
      const blob = new Blob([bytes], { type: 'application/pdf' });
      mergeState.previewUrl = Utils.revokeBlobUrl(mergeState.previewUrl);
      mergeState.previewUrl = URL.createObjectURL(blob);
      pdfPreview.src = mergeState.previewUrl;
      Utils.$('previewSection-pdf-merge').classList.remove('hidden');

      previewInfo.textContent = `${mergeState.pdfs.length} ファイル・${mergeState.pages.length} ページ`;
    } catch (err) {
      previewInfo.textContent = '結合に失敗しました';
      console.error(err);
    }
  }, 300);

  async function downloadMerged() {
    if (mergeState.pages.length === 0) return;
    Loading.show('PDFを結合中...');
    try {
      const merged = await mergePdfs();
      if (!merged) return;
      const bytes = await merged.save();
      const blob = new Blob([bytes], { type: 'application/pdf' });
      const url = URL.createObjectURL(blob);
      const a = document.createElement('a');
      a.href = url;
      const baseNames = mergeState.pdfs.map((p) => Utils.getBaseName(p.name));
      a.download = Utils.buildDownloadName(baseNames, 'pdf');
      document.body.appendChild(a);
      a.click();
      document.body.removeChild(a);
      URL.revokeObjectURL(url);
      Toast.show('結合PDFのダウンロードが完了しました！', 'success');
    } catch (err) {
      Toast.show('PDF結合中にエラーが発生しました', 'error');
      console.error(err);
    } finally {
      Loading.hide();
    }
  }

  // ==========================================
  // PDF 分割（ページ抽出）
  // ==========================================
  const splitState = {
    fileName: null,
    data: null,
    pageCount: 0,
    selectedPages: [],  // { uid, pageIndex (0-based), selected: true/false }
    previewUrl: null,
  };

  let splitOptions = null;
  let _splitSyncing = false; // 無限ループ防止フラグ

  async function handleSplitFile(files) {
    const file = files[0];
    if (!file) return;
    splitState.fileName = file.name;

    Loading.show('PDFを読み込み中...');
    try {
      splitState.data = new Uint8Array(await file.arrayBuffer());
      const pdfDoc = await PDFDocument.load(splitState.data, { ignoreEncryption: true });
      splitState.pageCount = pdfDoc.getPageCount();

      // Initialize selectedPages (all selected, in order)
      splitState.selectedPages = [];
      for (let i = 0; i < splitState.pageCount; i++) {
        splitState.selectedPages.push({
          uid: nextPageUid(),
          pageIndex: i,
          selected: true,
        });
      }

      Utils.$('splitOptions-pdf-split').classList.remove('hidden');
      Utils.$('splitPageInfo-pdf-split').textContent = `全 ${splitState.pageCount} ページ`;
      syncPageRangeText();
      if (splitOptions) splitOptions.show();
      Utils.$('actionsPanel-pdf-split').classList.remove('hidden');

      renderSplitPageGrid();
      debouncedSplitPreview();
    } catch (err) {
      Toast.show('PDFの読み込みに失敗しました', 'error');
      console.error(err);
    } finally {
      Loading.hide();
    }
  }

  function renderSplitPageGrid() {
    const gridEl = Utils.$('splitPageGridContainer-pdf-split');
    gridEl.innerHTML = '';

    splitState.selectedPages.forEach((page, idx) => {
      const card = document.createElement('div');
      card.className = 'page-card' + (page.selected ? ' selected' : '');
      card.dataset.pageUid = page.uid;
      card.draggable = true;
      card.innerHTML = `
        <input type="checkbox" class="page-card-checkbox" ${page.selected ? 'checked' : ''} data-page-uid="${page.uid}">
        <span class="page-card-number">${page.pageIndex + 1}</span>
        <span class="page-card-label">P${page.pageIndex + 1}</span>
        <div class="page-card-actions">
          <button class="sheet-btn page-move-up-btn" title="上に移動" ${idx === 0 ? 'disabled' : ''} data-page-uid="${page.uid}">
            ${UP_SVG}
          </button>
          <button class="sheet-btn page-move-down-btn" title="下に移動" ${idx === splitState.selectedPages.length - 1 ? 'disabled' : ''} data-page-uid="${page.uid}">
            ${DOWN_SVG}
          </button>
        </div>
      `;
      gridEl.appendChild(card);
    });

    bindSplitPageEvents(gridEl);
  }

  let draggedSplitPageUid = null;

  function bindSplitPageEvents(gridEl) {
    gridEl.querySelectorAll('.page-card-checkbox').forEach((cb) => {
      cb.addEventListener('change', (e) => {
        e.stopPropagation();
        const uid = parseInt(cb.dataset.pageUid, 10);
        const page = splitState.selectedPages.find((p) => p.uid === uid);
        if (page) {
          page.selected = cb.checked;
          cb.closest('.page-card').classList.toggle('selected', cb.checked);
          syncPageRangeText();
          debouncedSplitPreview();
        }
      });
    });

    gridEl.querySelectorAll('.page-move-up-btn').forEach((btn) => {
      btn.addEventListener('click', (e) => {
        e.stopPropagation();
        const uid = parseInt(btn.dataset.pageUid, 10);
        const idx = splitState.selectedPages.findIndex((p) => p.uid === uid);
        if (idx > 0) {
          [splitState.selectedPages[idx - 1], splitState.selectedPages[idx]] =
            [splitState.selectedPages[idx], splitState.selectedPages[idx - 1]];
          renderSplitPageGrid();
          syncPageRangeText();
          debouncedSplitPreview();
        }
      });
    });

    gridEl.querySelectorAll('.page-move-down-btn').forEach((btn) => {
      btn.addEventListener('click', (e) => {
        e.stopPropagation();
        const uid = parseInt(btn.dataset.pageUid, 10);
        const idx = splitState.selectedPages.findIndex((p) => p.uid === uid);
        if (idx < splitState.selectedPages.length - 1) {
          [splitState.selectedPages[idx], splitState.selectedPages[idx + 1]] =
            [splitState.selectedPages[idx + 1], splitState.selectedPages[idx]];
          renderSplitPageGrid();
          syncPageRangeText();
          debouncedSplitPreview();
        }
      });
    });

    // Page-level drag & drop
    gridEl.querySelectorAll('.page-card').forEach((card) => {
      card.addEventListener('dragstart', (e) => {
        draggedSplitPageUid = parseInt(card.dataset.pageUid, 10);
        card.classList.add('dragging');
        e.dataTransfer.effectAllowed = 'move';
        e.dataTransfer.setData('text/plain', 'split-page-' + card.dataset.pageUid);
      });
      card.addEventListener('dragend', () => {
        card.classList.remove('dragging');
        draggedSplitPageUid = null;
        gridEl.querySelectorAll('.page-card').forEach((c) => c.classList.remove('drag-over-card'));
      });
      card.addEventListener('dragover', (e) => {
        e.preventDefault();
        e.dataTransfer.dropEffect = 'move';
        if (draggedSplitPageUid !== null && draggedSplitPageUid !== parseInt(card.dataset.pageUid, 10)) {
          card.classList.add('drag-over-card');
        }
      });
      card.addEventListener('dragleave', () => card.classList.remove('drag-over-card'));
      card.addEventListener('drop', (e) => {
        e.preventDefault();
        card.classList.remove('drag-over-card');
        const targetUid = parseInt(card.dataset.pageUid, 10);
        if (draggedSplitPageUid === null || draggedSplitPageUid === targetUid) return;
        const fromIdx = splitState.selectedPages.findIndex((p) => p.uid === draggedSplitPageUid);
        const toIdx = splitState.selectedPages.findIndex((p) => p.uid === targetUid);
        if (fromIdx === -1 || toIdx === -1) return;
        const [moved] = splitState.selectedPages.splice(fromIdx, 1);
        splitState.selectedPages.splice(toIdx, 0, moved);
        renderSplitPageGrid();
        syncPageRangeText();
        debouncedSplitPreview();
      });
    });
  }

  // ── グリッド → テキスト 同期 ──
  function syncPageRangeText() {
    if (_splitSyncing) return;
    _splitSyncing = true;
    try {
      const selected = splitState.selectedPages
        .filter((p) => p.selected)
        .map((p) => p.pageIndex + 1); // 1-based

      Utils.$('pageRange-pdf-split').value = compactPageRange(selected);
    } finally {
      _splitSyncing = false;
    }
  }

  function compactPageRange(pages) {
    if (pages.length === 0) return '';
    const parts = [];
    let rangeStart = pages[0];
    let rangeEnd = pages[0];
    let prev = pages[0];

    for (let i = 1; i < pages.length; i++) {
      if (pages[i] === prev + 1) {
        rangeEnd = pages[i];
      } else {
        parts.push(rangeStart === rangeEnd ? `${rangeStart}` : `${rangeStart}-${rangeEnd}`);
        rangeStart = pages[i];
        rangeEnd = pages[i];
      }
      prev = pages[i];
    }
    parts.push(rangeStart === rangeEnd ? `${rangeStart}` : `${rangeStart}-${rangeEnd}`);
    return parts.join(', ');
  }

  // ── テキスト → グリッド 同期 ──
  function syncFromPageRangeText() {
    if (_splitSyncing) return;
    _splitSyncing = true;
    try {
      const rangeStr = Utils.$('pageRange-pdf-split').value;
      const pageNumbers = parsePageRangeOrdered(rangeStr, splitState.pageCount);

      // Update selected state based on text
      const selectedSet = new Set(pageNumbers);
      splitState.selectedPages.forEach((p) => {
        p.selected = selectedSet.has(p.pageIndex + 1);
      });

      // Reorder: move selected pages to match text order, keep unselected at end in original order
      const selectedMap = new Map();
      splitState.selectedPages.forEach((p) => {
        if (!selectedMap.has(p.pageIndex)) {
          selectedMap.set(p.pageIndex, p);
        }
      });

      const reordered = [];
      const used = new Set();
      for (const pageNum of pageNumbers) {
        const pi = pageNum - 1;
        const page = selectedMap.get(pi);
        if (page && !used.has(page.uid)) {
          reordered.push(page);
          used.add(page.uid);
        }
      }
      // Add unselected pages at end
      splitState.selectedPages.forEach((p) => {
        if (!used.has(p.uid)) {
          reordered.push(p);
          used.add(p.uid);
        }
      });

      splitState.selectedPages = reordered;
      renderSplitPageGrid();
    } finally {
      _splitSyncing = false;
    }
  }

  function parsePageRangeOrdered(rangeStr, maxPages) {
    const pages = [];
    const seen = new Set();
    const parts = rangeStr.split(',');
    for (let part of parts) {
      part = part.trim();
      if (!part) continue;
      const match = part.match(/^(\d+)\s*-\s*(\d+)$/);
      if (match) {
        const start = Math.max(1, parseInt(match[1], 10));
        const end = Math.min(maxPages, parseInt(match[2], 10));
        for (let i = start; i <= end; i++) {
          if (!seen.has(i)) { pages.push(i); seen.add(i); }
        }
      } else {
        const num = parseInt(part, 10);
        if (!isNaN(num) && num >= 1 && num <= maxPages && !seen.has(num)) {
          pages.push(num);
          seen.add(num);
        }
      }
    }
    return pages;
  }

  async function extractPages() {
    if (!splitState.data) return null;
    const selected = splitState.selectedPages.filter((p) => p.selected);
    if (selected.length === 0) return null;

    const src = await PDFDocument.load(splitState.data, { ignoreEncryption: true });
    const newDoc = await PDFDocument.create();
    const indices = selected.map((p) => p.pageIndex);
    const copiedPages = await newDoc.copyPages(src, indices);
    copiedPages.forEach((page) => newDoc.addPage(page));

    if (splitOptions) {
      _applyMetadata(newDoc, splitOptions.getOptions());
    }

    return { doc: newDoc, pageNumbers: selected.map((p) => p.pageIndex + 1) };
  }

  const debouncedSplitPreview = Utils.debounce(async () => {
    if (!splitState.data) return;
    const previewInfo = Utils.$('previewInfo-pdf-split');
    const pdfPreview = Utils.$('pdfPreview-pdf-split');

    try {
      const result = await extractPages();
      if (!result) {
        previewInfo.textContent = '有効なページ範囲を指定してください';
        return;
      }
      const bytes = await result.doc.save();
      const blob = new Blob([bytes], { type: 'application/pdf' });
      splitState.previewUrl = Utils.revokeBlobUrl(splitState.previewUrl);
      splitState.previewUrl = URL.createObjectURL(blob);
      pdfPreview.src = splitState.previewUrl;
      Utils.$('previewSection-pdf-split').classList.remove('hidden');
      previewInfo.textContent = `${result.pageNumbers.length} ページを抽出`;
    } catch (err) {
      previewInfo.textContent = 'プレビュー生成に失敗しました';
      console.error(err);
    }
  }, 500);

  async function downloadSplit() {
    if (!splitState.data) return;
    Loading.show('ページを抽出中...');
    try {
      const result = await extractPages();
      if (!result) { Toast.show('有効なページ範囲を指定してください', 'error'); Loading.hide(); return; }

      const bytes = await result.doc.save();
      const blob = new Blob([bytes], { type: 'application/pdf' });
      const url = URL.createObjectURL(blob);
      const a = document.createElement('a');
      a.href = url;
      a.download = Utils.buildDownloadName(Utils.getBaseName(splitState.fileName || 'extracted'), 'pdf');
      document.body.appendChild(a);
      a.click();
      document.body.removeChild(a);
      URL.revokeObjectURL(url);
      Toast.show('抽出PDFのダウンロードが完了しました！', 'success');
    } catch (err) {
      Toast.show('PDF抽出中にエラーが発生しました', 'error');
      console.error(err);
    } finally {
      Loading.hide();
    }
  }

  // ── Init ──
  function init() {
    initSubmode();

    // PDF 結合
    DropZone.init({
      elementId: 'dropZone-pdf-merge',
      inputId: 'fileInput-pdf-merge',
      acceptExtensions: ['.pdf'],
      multiple: true,
      onFiles: handleMergeFiles,
    });

    const mergeListEl = Utils.$('pdfList-pdf-merge');
    if (mergeListEl) {
      mergeOptions = PdfOutputOptions.create({
        container: mergeListEl,
        suffix: 'pdf-merge',
        features: { metadata: true, security: false, colorMode: false, imageQuality: false },
      });
    }

    Utils.$('clearAllPdfs-pdf-merge').addEventListener('click', resetMerge);
    Utils.$('convertBtn-pdf-merge').addEventListener('click', downloadMerged);

    // 全ページ展開 / すべて閉じる
    Utils.$('expandAllPages-pdf-merge').addEventListener('click', () => {
      mergeState.pdfs.forEach((pdf) => { pdf.expanded = true; });
      renderMergeList();
    });
    Utils.$('collapseAllPages-pdf-merge').addEventListener('click', () => {
      mergeState.pdfs.forEach((pdf) => { pdf.expanded = false; });
      renderMergeList();
    });

    // PDF 分割
    DropZone.init({
      elementId: 'dropZone-pdf-split',
      inputId: 'fileInput-pdf-split',
      acceptExtensions: ['.pdf'],
      multiple: false,
      onFiles: handleSplitFile,
    });

    const splitOptsEl = Utils.$('splitOptions-pdf-split');
    if (splitOptsEl) {
      splitOptions = PdfOutputOptions.create({
        container: splitOptsEl,
        suffix: 'pdf-split',
        features: { metadata: true, security: false, colorMode: false, imageQuality: false },
      });
    }

    // テキスト入力変更 → グリッド同期
    Utils.$('pageRange-pdf-split').addEventListener('input', () => {
      syncFromPageRangeText();
      debouncedSplitPreview();
    });

    Utils.$('convertBtn-pdf-split').addEventListener('click', downloadSplit);

    // 全選択 / 全解除
    Utils.$('selectAllPages-pdf-split').addEventListener('click', () => {
      splitState.selectedPages.forEach((p) => { p.selected = true; });
      renderSplitPageGrid();
      syncPageRangeText();
      debouncedSplitPreview();
    });
    Utils.$('deselectAllPages-pdf-split').addEventListener('click', () => {
      splitState.selectedPages.forEach((p) => { p.selected = false; });
      renderSplitPageGrid();
      syncPageRangeText();
      debouncedSplitPreview();
    });

    TabManager.register('pdf-tools', {
      onActivate() {},
      onDeactivate() {},
    });
  }

  init();
  return { mergeState, splitState };
})();
