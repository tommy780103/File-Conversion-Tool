/**
 * PDF 結合・分割 コンバーター
 * pdf-lib (操作) + pdf.js (サムネイルレンダリング)
 * CubePDF Utility風のサムネイルベースUI
 */
const PdfToolsConverter = (() => {
  const panel = Utils.$('panel-pdf-tools');
  const { PDFDocument } = PDFLib;

  // pdf.js が利用可能か
  const hasPdfJs = typeof pdfjsLib !== 'undefined';
  if (hasPdfJs) {
    pdfjsLib.GlobalWorkerOptions.workerSrc =
      'https://unpkg.com/pdfjs-dist@3.11.174/build/pdf.worker.min.js';
  }

  let _pageUidCounter = 0;
  function nextPageUid() { return ++_pageUidCounter; }

  // ── SVG定数 ──
  const DRAG_HANDLE_SVG = `<svg width="14" height="14" viewBox="0 0 14 14" fill="none"><circle cx="4" cy="3" r="1.2" fill="currentColor"/><circle cx="10" cy="3" r="1.2" fill="currentColor"/><circle cx="4" cy="7" r="1.2" fill="currentColor"/><circle cx="10" cy="7" r="1.2" fill="currentColor"/><circle cx="4" cy="11" r="1.2" fill="currentColor"/><circle cx="10" cy="11" r="1.2" fill="currentColor"/></svg>`;
  const UP_SVG = `<svg width="14" height="14" viewBox="0 0 14 14" fill="none"><path d="M7 2v10M3 6l4-4 4 4" stroke="currentColor" stroke-width="1.5" stroke-linecap="round" stroke-linejoin="round"/></svg>`;
  const DOWN_SVG = `<svg width="14" height="14" viewBox="0 0 14 14" fill="none"><path d="M7 12V2M3 8l4 4 4-4" stroke="currentColor" stroke-width="1.5" stroke-linecap="round" stroke-linejoin="round"/></svg>`;
  const REMOVE_SVG = `<svg width="14" height="14" viewBox="0 0 14 14" fill="none"><path d="M3 3l8 8M11 3l-8 8" stroke="currentColor" stroke-width="1.5" stroke-linecap="round"/></svg>`;

  // ==========================================
  // サムネイルレンダラー (pdf.js)
  // ==========================================
  const PageThumbnail = {
    _cache: new Map(),     // "cacheKey_pageIdx" -> dataURL
    _docCache: new Map(),  // cacheKey -> Promise<PDFDocumentProxy>

    _getDoc(cacheKey, data) {
      if (!this._docCache.has(cacheKey)) {
        this._docCache.set(cacheKey,
          pdfjsLib.getDocument({ data: data.slice(0) }).promise
        );
      }
      return this._docCache.get(cacheKey);
    },

    async render(cacheKey, pdfData, pageIndex, width) {
      width = width || 180;
      const key = `${cacheKey}_${pageIndex}`;
      if (this._cache.has(key)) return this._cache.get(key);

      const doc = await this._getDoc(cacheKey, pdfData);
      const page = await doc.getPage(pageIndex + 1);
      const vp = page.getViewport({ scale: 1 });
      const scale = width / vp.width;
      const svp = page.getViewport({ scale });

      const canvas = document.createElement('canvas');
      canvas.width = Math.floor(svp.width);
      canvas.height = Math.floor(svp.height);
      await page.render({
        canvasContext: canvas.getContext('2d'),
        viewport: svp,
      }).promise;

      const dataUrl = canvas.toDataURL('image/jpeg', 0.75);
      this._cache.set(key, dataUrl);
      return dataUrl;
    },

    getCached(cacheKey, pageIndex) {
      return this._cache.get(`${cacheKey}_${pageIndex}`) || null;
    },

    clearForKey(prefix) {
      for (const k of [...this._cache.keys()]) {
        if (k.startsWith(`${prefix}_`)) this._cache.delete(k);
      }
      this._docCache.delete(prefix);
    },

    clearAll() {
      this._cache.clear();
      this._docCache.clear();
    },
  };

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
    pdfs: [],    // { id, name, size, data (Uint8Array), pageCount }
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
        };
        mergeState.pdfs.push(pdf);
        mergeState.pages.push(...buildPagesForPdf(pdf));
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
    Utils.$('pageThumbnails-pdf-merge').classList.remove('hidden');
    if (mergeOptions) mergeOptions.show();
    Utils.$('actionsPanel-pdf-merge').classList.remove('hidden');
    renderMergeFileList();
    renderMergeThumbnailGrid();
    debouncedMergePreview();
  }

  function resetMerge() {
    mergeState.pdfs = [];
    mergeState.pages = [];
    mergeState.idCounter = 0;
    mergeState.previewUrl = Utils.revokeBlobUrl(mergeState.previewUrl);

    Utils.$('dropZone-pdf-merge').classList.remove('compact');
    Utils.$('pdfList-pdf-merge').classList.add('hidden');
    Utils.$('pageThumbnails-pdf-merge').classList.add('hidden');
    if (mergeOptions) mergeOptions.hide();
    Utils.$('previewSection-pdf-merge').classList.add('hidden');
    Utils.$('actionsPanel-pdf-merge').classList.add('hidden');
    Utils.$('pdfListContainer-pdf-merge').innerHTML = '';
    Utils.$('pageThumbnailGrid-pdf-merge').innerHTML = '';
    PageThumbnail.clearAll();
  }

  // ── ファイル一覧（シンプル版: 展開なし） ──
  function renderMergeFileList() {
    const container = Utils.$('pdfListContainer-pdf-merge');
    container.innerHTML = '';

    mergeState.pdfs.forEach((pdf, idx) => {
      const pdfPageCount = mergeState.pages.filter((p) => p.pdfId === pdf.id).length;
      const card = document.createElement('div');
      card.className = 'sheet-card';
      card.dataset.pdfId = pdf.id;
      card.draggable = true;
      card.innerHTML = `
        <div class="sheet-card-main">
          <div class="sheet-drag-handle" title="ドラッグして並び替え">${DRAG_HANDLE_SVG}</div>
          <div class="sheet-info">
            <span class="sheet-name">${Utils.escapeHtml(pdf.name)}</span>
            <span class="sheet-file-name">${pdfPageCount} / ${pdf.pageCount} ページ ・ ${Utils.formatFileSize(pdf.size)}</span>
          </div>
          <div class="sheet-card-actions">
            <button class="sheet-btn move-up-btn" title="上に移動" ${idx === 0 ? 'disabled' : ''} data-pdf-id="${pdf.id}">${UP_SVG}</button>
            <button class="sheet-btn move-down-btn" title="下に移動" ${idx === mergeState.pdfs.length - 1 ? 'disabled' : ''} data-pdf-id="${pdf.id}">${DOWN_SVG}</button>
            <button class="sheet-btn danger-remove-btn" title="削除" data-pdf-id="${pdf.id}">${REMOVE_SVG}</button>
          </div>
        </div>
      `;
      container.appendChild(card);
    });

    bindMergeFileEvents();
  }

  let _draggedFilePdfId = null;

  function bindMergeFileEvents() {
    const container = Utils.$('pdfListContainer-pdf-merge');

    container.querySelectorAll('.move-up-btn').forEach((btn) => {
      btn.addEventListener('click', () => {
        const idx = mergeState.pdfs.findIndex((p) => p.id === parseInt(btn.dataset.pdfId, 10));
        if (idx > 0) {
          [mergeState.pdfs[idx - 1], mergeState.pdfs[idx]] = [mergeState.pdfs[idx], mergeState.pdfs[idx - 1]];
          reorderPagesForFiles();
          renderMergeFileList();
          renderMergeThumbnailGrid();
          debouncedMergePreview();
        }
      });
    });

    container.querySelectorAll('.move-down-btn').forEach((btn) => {
      btn.addEventListener('click', () => {
        const idx = mergeState.pdfs.findIndex((p) => p.id === parseInt(btn.dataset.pdfId, 10));
        if (idx < mergeState.pdfs.length - 1) {
          [mergeState.pdfs[idx], mergeState.pdfs[idx + 1]] = [mergeState.pdfs[idx + 1], mergeState.pdfs[idx]];
          reorderPagesForFiles();
          renderMergeFileList();
          renderMergeThumbnailGrid();
          debouncedMergePreview();
        }
      });
    });

    container.querySelectorAll('.danger-remove-btn').forEach((btn) => {
      btn.addEventListener('click', () => {
        const pdfId = parseInt(btn.dataset.pdfId, 10);
        mergeState.pdfs = mergeState.pdfs.filter((p) => p.id !== pdfId);
        mergeState.pages = mergeState.pages.filter((p) => p.pdfId !== pdfId);
        if (hasPdfJs) PageThumbnail.clearForKey(`merge_${pdfId}`);
        if (mergeState.pdfs.length === 0) {
          resetMerge();
        } else {
          renderMergeFileList();
          renderMergeThumbnailGrid();
          debouncedMergePreview();
        }
      });
    });

    // ファイルレベル ドラッグ&ドロップ
    container.querySelectorAll('.sheet-card').forEach((card) => {
      card.addEventListener('dragstart', (e) => {
        _draggedFilePdfId = parseInt(card.dataset.pdfId, 10);
        card.classList.add('dragging');
        e.dataTransfer.effectAllowed = 'move';
        e.dataTransfer.setData('text/plain', 'file-' + card.dataset.pdfId);
      });
      card.addEventListener('dragend', () => {
        card.classList.remove('dragging');
        _draggedFilePdfId = null;
        container.querySelectorAll('.sheet-card').forEach((c) => c.classList.remove('drag-over-card'));
      });
      card.addEventListener('dragover', (e) => {
        e.preventDefault();
        e.dataTransfer.dropEffect = 'move';
        if (_draggedFilePdfId !== null && _draggedFilePdfId !== parseInt(card.dataset.pdfId, 10)) {
          card.classList.add('drag-over-card');
        }
      });
      card.addEventListener('dragleave', () => card.classList.remove('drag-over-card'));
      card.addEventListener('drop', (e) => {
        e.preventDefault();
        card.classList.remove('drag-over-card');
        const targetId = parseInt(card.dataset.pdfId, 10);
        if (_draggedFilePdfId === null || _draggedFilePdfId === targetId) return;
        const fromIdx = mergeState.pdfs.findIndex((p) => p.id === _draggedFilePdfId);
        const toIdx = mergeState.pdfs.findIndex((p) => p.id === targetId);
        if (fromIdx === -1 || toIdx === -1) return;
        const [moved] = mergeState.pdfs.splice(fromIdx, 1);
        mergeState.pdfs.splice(toIdx, 0, moved);
        reorderPagesForFiles();
        renderMergeFileList();
        renderMergeThumbnailGrid();
        debouncedMergePreview();
      });
    });
  }

  function reorderPagesForFiles() {
    const reordered = [];
    for (const pdf of mergeState.pdfs) {
      reordered.push(...mergeState.pages.filter((p) => p.pdfId === pdf.id));
    }
    mergeState.pages = reordered;
  }

  // ── 統一サムネイルグリッド（全ページ表示） ──
  let _mergeThumbGen = 0;
  let _draggedMergeThumbUid = null;

  function renderMergeThumbnailGrid() {
    const container = Utils.$('pageThumbnailGrid-pdf-merge');
    container.innerHTML = '';

    const info = Utils.$('pageThumbnailInfo-pdf-merge');
    info.textContent = `${mergeState.pdfs.length} ファイル・${mergeState.pages.length} ページ`;

    mergeState.pages.forEach((page) => {
      const pdf = mergeState.pdfs.find((p) => p.id === page.pdfId);
      if (!pdf) return;

      const cacheKey = `merge_${pdf.id}`;
      const cached = hasPdfJs ? PageThumbnail.getCached(cacheKey, page.pageIndex) : null;

      const card = document.createElement('div');
      card.className = 'page-thumb-card';
      card.dataset.pageUid = page.uid;
      card.draggable = true;
      card.innerHTML = `
        <div class="page-thumb-img-wrap">
          <span class="page-thumb-placeholder${hasPdfJs ? ' page-thumb-loading' : ''}" ${cached ? 'style="display:none"' : ''}>${page.pageIndex + 1}</span>
          <img class="page-thumb-img" ${cached ? `src="${cached}"` : 'style="display:none"'} alt="P${page.pageIndex + 1}">
        </div>
        <div class="page-thumb-info">
          <span class="page-thumb-label" title="${Utils.escapeHtml(pdf.name)}">${Utils.escapeHtml(pdf.name)}</span>
          <span class="page-thumb-page-num">P${page.pageIndex + 1}</span>
        </div>
        <button class="page-thumb-delete" title="削除" data-page-uid="${page.uid}">\u00D7</button>
      `;
      container.appendChild(card);
    });

    bindMergeThumbnailEvents();
    if (hasPdfJs) renderMergeThumbnailsAsync();
  }

  async function renderMergeThumbnailsAsync() {
    const gen = ++_mergeThumbGen;

    for (const page of mergeState.pages) {
      if (gen !== _mergeThumbGen) return;

      const pdf = mergeState.pdfs.find((p) => p.id === page.pdfId);
      if (!pdf) continue;

      const cacheKey = `merge_${pdf.id}`;
      if (PageThumbnail.getCached(cacheKey, page.pageIndex)) continue;

      try {
        const dataUrl = await PageThumbnail.render(cacheKey, pdf.data, page.pageIndex);
        if (gen !== _mergeThumbGen) return;

        const card = document.querySelector(`#pageThumbnailGrid-pdf-merge [data-page-uid="${page.uid}"]`);
        if (card) {
          const img = card.querySelector('.page-thumb-img');
          const ph = card.querySelector('.page-thumb-placeholder');
          if (img) { img.src = dataUrl; img.style.display = 'block'; }
          if (ph) ph.style.display = 'none';
        }
      } catch (err) {
        console.warn('Merge thumbnail render failed:', page.pageIndex, err);
      }

      await new Promise((r) => setTimeout(r, 10));
    }
  }

  function bindMergeThumbnailEvents() {
    const container = Utils.$('pageThumbnailGrid-pdf-merge');

    // 削除ボタン
    container.querySelectorAll('.page-thumb-delete').forEach((btn) => {
      btn.addEventListener('click', (e) => {
        e.stopPropagation();
        const uid = parseInt(btn.dataset.pageUid, 10);
        mergeState.pages = mergeState.pages.filter((p) => p.uid !== uid);
        renderMergeFileList();
        renderMergeThumbnailGrid();
        debouncedMergePreview();
      });
    });

    // ドラッグ&ドロップ
    container.querySelectorAll('.page-thumb-card').forEach((card) => {
      card.addEventListener('dragstart', (e) => {
        _draggedMergeThumbUid = parseInt(card.dataset.pageUid, 10);
        card.classList.add('dragging');
        e.dataTransfer.effectAllowed = 'move';
        e.dataTransfer.setData('text/plain', 'mpage-' + card.dataset.pageUid);
      });
      card.addEventListener('dragend', () => {
        card.classList.remove('dragging');
        _draggedMergeThumbUid = null;
        container.querySelectorAll('.page-thumb-card').forEach((c) => c.classList.remove('drag-over-card'));
      });
      card.addEventListener('dragover', (e) => {
        e.preventDefault();
        e.dataTransfer.dropEffect = 'move';
        if (_draggedMergeThumbUid !== null && _draggedMergeThumbUid !== parseInt(card.dataset.pageUid, 10)) {
          card.classList.add('drag-over-card');
        }
      });
      card.addEventListener('dragleave', () => card.classList.remove('drag-over-card'));
      card.addEventListener('drop', (e) => {
        e.preventDefault();
        card.classList.remove('drag-over-card');
        const targetUid = parseInt(card.dataset.pageUid, 10);
        if (_draggedMergeThumbUid === null || _draggedMergeThumbUid === targetUid) return;
        const fromIdx = mergeState.pages.findIndex((p) => p.uid === _draggedMergeThumbUid);
        const toIdx = mergeState.pages.findIndex((p) => p.uid === targetUid);
        if (fromIdx === -1 || toIdx === -1) return;
        const [moved] = mergeState.pages.splice(fromIdx, 1);
        mergeState.pages.splice(toIdx, 0, moved);
        renderMergeThumbnailGrid();
        debouncedMergePreview();
      });
    });
  }

  // ── PDF結合ロジック ──
  function _applyMetadata(pdfDoc, opts) {
    if (opts.title) pdfDoc.setTitle(opts.title);
    if (opts.author) pdfDoc.setAuthor(opts.author);
    if (opts.subject) pdfDoc.setSubject(opts.subject);
    if (opts.keywords) pdfDoc.setKeywords(opts.keywords.split(',').map((k) => k.trim()));
  }

  async function mergePdfs() {
    if (mergeState.pages.length === 0) return null;
    const merged = await PDFDocument.create();

    const pdfCache = new Map();
    for (const pdf of mergeState.pdfs) {
      pdfCache.set(pdf.id, await PDFDocument.load(pdf.data, { ignoreEncryption: true }));
    }

    for (const page of mergeState.pages) {
      const srcDoc = pdfCache.get(page.pdfId);
      if (!srcDoc) continue;
      const [copiedPage] = await merged.copyPages(srcDoc, [page.pageIndex]);
      merged.addPage(copiedPage);
    }

    if (mergeOptions) _applyMetadata(merged, mergeOptions.getOptions());
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
      a.download = Utils.buildDownloadName(mergeState.pdfs.map((p) => Utils.getBaseName(p.name)), 'pdf');
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
    selectedPages: [], // { uid, pageIndex (0-based), selected }
    previewUrl: null,
  };

  let splitOptions = null;
  let _splitSyncing = false;

  async function handleSplitFile(files) {
    const file = files[0];
    if (!file) return;
    splitState.fileName = file.name;

    Loading.show('PDFを読み込み中...');
    try {
      if (hasPdfJs) PageThumbnail.clearForKey('split');
      splitState.data = new Uint8Array(await file.arrayBuffer());
      const pdfDoc = await PDFDocument.load(splitState.data, { ignoreEncryption: true });
      splitState.pageCount = pdfDoc.getPageCount();

      splitState.selectedPages = [];
      for (let i = 0; i < splitState.pageCount; i++) {
        splitState.selectedPages.push({ uid: nextPageUid(), pageIndex: i, selected: true });
      }

      Utils.$('splitOptions-pdf-split').classList.remove('hidden');
      Utils.$('splitPageInfo-pdf-split').textContent = `全 ${splitState.pageCount} ページ`;
      syncPageRangeText();
      if (splitOptions) splitOptions.show();
      Utils.$('actionsPanel-pdf-split').classList.remove('hidden');

      renderSplitThumbnailGrid();
      debouncedSplitPreview();
    } catch (err) {
      Toast.show('PDFの読み込みに失敗しました', 'error');
      console.error(err);
    } finally {
      Loading.hide();
    }
  }

  // ── 分割サムネイルグリッド ──
  let _splitThumbGen = 0;
  let _draggedSplitThumbUid = null;

  function renderSplitThumbnailGrid() {
    const container = Utils.$('splitPageGridContainer-pdf-split');
    container.innerHTML = '';

    splitState.selectedPages.forEach((page) => {
      const cached = hasPdfJs ? PageThumbnail.getCached('split', page.pageIndex) : null;

      const card = document.createElement('div');
      card.className = 'page-thumb-card' + (page.selected ? ' selected' : '');
      card.dataset.pageUid = page.uid;
      card.draggable = true;
      card.innerHTML = `
        <input type="checkbox" class="page-thumb-checkbox" ${page.selected ? 'checked' : ''} data-page-uid="${page.uid}">
        <div class="page-thumb-img-wrap">
          <span class="page-thumb-placeholder${hasPdfJs ? ' page-thumb-loading' : ''}" ${cached ? 'style="display:none"' : ''}>${page.pageIndex + 1}</span>
          <img class="page-thumb-img" ${cached ? `src="${cached}"` : 'style="display:none"'} alt="P${page.pageIndex + 1}">
        </div>
        <div class="page-thumb-info">
          <span class="page-thumb-page-num">P${page.pageIndex + 1}</span>
        </div>
      `;
      container.appendChild(card);
    });

    bindSplitThumbnailEvents();
    if (hasPdfJs) renderSplitThumbnailsAsync();
  }

  async function renderSplitThumbnailsAsync() {
    const gen = ++_splitThumbGen;

    for (const page of splitState.selectedPages) {
      if (gen !== _splitThumbGen) return;
      if (PageThumbnail.getCached('split', page.pageIndex)) continue;

      try {
        const dataUrl = await PageThumbnail.render('split', splitState.data, page.pageIndex);
        if (gen !== _splitThumbGen) return;

        const card = document.querySelector(`#splitPageGridContainer-pdf-split [data-page-uid="${page.uid}"]`);
        if (card) {
          const img = card.querySelector('.page-thumb-img');
          const ph = card.querySelector('.page-thumb-placeholder');
          if (img) { img.src = dataUrl; img.style.display = 'block'; }
          if (ph) ph.style.display = 'none';
        }
      } catch (err) {
        console.warn('Split thumbnail render failed:', page.pageIndex, err);
      }

      await new Promise((r) => setTimeout(r, 10));
    }
  }

  function bindSplitThumbnailEvents() {
    const container = Utils.$('splitPageGridContainer-pdf-split');

    // チェックボックス
    container.querySelectorAll('.page-thumb-checkbox').forEach((cb) => {
      cb.addEventListener('change', (e) => {
        e.stopPropagation();
        const uid = parseInt(cb.dataset.pageUid, 10);
        const page = splitState.selectedPages.find((p) => p.uid === uid);
        if (page) {
          page.selected = cb.checked;
          cb.closest('.page-thumb-card').classList.toggle('selected', cb.checked);
          syncPageRangeText();
          debouncedSplitPreview();
        }
      });
    });

    // ドラッグ&ドロップ
    container.querySelectorAll('.page-thumb-card').forEach((card) => {
      card.addEventListener('dragstart', (e) => {
        _draggedSplitThumbUid = parseInt(card.dataset.pageUid, 10);
        card.classList.add('dragging');
        e.dataTransfer.effectAllowed = 'move';
        e.dataTransfer.setData('text/plain', 'spage-' + card.dataset.pageUid);
      });
      card.addEventListener('dragend', () => {
        card.classList.remove('dragging');
        _draggedSplitThumbUid = null;
        container.querySelectorAll('.page-thumb-card').forEach((c) => c.classList.remove('drag-over-card'));
      });
      card.addEventListener('dragover', (e) => {
        e.preventDefault();
        e.dataTransfer.dropEffect = 'move';
        if (_draggedSplitThumbUid !== null && _draggedSplitThumbUid !== parseInt(card.dataset.pageUid, 10)) {
          card.classList.add('drag-over-card');
        }
      });
      card.addEventListener('dragleave', () => card.classList.remove('drag-over-card'));
      card.addEventListener('drop', (e) => {
        e.preventDefault();
        card.classList.remove('drag-over-card');
        const targetUid = parseInt(card.dataset.pageUid, 10);
        if (_draggedSplitThumbUid === null || _draggedSplitThumbUid === targetUid) return;
        const fromIdx = splitState.selectedPages.findIndex((p) => p.uid === _draggedSplitThumbUid);
        const toIdx = splitState.selectedPages.findIndex((p) => p.uid === targetUid);
        if (fromIdx === -1 || toIdx === -1) return;
        const [moved] = splitState.selectedPages.splice(fromIdx, 1);
        splitState.selectedPages.splice(toIdx, 0, moved);
        renderSplitThumbnailGrid();
        syncPageRangeText();
        debouncedSplitPreview();
      });
    });
  }

  // ── グリッド → テキスト同期 ──
  function syncPageRangeText() {
    if (_splitSyncing) return;
    _splitSyncing = true;
    try {
      const selected = splitState.selectedPages.filter((p) => p.selected).map((p) => p.pageIndex + 1);
      Utils.$('pageRange-pdf-split').value = compactPageRange(selected);
    } finally {
      _splitSyncing = false;
    }
  }

  function compactPageRange(pages) {
    if (pages.length === 0) return '';
    const parts = [];
    let start = pages[0], end = pages[0], prev = pages[0];
    for (let i = 1; i < pages.length; i++) {
      if (pages[i] === prev + 1) { end = pages[i]; }
      else { parts.push(start === end ? `${start}` : `${start}-${end}`); start = pages[i]; end = pages[i]; }
      prev = pages[i];
    }
    parts.push(start === end ? `${start}` : `${start}-${end}`);
    return parts.join(', ');
  }

  // ── テキスト → グリッド同期 ──
  function syncFromPageRangeText() {
    if (_splitSyncing) return;
    _splitSyncing = true;
    try {
      const rangeStr = Utils.$('pageRange-pdf-split').value;
      const pageNumbers = parsePageRangeOrdered(rangeStr, splitState.pageCount);
      const selectedSet = new Set(pageNumbers);

      splitState.selectedPages.forEach((p) => { p.selected = selectedSet.has(p.pageIndex + 1); });

      const selectedMap = new Map();
      splitState.selectedPages.forEach((p) => { if (!selectedMap.has(p.pageIndex)) selectedMap.set(p.pageIndex, p); });

      const reordered = [];
      const used = new Set();
      for (const num of pageNumbers) {
        const page = selectedMap.get(num - 1);
        if (page && !used.has(page.uid)) { reordered.push(page); used.add(page.uid); }
      }
      splitState.selectedPages.forEach((p) => { if (!used.has(p.uid)) { reordered.push(p); used.add(p.uid); } });

      splitState.selectedPages = reordered;
      renderSplitThumbnailGrid();
    } finally {
      _splitSyncing = false;
    }
  }

  function parsePageRangeOrdered(rangeStr, maxPages) {
    const pages = [];
    const seen = new Set();
    for (let part of rangeStr.split(',')) {
      part = part.trim();
      if (!part) continue;
      const match = part.match(/^(\d+)\s*-\s*(\d+)$/);
      if (match) {
        const s = Math.max(1, parseInt(match[1], 10));
        const e = Math.min(maxPages, parseInt(match[2], 10));
        for (let i = s; i <= e; i++) { if (!seen.has(i)) { pages.push(i); seen.add(i); } }
      } else {
        const num = parseInt(part, 10);
        if (!isNaN(num) && num >= 1 && num <= maxPages && !seen.has(num)) { pages.push(num); seen.add(num); }
      }
    }
    return pages;
  }

  // ── 分割ロジック ──
  async function extractPages() {
    if (!splitState.data) return null;
    const selected = splitState.selectedPages.filter((p) => p.selected);
    if (selected.length === 0) return null;

    const src = await PDFDocument.load(splitState.data, { ignoreEncryption: true });
    const newDoc = await PDFDocument.create();
    const indices = selected.map((p) => p.pageIndex);
    const copiedPages = await newDoc.copyPages(src, indices);
    copiedPages.forEach((page) => newDoc.addPage(page));

    if (splitOptions) _applyMetadata(newDoc, splitOptions.getOptions());
    return { doc: newDoc, pageNumbers: selected.map((p) => p.pageIndex + 1) };
  }

  const debouncedSplitPreview = Utils.debounce(async () => {
    if (!splitState.data) return;
    const previewInfo = Utils.$('previewInfo-pdf-split');
    const pdfPreview = Utils.$('pdfPreview-pdf-split');

    try {
      const result = await extractPages();
      if (!result) { previewInfo.textContent = '有効なページ範囲を指定してください'; return; }
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

    Utils.$('pageRange-pdf-split').addEventListener('input', () => {
      syncFromPageRangeText();
      debouncedSplitPreview();
    });
    Utils.$('convertBtn-pdf-split').addEventListener('click', downloadSplit);

    Utils.$('selectAllPages-pdf-split').addEventListener('click', () => {
      splitState.selectedPages.forEach((p) => { p.selected = true; });
      renderSplitThumbnailGrid();
      syncPageRangeText();
      debouncedSplitPreview();
    });
    Utils.$('deselectAllPages-pdf-split').addEventListener('click', () => {
      splitState.selectedPages.forEach((p) => { p.selected = false; });
      renderSplitThumbnailGrid();
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
