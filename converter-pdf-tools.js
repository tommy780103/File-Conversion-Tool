/**
 * PDF 結合・分割 コンバーター
 * pdf-lib を使用。結合と分割（ページ抽出）のサブモード切替
 */
const PdfToolsConverter = (() => {
  const panel = Utils.$('panel-pdf-tools');
  const { PDFDocument } = PDFLib;

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
    idCounter: 0,
    previewUrl: null,
  };

  let mergeOptions = null; // PdfOutputOptions インスタンス

  async function handleMergeFiles(files) {
    Loading.show('PDFを読み込み中...');
    try {
      for (const file of files) {
        const data = new Uint8Array(await file.arrayBuffer());
        const pdfDoc = await PDFDocument.load(data, { ignoreEncryption: true });
        mergeState.pdfs.push({
          id: mergeState.idCounter++,
          name: file.name,
          size: file.size,
          data,
          pageCount: pdfDoc.getPageCount(),
        });
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
    mergeState.idCounter = 0;
    mergeState.previewUrl = Utils.revokeBlobUrl(mergeState.previewUrl);

    Utils.$('dropZone-pdf-merge').classList.remove('compact');
    Utils.$('pdfList-pdf-merge').classList.add('hidden');
    if (mergeOptions) mergeOptions.hide();
    Utils.$('previewSection-pdf-merge').classList.add('hidden');
    Utils.$('actionsPanel-pdf-merge').classList.add('hidden');
    Utils.$('pdfListContainer-pdf-merge').innerHTML = '';
  }

  function renderMergeList() {
    const container = Utils.$('pdfListContainer-pdf-merge');
    container.innerHTML = '';

    mergeState.pdfs.forEach((pdf, idx) => {
      const card = document.createElement('div');
      card.className = 'sheet-card';
      card.dataset.pdfId = pdf.id;
      card.draggable = true;
      card.innerHTML = `
        <div class="sheet-card-main">
          <div class="sheet-drag-handle" title="ドラッグして並び替え">
            <svg width="14" height="14" viewBox="0 0 14 14" fill="none">
              <circle cx="4" cy="3" r="1.2" fill="currentColor"/>
              <circle cx="10" cy="3" r="1.2" fill="currentColor"/>
              <circle cx="4" cy="7" r="1.2" fill="currentColor"/>
              <circle cx="10" cy="7" r="1.2" fill="currentColor"/>
              <circle cx="4" cy="11" r="1.2" fill="currentColor"/>
              <circle cx="10" cy="11" r="1.2" fill="currentColor"/>
            </svg>
          </div>
          <div class="sheet-info">
            <span class="sheet-name">${Utils.escapeHtml(pdf.name)}</span>
            <span class="sheet-file-name">${pdf.pageCount} ページ ・ ${Utils.formatFileSize(pdf.size)}</span>
          </div>
          <div class="sheet-card-actions">
            <button class="sheet-btn move-up-btn" title="上に移動" ${idx === 0 ? 'disabled' : ''} data-pdf-id="${pdf.id}">
              <svg width="14" height="14" viewBox="0 0 14 14" fill="none"><path d="M7 2v10M3 6l4-4 4 4" stroke="currentColor" stroke-width="1.5" stroke-linecap="round" stroke-linejoin="round"/></svg>
            </button>
            <button class="sheet-btn move-down-btn" title="下に移動" ${idx === mergeState.pdfs.length - 1 ? 'disabled' : ''} data-pdf-id="${pdf.id}">
              <svg width="14" height="14" viewBox="0 0 14 14" fill="none"><path d="M7 12V2M3 8l4 4 4-4" stroke="currentColor" stroke-width="1.5" stroke-linecap="round" stroke-linejoin="round"/></svg>
            </button>
            <button class="sheet-btn danger-remove-btn" title="削除" data-pdf-id="${pdf.id}">
              <svg width="14" height="14" viewBox="0 0 14 14" fill="none"><path d="M3 3l8 8M11 3l-8 8" stroke="currentColor" stroke-width="1.5" stroke-linecap="round"/></svg>
            </button>
          </div>
        </div>
      `;
      container.appendChild(card);
    });

    bindMergeEvents();
  }

  let draggedPdfId = null;

  function bindMergeEvents() {
    const container = Utils.$('pdfListContainer-pdf-merge');

    container.querySelectorAll('.move-up-btn').forEach((btn) => {
      btn.addEventListener('click', () => {
        const idx = mergeState.pdfs.findIndex((p) => p.id === parseInt(btn.dataset.pdfId, 10));
        if (idx > 0) {
          [mergeState.pdfs[idx - 1], mergeState.pdfs[idx]] = [mergeState.pdfs[idx], mergeState.pdfs[idx - 1]];
          renderMergeList(); debouncedMergePreview();
        }
      });
    });

    container.querySelectorAll('.move-down-btn').forEach((btn) => {
      btn.addEventListener('click', () => {
        const idx = mergeState.pdfs.findIndex((p) => p.id === parseInt(btn.dataset.pdfId, 10));
        if (idx < mergeState.pdfs.length - 1) {
          [mergeState.pdfs[idx], mergeState.pdfs[idx + 1]] = [mergeState.pdfs[idx + 1], mergeState.pdfs[idx]];
          renderMergeList(); debouncedMergePreview();
        }
      });
    });

    container.querySelectorAll('.danger-remove-btn').forEach((btn) => {
      btn.addEventListener('click', () => {
        mergeState.pdfs = mergeState.pdfs.filter((p) => p.id !== parseInt(btn.dataset.pdfId, 10));
        if (mergeState.pdfs.length === 0) { resetMerge(); } else { renderMergeList(); debouncedMergePreview(); }
      });
    });

    // Drag & Drop
    container.querySelectorAll('.sheet-card').forEach((card) => {
      card.addEventListener('dragstart', (e) => {
        draggedPdfId = parseInt(card.dataset.pdfId, 10);
        card.classList.add('dragging');
        e.dataTransfer.effectAllowed = 'move';
        e.dataTransfer.setData('text/plain', card.dataset.pdfId);
      });
      card.addEventListener('dragend', () => {
        card.classList.remove('dragging');
        draggedPdfId = null;
        container.querySelectorAll('.sheet-card').forEach((c) => c.classList.remove('drag-over-card'));
      });
      card.addEventListener('dragover', (e) => {
        e.preventDefault();
        e.dataTransfer.dropEffect = 'move';
        if (draggedPdfId !== null && draggedPdfId !== parseInt(card.dataset.pdfId, 10)) {
          card.classList.add('drag-over-card');
        }
      });
      card.addEventListener('dragleave', () => card.classList.remove('drag-over-card'));
      card.addEventListener('drop', (e) => {
        e.preventDefault();
        card.classList.remove('drag-over-card');
        const targetId = parseInt(card.dataset.pdfId, 10);
        if (draggedPdfId === null || draggedPdfId === targetId) return;
        const fromIdx = mergeState.pdfs.findIndex((p) => p.id === draggedPdfId);
        const toIdx = mergeState.pdfs.findIndex((p) => p.id === targetId);
        if (fromIdx === -1 || toIdx === -1) return;
        const [moved] = mergeState.pdfs.splice(fromIdx, 1);
        mergeState.pdfs.splice(toIdx, 0, moved);
        renderMergeList(); debouncedMergePreview();
      });
    });
  }

  function _applyMetadata(pdfDoc, opts) {
    if (opts.title) pdfDoc.setTitle(opts.title);
    if (opts.author) pdfDoc.setAuthor(opts.author);
    if (opts.subject) pdfDoc.setSubject(opts.subject);
    if (opts.keywords) pdfDoc.setKeywords(opts.keywords.split(',').map((k) => k.trim()));
  }

  async function mergePdfs() {
    if (mergeState.pdfs.length === 0) return null;
    const merged = await PDFDocument.create();
    for (const pdf of mergeState.pdfs) {
      const src = await PDFDocument.load(pdf.data, { ignoreEncryption: true });
      const pages = await merged.copyPages(src, src.getPageIndices());
      pages.forEach((page) => merged.addPage(page));
    }
    // メタデータ設定
    if (mergeOptions) {
      _applyMetadata(merged, mergeOptions.getOptions());
    }
    return merged;
  }

  const debouncedMergePreview = Utils.debounce(async () => {
    if (mergeState.pdfs.length === 0) return;
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

      const totalPages = mergeState.pdfs.reduce((sum, p) => sum + p.pageCount, 0);
      previewInfo.textContent = `${mergeState.pdfs.length} ファイル・${totalPages} ページ`;
    } catch (err) {
      previewInfo.textContent = '結合に失敗しました';
      console.error(err);
    }
  }, 300);

  async function downloadMerged() {
    if (mergeState.pdfs.length === 0) return;
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
    previewUrl: null,
  };

  let splitOptions = null; // PdfOutputOptions インスタンス

  async function handleSplitFile(files) {
    const file = files[0];
    if (!file) return;
    splitState.fileName = file.name;

    Loading.show('PDFを読み込み中...');
    try {
      splitState.data = new Uint8Array(await file.arrayBuffer());
      const pdfDoc = await PDFDocument.load(splitState.data, { ignoreEncryption: true });
      splitState.pageCount = pdfDoc.getPageCount();

      Utils.$('splitOptions-pdf-split').classList.remove('hidden');
      Utils.$('splitPageInfo-pdf-split').textContent = `全 ${splitState.pageCount} ページ`;
      Utils.$('pageRange-pdf-split').value = `1-${splitState.pageCount}`;
      if (splitOptions) splitOptions.show();
      Utils.$('actionsPanel-pdf-split').classList.remove('hidden');

      debouncedSplitPreview();
    } catch (err) {
      Toast.show('PDFの読み込みに失敗しました', 'error');
      console.error(err);
    } finally {
      Loading.hide();
    }
  }

  function parsePageRange(rangeStr, maxPages) {
    const pages = new Set();
    const parts = rangeStr.split(',');
    for (let part of parts) {
      part = part.trim();
      if (!part) continue;
      const match = part.match(/^(\d+)\s*-\s*(\d+)$/);
      if (match) {
        const start = Math.max(1, parseInt(match[1], 10));
        const end = Math.min(maxPages, parseInt(match[2], 10));
        for (let i = start; i <= end; i++) pages.add(i);
      } else {
        const num = parseInt(part, 10);
        if (!isNaN(num) && num >= 1 && num <= maxPages) pages.add(num);
      }
    }
    return Array.from(pages).sort((a, b) => a - b);
  }

  async function extractPages() {
    if (!splitState.data) return null;
    const rangeStr = Utils.$('pageRange-pdf-split').value;
    const pageNumbers = parsePageRange(rangeStr, splitState.pageCount);
    if (pageNumbers.length === 0) return null;

    const src = await PDFDocument.load(splitState.data, { ignoreEncryption: true });
    const newDoc = await PDFDocument.create();
    const indices = pageNumbers.map((p) => p - 1); // 0-based
    const copiedPages = await newDoc.copyPages(src, indices);
    copiedPages.forEach((page) => newDoc.addPage(page));

    // メタデータ設定
    if (splitOptions) {
      _applyMetadata(newDoc, splitOptions.getOptions());
    }

    return { doc: newDoc, pageNumbers };
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

    // 結合用オプションパネル（メタデータのみ）
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

    // 分割用オプションパネル（メタデータのみ）
    const splitOptsEl = Utils.$('splitOptions-pdf-split');
    if (splitOptsEl) {
      splitOptions = PdfOutputOptions.create({
        container: splitOptsEl,
        suffix: 'pdf-split',
        features: { metadata: true, security: false, colorMode: false, imageQuality: false },
      });
    }

    Utils.$('pageRange-pdf-split').addEventListener('input', debouncedSplitPreview);
    Utils.$('convertBtn-pdf-split').addEventListener('click', downloadSplit);

    TabManager.register('pdf-tools', {
      onActivate() {},
      onDeactivate() {},
    });
  }

  init();
  return { mergeState, splitState };
})();
