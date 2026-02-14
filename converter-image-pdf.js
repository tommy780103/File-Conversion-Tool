/**
 * 画像 → PDF コンバーター
 * 複数画像を1つのPDFに結合。サムネイルグリッド、並び替え、設定対応
 */
const ImagePdfConverter = (() => {
  const S = (id) => Utils.$(`${id}-image-pdf`) || Utils.$(id);

  const state = {
    images: [],   // { id, file, name, thumbnailUrl, dataUrl, width, height }
    idCounter: 0,
    pdfBlobUrl: null,
    pdfBytes: null,       // Uint8Array — 生成済みPDF
    pageEntries: [],      // { uid, pageIndex, label }
    pageUidCounter: 0,
  };

  const hasPdfJs = typeof pdfjsLib !== 'undefined';

  let pdfOptions = null; // PdfOutputOptions インスタンス

  function handleFiles(files) {
    const promises = Array.from(files).map((file) => loadImage(file));
    Promise.all(promises).then(() => {
      showUI();
    });
  }

  function loadImage(file) {
    return new Promise((resolve) => {
      const thumbnailUrl = URL.createObjectURL(file);
      const reader = new FileReader();
      reader.onload = (e) => {
        const img = new Image();
        img.onload = () => {
          state.images.push({
            id: state.idCounter++,
            file,
            name: file.name,
            size: file.size,
            thumbnailUrl,
            dataUrl: e.target.result,
            width: img.naturalWidth,
            height: img.naturalHeight,
          });
          resolve();
        };
        img.onerror = () => {
          URL.revokeObjectURL(thumbnailUrl);
          Toast.show(`${Utils.escapeHtml(file.name)} の読み込みに失敗しました`, 'error');
          resolve();
        };
        img.src = e.target.result;
      };
      reader.readAsDataURL(file);
    });
  }

  function showUI() {
    if (state.images.length === 0) return;
    S('dropZone').classList.add('compact');
    S('imageList').classList.remove('hidden');
    S('defaultSettings').classList.remove('hidden');
    if (pdfOptions) pdfOptions.show();
    S('previewSection').classList.remove('hidden');
    S('actionsPanel').classList.remove('hidden');
    renderImageGrid();
    debouncedUpdatePreview();
  }

  function resetAll() {
    state.images.forEach((img) => URL.revokeObjectURL(img.thumbnailUrl));
    state.images = [];
    state.idCounter = 0;
    state.pdfBlobUrl = Utils.revokeBlobUrl(state.pdfBlobUrl);
    state.pdfBytes = null;
    state.pageEntries = [];
    state.pageUidCounter = 0;
    if (hasPdfJs) PageThumbnail.clearForKey('image-pdf');

    S('dropZone').classList.remove('compact');
    S('imageList').classList.add('hidden');
    S('defaultSettings').classList.add('hidden');
    if (pdfOptions) pdfOptions.hide();
    S('previewSection').classList.add('hidden');
    S('actionsPanel').classList.add('hidden');
    hidePageThumbnails();
    S('imageGrid').innerHTML = '';
  }

  // ── Image Grid ──
  function renderImageGrid() {
    const grid = S('imageGrid');
    grid.innerHTML = '';

    state.images.forEach((img, idx) => {
      const card = document.createElement('div');
      card.className = 'image-card';
      card.dataset.imageId = img.id;
      card.draggable = true;
      card.innerHTML = `
        <img class="image-card-thumbnail" src="${img.thumbnailUrl}" alt="${Utils.escapeHtml(img.name)}" loading="lazy">
        <div class="image-card-info">
          <span class="image-card-name" title="${Utils.escapeHtml(img.name)}">${Utils.escapeHtml(img.name)}</span>
          <span class="image-card-size">${img.width}×${img.height}</span>
        </div>
        <div class="image-card-actions">
          <button class="sheet-btn move-up-btn" title="前に移動" ${idx === 0 ? 'disabled' : ''} data-image-id="${img.id}">
            <svg width="14" height="14" viewBox="0 0 14 14" fill="none"><path d="M7 2v10M3 6l4-4 4 4" stroke="currentColor" stroke-width="1.5" stroke-linecap="round" stroke-linejoin="round"/></svg>
          </button>
          <button class="sheet-btn move-down-btn" title="次に移動" ${idx === state.images.length - 1 ? 'disabled' : ''} data-image-id="${img.id}">
            <svg width="14" height="14" viewBox="0 0 14 14" fill="none"><path d="M7 12V2M3 8l4 4 4-4" stroke="currentColor" stroke-width="1.5" stroke-linecap="round" stroke-linejoin="round"/></svg>
          </button>
          <button class="sheet-btn danger-remove-btn" title="削除" data-image-id="${img.id}">
            <svg width="14" height="14" viewBox="0 0 14 14" fill="none"><path d="M3 3l8 8M11 3l-8 8" stroke="currentColor" stroke-width="1.5" stroke-linecap="round"/></svg>
          </button>
        </div>
      `;
      grid.appendChild(card);
    });

    bindImageEvents();
  }

  let draggedImageId = null;

  function bindImageEvents() {
    const grid = S('imageGrid');

    grid.querySelectorAll('.move-up-btn').forEach((btn) => {
      btn.addEventListener('click', () => {
        const idx = state.images.findIndex((i) => i.id === parseInt(btn.dataset.imageId, 10));
        if (idx > 0) {
          [state.images[idx - 1], state.images[idx]] = [state.images[idx], state.images[idx - 1]];
          renderImageGrid(); debouncedUpdatePreview();
        }
      });
    });

    grid.querySelectorAll('.move-down-btn').forEach((btn) => {
      btn.addEventListener('click', () => {
        const idx = state.images.findIndex((i) => i.id === parseInt(btn.dataset.imageId, 10));
        if (idx < state.images.length - 1) {
          [state.images[idx], state.images[idx + 1]] = [state.images[idx + 1], state.images[idx]];
          renderImageGrid(); debouncedUpdatePreview();
        }
      });
    });

    grid.querySelectorAll('.danger-remove-btn').forEach((btn) => {
      btn.addEventListener('click', () => {
        const id = parseInt(btn.dataset.imageId, 10);
        const img = state.images.find((i) => i.id === id);
        if (img) URL.revokeObjectURL(img.thumbnailUrl);
        state.images = state.images.filter((i) => i.id !== id);
        if (state.images.length === 0) { resetAll(); } else { renderImageGrid(); debouncedUpdatePreview(); }
      });
    });

    // Drag & Drop for reordering
    grid.querySelectorAll('.image-card').forEach((card) => {
      card.addEventListener('dragstart', (e) => {
        draggedImageId = parseInt(card.dataset.imageId, 10);
        card.classList.add('dragging');
        e.dataTransfer.effectAllowed = 'move';
        e.dataTransfer.setData('text/plain', card.dataset.imageId);
      });
      card.addEventListener('dragend', () => {
        card.classList.remove('dragging');
        draggedImageId = null;
        grid.querySelectorAll('.image-card').forEach((c) => c.classList.remove('drag-over-card'));
      });
      card.addEventListener('dragover', (e) => {
        e.preventDefault();
        e.dataTransfer.dropEffect = 'move';
        if (draggedImageId !== null && draggedImageId !== parseInt(card.dataset.imageId, 10)) {
          card.classList.add('drag-over-card');
        }
      });
      card.addEventListener('dragleave', () => card.classList.remove('drag-over-card'));
      card.addEventListener('drop', (e) => {
        e.preventDefault();
        card.classList.remove('drag-over-card');
        const targetId = parseInt(card.dataset.imageId, 10);
        if (draggedImageId === null || draggedImageId === targetId) return;
        const fromIdx = state.images.findIndex((i) => i.id === draggedImageId);
        const toIdx = state.images.findIndex((i) => i.id === targetId);
        if (fromIdx === -1 || toIdx === -1) return;
        const [moved] = state.images.splice(fromIdx, 1);
        state.images.splice(toIdx, 0, moved);
        renderImageGrid(); debouncedUpdatePreview();
      });
    });
  }

  // ── PDF Generation ──
  function getSettings() {
    return {
      pageSize: S('pageSize').value,
      orientation: S('orientation').value,
      margin: parseInt(S('margin').value, 10),
    };
  }

  /**
   * @param {boolean} isPreview - true ならプレビュー用（暗号化なし）
   */
  async function generatePDF(isPreview) {
    if (state.images.length === 0) return null;
    const { jsPDF } = window.jspdf;
    const settings = getSettings();
    const opts = pdfOptions ? pdfOptions.getOptions() : {};

    // ページサイズ定義 (mm)
    const pageSizes = {
      a4: [210, 297], a3: [297, 420], letter: [215.9, 279.4], legal: [215.9, 355.6],
    };

    let doc = null;
    const colorMode = opts.colorMode || 'color';
    const jpegQuality = opts.imageQuality || 0.75;

    for (let idx = 0; idx < state.images.length; idx++) {
      const img = state.images[idx];

      const isLandscape = settings.orientation === 'landscape' ||
        (settings.orientation === 'auto' && img.width > img.height);
      const orient = isLandscape ? 'l' : 'p';

      if (idx === 0) {
        // jsPDF暗号化はコンストラクタ時のみ設定可能
        const jsPdfOpts = { orientation: orient, unit: 'mm', format: settings.pageSize };

        if (!isPreview && (opts.userPassword || opts.ownerPassword)) {
          jsPdfOpts.encryption = {
            userPassword: opts.userPassword || '',
            ownerPassword: opts.ownerPassword || opts.userPassword || '',
            userPermissions: [],
          };
          if (opts.allowPrint) jsPdfOpts.encryption.userPermissions.push('print');
          if (opts.allowCopy) jsPdfOpts.encryption.userPermissions.push('copy');
        }

        doc = new jsPDF(jsPdfOpts);
      } else {
        doc.addPage(settings.pageSize, orient);
      }

      const [baseW, baseH] = pageSizes[settings.pageSize] || pageSizes.a4;
      const pageW = isLandscape ? Math.max(baseW, baseH) : Math.min(baseW, baseH);
      const pageH = isLandscape ? Math.min(baseW, baseH) : Math.max(baseW, baseH);
      const margin = settings.margin;

      const availW = pageW - margin * 2;
      const availH = pageH - margin * 2;

      const imgRatio = img.width / img.height;
      const areaRatio = availW / availH;

      let drawW, drawH;
      if (imgRatio > areaRatio) {
        drawW = availW;
        drawH = availW / imgRatio;
      } else {
        drawH = availH;
        drawW = availH * imgRatio;
      }

      const x = margin + (availW - drawW) / 2;
      const y = margin + (availH - drawH) / 2;

      // 色変換+品質調整
      let imageData = img.dataUrl;
      if (colorMode !== 'color' || jpegQuality < 0.9) {
        imageData = await PdfOutputOptions.utils.processImage(
          img.dataUrl, colorMode, jpegQuality, img.width, img.height
        );
      }

      const format = (colorMode !== 'color' || jpegQuality < 0.9) ? 'JPEG' : _getImageFormat(img.name);
      doc.addImage(imageData, format, x, y, drawW, drawH);
    }

    // メタデータ設定
    if (doc && (opts.title || opts.author || opts.subject || opts.keywords)) {
      doc.setProperties({
        title: opts.title || '',
        author: opts.author || '',
        subject: opts.subject || '',
        keywords: opts.keywords || '',
      });
    }

    return doc;
  }

  function _getImageFormat(filename) {
    const ext = filename.split('.').pop().toLowerCase();
    if (ext === 'jpg' || ext === 'jpeg') return 'JPEG';
    if (ext === 'png') return 'PNG';
    if (ext === 'webp') return 'WEBP';
    if (ext === 'gif') return 'GIF';
    if (ext === 'bmp') return 'BMP';
    return 'JPEG';
  }

  // ── Preview ──
  const debouncedUpdatePreview = Utils.debounce(updatePreview, 300);

  async function updatePreview() {
    if (state.images.length === 0) return;
    const previewInfo = S('previewInfo');
    const pdfPreview = S('pdfPreview');
    previewInfo.textContent = '生成中...';

    try {
      const doc = await generatePDF(true);
      if (!doc) return;
      const ab = doc.output('arraybuffer');
      state.pdfBytes = new Uint8Array(ab);

      state.pdfBlobUrl = Utils.revokeBlobUrl(state.pdfBlobUrl);
      const pageCount = doc.internal.getNumberOfPages();
      const firstName = state.images.length > 0 ? Utils.getBaseName(state.images[0].name) : 'preview';
      state.pdfBlobUrl = Utils.createPdfUrl(state.pdfBytes, Utils.buildDownloadName(firstName, 'pdf'));
      pdfPreview.src = state.pdfBlobUrl;
      previewInfo.textContent = `${state.images.length} 画像・${pageCount} ページ`;

      // ページエントリ初期化
      state.pageEntries = [];
      state.pageUidCounter = 0;
      for (let i = 0; i < pageCount; i++) {
        state.pageEntries.push({ uid: ++state.pageUidCounter, pageIndex: i, label: `P${i + 1}` });
      }
      if (hasPdfJs) {
        PageThumbnail.clearForKey('image-pdf');
        renderPageThumbnailGrid();
      }
    } catch (err) {
      previewInfo.textContent = 'プレビュー生成に失敗しました';
      console.error(err);
    }
  }

  // ── Download ──
  async function download() {
    if (state.images.length === 0) return;
    try {
      Loading.show('PDF変換中...');
      const baseNames = state.images.map((img) => Utils.getBaseName(img.name));
      const pdfFileName = Utils.buildDownloadName(baseNames, 'pdf');

      // ページ管理済みの場合は pdf-lib で再構築
      if (state.pdfBytes && state.pageEntries.length > 0) {
        const finalBytes = await rebuildPdfFromPages();
        if (!finalBytes) { Loading.hide(); return; }
        const url = Utils.createPdfUrl(finalBytes, pdfFileName);
        const a = document.createElement('a');
        a.href = url;
        a.download = pdfFileName;
        document.body.appendChild(a);
        a.click();
        document.body.removeChild(a);
        URL.revokeObjectURL(url);
      } else {
        const doc = await generatePDF(false);
        if (!doc) { Loading.hide(); return; }
        doc.save(pdfFileName);
      }
      Toast.show('PDFのダウンロードが完了しました！', 'success');
    } catch (err) {
      Toast.show('PDF変換中にエラーが発生しました', 'error');
      console.error(err);
    } finally {
      Loading.hide();
    }
  }

  // ── ページサムネイル管理 ──
  let _imageThumbGen = 0;
  let _draggedImageThumbUid = null;

  function hidePageThumbnails() {
    const section = Utils.$('pageThumbnails-image-pdf');
    if (section) section.classList.add('hidden');
    const grid = Utils.$('pageThumbnailGrid-image-pdf');
    if (grid) grid.innerHTML = '';
  }

  function renderPageThumbnailGrid() {
    const section = Utils.$('pageThumbnails-image-pdf');
    const container = Utils.$('pageThumbnailGrid-image-pdf');
    const info = Utils.$('pageThumbnailInfo-image-pdf');
    if (!section || !container) return;

    section.classList.remove('hidden');
    container.innerHTML = '';
    info.textContent = `${state.pageEntries.length} ページ`;

    state.pageEntries.forEach((page) => {
      const cached = PageThumbnail.getCached('image-pdf', page.pageIndex);
      const card = document.createElement('div');
      card.className = 'page-thumb-card';
      card.dataset.pageUid = page.uid;
      card.draggable = true;
      card.innerHTML = `
        <div class="page-thumb-img-wrap">
          <span class="page-thumb-placeholder page-thumb-loading" ${cached ? 'style="display:none"' : ''}>${page.pageIndex + 1}</span>
          <img class="page-thumb-img" ${cached ? `src="${cached}"` : 'style="display:none"'} alt="P${page.pageIndex + 1}">
        </div>
        <div class="page-thumb-info">
          <span class="page-thumb-page-num">${page.label}</span>
        </div>
        <button class="page-thumb-delete" title="削除" data-page-uid="${page.uid}">\u00D7</button>
      `;
      container.appendChild(card);
    });

    bindPageThumbnailEvents();
    renderPageThumbnailsAsync();
  }

  async function renderPageThumbnailsAsync() {
    const gen = ++_imageThumbGen;

    for (const page of state.pageEntries) {
      if (gen !== _imageThumbGen) return;
      if (PageThumbnail.getCached('image-pdf', page.pageIndex)) continue;

      try {
        const dataUrl = await PageThumbnail.render('image-pdf', state.pdfBytes, page.pageIndex);
        if (gen !== _imageThumbGen) return;

        const card = document.querySelector(`#pageThumbnailGrid-image-pdf [data-page-uid="${page.uid}"]`);
        if (card) {
          const img = card.querySelector('.page-thumb-img');
          const ph = card.querySelector('.page-thumb-placeholder');
          if (img) { img.src = dataUrl; img.style.display = 'block'; }
          if (ph) ph.style.display = 'none';
        }
      } catch (err) {
        console.warn('Image-PDF thumbnail render failed:', page.pageIndex, err);
      }

      await new Promise((r) => setTimeout(r, 10));
    }
  }

  function bindPageThumbnailEvents() {
    const container = Utils.$('pageThumbnailGrid-image-pdf');

    // 削除ボタン
    container.querySelectorAll('.page-thumb-delete').forEach((btn) => {
      btn.addEventListener('click', (e) => {
        e.stopPropagation();
        const uid = parseInt(btn.dataset.pageUid, 10);
        state.pageEntries = state.pageEntries.filter((p) => p.uid !== uid);
        if (state.pageEntries.length === 0) {
          hidePageThumbnails();
        } else {
          renderPageThumbnailGrid();
          rebuildPdfPreview();
        }
      });
    });

    // クリック → プレビューモーダル
    container.querySelectorAll('.page-thumb-card').forEach((card) => {
      card.addEventListener('click', (e) => {
        if (e.target.closest('.page-thumb-delete')) return;
        if (card.classList.contains('dragging')) return;
        const uid = parseInt(card.dataset.pageUid, 10);
        const idx = state.pageEntries.findIndex((p) => p.uid === uid);
        if (idx === -1) return;
        PagePreviewModal.open({
          pages: state.pageEntries,
          currentIndex: idx,
          cacheKey: 'image-pdf',
          pdfData: state.pdfBytes,
          getTitle: (page) => page.label,
        });
      });
    });

    // ドラッグ&ドロップ
    container.querySelectorAll('.page-thumb-card').forEach((card) => {
      card.addEventListener('dragstart', (e) => {
        _draggedImageThumbUid = parseInt(card.dataset.pageUid, 10);
        card.classList.add('dragging');
        e.dataTransfer.effectAllowed = 'move';
        e.dataTransfer.setData('text/plain', 'ipage-' + card.dataset.pageUid);
      });
      card.addEventListener('dragend', () => {
        card.classList.remove('dragging');
        _draggedImageThumbUid = null;
        container.querySelectorAll('.page-thumb-card').forEach((c) => c.classList.remove('drag-over-card'));
      });
      card.addEventListener('dragover', (e) => {
        e.preventDefault();
        e.dataTransfer.dropEffect = 'move';
        if (_draggedImageThumbUid !== null && _draggedImageThumbUid !== parseInt(card.dataset.pageUid, 10)) {
          card.classList.add('drag-over-card');
        }
      });
      card.addEventListener('dragleave', () => card.classList.remove('drag-over-card'));
      card.addEventListener('drop', (e) => {
        e.preventDefault();
        card.classList.remove('drag-over-card');
        const targetUid = parseInt(card.dataset.pageUid, 10);
        if (_draggedImageThumbUid === null || _draggedImageThumbUid === targetUid) return;
        const fromIdx = state.pageEntries.findIndex((p) => p.uid === _draggedImageThumbUid);
        const toIdx = state.pageEntries.findIndex((p) => p.uid === targetUid);
        if (fromIdx === -1 || toIdx === -1) return;
        const [moved] = state.pageEntries.splice(fromIdx, 1);
        state.pageEntries.splice(toIdx, 0, moved);
        renderPageThumbnailGrid();
        rebuildPdfPreview();
      });
    });
  }

  async function rebuildPdfFromPages() {
    if (!state.pdfBytes || state.pageEntries.length === 0) return null;
    const { PDFDocument } = PDFLib;
    const src = await PDFDocument.load(state.pdfBytes, { ignoreEncryption: true });
    const newDoc = await PDFDocument.create();
    const indices = state.pageEntries.map((p) => p.pageIndex);
    const copiedPages = await newDoc.copyPages(src, indices);
    copiedPages.forEach((page) => newDoc.addPage(page));
    return newDoc.save();
  }

  const rebuildPdfPreview = Utils.debounce(async () => {
    try {
      const bytes = await rebuildPdfFromPages();
      if (!bytes) return;
      state.pdfBlobUrl = Utils.revokeBlobUrl(state.pdfBlobUrl);
      const firstName = state.images.length > 0 ? Utils.getBaseName(state.images[0].name) : 'preview';
      state.pdfBlobUrl = Utils.createPdfUrl(bytes, Utils.buildDownloadName(firstName, 'pdf'));
      S('pdfPreview').src = state.pdfBlobUrl;
      S('previewInfo').textContent = `${state.pageEntries.length} ページ`;
    } catch (err) {
      console.error('rebuildPdfPreview failed:', err);
    }
  }, 300);

  // ── Init ──
  function init() {
    DropZone.init({
      elementId: 'dropZone-image-pdf',
      inputId: 'fileInput-image-pdf',
      acceptExtensions: ['.jpg', '.jpeg', '.png', '.webp', '.gif', '.bmp'],
      multiple: true,
      onFiles: handleFiles,
    });

    // PDF出力オプションパネルを defaultSettings の後に挿入
    const settingsEl = S('defaultSettings');
    if (settingsEl) {
      pdfOptions = PdfOutputOptions.create({
        container: settingsEl,
        suffix: 'image-pdf',
        features: { metadata: true, security: true, colorMode: true, imageQuality: true },
      });
    }

    S('clearAllImages').addEventListener('click', resetAll);
    S('convertBtn').addEventListener('click', download);

    // 設定変更時にプレビュー更新
    ['pageSize', 'orientation', 'margin'].forEach((key) => {
      S(key).addEventListener('change', debouncedUpdatePreview);
    });

    TabManager.register('image-pdf', {
      onActivate() {},
      onDeactivate() {},
    });
  }

  init();
  return { state };
})();
