/**
 * Excel → PDF コンバーター
 * 既存ロジックを ExcelPdfConverter オブジェクトに再構成
 */
const ExcelPdfConverter = (() => {
  const S = (id) => Utils.$(`${id}-excel-pdf`) || Utils.$(id);

  const state = {
    files: {},
    sheets: [],
    fileIdCounter: 0,
    sheetIdCounter: 0,
    pdfBlobUrl: null,
    pdfBytes: null,       // Uint8Array — 生成済みPDF
    pageEntries: [],      // { uid, pageIndex, label }
    pageUidCounter: 0,
  };

  const hasPdfJs = typeof pdfjsLib !== 'undefined';

  let pdfOptions = null; // PdfOutputOptions インスタンス
  let previewTimer = null;

  function getDefaults() {
    return {
      pageSize: S('defaultPageSize').value,
      orientation: S('defaultOrientation').value,
      fontSize: parseInt(S('defaultFontSize').value, 10),
    };
  }

  function handleFiles(files) {
    for (const file of files) {
      handleFile(file);
    }
  }

  function handleFile(file) {
    const reader = new FileReader();
    reader.onload = (e) => {
      try {
        const data = new Uint8Array(e.target.result);
        const workbook = XLSX.read(data, { type: 'array', cellStyles: true });

        const fileId = state.fileIdCounter++;
        state.files[fileId] = {
          id: fileId,
          name: file.name,
          size: file.size,
          workbook,
          sheetNames: workbook.SheetNames,
        };

        const defaults = getDefaults();
        workbook.SheetNames.forEach((sheetName, idx) => {
          state.sheets.push({
            id: state.sheetIdCounter++,
            fileId,
            sheetIndex: idx,
            sheetName,
            selected: true,
            orientation: defaults.orientation,
            pageSize: defaults.pageSize,
            fontSize: defaults.fontSize,
          });
        });

        showUI();
      } catch (err) {
        Toast.show('ファイルの読み込みに失敗しました', 'error');
        console.error(err);
      }
    };
    reader.readAsArrayBuffer(file);
  }

  function showUI() {
    S('dropZone').classList.add('compact');
    S('fileChips').classList.remove('hidden');
    S('sheetList').classList.remove('hidden');
    S('defaultSettings').classList.remove('hidden');
    if (pdfOptions) pdfOptions.show();
    S('previewSection').classList.remove('hidden');
    S('actionsPanel').classList.remove('hidden');

    renderFileChips();
    renderSheetList();
    debouncedUpdatePreview();
  }

  function resetAll() {
    state.files = {};
    state.sheets = [];
    state.fileIdCounter = 0;
    state.sheetIdCounter = 0;
    state.pdfBytes = null;
    state.pageEntries = [];
    state.pageUidCounter = 0;
    revokePdfUrl();

    S('dropZone').classList.remove('compact');
    S('fileChips').classList.add('hidden');
    S('sheetList').classList.add('hidden');
    S('defaultSettings').classList.add('hidden');
    if (pdfOptions) pdfOptions.hide();
    S('previewSection').classList.add('hidden');
    S('actionsPanel').classList.add('hidden');
    hidePageThumbnails();

    S('fileChipsContainer').innerHTML = '';
    S('sheetListContainer').innerHTML = '';
  }

  // ── File Chips ──
  function renderFileChips() {
    const container = S('fileChipsContainer');
    container.innerHTML = '';

    const fileIds = Object.keys(state.files);
    if (fileIds.length === 0) { resetAll(); return; }

    fileIds.forEach((fid) => {
      const file = state.files[fid];
      const chip = document.createElement('div');
      chip.className = 'file-chip';
      chip.innerHTML = `
        <span class="file-chip-icon">
          <svg width="16" height="16" viewBox="0 0 16 16" fill="none">
            <rect x="2" y="1" width="12" height="14" rx="2" stroke="currentColor" stroke-width="1.5"/>
            <path d="M5 5h6M5 8h6M5 11h3" stroke="currentColor" stroke-width="1" stroke-linecap="round"/>
          </svg>
        </span>
        <span class="file-chip-name" title="${Utils.escapeHtml(file.name)}">${Utils.escapeHtml(file.name)}</span>
        <span class="file-chip-size">${Utils.formatFileSize(file.size)}</span>
        <button class="file-chip-remove" title="削除" data-file-id="${file.id}">
          <svg width="12" height="12" viewBox="0 0 12 12" fill="none">
            <path d="M3 3l6 6M9 3l-6 6" stroke="currentColor" stroke-width="1.5" stroke-linecap="round"/>
          </svg>
        </button>
      `;
      container.appendChild(chip);
    });

    container.querySelectorAll('.file-chip-remove').forEach((btn) => {
      btn.addEventListener('click', (e) => {
        e.stopPropagation();
        removeFile(parseInt(btn.dataset.fileId, 10));
      });
    });
  }

  function removeFile(fileId) {
    delete state.files[fileId];
    state.sheets = state.sheets.filter((s) => s.fileId !== fileId);
    if (Object.keys(state.files).length === 0) {
      resetAll();
    } else {
      renderFileChips();
      renderSheetList();
      debouncedUpdatePreview();
    }
  }

  // ── Sheet List ──
  function renderSheetList() {
    const container = S('sheetListContainer');
    container.innerHTML = '';
    if (state.sheets.length === 0) return;

    const multipleFiles = Object.keys(state.files).length > 1;

    state.sheets.forEach((sheet, idx) => {
      const file = state.files[sheet.fileId];
      const card = document.createElement('div');
      card.className = 'sheet-card';
      card.dataset.sheetId = sheet.id;
      card.draggable = true;

      const pageSizeLabel = sheet.pageSize.toUpperCase();
      const orientLabel = sheet.orientation === 'landscape' ? '横' : '縦';
      const fontLabel = sheet.fontSize + 'pt';

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
          <input type="checkbox" class="sheet-checkbox" ${sheet.selected ? 'checked' : ''} data-sheet-id="${sheet.id}">
          <div class="sheet-info">
            <span class="sheet-name">${Utils.escapeHtml(sheet.sheetName)}</span>
            ${multipleFiles && file ? `<span class="sheet-file-name">${Utils.escapeHtml(file.name)}</span>` : ''}
          </div>
          <div class="sheet-badges">
            <span class="sheet-badge">${pageSizeLabel}</span>
            <span class="sheet-badge">${orientLabel}</span>
            <span class="sheet-badge">${fontLabel}</span>
          </div>
          <div class="sheet-card-actions">
            <button class="sheet-btn move-up-btn" title="上に移動" ${idx === 0 ? 'disabled' : ''} data-sheet-id="${sheet.id}">
              <svg width="14" height="14" viewBox="0 0 14 14" fill="none"><path d="M7 2v10M3 6l4-4 4 4" stroke="currentColor" stroke-width="1.5" stroke-linecap="round" stroke-linejoin="round"/></svg>
            </button>
            <button class="sheet-btn move-down-btn" title="下に移動" ${idx === state.sheets.length - 1 ? 'disabled' : ''} data-sheet-id="${sheet.id}">
              <svg width="14" height="14" viewBox="0 0 14 14" fill="none"><path d="M7 12V2M3 8l4 4 4-4" stroke="currentColor" stroke-width="1.5" stroke-linecap="round" stroke-linejoin="round"/></svg>
            </button>
            <button class="sheet-btn settings-btn" title="設定" data-sheet-id="${sheet.id}">
              <svg width="14" height="14" viewBox="0 0 14 14" fill="none"><circle cx="7" cy="7" r="2" stroke="currentColor" stroke-width="1.5"/><path d="M7 1v2M7 11v2M1 7h2M11 7h2M2.8 2.8l1.4 1.4M9.8 9.8l1.4 1.4M11.2 2.8l-1.4 1.4M4.2 9.8l-1.4 1.4" stroke="currentColor" stroke-width="1" stroke-linecap="round"/></svg>
            </button>
          </div>
        </div>
        <div class="sheet-settings" data-settings-for="${sheet.id}">
          <div class="options-grid">
            <div class="option-group">
              <label>用紙サイズ</label>
              <div class="select-wrapper">
                <select class="sheet-pageSize" data-sheet-id="${sheet.id}">
                  <option value="a4" ${sheet.pageSize === 'a4' ? 'selected' : ''}>A4</option>
                  <option value="a3" ${sheet.pageSize === 'a3' ? 'selected' : ''}>A3</option>
                  <option value="letter" ${sheet.pageSize === 'letter' ? 'selected' : ''}>Letter</option>
                  <option value="legal" ${sheet.pageSize === 'legal' ? 'selected' : ''}>Legal</option>
                </select>
              </div>
            </div>
            <div class="option-group">
              <label>向き</label>
              <div class="select-wrapper">
                <select class="sheet-orientation" data-sheet-id="${sheet.id}">
                  <option value="portrait" ${sheet.orientation === 'portrait' ? 'selected' : ''}>縦</option>
                  <option value="landscape" ${sheet.orientation === 'landscape' ? 'selected' : ''}>横</option>
                </select>
              </div>
            </div>
            <div class="option-group">
              <label>フォントサイズ</label>
              <div class="select-wrapper">
                <select class="sheet-fontSize" data-sheet-id="${sheet.id}">
                  <option value="7" ${sheet.fontSize === 7 ? 'selected' : ''}>7pt</option>
                  <option value="8" ${sheet.fontSize === 8 ? 'selected' : ''}>8pt</option>
                  <option value="9" ${sheet.fontSize === 9 ? 'selected' : ''}>9pt</option>
                  <option value="10" ${sheet.fontSize === 10 ? 'selected' : ''}>10pt</option>
                  <option value="12" ${sheet.fontSize === 12 ? 'selected' : ''}>12pt</option>
                </select>
              </div>
            </div>
          </div>
        </div>
      `;
      container.appendChild(card);
    });

    bindSheetEvents();
  }

  function bindSheetEvents() {
    const container = S('sheetListContainer');

    container.querySelectorAll('.sheet-checkbox').forEach((cb) => {
      cb.addEventListener('change', () => {
        const id = parseInt(cb.dataset.sheetId, 10);
        const sheet = state.sheets.find((s) => s.id === id);
        if (sheet) { sheet.selected = cb.checked; debouncedUpdatePreview(); }
      });
    });

    container.querySelectorAll('.move-up-btn').forEach((btn) => {
      btn.addEventListener('click', () => {
        const id = parseInt(btn.dataset.sheetId, 10);
        const idx = state.sheets.findIndex((s) => s.id === id);
        if (idx > 0) {
          [state.sheets[idx - 1], state.sheets[idx]] = [state.sheets[idx], state.sheets[idx - 1]];
          renderSheetList(); debouncedUpdatePreview();
        }
      });
    });

    container.querySelectorAll('.move-down-btn').forEach((btn) => {
      btn.addEventListener('click', () => {
        const id = parseInt(btn.dataset.sheetId, 10);
        const idx = state.sheets.findIndex((s) => s.id === id);
        if (idx < state.sheets.length - 1) {
          [state.sheets[idx], state.sheets[idx + 1]] = [state.sheets[idx + 1], state.sheets[idx]];
          renderSheetList(); debouncedUpdatePreview();
        }
      });
    });

    container.querySelectorAll('.settings-btn').forEach((btn) => {
      btn.addEventListener('click', () => {
        const id = btn.dataset.sheetId;
        const settingsEl = container.querySelector(`[data-settings-for="${id}"]`);
        if (settingsEl) { settingsEl.classList.toggle('open'); btn.classList.toggle('active'); }
      });
    });

    container.querySelectorAll('.sheet-pageSize').forEach((sel) => {
      sel.addEventListener('change', () => {
        const sheet = state.sheets.find((s) => s.id === parseInt(sel.dataset.sheetId, 10));
        if (sheet) { sheet.pageSize = sel.value; renderSheetList(); debouncedUpdatePreview(); }
      });
    });

    container.querySelectorAll('.sheet-orientation').forEach((sel) => {
      sel.addEventListener('change', () => {
        const sheet = state.sheets.find((s) => s.id === parseInt(sel.dataset.sheetId, 10));
        if (sheet) { sheet.orientation = sel.value; renderSheetList(); debouncedUpdatePreview(); }
      });
    });

    container.querySelectorAll('.sheet-fontSize').forEach((sel) => {
      sel.addEventListener('change', () => {
        const sheet = state.sheets.find((s) => s.id === parseInt(sel.dataset.sheetId, 10));
        if (sheet) { sheet.fontSize = parseInt(sel.value, 10); renderSheetList(); debouncedUpdatePreview(); }
      });
    });

    bindDragAndDrop();
  }

  // ── Drag & Drop ──
  let draggedSheetId = null;

  function bindDragAndDrop() {
    const container = S('sheetListContainer');
    const cards = container.querySelectorAll('.sheet-card');
    cards.forEach((card) => {
      card.addEventListener('dragstart', (e) => {
        draggedSheetId = parseInt(card.dataset.sheetId, 10);
        card.classList.add('dragging');
        e.dataTransfer.effectAllowed = 'move';
        e.dataTransfer.setData('text/plain', card.dataset.sheetId);
      });
      card.addEventListener('dragend', () => {
        card.classList.remove('dragging');
        draggedSheetId = null;
        container.querySelectorAll('.sheet-card').forEach((c) => c.classList.remove('drag-over-card'));
      });
      card.addEventListener('dragover', (e) => {
        e.preventDefault();
        e.dataTransfer.dropEffect = 'move';
        const targetId = parseInt(card.dataset.sheetId, 10);
        if (draggedSheetId !== null && draggedSheetId !== targetId) card.classList.add('drag-over-card');
      });
      card.addEventListener('dragleave', () => card.classList.remove('drag-over-card'));
      card.addEventListener('drop', (e) => {
        e.preventDefault();
        card.classList.remove('drag-over-card');
        const targetId = parseInt(card.dataset.sheetId, 10);
        if (draggedSheetId === null || draggedSheetId === targetId) return;
        const fromIdx = state.sheets.findIndex((s) => s.id === draggedSheetId);
        const toIdx = state.sheets.findIndex((s) => s.id === targetId);
        if (fromIdx === -1 || toIdx === -1) return;
        const [moved] = state.sheets.splice(fromIdx, 1);
        state.sheets.splice(toIdx, 0, moved);
        renderSheetList(); debouncedUpdatePreview();
      });
    });
  }

  // ── Excel Style Helpers ──
  function parseExcelColor(colorObj) {
    if (!colorObj || !colorObj.rgb) return null;
    const hex = colorObj.rgb.length === 8 ? colorObj.rgb.substr(2) : colorObj.rgb;
    const r = parseInt(hex.substr(0, 2), 16);
    const g = parseInt(hex.substr(2, 2), 16);
    const b = parseInt(hex.substr(4, 2), 16);
    if (isNaN(r) || isNaN(g) || isNaN(b)) return null;
    return [r, g, b];
  }

  function getExcelCellStyle(ws, row, col) {
    const ref = XLSX.utils.encode_cell({ r: row, c: col });
    const cell = ws[ref];
    if (!cell || !cell.s) return null;
    const s = cell.s;
    const result = {};
    if (s.fill && s.fill.fgColor) {
      const bg = parseExcelColor(s.fill.fgColor);
      if (bg) result.fillColor = bg;
    }
    if (s.font) {
      if (s.font.color) {
        const fc = parseExcelColor(s.font.color);
        if (fc) result.textColor = fc;
      }
      if (s.font.bold) result.fontStyle = 'bold';
    }
    if (s.alignment && s.alignment.horizontal) {
      result.halign = s.alignment.horizontal;
    }
    return Object.keys(result).length > 0 ? result : null;
  }

  // ── PDF Generation ──
  /**
   * @param {boolean} isPreview - true ならプレビュー用（暗号化なし）
   */
  async function generatePDF(isPreview) {
    const { jsPDF } = window.jspdf;
    const selectedSheets = state.sheets.filter((s) => s.selected);
    if (selectedSheets.length === 0) return null;

    const opts = pdfOptions ? pdfOptions.getOptions() : {};
    const colorMode = opts.colorMode || 'color';

    // 日本語フォントを読み込み
    await JapaneseFont.load();

    const multipleFiles = Object.keys(state.files).length > 1;
    const first = selectedSheets[0];

    // jsPDF暗号化はコンストラクタ時のみ設定可能
    const jsPdfOpts = {
      orientation: first.orientation === 'landscape' ? 'l' : 'p',
      unit: 'mm',
      format: first.pageSize,
    };

    if (!isPreview && (opts.userPassword || opts.ownerPassword)) {
      jsPdfOpts.encryption = {
        userPassword: opts.userPassword || '',
        ownerPassword: opts.ownerPassword || opts.userPassword || '',
        userPermissions: [],
      };
      if (opts.allowPrint) jsPdfOpts.encryption.userPermissions.push('print');
      if (opts.allowCopy) jsPdfOpts.encryption.userPermissions.push('copy');
    }

    const doc = new jsPDF(jsPdfOpts);

    // 日本語フォントを登録
    const fontAvailable = JapaneseFont.register(doc);
    const fontName = fontAvailable ? JapaneseFont.FONT_NAME : undefined;

    for (let s = 0; s < selectedSheets.length; s++) {
      const sheetConf = selectedSheets[s];
      const file = state.files[sheetConf.fileId];
      if (!file) continue;

      if (s > 0) doc.addPage(sheetConf.pageSize, sheetConf.orientation === 'landscape' ? 'l' : 'p');

      const ws = file.workbook.Sheets[sheetConf.sheetName];
      const jsonData = XLSX.utils.sheet_to_json(ws, { header: 1, defval: '' });
      if (jsonData.length === 0) continue;

      const headers = jsonData[0].map((h) => String(h));
      const body = jsonData.slice(1).map((row) =>
        headers.map((_, i) => String(row[i] !== undefined ? row[i] : ''))
      );

      let titleText = sheetConf.sheetName;
      if (multipleFiles && file) titleText = `${sheetConf.sheetName}（${file.name}）`;

      const showTitle = selectedSheets.length > 1 || multipleFiles;
      if (showTitle) {
        doc.setFontSize(12);
        doc.setTextColor(100, 100, 100);
        if (fontAvailable) doc.setFont(fontName);
        doc.text(titleText, 14, 12);
      }

      doc.autoTable({
        head: [headers],
        body,
        startY: showTitle ? 16 : 10,
        styles: { font: fontName, fontSize: sheetConf.fontSize, cellPadding: 2, overflow: 'linebreak', lineColor: [200, 200, 200], lineWidth: 0.1 },
        headStyles: { font: fontName, fillColor: [255, 255, 255], textColor: [0, 0, 0] },
        margin: { top: 10, right: 10, bottom: 10, left: 10 },
        didParseCell: (data) => {
          const excelRow = data.section === 'head' ? 0 : data.row.index + 1;
          const excelCol = data.column.index;
          const cellStyle = getExcelCellStyle(ws, excelRow, excelCol);
          if (cellStyle) {
            if (cellStyle.fillColor) data.cell.styles.fillColor = cellStyle.fillColor;
            if (cellStyle.textColor) data.cell.styles.textColor = cellStyle.textColor;
            if (cellStyle.fontStyle) data.cell.styles.fontStyle = cellStyle.fontStyle;
            if (cellStyle.halign) data.cell.styles.halign = cellStyle.halign;
          }
          // 色変換を適用
          if (colorMode !== 'color') {
            if (data.cell.styles.fillColor && Array.isArray(data.cell.styles.fillColor)) {
              data.cell.styles.fillColor = PdfOutputOptions.utils.convertColor(data.cell.styles.fillColor, colorMode);
            }
            if (data.cell.styles.textColor && Array.isArray(data.cell.styles.textColor)) {
              data.cell.styles.textColor = PdfOutputOptions.utils.convertColor(data.cell.styles.textColor, colorMode);
            }
          }
        },
        didDrawPage: (data) => {
          const pageCount = doc.internal.getNumberOfPages();
          doc.setFontSize(8);
          const footerColor = colorMode !== 'color'
            ? PdfOutputOptions.utils.convertColor([150, 150, 150], colorMode)
            : [150, 150, 150];
          doc.setTextColor(...footerColor);
          if (fontAvailable) doc.setFont(fontName);
          const pw = doc.internal.pageSize.getWidth();
          const ph = doc.internal.pageSize.getHeight();
          doc.text(`${data.pageNumber} / ${pageCount}`, pw / 2, ph - 6, { align: 'center' });
        },
      });
    }

    // メタデータ設定
    if (opts.title || opts.author || opts.subject || opts.keywords) {
      doc.setProperties({
        title: opts.title || '',
        author: opts.author || '',
        subject: opts.subject || '',
        keywords: opts.keywords || '',
      });
    }

    return doc;
  }

  // ── Preview ──
  function revokePdfUrl() {
    state.pdfBlobUrl = Utils.revokeBlobUrl(state.pdfBlobUrl);
  }

  const debouncedUpdatePreview = Utils.debounce(updatePreview, 300);

  function updatePreview() {
    if (Object.keys(state.files).length === 0) return;
    const selectedCount = state.sheets.filter((s) => s.selected).length;
    const previewInfo = S('previewInfo');
    const pdfPreview = S('pdfPreview');

    if (selectedCount === 0) {
      previewInfo.textContent = 'シートが選択されていません';
      revokePdfUrl();
      pdfPreview.src = '';
      hidePageThumbnails();
      return;
    }
    previewInfo.textContent = JapaneseFont.isLoaded() ? '生成中...' : 'フォント読み込み中...';

    setTimeout(async () => {
      try {
        const doc = await generatePDF(true);
        if (!doc) { previewInfo.textContent = 'シートが選択されていません'; return; }
        const ab = doc.output('arraybuffer');
        state.pdfBytes = new Uint8Array(ab);

        revokePdfUrl();
        const pageCount = doc.internal.getNumberOfPages();
        const fileIds = Object.keys(state.files);
        const firstName = fileIds.length > 0 ? Utils.getBaseName(state.files[fileIds[0]].name) : 'preview';
        state.pdfBlobUrl = Utils.createPdfUrl(state.pdfBytes, Utils.buildDownloadName(firstName, 'pdf'));
        pdfPreview.src = state.pdfBlobUrl;
        previewInfo.textContent = `${selectedCount} シート・${pageCount} ページ`;

        // ページエントリ初期化
        state.pageEntries = [];
        state.pageUidCounter = 0;
        for (let i = 0; i < pageCount; i++) {
          state.pageEntries.push({ uid: ++state.pageUidCounter, pageIndex: i, label: `P${i + 1}` });
        }
        if (hasPdfJs) {
          PageThumbnail.clearForKey('excel-pdf');
          renderPageThumbnailGrid();
        }
      } catch (err) {
        previewInfo.textContent = 'プレビュー生成に失敗しました';
        console.error(err);
      }
    }, 50);
  }

  // ── Download ──
  async function download() {
    if (Object.keys(state.files).length === 0) return;
    try {
      Loading.show('PDF変換中...');
      const fileIds = Object.keys(state.files);
      const baseNames = fileIds.map((fid) => Utils.getBaseName(state.files[fid].name));
      const pdfFileName = Utils.buildDownloadName(baseNames, 'pdf');

      // ページ管理済みの場合は pdf-lib で再構築
      if (state.pdfBytes && state.pageEntries.length > 0) {
        const finalBytes = await rebuildPdfFromPages();
        if (!finalBytes) { Toast.show('シートが選択されていません', 'error'); Loading.hide(); return; }
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
        if (!doc) { Toast.show('シートが選択されていません', 'error'); Loading.hide(); return; }
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
  let _excelThumbGen = 0;
  let _draggedExcelThumbUid = null;

  function hidePageThumbnails() {
    const section = Utils.$('pageThumbnails-excel-pdf');
    if (section) section.classList.add('hidden');
    const grid = Utils.$('pageThumbnailGrid-excel-pdf');
    if (grid) grid.innerHTML = '';
  }

  function renderPageThumbnailGrid() {
    const section = Utils.$('pageThumbnails-excel-pdf');
    const container = Utils.$('pageThumbnailGrid-excel-pdf');
    const info = Utils.$('pageThumbnailInfo-excel-pdf');
    if (!section || !container) return;

    section.classList.remove('hidden');
    container.innerHTML = '';
    info.textContent = `${state.pageEntries.length} ページ`;

    state.pageEntries.forEach((page) => {
      const cached = PageThumbnail.getCached('excel-pdf', page.pageIndex);
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
    const gen = ++_excelThumbGen;

    for (const page of state.pageEntries) {
      if (gen !== _excelThumbGen) return;
      if (PageThumbnail.getCached('excel-pdf', page.pageIndex)) continue;

      try {
        const dataUrl = await PageThumbnail.render('excel-pdf', state.pdfBytes, page.pageIndex);
        if (gen !== _excelThumbGen) return;

        const card = document.querySelector(`#pageThumbnailGrid-excel-pdf [data-page-uid="${page.uid}"]`);
        if (card) {
          const img = card.querySelector('.page-thumb-img');
          const ph = card.querySelector('.page-thumb-placeholder');
          if (img) { img.src = dataUrl; img.style.display = 'block'; }
          if (ph) ph.style.display = 'none';
        }
      } catch (err) {
        console.warn('Excel-PDF thumbnail render failed:', page.pageIndex, err);
      }

      await new Promise((r) => setTimeout(r, 10));
    }
  }

  function bindPageThumbnailEvents() {
    const container = Utils.$('pageThumbnailGrid-excel-pdf');

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
          cacheKey: 'excel-pdf',
          pdfData: state.pdfBytes,
          getTitle: (page) => page.label,
        });
      });
    });

    // ドラッグ&ドロップ
    container.querySelectorAll('.page-thumb-card').forEach((card) => {
      card.addEventListener('dragstart', (e) => {
        _draggedExcelThumbUid = parseInt(card.dataset.pageUid, 10);
        card.classList.add('dragging');
        e.dataTransfer.effectAllowed = 'move';
        e.dataTransfer.setData('text/plain', 'epage-' + card.dataset.pageUid);
      });
      card.addEventListener('dragend', () => {
        card.classList.remove('dragging');
        _draggedExcelThumbUid = null;
        container.querySelectorAll('.page-thumb-card').forEach((c) => c.classList.remove('drag-over-card'));
      });
      card.addEventListener('dragover', (e) => {
        e.preventDefault();
        e.dataTransfer.dropEffect = 'move';
        if (_draggedExcelThumbUid !== null && _draggedExcelThumbUid !== parseInt(card.dataset.pageUid, 10)) {
          card.classList.add('drag-over-card');
        }
      });
      card.addEventListener('dragleave', () => card.classList.remove('drag-over-card'));
      card.addEventListener('drop', (e) => {
        e.preventDefault();
        card.classList.remove('drag-over-card');
        const targetUid = parseInt(card.dataset.pageUid, 10);
        if (_draggedExcelThumbUid === null || _draggedExcelThumbUid === targetUid) return;
        const fromIdx = state.pageEntries.findIndex((p) => p.uid === _draggedExcelThumbUid);
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
      revokePdfUrl();
      const fileIds = Object.keys(state.files);
      const firstName = fileIds.length > 0 ? Utils.getBaseName(state.files[fileIds[0]].name) : 'preview';
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
      elementId: 'dropZone-excel-pdf',
      inputId: 'fileInput-excel-pdf',
      acceptExtensions: ['.xlsx', '.xls', '.csv'],
      multiple: true,
      onFiles: handleFiles,
    });

    // PDF出力オプションパネルを defaultSettings の後に挿入
    const settingsEl = S('defaultSettings');
    if (settingsEl) {
      pdfOptions = PdfOutputOptions.create({
        container: settingsEl,
        suffix: 'excel-pdf',
        features: { metadata: true, security: true, colorMode: true, imageQuality: false },
      });
    }

    S('clearAllFiles').addEventListener('click', resetAll);

    S('selectAllSheets').addEventListener('click', () => {
      state.sheets.forEach((s) => { s.selected = true; });
      renderSheetList(); debouncedUpdatePreview();
    });

    S('deselectAllSheets').addEventListener('click', () => {
      state.sheets.forEach((s) => { s.selected = false; });
      renderSheetList(); debouncedUpdatePreview();
    });

    S('applyToAll').addEventListener('click', () => {
      const defaults = getDefaults();
      state.sheets.forEach((sheet) => {
        sheet.pageSize = defaults.pageSize;
        sheet.orientation = defaults.orientation;
        sheet.fontSize = defaults.fontSize;
      });
      renderSheetList(); debouncedUpdatePreview();
    });

    S('convertBtn').addEventListener('click', download);

    TabManager.register('excel-pdf', {
      onActivate() {},
      onDeactivate() {},
    });
  }

  init();
  return { state };
})();
