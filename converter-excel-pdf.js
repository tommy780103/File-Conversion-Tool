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
  };

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
    revokePdfUrl();

    S('dropZone').classList.remove('compact');
    S('fileChips').classList.add('hidden');
    S('sheetList').classList.add('hidden');
    S('defaultSettings').classList.add('hidden');
    S('previewSection').classList.add('hidden');
    S('actionsPanel').classList.add('hidden');

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

  // ── PDF Generation ──
  async function generatePDF() {
    const { jsPDF } = window.jspdf;
    const selectedSheets = state.sheets.filter((s) => s.selected);
    if (selectedSheets.length === 0) return null;

    // 日本語フォントを読み込み
    await JapaneseFont.load();

    const multipleFiles = Object.keys(state.files).length > 1;
    const first = selectedSheets[0];
    const doc = new jsPDF({
      orientation: first.orientation === 'landscape' ? 'l' : 'p',
      unit: 'mm',
      format: first.pageSize,
    });

    // 日本語フォントを登録
    const fontAvailable = JapaneseFont.register(doc);
    const fontName = fontAvailable ? JapaneseFont.FONT_NAME : undefined;

    for (let s = 0; s < selectedSheets.length; s++) {
      const sheetConf = selectedSheets[s];
      const file = state.files[sheetConf.fileId];
      if (!file) continue;

      if (s > 0) doc.addPage(sheetConf.pageSize, sheetConf.orientation === 'landscape' ? 'l' : 'p');

      const sheet = file.workbook.Sheets[sheetConf.sheetName];
      const jsonData = XLSX.utils.sheet_to_json(sheet, { header: 1, defval: '' });
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
        headStyles: { font: fontName, fillColor: [99, 102, 241], textColor: [255, 255, 255], fontStyle: 'bold', halign: 'center' },
        alternateRowStyles: { fillColor: [245, 245, 255] },
        margin: { top: 10, right: 10, bottom: 10, left: 10 },
        didDrawPage: (data) => {
          const pageCount = doc.internal.getNumberOfPages();
          doc.setFontSize(8);
          doc.setTextColor(150, 150, 150);
          if (fontAvailable) doc.setFont(fontName);
          const pw = doc.internal.pageSize.getWidth();
          const ph = doc.internal.pageSize.getHeight();
          doc.text(`${data.pageNumber} / ${pageCount}`, pw / 2, ph - 6, { align: 'center' });
        },
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
      return;
    }
    previewInfo.textContent = JapaneseFont.isLoaded() ? '生成中...' : 'フォント読み込み中...';

    setTimeout(async () => {
      try {
        const doc = await generatePDF();
        if (!doc) { previewInfo.textContent = 'シートが選択されていません'; return; }
        const blob = doc.output('blob');
        revokePdfUrl();
        state.pdfBlobUrl = URL.createObjectURL(blob);
        pdfPreview.src = state.pdfBlobUrl;
        const pageCount = doc.internal.getNumberOfPages();
        previewInfo.textContent = `${selectedCount} シート・${pageCount} ページ`;
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
      const doc = await generatePDF();
      if (!doc) { Toast.show('シートが選択されていません', 'error'); return; }
      const fileIds = Object.keys(state.files);
      let pdfFileName;
      if (fileIds.length === 1) {
        pdfFileName = Utils.getBaseName(state.files[fileIds[0]].name) + '.pdf';
      } else {
        pdfFileName = 'combined_output.pdf';
      }
      doc.save(pdfFileName);
      Toast.show('PDFのダウンロードが完了しました！', 'success');
    } catch (err) {
      Toast.show('PDF変換中にエラーが発生しました', 'error');
      console.error(err);
    } finally {
      Loading.hide();
    }
  }

  // ── Init ──
  function init() {
    DropZone.init({
      elementId: 'dropZone-excel-pdf',
      inputId: 'fileInput-excel-pdf',
      acceptExtensions: ['.xlsx', '.xls', '.csv'],
      multiple: true,
      onFiles: handleFiles,
    });

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
