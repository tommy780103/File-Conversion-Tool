/**
 * CSV ⇔ Excel コンバーター
 * CSV→Excel / Excel→CSV のサブモード切替
 */
const CsvExcelConverter = (() => {
  const panel = Utils.$('panel-csv-excel');

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
  // CSV → Excel
  // ==========================================
  const csvState = {
    fileName: null,
    rawData: null,     // ArrayBuffer
    parsedData: null,  // 2D array
    detectedEncoding: null,
    detectedDelimiter: null,
  };

  function handleCsvFile(files) {
    const file = files[0];
    if (!file) return;
    csvState.fileName = file.name;

    const reader = new FileReader();
    reader.onload = (e) => {
      csvState.rawData = e.target.result;
      csvState.detectedEncoding = detectEncoding(new Uint8Array(csvState.rawData));
      parseCsvAndShow();
    };
    reader.readAsArrayBuffer(file);
  }

  function detectEncoding(bytes) {
    // BOM check
    if (bytes[0] === 0xEF && bytes[1] === 0xBB && bytes[2] === 0xBF) return 'utf-8';
    if (bytes[0] === 0xFF && bytes[1] === 0xFE) return 'utf-16le';
    if (bytes[0] === 0xFE && bytes[1] === 0xFF) return 'utf-16be';

    // Shift_JIS heuristic: look for 0x80-0x9F or 0xE0-0xEF as lead bytes
    let sjisScore = 0;
    let eucScore = 0;
    for (let i = 0; i < Math.min(bytes.length, 4096); i++) {
      const b = bytes[i];
      if ((b >= 0x81 && b <= 0x9F) || (b >= 0xE0 && b <= 0xEF)) sjisScore++;
      if (b === 0x8E || (b >= 0xA1 && b <= 0xFE)) eucScore++;
    }
    if (sjisScore > 10 && sjisScore > eucScore) return 'shift_jis';
    if (eucScore > 10 && eucScore > sjisScore) return 'euc-jp';
    return 'utf-8';
  }

  function detectDelimiter(text) {
    const lines = text.split('\n').slice(0, 5).filter((l) => l.trim());
    const delimiters = [',', '\t', ';'];
    let best = ',';
    let bestScore = -1;

    for (const d of delimiters) {
      const counts = lines.map((l) => l.split(d).length - 1);
      const avg = counts.reduce((a, b) => a + b, 0) / counts.length;
      const consistent = counts.every((c) => c === counts[0]) && counts[0] > 0;
      const score = avg * (consistent ? 2 : 1);
      if (score > bestScore) { bestScore = score; best = d; }
    }
    return best;
  }

  function parseCsvAndShow() {
    const encodingSelect = Utils.$('encoding-csv-to-excel');
    const delimiterSelect = Utils.$('delimiter-csv-to-excel');
    const encoding = encodingSelect.value === 'auto' ? csvState.detectedEncoding : encodingSelect.value;

    let text;
    try {
      const decoder = new TextDecoder(encoding);
      text = decoder.decode(csvState.rawData);
    } catch {
      text = new TextDecoder('utf-8').decode(csvState.rawData);
    }

    const delimiter = delimiterSelect.value === 'auto' ? detectDelimiter(text) : delimiterSelect.value;
    csvState.detectedDelimiter = delimiter;

    // Parse CSV manually respecting quotes
    csvState.parsedData = parseCsvText(text, delimiter);

    Utils.$('options-csv-to-excel').classList.remove('hidden');
    Utils.$('previewSection-csv-to-excel').classList.remove('hidden');
    Utils.$('actionsPanel-csv-to-excel').classList.remove('hidden');

    renderCsvPreview();
  }

  function parseCsvText(text, delimiter) {
    const rows = [];
    let row = [];
    let cell = '';
    let inQuote = false;

    for (let i = 0; i < text.length; i++) {
      const ch = text[i];
      if (inQuote) {
        if (ch === '"') {
          if (i + 1 < text.length && text[i + 1] === '"') {
            cell += '"';
            i++;
          } else {
            inQuote = false;
          }
        } else {
          cell += ch;
        }
      } else {
        if (ch === '"') {
          inQuote = true;
        } else if (ch === delimiter) {
          row.push(cell);
          cell = '';
        } else if (ch === '\r') {
          // skip
        } else if (ch === '\n') {
          row.push(cell);
          cell = '';
          rows.push(row);
          row = [];
        } else {
          cell += ch;
        }
      }
    }
    if (cell || row.length > 0) {
      row.push(cell);
      rows.push(row);
    }
    return rows;
  }

  function renderCsvPreview() {
    const data = csvState.parsedData;
    if (!data || data.length === 0) return;

    const maxRows = 100;
    const previewData = data.slice(0, maxRows + 1);
    const info = Utils.$('previewInfo-csv-to-excel');
    info.textContent = `${data.length} 行（先頭 ${Math.min(data.length, maxRows + 1)} 行を表示）`;

    const table = Utils.$('dataTable-csv-to-excel');
    let html = '<thead><tr>';
    if (previewData[0]) {
      previewData[0].forEach((h, i) => { html += `<th>${Utils.escapeHtml(h || `列${i + 1}`)}</th>`; });
    }
    html += '</tr></thead><tbody>';
    for (let r = 1; r < previewData.length; r++) {
      html += '<tr>';
      const maxCols = previewData[0] ? previewData[0].length : 0;
      for (let c = 0; c < maxCols; c++) {
        html += `<td>${Utils.escapeHtml(previewData[r] && previewData[r][c] !== undefined ? previewData[r][c] : '')}</td>`;
      }
      html += '</tr>';
    }
    html += '</tbody>';
    table.innerHTML = html;
  }

  function downloadExcel() {
    if (!csvState.parsedData) return;
    try {
      const ws = XLSX.utils.aoa_to_sheet(csvState.parsedData);
      const wb = XLSX.utils.book_new();
      XLSX.utils.book_append_sheet(wb, ws, 'Sheet1');
      const outName = Utils.buildDownloadName(Utils.getBaseName(csvState.fileName || 'data'), 'xlsx');
      XLSX.writeFile(wb, outName);
      Toast.show('Excelファイルのダウンロードが完了しました！', 'success');
    } catch (err) {
      Toast.show('変換中にエラーが発生しました', 'error');
      console.error(err);
    }
  }

  // ==========================================
  // Excel → CSV
  // ==========================================
  const excelState = {
    fileName: null,
    workbook: null,
    currentSheet: null,
    parsedData: null,
  };

  function handleExcelFile(files) {
    const file = files[0];
    if (!file) return;
    excelState.fileName = file.name;

    const reader = new FileReader();
    reader.onload = (e) => {
      try {
        const data = new Uint8Array(e.target.result);
        excelState.workbook = XLSX.read(data, { type: 'array' });

        const sheetSelect = Utils.$('sheet-excel-to-csv');
        sheetSelect.innerHTML = '';
        excelState.workbook.SheetNames.forEach((name) => {
          const opt = document.createElement('option');
          opt.value = name;
          opt.textContent = name;
          sheetSelect.appendChild(opt);
        });

        excelState.currentSheet = excelState.workbook.SheetNames[0];
        parseExcelAndShow();
      } catch (err) {
        Toast.show('ファイルの読み込みに失敗しました', 'error');
        console.error(err);
      }
    };
    reader.readAsArrayBuffer(file);
  }

  function parseExcelAndShow() {
    const sheet = excelState.workbook.Sheets[excelState.currentSheet];
    excelState.parsedData = XLSX.utils.sheet_to_json(sheet, { header: 1, defval: '' });

    Utils.$('options-excel-to-csv').classList.remove('hidden');
    Utils.$('previewSection-excel-to-csv').classList.remove('hidden');
    Utils.$('actionsPanel-excel-to-csv').classList.remove('hidden');

    renderExcelPreview();
  }

  function renderExcelPreview() {
    const data = excelState.parsedData;
    if (!data || data.length === 0) return;

    const maxRows = 100;
    const previewData = data.slice(0, maxRows + 1);
    const info = Utils.$('previewInfo-excel-to-csv');
    info.textContent = `${data.length} 行（先頭 ${Math.min(data.length, maxRows + 1)} 行を表示）`;

    const table = Utils.$('dataTable-excel-to-csv');
    let html = '<thead><tr>';
    if (previewData[0]) {
      previewData[0].forEach((h, i) => { html += `<th>${Utils.escapeHtml(String(h || `列${i + 1}`))}</th>`; });
    }
    html += '</tr></thead><tbody>';
    for (let r = 1; r < previewData.length; r++) {
      html += '<tr>';
      const maxCols = previewData[0] ? previewData[0].length : 0;
      for (let c = 0; c < maxCols; c++) {
        html += `<td>${Utils.escapeHtml(String(previewData[r] && previewData[r][c] !== undefined ? previewData[r][c] : ''))}</td>`;
      }
      html += '</tr>';
    }
    html += '</tbody>';
    table.innerHTML = html;
  }

  function downloadCsv() {
    if (!excelState.workbook || !excelState.currentSheet) return;
    try {
      const sheet = excelState.workbook.Sheets[excelState.currentSheet];
      const delimiterSel = Utils.$('delimiter-excel-to-csv');
      const delimiter = delimiterSel.value;

      // Use XLSX to generate CSV with specified separator
      const csvText = XLSX.utils.sheet_to_csv(sheet, { FS: delimiter });

      const encodingSel = Utils.$('encoding-excel-to-csv');
      let blob;
      if (encodingSel.value === 'utf-8-bom') {
        const bom = new Uint8Array([0xEF, 0xBB, 0xBF]);
        blob = new Blob([bom, csvText], { type: 'text/csv;charset=utf-8' });
      } else {
        blob = new Blob([csvText], { type: 'text/csv;charset=utf-8' });
      }

      const url = URL.createObjectURL(blob);
      const a = document.createElement('a');
      a.href = url;
      a.download = Utils.buildDownloadName(Utils.getBaseName(excelState.fileName || 'data'), 'csv');
      document.body.appendChild(a);
      a.click();
      document.body.removeChild(a);
      URL.revokeObjectURL(url);

      Toast.show('CSVファイルのダウンロードが完了しました！', 'success');
    } catch (err) {
      Toast.show('変換中にエラーが発生しました', 'error');
      console.error(err);
    }
  }

  // ── Init ──
  function init() {
    initSubmode();

    // CSV → Excel
    DropZone.init({
      elementId: 'dropZone-csv-to-excel',
      inputId: 'fileInput-csv-to-excel',
      acceptExtensions: ['.csv', '.tsv', '.txt'],
      multiple: false,
      onFiles: handleCsvFile,
    });

    Utils.$('encoding-csv-to-excel').addEventListener('change', () => {
      if (csvState.rawData) parseCsvAndShow();
    });
    Utils.$('delimiter-csv-to-excel').addEventListener('change', () => {
      if (csvState.rawData) parseCsvAndShow();
    });
    Utils.$('convertBtn-csv-to-excel').addEventListener('click', downloadExcel);

    // Excel → CSV
    DropZone.init({
      elementId: 'dropZone-excel-to-csv',
      inputId: 'fileInput-excel-to-csv',
      acceptExtensions: ['.xlsx', '.xls'],
      multiple: false,
      onFiles: handleExcelFile,
    });

    Utils.$('sheet-excel-to-csv').addEventListener('change', (e) => {
      excelState.currentSheet = e.target.value;
      parseExcelAndShow();
    });
    Utils.$('convertBtn-excel-to-csv').addEventListener('click', downloadCsv);

    TabManager.register('csv-excel', {
      onActivate() {},
      onDeactivate() {},
    });
  }

  init();
  return { csvState, excelState };
})();
