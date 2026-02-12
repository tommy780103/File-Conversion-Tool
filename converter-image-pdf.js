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
  };

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
          Toast.show(`${file.name} の読み込みに失敗しました`, 'error');
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

    S('dropZone').classList.remove('compact');
    S('imageList').classList.add('hidden');
    S('defaultSettings').classList.add('hidden');
    S('previewSection').classList.add('hidden');
    S('actionsPanel').classList.add('hidden');
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
        <img class="image-card-thumbnail" src="${img.thumbnailUrl}" alt="${img.name}" loading="lazy">
        <div class="image-card-info">
          <span class="image-card-name" title="${img.name}">${img.name}</span>
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

  function generatePDF() {
    if (state.images.length === 0) return null;
    const { jsPDF } = window.jspdf;
    const settings = getSettings();

    // ページサイズ定義 (mm)
    const pageSizes = {
      a4: [210, 297], a3: [297, 420], letter: [215.9, 279.4], legal: [215.9, 355.6],
    };

    let doc = null;

    state.images.forEach((img, idx) => {
      const isLandscape = settings.orientation === 'landscape' ||
        (settings.orientation === 'auto' && img.width > img.height);
      const orient = isLandscape ? 'l' : 'p';

      if (idx === 0) {
        doc = new jsPDF({ orientation: orient, unit: 'mm', format: settings.pageSize });
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

      const format = _getImageFormat(img.name);
      doc.addImage(img.dataUrl, format, x, y, drawW, drawH);
    });

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

  function updatePreview() {
    if (state.images.length === 0) return;
    const previewInfo = S('previewInfo');
    const pdfPreview = S('pdfPreview');
    previewInfo.textContent = '生成中...';

    setTimeout(() => {
      try {
        const doc = generatePDF();
        if (!doc) return;
        const blob = doc.output('blob');
        state.pdfBlobUrl = Utils.revokeBlobUrl(state.pdfBlobUrl);
        state.pdfBlobUrl = URL.createObjectURL(blob);
        pdfPreview.src = state.pdfBlobUrl;
        previewInfo.textContent = `${state.images.length} 画像・${doc.internal.getNumberOfPages()} ページ`;
      } catch (err) {
        previewInfo.textContent = 'プレビュー生成に失敗しました';
        console.error(err);
      }
    }, 50);
  }

  // ── Download ──
  function download() {
    if (state.images.length === 0) return;
    try {
      const doc = generatePDF();
      if (!doc) return;
      doc.save('images.pdf');
      Toast.show('PDFのダウンロードが完了しました！', 'success');
    } catch (err) {
      Toast.show('PDF変換中にエラーが発生しました', 'error');
      console.error(err);
    }
  }

  // ── Init ──
  function init() {
    DropZone.init({
      elementId: 'dropZone-image-pdf',
      inputId: 'fileInput-image-pdf',
      acceptExtensions: ['.jpg', '.jpeg', '.png', '.webp', '.gif', '.bmp'],
      multiple: true,
      onFiles: handleFiles,
    });

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
