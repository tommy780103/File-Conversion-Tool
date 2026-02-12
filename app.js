/**
 * ファイル変換ツール — 共通基盤
 * TabManager, DropZone, Toast, Loading, Utils
 */

// ==========================================
// Utils
// ==========================================
const Utils = {
  formatFileSize(bytes) {
    if (bytes < 1024) return bytes + ' B';
    if (bytes < 1024 * 1024) return (bytes / 1024).toFixed(1) + ' KB';
    return (bytes / (1024 * 1024)).toFixed(1) + ' MB';
  },

  debounce(fn, ms) {
    let timer = null;
    return (...args) => {
      if (timer) clearTimeout(timer);
      timer = setTimeout(() => fn(...args), ms);
    };
  },

  revokeBlobUrl(url) {
    if (url) URL.revokeObjectURL(url);
    return null;
  },

  getBaseName(filename) {
    return filename.replace(/\.[^.]+$/, '');
  },

  $(id) {
    return document.getElementById(id);
  },

  escapeHtml(str) {
    const map = { '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;' };
    return String(str).replace(/[&<>"']/g, (c) => map[c]);
  },
};

// ==========================================
// Toast
// ==========================================
const Toast = {
  _el: null,
  _timer: null,

  _getEl() {
    if (!this._el) this._el = Utils.$('toast');
    return this._el;
  },

  show(message, type = 'success') {
    const el = this._getEl();
    if (!el) return;

    if (this._timer) {
      clearTimeout(this._timer);
      this._timer = null;
    }

    el.className = 'toast';
    el.classList.add(`toast-${type}`);

    const iconMap = {
      success: `<svg width="24" height="24" viewBox="0 0 24 24" fill="none">
        <circle cx="12" cy="12" r="10" stroke="#10b981" stroke-width="2"/>
        <path d="M8 12l3 3 5-5" stroke="#10b981" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"/>
      </svg>`,
      error: `<svg width="24" height="24" viewBox="0 0 24 24" fill="none">
        <circle cx="12" cy="12" r="10" stroke="#ef4444" stroke-width="2"/>
        <path d="M15 9l-6 6M9 9l6 6" stroke="#ef4444" stroke-width="2" stroke-linecap="round"/>
      </svg>`,
      info: `<svg width="24" height="24" viewBox="0 0 24 24" fill="none">
        <circle cx="12" cy="12" r="10" stroke="#6366f1" stroke-width="2"/>
        <path d="M12 8v0M12 12v4" stroke="#6366f1" stroke-width="2" stroke-linecap="round"/>
      </svg>`,
    };

    el.innerHTML = `${iconMap[type] || iconMap.info}<span>${message}</span>`;
    el.classList.add('show');

    this._timer = setTimeout(() => {
      el.classList.remove('show');
      this._timer = null;
    }, 3000);
  },
};

// ==========================================
// Loading
// ==========================================
const Loading = {
  _el: null,
  _textEl: null,
  _progressEl: null,

  _getEls() {
    if (!this._el) {
      this._el = Utils.$('loadingOverlay');
      this._textEl = this._el ? this._el.querySelector('.loading-text') : null;
      this._progressEl = Utils.$('loadingProgress');
    }
  },

  show(text = '処理中...') {
    this._getEls();
    if (!this._el) return;
    if (this._textEl) this._textEl.textContent = text;
    if (this._progressEl) this._progressEl.style.width = '0%';
    this._el.classList.remove('hidden');
  },

  setProgress(percent) {
    this._getEls();
    if (this._progressEl) this._progressEl.style.width = percent + '%';
  },

  hide() {
    this._getEls();
    if (!this._el) return;
    this._el.classList.add('hidden');
  },
};

// ==========================================
// DropZone Factory
// ==========================================
const DropZone = {
  /**
   * @param {Object} opts
   * @param {string} opts.elementId  - ドロップゾーンのDOM要素ID
   * @param {string} opts.inputId    - file inputのID
   * @param {string[]} opts.acceptExtensions - 例: ['.jpg','.png']
   * @param {boolean} opts.multiple  - 複数ファイル許可
   * @param {Function} opts.onFiles  - ファイル受け取りコールバック (FileList) => void
   */
  init(opts) {
    const zone = Utils.$(opts.elementId);
    const input = Utils.$(opts.inputId);
    if (!zone || !input) return;

    input.multiple = opts.multiple !== false;

    zone.addEventListener('dragover', (e) => {
      e.preventDefault();
      zone.classList.add('drag-over');
    });

    zone.addEventListener('dragleave', () => {
      zone.classList.remove('drag-over');
    });

    zone.addEventListener('drop', (e) => {
      e.preventDefault();
      zone.classList.remove('drag-over');
      const files = _filterFiles(e.dataTransfer.files, opts.acceptExtensions);
      if (files.length > 0) {
        opts.onFiles(files);
      } else if (e.dataTransfer.files.length > 0) {
        Toast.show('対応していないファイル形式です', 'error');
      }
    });

    zone.addEventListener('click', (e) => {
      if (e.target.closest('.file-select-btn') || e.target === input) return;
      input.click();
    });

    input.addEventListener('change', () => {
      const files = _filterFiles(input.files, opts.acceptExtensions);
      if (files.length > 0) {
        opts.onFiles(files);
      }
      input.value = '';
    });

    function _filterFiles(fileList, extensions) {
      if (!extensions || extensions.length === 0) return Array.from(fileList);
      return Array.from(fileList).filter((f) => {
        const ext = '.' + f.name.split('.').pop().toLowerCase();
        return extensions.includes(ext);
      });
    }
  },
};

// ==========================================
// TabManager
// ==========================================
const TabManager = {
  _converters: {},
  _activeTab: null,

  register(name, converter) {
    this._converters[name] = converter;
  },

  init() {
    const tabBtns = document.querySelectorAll('.tab-btn');
    tabBtns.forEach((btn) => {
      btn.addEventListener('click', () => {
        this.switchTab(btn.dataset.tab);
      });
    });

    // 初期タブ
    const firstBtn = tabBtns[0];
    if (firstBtn) {
      this.switchTab(firstBtn.dataset.tab);
    }
  },

  switchTab(name) {
    if (this._activeTab === name) return;

    // 前タブの deactivate
    if (this._activeTab && this._converters[this._activeTab]) {
      const prev = this._converters[this._activeTab];
      if (typeof prev.onDeactivate === 'function') prev.onDeactivate();
    }

    // タブボタンの切り替え
    document.querySelectorAll('.tab-btn').forEach((btn) => {
      btn.classList.toggle('active', btn.dataset.tab === name);
    });

    // パネルの切り替え
    document.querySelectorAll('.tab-panel').forEach((panel) => {
      panel.classList.toggle('hidden', panel.id !== `panel-${name}`);
    });

    this._activeTab = name;

    // 新タブの activate
    if (this._converters[name]) {
      const cur = this._converters[name];
      if (typeof cur.onActivate === 'function') cur.onActivate();
    }
  },
};

// ==========================================
// App Init
// ==========================================
document.addEventListener('DOMContentLoaded', () => {
  // 各コンバーターの init（各ファイルで呼ばれる）後にタブを初期化
  // converterファイルが先に読み込まれるのでDOMContentLoadedで実行
  TabManager.init();
});
