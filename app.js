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

  getDateStr() {
    const d = new Date();
    const yyyy = d.getFullYear();
    const mm = String(d.getMonth() + 1).padStart(2, '0');
    const dd = String(d.getDate()).padStart(2, '0');
    return `${yyyy}${mm}${dd}`;
  },

  buildDownloadName(baseNames, ext) {
    const date = this.getDateStr();
    const name = Array.isArray(baseNames) ? baseNames.join('_') : baseNames;
    return `${name}_${date}.${ext}`;
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
// Font Cache (IndexedDB)
// ==========================================
const FontCache = {
  DB_NAME: 'fileConverterFontCache',
  STORE_NAME: 'fonts',
  DB_VERSION: 1,

  _openDB() {
    return new Promise((resolve, reject) => {
      const request = indexedDB.open(this.DB_NAME, this.DB_VERSION);
      request.onupgradeneeded = (e) => {
        const db = e.target.result;
        if (!db.objectStoreNames.contains(this.STORE_NAME)) {
          db.createObjectStore(this.STORE_NAME);
        }
      };
      request.onsuccess = () => resolve(request.result);
      request.onerror = () => reject(request.error);
    });
  },

  async get(key) {
    try {
      const db = await this._openDB();
      return new Promise((resolve) => {
        const tx = db.transaction(this.STORE_NAME, 'readonly');
        const store = tx.objectStore(this.STORE_NAME);
        const req = store.get(key);
        req.onsuccess = () => resolve(req.result || null);
        req.onerror = () => resolve(null);
      });
    } catch {
      return null;
    }
  },

  async set(key, value) {
    try {
      const db = await this._openDB();
      return new Promise((resolve) => {
        const tx = db.transaction(this.STORE_NAME, 'readwrite');
        const store = tx.objectStore(this.STORE_NAME);
        store.put(value, key);
        tx.oncomplete = () => resolve(true);
        tx.onerror = () => resolve(false);
      });
    } catch {
      return false;
    }
  },
};

// ==========================================
// Japanese Font Loader (Noto Sans JP)
// ==========================================
const JapaneseFont = {
  _data: null,
  _loadPromise: null,
  FONT_NAME: 'NotoSansJP',
  FONT_FILE: 'NotoSansJP.ttf',
  CACHE_KEY: 'NotoSansJP-v1',
  FONT_URL: 'https://cdn.jsdelivr.net/gh/google/fonts@main/ofl/notosansjp/NotoSansJP%5Bwght%5D.ttf',

  isLoaded() {
    return this._data !== null;
  },

  async load() {
    if (this._data) return true;
    if (this._loadPromise) return this._loadPromise;
    this._loadPromise = this._doLoad();
    const result = await this._loadPromise;
    this._loadPromise = null;
    return result;
  },

  async _doLoad() {
    // Try IndexedDB cache first
    try {
      const cached = await FontCache.get(this.CACHE_KEY);
      if (cached) {
        this._data = cached;
        return true;
      }
    } catch (e) {
      console.warn('Font cache read failed:', e);
    }

    // Download from CDN
    try {
      const response = await fetch(this.FONT_URL);
      if (!response.ok) throw new Error(`HTTP ${response.status}`);
      const arrayBuffer = await response.arrayBuffer();
      const base64 = this._arrayBufferToBase64(arrayBuffer);
      this._data = base64;
      // Cache for next time (fire and forget)
      FontCache.set(this.CACHE_KEY, base64).catch(() => {});
      return true;
    } catch (e) {
      console.warn('Japanese font download failed:', e);
      return false;
    }
  },

  _arrayBufferToBase64(buffer) {
    const bytes = new Uint8Array(buffer);
    const chunks = [];
    const chunkSize = 8192;
    for (let i = 0; i < bytes.length; i += chunkSize) {
      const chunk = bytes.subarray(i, Math.min(i + chunkSize, bytes.length));
      chunks.push(String.fromCharCode.apply(null, chunk));
    }
    return btoa(chunks.join(''));
  },

  register(doc) {
    if (!this._data) return false;
    doc.addFileToVFS(this.FONT_FILE, this._data);
    doc.addFont(this.FONT_FILE, this.FONT_NAME, 'normal');
    doc.addFont(this.FONT_FILE, this.FONT_NAME, 'bold');
    doc.setFont(this.FONT_NAME);
    return true;
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
  // 日本語フォントをバックグラウンドでプリロード
  JapaneseFont.load().catch(() => {});
});
