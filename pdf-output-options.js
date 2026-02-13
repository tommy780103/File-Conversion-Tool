/**
 * PDF出力オプション — 共通モジュール
 * メタデータ、セキュリティ、色変換、画像品質の設定UIを動的生成
 */
const PdfOutputOptions = (() => {

  /**
   * オプションパネルを生成して指定コンテナの後に挿入
   * @param {Object} config
   * @param {HTMLElement} config.container - 挿入先の兄弟要素（この直後に挿入）
   * @param {string} config.suffix - IDサフィックス（タブ間のID衝突回避）
   * @param {Object} config.features - { metadata, security, colorMode, imageQuality }
   * @returns {{ getOptions: Function, show: Function, hide: Function }}
   */
  function create(config) {
    const { container, suffix, features = {} } = config;
    const feat = {
      metadata: features.metadata !== false,
      security: !!features.security,
      colorMode: !!features.colorMode,
      imageQuality: !!features.imageQuality,
    };

    const wrapper = document.createElement('section');
    wrapper.className = 'pdf-advanced-options hidden';
    wrapper.id = `pdfAdvancedOptions-${suffix}`;

    // トグルボタン
    const toggle = document.createElement('button');
    toggle.type = 'button';
    toggle.className = 'pdf-advanced-options-toggle';
    toggle.innerHTML = `
      <svg width="16" height="16" viewBox="0 0 16 16" fill="none">
        <path d="M8 2a6 6 0 100 12A6 6 0 008 2z" stroke="currentColor" stroke-width="1.5"/>
        <path d="M6 7l2 2 2-2" stroke="currentColor" stroke-width="1.5" stroke-linecap="round" stroke-linejoin="round"/>
      </svg>
      <span>PDF出力の詳細設定</span>
      <svg class="pdf-advanced-options-chevron" width="14" height="14" viewBox="0 0 14 14" fill="none">
        <path d="M4 5.5l3 3 3-3" stroke="currentColor" stroke-width="1.5" stroke-linecap="round" stroke-linejoin="round"/>
      </svg>
    `;
    wrapper.appendChild(toggle);

    // ボディ
    const body = document.createElement('div');
    body.className = 'pdf-advanced-options-body';

    // --- メタデータセクション ---
    if (feat.metadata) {
      body.appendChild(_createMetadataSection(suffix));
    }

    // --- セキュリティセクション ---
    if (feat.security) {
      body.appendChild(_createSecuritySection(suffix));
    }

    // --- 色変換セクション ---
    if (feat.colorMode) {
      body.appendChild(_createColorModeSection(suffix));
    }

    // --- 画像品質セクション ---
    if (feat.imageQuality) {
      body.appendChild(_createImageQualitySection(suffix));
    }

    wrapper.appendChild(body);

    // トグル動作
    toggle.addEventListener('click', () => {
      wrapper.classList.toggle('expanded');
    });

    // DOMに挿入
    container.insertAdjacentElement('afterend', wrapper);

    return {
      getOptions() {
        return _collectOptions(suffix, feat);
      },
      show() {
        wrapper.classList.remove('hidden');
      },
      hide() {
        wrapper.classList.add('hidden');
      },
    };
  }

  // ── セクション生成 ──

  function _createMetadataSection(suffix) {
    const section = document.createElement('div');
    section.className = 'pdf-opts-section';
    section.innerHTML = `
      <h4 class="pdf-opts-section-title">文書プロパティ</h4>
      <div class="options-grid">
        <div class="option-group">
          <label for="pdfTitle-${suffix}">タイトル</label>
          <input type="text" id="pdfTitle-${suffix}" class="text-input" placeholder="文書タイトル">
        </div>
        <div class="option-group">
          <label for="pdfAuthor-${suffix}">作成者</label>
          <input type="text" id="pdfAuthor-${suffix}" class="text-input" placeholder="作成者名">
        </div>
        <div class="option-group">
          <label for="pdfSubject-${suffix}">サブタイトル</label>
          <input type="text" id="pdfSubject-${suffix}" class="text-input" placeholder="サブタイトル">
        </div>
        <div class="option-group">
          <label for="pdfKeywords-${suffix}">キーワード</label>
          <input type="text" id="pdfKeywords-${suffix}" class="text-input" placeholder="カンマ区切り">
        </div>
      </div>
    `;
    return section;
  }

  function _createSecuritySection(suffix) {
    const section = document.createElement('div');
    section.className = 'pdf-opts-section';
    section.innerHTML = `
      <h4 class="pdf-opts-section-title">セキュリティ</h4>
      <div class="options-grid">
        <div class="option-group">
          <label for="pdfUserPassword-${suffix}">開くパスワード</label>
          <div class="password-wrapper">
            <input type="password" id="pdfUserPassword-${suffix}" class="text-input" placeholder="未設定">
            <button type="button" class="password-toggle-btn" data-target="pdfUserPassword-${suffix}" title="パスワードを表示">
              <svg width="16" height="16" viewBox="0 0 16 16" fill="none">
                <path d="M1 8s2.5-5 7-5 7 5 7 5-2.5 5-7 5-7-5-7-5z" stroke="currentColor" stroke-width="1.2"/>
                <circle cx="8" cy="8" r="2" stroke="currentColor" stroke-width="1.2"/>
              </svg>
            </button>
          </div>
        </div>
        <div class="option-group">
          <label for="pdfOwnerPassword-${suffix}">権限パスワード</label>
          <div class="password-wrapper">
            <input type="password" id="pdfOwnerPassword-${suffix}" class="text-input" placeholder="未設定">
            <button type="button" class="password-toggle-btn" data-target="pdfOwnerPassword-${suffix}" title="パスワードを表示">
              <svg width="16" height="16" viewBox="0 0 16 16" fill="none">
                <path d="M1 8s2.5-5 7-5 7 5 7 5-2.5 5-7 5-7-5-7-5z" stroke="currentColor" stroke-width="1.2"/>
                <circle cx="8" cy="8" r="2" stroke="currentColor" stroke-width="1.2"/>
              </svg>
            </button>
          </div>
        </div>
      </div>
      <div class="pdf-permissions">
        <label class="pdf-permission-label">
          <input type="checkbox" id="pdfAllowPrint-${suffix}" checked>
          <span>印刷を許可</span>
        </label>
        <label class="pdf-permission-label">
          <input type="checkbox" id="pdfAllowCopy-${suffix}" checked>
          <span>コピーを許可</span>
        </label>
      </div>
    `;

    // パスワード表示切替
    setTimeout(() => {
      section.querySelectorAll('.password-toggle-btn').forEach((btn) => {
        btn.addEventListener('click', () => {
          const input = document.getElementById(btn.dataset.target);
          if (!input) return;
          const isPassword = input.type === 'password';
          input.type = isPassword ? 'text' : 'password';
          btn.title = isPassword ? 'パスワードを隠す' : 'パスワードを表示';
          btn.classList.toggle('active', isPassword);
        });
      });
    }, 0);

    return section;
  }

  function _createColorModeSection(suffix) {
    const section = document.createElement('div');
    section.className = 'pdf-opts-section';
    section.innerHTML = `
      <h4 class="pdf-opts-section-title">色変換</h4>
      <div class="options-grid">
        <div class="option-group">
          <label for="pdfColorMode-${suffix}">カラーモード</label>
          <div class="select-wrapper">
            <select id="pdfColorMode-${suffix}">
              <option value="color" selected>カラー（そのまま）</option>
              <option value="grayscale">グレースケール</option>
              <option value="mono">モノクロ（白黒）</option>
            </select>
          </div>
        </div>
      </div>
    `;
    return section;
  }

  function _createImageQualitySection(suffix) {
    const section = document.createElement('div');
    section.className = 'pdf-opts-section';
    section.innerHTML = `
      <h4 class="pdf-opts-section-title">画像品質</h4>
      <div class="options-grid">
        <div class="option-group">
          <label for="pdfImageQuality-${suffix}">JPEG圧縮品質</label>
          <div class="select-wrapper">
            <select id="pdfImageQuality-${suffix}">
              <option value="0.5">低（ファイルサイズ小）</option>
              <option value="0.75" selected>中（バランス）</option>
              <option value="0.92">高（高画質）</option>
            </select>
          </div>
        </div>
      </div>
    `;
    return section;
  }

  // ── 値収集 ──

  function _collectOptions(suffix, feat) {
    const opts = {};

    if (feat.metadata) {
      opts.title = (document.getElementById(`pdfTitle-${suffix}`) || {}).value || '';
      opts.author = (document.getElementById(`pdfAuthor-${suffix}`) || {}).value || '';
      opts.subject = (document.getElementById(`pdfSubject-${suffix}`) || {}).value || '';
      opts.keywords = (document.getElementById(`pdfKeywords-${suffix}`) || {}).value || '';
    }

    if (feat.security) {
      opts.userPassword = (document.getElementById(`pdfUserPassword-${suffix}`) || {}).value || '';
      opts.ownerPassword = (document.getElementById(`pdfOwnerPassword-${suffix}`) || {}).value || '';
      opts.allowPrint = (document.getElementById(`pdfAllowPrint-${suffix}`) || {}).checked !== false;
      opts.allowCopy = (document.getElementById(`pdfAllowCopy-${suffix}`) || {}).checked !== false;
    }

    if (feat.colorMode) {
      opts.colorMode = (document.getElementById(`pdfColorMode-${suffix}`) || {}).value || 'color';
    }

    if (feat.imageQuality) {
      opts.imageQuality = parseFloat((document.getElementById(`pdfImageQuality-${suffix}`) || {}).value || '0.75');
    }

    return opts;
  }

  // ── ユーティリティ ──

  const utils = {
    /**
     * RGB配列をグレースケール/モノクロに変換
     * @param {number[]} rgb - [R, G, B] (0-255)
     * @param {string} mode - 'color' | 'grayscale' | 'mono'
     * @returns {number[]} [R, G, B]
     */
    convertColor(rgb, mode) {
      if (!rgb || mode === 'color') return rgb;
      const [r, g, b] = rgb;
      // 知覚輝度ベースのグレースケール
      const gray = Math.round(0.299 * r + 0.587 * g + 0.114 * b);
      if (mode === 'mono') {
        const bw = gray > 127 ? 255 : 0;
        return [bw, bw, bw];
      }
      // grayscale
      return [gray, gray, gray];
    },

    /**
     * Canvas経由で画像を色変換+JPEG品質調整
     * @param {string} dataUrl - 画像のdataURL
     * @param {string} colorMode - 'color' | 'grayscale' | 'mono'
     * @param {number} jpegQuality - 0.0-1.0
     * @param {number} width - 画像の幅
     * @param {number} height - 画像の高さ
     * @returns {Promise<string>} 変換済みdataURL
     */
    processImage(dataUrl, colorMode, jpegQuality, width, height) {
      return new Promise((resolve) => {
        // 変換不要なケース
        if (colorMode === 'color' && jpegQuality >= 0.9) {
          resolve(dataUrl);
          return;
        }

        const img = new Image();
        img.onload = () => {
          const w = width || img.naturalWidth;
          const h = height || img.naturalHeight;
          const canvas = document.createElement('canvas');
          canvas.width = w;
          canvas.height = h;
          const ctx = canvas.getContext('2d');

          // PNG透過→JPEG変換時に白背景で塗りつぶし
          ctx.fillStyle = '#ffffff';
          ctx.fillRect(0, 0, w, h);
          ctx.drawImage(img, 0, 0, w, h);

          // 色変換が必要な場合
          if (colorMode !== 'color') {
            const imageData = ctx.getImageData(0, 0, w, h);
            const data = imageData.data;
            for (let i = 0; i < data.length; i += 4) {
              const gray = Math.round(0.299 * data[i] + 0.587 * data[i + 1] + 0.114 * data[i + 2]);
              if (colorMode === 'mono') {
                const bw = gray > 127 ? 255 : 0;
                data[i] = data[i + 1] = data[i + 2] = bw;
              } else {
                data[i] = data[i + 1] = data[i + 2] = gray;
              }
            }
            ctx.putImageData(imageData, 0, 0);
          }

          resolve(canvas.toDataURL('image/jpeg', jpegQuality));
        };
        img.onerror = () => resolve(dataUrl);
        img.src = dataUrl;
      });
    },
  };

  return { create, utils };
})();
