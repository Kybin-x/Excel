/* ============================================
   Excel 快捷键学习页面 - 公用脚本
   ============================================ */

(function () {
  'use strict';

  /* ========== 1. 主题切换 ========== */
  const ThemeManager = {
    STORAGE_KEY: 'excel-shortcut-theme',

    init() {
      const saved = localStorage.getItem(this.STORAGE_KEY);
      if (saved) {
        document.documentElement.setAttribute('data-theme', saved);
      } else {
        // 跟随系统偏好
        const prefersDark = window.matchMedia('(prefers-color-scheme: dark)').matches;
        document.documentElement.setAttribute('data-theme', prefersDark ? 'dark' : 'light');
      }
      this.updateIcon();
    },

    toggle() {
      const current = document.documentElement.getAttribute('data-theme');
      const next = current === 'dark' ? 'light' : 'dark';
      document.documentElement.setAttribute('data-theme', next);
      localStorage.setItem(this.STORAGE_KEY, next);
      this.updateIcon();
    },

    updateIcon() {
      const btn = document.getElementById('theme-toggle-btn');
      if (!btn) return;
      const isDark = document.documentElement.getAttribute('data-theme') === 'dark';
      btn.innerHTML = isDark ? '☀️' : '🌙';
      btn.setAttribute('data-tooltip', isDark ? '切换到浅色模式' : '切换到深色模式');
    }
  };

  /* ========== 2. 收藏管理 ========== */
  const FavoritesManager = {
    STORAGE_KEY: 'excel-shortcut-favorites',

    getAll() {
      try {
        return JSON.parse(localStorage.getItem(this.STORAGE_KEY)) || [];
      } catch {
        return [];
      }
    },

    toggle(id) {
      const favs = this.getAll();
      const index = favs.indexOf(id);
      if (index > -1) {
        favs.splice(index, 1);
      } else {
        favs.push(id);
      }
      localStorage.setItem(this.STORAGE_KEY, JSON.stringify(favs));
      return index === -1; // 返回是否为收藏状态
    },

    isFavorite(id) {
      return this.getAll().indexOf(id) > -1;
    },

    count() {
      return this.getAll().length;
    }
  };

  /* ========== 3. 搜索筛选 ========== */
  const SearchManager = {
    init() {
      const input = document.getElementById('search-input');
      if (!input) return;

      let debounceTimer;
      input.addEventListener('input', () => {
        clearTimeout(debounceTimer);
        debounceTimer = setTimeout(() => {
          this.filter(input.value.trim().toLowerCase());
        }, 200);
      });

      // 支持 Ctrl+F / Cmd+F 聚焦搜索框
      document.addEventListener('keydown', (e) => {
        if ((e.ctrlKey || e.metaKey) && e.key === 'f') {
          // 仅当搜索框存在时拦截
          if (input) {
            e.preventDefault();
            input.focus();
            input.select();
          }
        }
        // ESC 清空搜索
        if (e.key === 'Escape' && document.activeElement === input) {
          input.value = '';
          this.filter('');
          input.blur();
        }
      });
    },

    filter(query) {
      const cards = document.querySelectorAll('.shortcut-card');
      const noResults = document.querySelector('.no-results');
      let visibleCount = 0;

      cards.forEach(card => {
        const keys = (card.getAttribute('data-keys') || '').toLowerCase();
        const desc = (card.getAttribute('data-desc') || '').toLowerCase();
        const sub = (card.getAttribute('data-subcategory') || '').toLowerCase();

        const match = !query || keys.includes(query) || desc.includes(query) || sub.includes(query);
        card.classList.toggle('hidden', !match);
        if (match) visibleCount++;
      });

      if (noResults) {
        noResults.classList.toggle('visible', visibleCount === 0 && query.length > 0);
      }

      // 更新计数
      const countEl = document.getElementById('visible-count');
      if (countEl) {
        countEl.textContent = visibleCount;
      }
    }
  };

  /* ========== 4. 子分类筛选 ========== */
  const SubFilterManager = {
    init() {
      const btns = document.querySelectorAll('.sub-filter-btn');
      if (!btns.length) return;

      btns.forEach(btn => {
        btn.addEventListener('click', () => {
          // 切换 active 状态
          btns.forEach(b => b.classList.remove('active'));
          btn.classList.add('active');

          const category = btn.getAttribute('data-filter');
          this.filter(category);
        });
      });
    },

    filter(category) {
      const cards = document.querySelectorAll('.shortcut-card');
      const noResults = document.querySelector('.no-results');
      let visibleCount = 0;

      cards.forEach(card => {
        const sub = card.getAttribute('data-subcategory') || '';
        const match = !category || category === 'all' || sub === category;
        card.classList.toggle('hidden', !match);
        if (match) visibleCount++;
      });

      if (noResults) {
        noResults.classList.toggle('visible', visibleCount === 0);
      }

      const countEl = document.getElementById('visible-count');
      if (countEl) countEl.textContent = visibleCount;

      // 清空搜索
      const input = document.getElementById('search-input');
      if (input) input.value = '';
    }
  };

  /* ========== 5. 虚拟键盘高亮 ========== */
  const KeyboardManager = {
    keyMap: {}, // DOM 元素映射

    init() {
      const container = document.querySelector('.keyboard-container');
      if (!container) return;

      // 构建键位映射
      container.querySelectorAll('.kb-key').forEach(el => {
        const keyName = el.getAttribute('data-key');
        if (keyName) {
          this.keyMap[keyName.toLowerCase()] = el;
        }
      });

      // 监听卡片悬停 / 点击
      document.querySelectorAll('.shortcut-card').forEach(card => {
        card.addEventListener('mouseenter', () => this.highlightFromCard(card));
        card.addEventListener('mouseleave', () => this.clearAll());
        card.addEventListener('focus', () => this.highlightFromCard(card));
        card.addEventListener('blur', () => this.clearAll());
      });
    },

    highlightFromCard(card) {
      this.clearAll();
      const keys = card.getAttribute('data-keys');
      if (!keys) return;

      // 解析按键：如 "Ctrl+Shift+Home"
      const parts = keys.split('+').map(k => k.trim().toLowerCase());
      parts.forEach(key => {
        // 处理别名
        const aliases = this.getAliases(key);
        aliases.forEach(alias => {
          if (this.keyMap[alias]) {
            this.keyMap[alias].classList.add('highlight');
          }
        });
      });
    },

    getAliases(key) {
      const map = {
        'ctrl': ['ctrl', 'control'],
        'control': ['ctrl', 'control'],
        'shift': ['shift'],
        'alt': ['alt'],
        'enter': ['enter', 'return'],
        'return': ['enter', 'return'],
        'esc': ['esc', 'escape'],
        'escape': ['esc', 'escape'],
        'del': ['del', 'delete'],
        'delete': ['del', 'delete'],
        'backspace': ['backspace'],
        'tab': ['tab'],
        'space': ['space'],
        '空格': ['space'],
        'home': ['home'],
        'end': ['end'],
        'pageup': ['pageup'],
        'pagedown': ['pagedown'],
        'pgup': ['pageup'],
        'pgdn': ['pagedown'],
        '向上键': ['arrowup', 'up', '↑'],
        '向下键': ['arrowdown', 'down', '↓'],
        '向左键': ['arrowleft', 'left', '←'],
        '向右键': ['arrowright', 'right', '→'],
        '↑': ['arrowup', 'up', '↑'],
        '↓': ['arrowdown', 'down', '↓'],
        '←': ['arrowleft', 'left', '←'],
        '→': ['arrowright', 'right', '→'],
        'up': ['arrowup', 'up', '↑'],
        'down': ['arrowdown', 'down', '↓'],
        'left': ['arrowleft', 'left', '←'],
        'right': ['arrowright', 'right', '→'],
        'f1': ['f1'], 'f2': ['f2'], 'f3': ['f3'], 'f4': ['f4'],
        'f5': ['f5'], 'f6': ['f6'], 'f7': ['f7'], 'f8': ['f8'],
        'f9': ['f9'], 'f10': ['f10'], 'f11': ['f11'], 'f12': ['f12'],
        'scrolllock': ['scrolllock'],
        'capslock': ['capslock'],
        'numlock': ['numlock'],
        'insert': ['insert'],
        'pause': ['pause'],
        'printscreen': ['printscreen'],
      };
      return map[key] || [key];
    },

    clearAll() {
      Object.values(this.keyMap).forEach(el => el.classList.remove('highlight'));
    },

    highlightKeys(keyArray) {
      this.clearAll();
      keyArray.forEach(key => {
        const aliases = this.getAliases(key.toLowerCase());
        aliases.forEach(alias => {
          if (this.keyMap[alias]) {
            this.keyMap[alias].classList.add('highlight');
          }
        });
      });
    }
  };

  /* ========== 6. 测验模式 ========== */
  const QuizManager = {
    questions: [],
    currentIndex: 0,
    score: 0,
    total: 0,
    isActive: false,

    init() {
      // 从页面卡片中收集题目
      document.querySelectorAll('.shortcut-card').forEach(card => {
        const keys = card.getAttribute('data-keys');
        const desc = card.getAttribute('data-desc');
        if (keys && desc) {
          this.questions.push({ keys, desc });
        }
      });
    },

    start() {
      if (this.questions.length === 0) {
        alert('当前页面没有可用的测验题目。');
        return;
      }
      this.isActive = true;
      this.currentIndex = 0;
      this.score = 0;
      this.total = 0;
      this.shuffle();
      this.showOverlay();
      this.showQuestion();
    },

    shuffle() {
      // Fisher-Yates 洗牌
      for (let i = this.questions.length - 1; i > 0; i--) {
        const j = Math.floor(Math.random() * (i + 1));
        [this.questions[i], this.questions[j]] = [this.questions[j], this.questions[i]];
      }
    },

    showOverlay() {
      const overlay = document.getElementById('quiz-overlay');
      if (overlay) overlay.classList.add('active');
    },

    hideOverlay() {
      const overlay = document.getElementById('quiz-overlay');
      if (overlay) overlay.classList.remove('active');
      this.isActive = false;
    },

    showQuestion() {
      if (this.currentIndex >= this.questions.length) {
        this.showResult();
        return;
      }

      const q = this.questions[this.currentIndex];
      const descEl = document.getElementById('quiz-desc');
      const inputEl = document.getElementById('quiz-input');
      const feedbackEl = document.getElementById('quiz-feedback');
      const scoreEl = document.getElementById('quiz-score');
      const questionEl = document.getElementById('quiz-question');

      if (questionEl) questionEl.textContent = `第 ${this.currentIndex + 1} / ${this.questions.length} 题`;
      if (descEl) descEl.textContent = q.desc;
      if (inputEl) {
        inputEl.value = '';
        inputEl.focus();
        inputEl.disabled = false;
      }
      if (feedbackEl) {
        feedbackEl.className = 'quiz-feedback';
        feedbackEl.style.display = 'none';
      }
      if (scoreEl) scoreEl.textContent = `得分：${this.score} / ${this.total}`;

      // 更新按钮
      const submitBtn = document.getElementById('quiz-submit-btn');
      const nextBtn = document.getElementById('quiz-next-btn');
      if (submitBtn) submitBtn.style.display = '';
      if (nextBtn) nextBtn.style.display = 'none';
    },

    checkAnswer() {
      const q = this.questions[this.currentIndex];
      const inputEl = document.getElementById('quiz-input');
      const feedbackEl = document.getElementById('quiz-feedback');

      if (!inputEl || !feedbackEl) return;

      const userAnswer = inputEl.value.trim().toLowerCase().replace(/\s+/g, '');
      const correctAnswer = q.keys.toLowerCase().replace(/\s+/g, '');

      this.total++;

      // 宽松比较：忽略大小写和空格
      const isCorrect = userAnswer === correctAnswer ||
        userAnswer === correctAnswer.replace(/\+/g, '') ||
        this.normalizeAnswer(userAnswer) === this.normalizeAnswer(correctAnswer);

      if (isCorrect) {
        this.score++;
        feedbackEl.className = 'quiz-feedback correct';
        feedbackEl.textContent = '✅ 正确！';
        feedbackEl.style.display = 'block';
      } else {
        feedbackEl.className = 'quiz-feedback wrong';
        feedbackEl.textContent = `❌ 错误！正确答案是：${q.keys}`;
        feedbackEl.style.display = 'block';
      }

      // 键盘高亮正确答案
      const keyParts = q.keys.split('+').map(k => k.trim());
      KeyboardManager.highlightKeys(keyParts);

      inputEl.disabled = true;

      const submitBtn = document.getElementById('quiz-submit-btn');
      const nextBtn = document.getElementById('quiz-next-btn');
      if (submitBtn) submitBtn.style.display = 'none';
      if (nextBtn) {
        nextBtn.style.display = '';
        nextBtn.focus();
      }

      const scoreEl = document.getElementById('quiz-score');
      if (scoreEl) scoreEl.textContent = `得分：${this.score} / ${this.total}`;
    },

    normalizeAnswer(str) {
      return str.replace(/ctrl|control/gi, 'ctrl')
        .replace(/shift/gi, 'shift')
        .replace(/alt/gi, 'alt')
        .replace(/enter|return/gi, 'enter')
        .replace(/esc|escape/gi, 'esc')
        .replace(/delete|del/gi, 'del')
        .replace(/\s+/g, '')
        .replace(/\+/g, '+');
    },

    nextQuestion() {
      this.currentIndex++;
      KeyboardManager.clearAll();
      this.showQuestion();
    },

    showResult() {
      const descEl = document.getElementById('quiz-desc');
      const inputEl = document.getElementById('quiz-input');
      const feedbackEl = document.getElementById('quiz-feedback');
      const questionEl = document.getElementById('quiz-question');
      const submitBtn = document.getElementById('quiz-submit-btn');
      const nextBtn = document.getElementById('quiz-next-btn');

      if (questionEl) questionEl.textContent = '🎉 测验完成！';
      if (descEl) {
        const percent = Math.round((this.score / this.total) * 100);
        let emoji = '🏆';
        if (percent < 60) emoji = '💪';
        else if (percent < 80) emoji = '👍';
        descEl.textContent = `${emoji} 你答对了 ${this.score} / ${this.total} 题 (${percent}%)`;
      }
      if (inputEl) inputEl.style.display = 'none';
      if (feedbackEl) feedbackEl.style.display = 'none';
      if (submitBtn) submitBtn.style.display = 'none';
      if (nextBtn) nextBtn.style.display = 'none';

      KeyboardManager.clearAll();
    },

    stop() {
      this.hideOverlay();
      KeyboardManager.clearAll();
      // 重置 input 显示
      const inputEl = document.getElementById('quiz-input');
      if (inputEl) inputEl.style.display = '';
    }
  };

  /* ========== 7. 卡片收藏交互 ========== */
  function initCardStars() {
    document.querySelectorAll('.card-star').forEach(btn => {
      const cardId = btn.getAttribute('data-id');
      if (!cardId) return;

      // 初始化状态
      if (FavoritesManager.isFavorite(cardId)) {
        btn.classList.add('starred');
        btn.textContent = '★';
      } else {
        btn.textContent = '☆';
      }

      btn.addEventListener('click', (e) => {
        e.stopPropagation();
        const isNowFav = FavoritesManager.toggle(cardId);
        btn.classList.toggle('starred', isNowFav);
        btn.textContent = isNowFav ? '★' : '☆';
        btn.classList.add('pop');
        setTimeout(() => btn.classList.remove('pop'), 300);
      });
    });
  }

  /* ========== 8. 打印功能 ========== */
  function handlePrint() {
    window.print();
  }

  /* ========== 9. 虚拟键盘 HTML 生成 ========== */
  function generateKeyboard() {
    const container = document.querySelector('.keyboard-container');
    if (!container) return;

    const rows = [
      [
        { key: 'Esc', label: 'Esc' },
        { key: 'F1', label: 'F1' }, { key: 'F2', label: 'F2' },
        { key: 'F3', label: 'F3' }, { key: 'F4', label: 'F4' },
        { key: 'F5', label: 'F5' }, { key: 'F6', label: 'F6' },
        { key: 'F7', label: 'F7' }, { key: 'F8', label: 'F8' },
        { key: 'F9', label: 'F9' }, { key: 'F10', label: 'F10' },
        { key: 'F11', label: 'F11' }, { key: 'F12', label: 'F12' },
        { key: 'PrintScreen', label: 'PrtSc' },
        { key: 'ScrollLock', label: 'ScrLk' },
        { key: 'Pause', label: 'Pause' }
      ],
      [
        { key: '`', label: '`' },
        { key: '1', label: '1' }, { key: '2', label: '2' },
        { key: '3', label: '3' }, { key: '4', label: '4' },
        { key: '5', label: '5' }, { key: '6', label: '6' },
        { key: '7', label: '7' }, { key: '8', label: '8' },
        { key: '9', label: '9' }, { key: '0', label: '0' },
        { key: '-', label: '-' }, { key: '=', label: '=' },
        { key: 'Backspace', label: '⌫', className: 'w-backspace' }
      ],
      [
        { key: 'Tab', label: 'Tab', className: 'w-tab' },
        { key: 'Q', label: 'Q' }, { key: 'W', label: 'W' },
        { key: 'E', label: 'E' }, { key: 'R', label: 'R' },
        { key: 'T', label: 'T' }, { key: 'Y', label: 'Y' },
        { key: 'U', label: 'U' }, { key: 'I', label: 'I' },
        { key: 'O', label: 'O' }, { key: 'P', label: 'P' },
        { key: '[', label: '[' }, { key: ']', label: ']' },
        { key: '\\', label: '\\' }
      ],
      [
        { key: 'CapsLock', label: 'Caps', className: 'w-caps' },
        { key: 'A', label: 'A' }, { key: 'S', label: 'S' },
        { key: 'D', label: 'D' }, { key: 'F', label: 'F' },
        { key: 'G', label: 'G' }, { key: 'H', label: 'H' },
        { key: 'J', label: 'J' }, { key: 'K', label: 'K' },
        { key: 'L', label: 'L' }, { key: ';', label: ';' },
        { key: "'", label: "'" },
        { key: 'Enter', label: 'Enter', className: 'w-enter' }
      ],
      [
        { key: 'Shift', label: 'Shift', className: 'w-shift' },
        { key: 'Z', label: 'Z' }, { key: 'X', label: 'X' },
        { key: 'C', label: 'C' }, { key: 'V', label: 'V' },
        { key: 'B', label: 'B' }, { key: 'N', label: 'N' },
        { key: 'M', label: 'M' }, { key: ',', label: ',' },
        { key: '.', label: '.' }, { key: '/', label: '/' },
        { key: 'Shift', label: 'Shift', className: 'w-shift' }
      ],
      [
        { key: 'Ctrl', label: 'Ctrl', className: 'w-ctrl' },
        { key: 'Win', label: 'Win' },
        { key: 'Alt', label: 'Alt', className: 'w-alt' },
        { key: 'Space', label: '', className: 'w-space' },
        { key: 'Alt', label: 'Alt', className: 'w-alt' },
        { key: 'Win', label: 'Win' },
        { key: 'Menu', label: 'Menu' },
        { key: 'Ctrl', label: 'Ctrl', className: 'w-ctrl' }
      ]
    ];

    // 额外的导航键区域
    const navKeys = [
      { key: 'Insert', label: 'Ins' },
      { key: 'Home', label: 'Home' },
      { key: 'PageUp', label: 'PgUp' },
      { key: 'Delete', label: 'Del' },
      { key: 'End', label: 'End' },
      { key: 'PageDown', label: 'PgDn' }
    ];

    const arrowKeys = [
      { key: '↑', label: '↑' },
      { key: '←', label: '←' },
      { key: '↓', label: '↓' },
      { key: '→', label: '→' }
    ];

    let html = '';

    // 主键盘行
    rows.forEach(row => {
      html += '<div class="keyboard-row">';
      row.forEach(k => {
        const cls = k.className ? ` ${k.className}` : '';
        html += `<div class="kb-key${cls}" data-key="${k.key.toLowerCase()}">${k.label}</div>`;
      });
      html += '</div>';
    });

    // 导航键（另起一个小区域）
    html += '<div class="keyboard-row" style="margin-top:8px; gap:4px;">';
    navKeys.forEach(k => {
      html += `<div class="kb-key" data-key="${k.key.toLowerCase()}" style="min-width:48px;">${k.label}</div>`;
    });
    html += '</div>';

    // 方向键
    html += '<div class="keyboard-row" style="gap:4px;">';
    html += '<div style="width:52px;"></div>'; // 占位
    arrowKeys.forEach(k => {
      html += `<div class="kb-key" data-key="${k.key}" style="min-width:40px;">${k.label}</div>`;
    });
    html += '</div>';

    container.innerHTML = html;
  }

  /* ========== 10. 首页收藏展示 ========== */
  function renderFavorites() {
    const container = document.getElementById('favorites-list');
    if (!container) return;

    const favs = FavoritesManager.getAll();
    if (favs.length === 0) {
      container.innerHTML = '<div class="empty-state">⭐ 还没有收藏任何快捷键<br>浏览各分类页面，点击星标即可收藏</div>';
      return;
    }

    // 从所有页面数据中找到已收藏的项（需要全局数据支持）
    if (typeof window.ALL_SHORTCUTS === 'undefined') {
      container.innerHTML = '<div class="empty-state">⭐ 已收藏 ' + favs.length + ' 个快捷键<br>请前往各分类页面查看</div>';
      return;
    }

    let html = '<div class="cards-grid">';
    favs.forEach(id => {
      const item = window.ALL_SHORTCUTS.find(s => s.id === id);
      if (!item) return;
      html += buildCardHTML(item);
    });
    html += '</div>';
    container.innerHTML = html;

    // 重新绑定星标事件
    initCardStars();
  }

  /* ========== 11. 卡片 HTML 构建工具 ========== */
  function buildCardHTML(item) {
    const keyParts = item.keys.split('+').map(k => k.trim());
    const keysHtml = keyParts.map((k, i) => {
      const plus = i < keyParts.length - 1 ? '<span class="key-plus">+</span>' : '';
      return `<span class="key-tag">${escapeHtml(k)}</span>${plus}`;
    }).join('');

    const isFav = FavoritesManager.isFavorite(item.id);

    return `
      <div class="shortcut-card"
           data-keys="${escapeAttr(item.keys)}"
           data-desc="${escapeAttr(item.desc)}"
           data-subcategory="${escapeAttr(item.subcategory || '')}">
        <div class="card-header">
          <div class="card-keys">${keysHtml}</div>
          <button class="card-star ${isFav ? 'starred' : ''}" data-id="${escapeAttr(item.id)}" title="收藏">
            ${isFav ? '★' : '☆'}
          </button>
        </div>
        <div class="card-desc">${escapeHtml(item.desc)}</div>
        ${item.subcategory ? `<div class="card-subcategory">${escapeHtml(item.subcategory)}</div>` : ''}
      </div>
    `;
  }

  /* ========== 12. 工具函数 ========== */
  function escapeHtml(str) {
    const div = document.createElement('div');
    div.textContent = str;
    return div.innerHTML;
  }

  function escapeAttr(str) {
    return str.replace(/&/g, '&amp;')
      .replace(/"/g, '&quot;')
      .replace(/'/g, '&#39;')
      .replace(/</g, '&lt;')
      .replace(/>/g, '&gt;');
  }

  /* ========== 13. 页面数据渲染 ========== */
  function renderShortcuts() {
    const container = document.getElementById('shortcuts-container');
    if (!container || typeof window.PAGE_SHORTCUTS === 'undefined') return;

    let html = '';
    window.PAGE_SHORTCUTS.forEach(item => {
      html += buildCardHTML(item);
    });
    container.innerHTML = html;

    // 更新计数
    const countEl = document.getElementById('visible-count');
    const totalEl = document.getElementById('total-count');
    if (countEl) countEl.textContent = window.PAGE_SHORTCUTS.length;
    if (totalEl) totalEl.textContent = window.PAGE_SHORTCUTS.length;
  }

  /* ========== 14. 初始化入口 ========== */
  function init() {
    // 主题
    ThemeManager.init();

    // 主题切换按钮
    const themeBtn = document.getElementById('theme-toggle-btn');
    if (themeBtn) {
      themeBtn.addEventListener('click', () => ThemeManager.toggle());
    }

    // 打印按钮
    const printBtn = document.getElementById('print-btn');
    if (printBtn) {
      printBtn.addEventListener('click', handlePrint);
    }

    // 生成虚拟键盘
    generateKeyboard();

    // 渲染快捷键卡片
    renderShortcuts();

    // 初始化收藏星标
    initCardStars();

    // 搜索
    SearchManager.init();

    // 子分类筛选
    SubFilterManager.init();

    // 键盘高亮
    KeyboardManager.init();

    // 测验
    QuizManager.init();

    // 测验按钮
    const quizStartBtn = document.getElementById('quiz-start-btn');
    if (quizStartBtn) {
      quizStartBtn.addEventListener('click', () => QuizManager.start());
    }

    const quizCloseBtn = document.getElementById('quiz-close-btn');
    if (quizCloseBtn) {
      quizCloseBtn.addEventListener('click', () => QuizManager.stop());
    }

    const quizSubmitBtn = document.getElementById('quiz-submit-btn');
    if (quizSubmitBtn) {
      quizSubmitBtn.addEventListener('click', () => QuizManager.checkAnswer());
    }

    const quizNextBtn = document.getElementById('quiz-next-btn');
    if (quizNextBtn) {
      quizNextBtn.addEventListener('click', () => QuizManager.nextQuestion());
    }

    // 测验输入框 Enter 提交
    const quizInput = document.getElementById('quiz-input');
    if (quizInput) {
      quizInput.addEventListener('keydown', (e) => {
        if (e.key === 'Enter') {
          e.preventDefault();
          if (!quizInput.disabled) {
            QuizManager.checkAnswer();
          } else {
            QuizManager.nextQuestion();
          }
        }
      });
    }

    // 首页收藏展示
    renderFavorites();

    // 高亮当前分类标签
    highlightCurrentTab();
  }

  /* ========== 15. 高亮当前分类标签 ========== */
  function highlightCurrentTab() {
    const currentPage = window.location.pathname.split('/').pop() || 'index.html';
    document.querySelectorAll('.category-tab').forEach(tab => {
      const href = tab.getAttribute('href') || '';
      if (href === currentPage || (currentPage === '' && href === 'index.html')) {
        tab.classList.add('active');
      } else {
        tab.classList.remove('active');
      }
    });
  }

  /* ========== 暴露全局接口 ========== */
  window.ExcelShortcuts = {
    ThemeManager,
    FavoritesManager,
    SearchManager,
    SubFilterManager,
    KeyboardManager,
    QuizManager,
    buildCardHTML,
    renderFavorites,
    renderShortcuts,
    initCardStars
  };

  /* ========== DOM Ready ========== */
  if (document.readyState === 'loading') {
    document.addEventListener('DOMContentLoaded', init);
  } else {
    init();
  }

})();