/*============================================================common.js — Excel 快捷键学习站· 全局逻辑
   ============================================================ */
(function () {
  'use strict';

  const $ = (sel, ctx = document) => ctx.querySelector(sel);
  const $$ = (sel, ctx = document) => [...ctx.querySelectorAll(sel)];

  const shortcuts = window.PAGE_SHORTCUTS || [];

  /* --------------------------------------------------
     主题切换
  -------------------------------------------------- */
  function initTheme() {
    const saved = localStorage.getItem('excel-kb-theme');
    if (saved === 'dark') document.documentElement.setAttribute('data-theme', 'dark');
    const btn = $('#theme-toggle-btn');
    if (!btn) return;
    const update = () => {
      const isDark = document.documentElement.getAttribute('data-theme') === 'dark';
      btn.textContent = isDark ? '☀️' : '🌙';
    };
    update();
    btn.addEventListener('click', () => {
      const isDark = document.documentElement.getAttribute('data-theme') === 'dark';
      if (isDark) {
        document.documentElement.removeAttribute('data-theme');
        localStorage.setItem('excel-kb-theme', 'light');
      } else {
        document.documentElement.setAttribute('data-theme', 'dark');
        localStorage.setItem('excel-kb-theme', 'dark');
      }
      update();
    });
  }

  /* --------------------------------------------------
     打印
  -------------------------------------------------- */
  function initPrint() {
    const btn = $('#print-btn');
    if (btn) btn.addEventListener('click', () => window.print());
  }

  /* --------------------------------------------------
     渲染快捷键卡片
  -------------------------------------------------- */
  let selectedCard = null;

  function parseKeys(keyStr) {
    return keyStr
      .replace(/\+/g, ' + ')
      .split(' ')
      .filter(k => k && k !== '+')
      .map(k => k.trim());
  }

  function renderCards(list) {
    const container = $('#shortcuts-container');
    if (!container) return;
    container.innerHTML = '';

    list.forEach(item => {
      const card = document.createElement('div');
      card.className = 'shortcut-card';
      card.dataset.id = item.id;
      card.dataset.keys = item.keys;
      card.dataset.subcategory = item.subcategory || '';

      const keysArr = parseKeys(item.keys);
      const badgesHTML = keysArr.map(k => `<span class="key-badge">${k}</span>`).join('<span class="key-plus">+</span>');

      card.innerHTML = `
        <div class="card-keys">${badgesHTML}</div>
        <div class="card-desc">${item.desc}</div>
        <div class="card-sub">${item.subcategory || ''}</div>
      `;

      //悬浮临时高亮（无选中时）
      card.addEventListener('mouseenter', () => {
        if (!selectedCard) highlightKeys(item.keys);
      });
      card.addEventListener('mouseleave', () => {
        if (!selectedCard) clearHighlight();
      });

      // 点击选中/取消
      card.addEventListener('click', () => {
        if (selectedCard === card) {
          card.classList.remove('selected');
          selectedCard = null;
          clearHighlight();
        } else {
          if (selectedCard) selectedCard.classList.remove('selected');
          card.classList.add('selected');
          selectedCard = card;
          highlightKeys(item.keys);
        }
      });

      container.appendChild(card);
    });

    updateCount(list.length);
    selectedCard = null;
  }

  function updateCount(n) {
    const el = $('#visible-count');
    if (el) el.textContent = n;
  }

  /* --------------------------------------------------
     虚拟键盘 — 经典87键TKL布局
  -------------------------------------------------- */
  const keyboardLayout = [
    // Row 0: 功能键行
    {
      keys: [
        { label: 'Esc', id: 'Esc' },
        { id: '_sp0', spacer: true, width: 1 },
        { label: 'F1', id: 'F1' },
        { label: 'F2', id: 'F2' },
        { label: 'F3', id: 'F3' },
        { label: 'F4', id: 'F4' },
        { id: '_sp1', spacer: true, width: 0.5 },
        { label: 'F5', id: 'F5' },
        { label: 'F6', id: 'F6' },
        { label: 'F7', id: 'F7' },
        { label: 'F8', id: 'F8' },
        { id: '_sp2', spacer: true, width: 0.5 },
        { label: 'F9', id: 'F9' },
        { label: 'F10', id: 'F10' },
        { label: 'F11', id: 'F11' },
        { label: 'F12', id: 'F12' },
        { id: '_sp3', spacer: true, width: 0.25 },
        { label: 'PrtSc', id: 'PrtSc' },
        { label: 'ScrLk', id: 'ScrollLock' },
        { label: 'Pause', id: 'Pause' }
      ]
    },
    // Row 1: 数字行
    {
      keys: [
        { label: '`', id: '`' },
        { label: '1', id: '1' },
        { label: '2', id: '2' },
        { label: '3', id: '3' },
        { label: '4', id: '4' },
        { label: '5', id: '5' },
        { label: '6', id: '6' },
        { label: '7', id: '7' },
        { label: '8', id: '8' },
        { label: '9', id: '9' },
        { label: '0', id: '0' },
        { label: '-', id: '-' },
        { label: '=', id: '=' },
        { label: 'Backspace', id: 'Backspace', width: 2 },
        { id: '_sp4', spacer: true, width: 0.25 },
        { label: 'Ins', id: 'Insert' },
        { label: 'Home', id: 'Home' },
        { label: 'PgUp', id: 'PageUp' }
      ]
    },
    // Row 2: QWERTY行
    {
      keys: [
        { label: 'Tab', id: 'Tab', width: 1.5 },
        { label: 'Q', id: 'Q' },
        { label: 'W', id: 'W' },
        { label: 'E', id: 'E' },
        { label: 'R', id: 'R' },
        { label: 'T', id: 'T' },
        { label: 'Y', id: 'Y' },
        { label: 'U', id: 'U' },
        { label: 'I', id: 'I' },
        { label: 'O', id: 'O' },
        { label: 'P', id: 'P' },
        { label: '[', id: '[' },
        { label: ']', id: ']' },
        { label: '\\', id: '\\', width: 1.5 },
        { id: '_sp5', spacer: true, width: 0.25 },
        { label: 'Del', id: 'Del' },
        { label: 'End', id: 'End' },
        { label: 'PgDn', id: 'PageDown' }
      ]
    },
    // Row 3: HOME行
    {
      keys: [
        { label: 'CapsLk', id: 'CapsLock', width: 1.75 },
        { label: 'A', id: 'A' },
        { label: 'S', id: 'S' },
        { label: 'D', id: 'D' },
        { label: 'F', id: 'F' },
        { label: 'G', id: 'G' },
        { label: 'H', id: 'H' },
        { label: 'J', id: 'J' },
        { label: 'K', id: 'K' },
        { label: 'L', id: 'L' },
        { label: ';', id: ';' },
        { label: "'", id: "'" },
        { label: 'Enter', id: 'Enter', width: 2.25 }
      ]
    },
    // Row 4: Shift行
    {
      keys: [
        { label: 'Shift', id: 'LShift', width: 2.25 },
        { label: 'Z', id: 'Z' },
        { label: 'X', id: 'X' },
        { label: 'C', id: 'C' },
        { label: 'V', id: 'V' },
        { label: 'B', id: 'B' },
        { label: 'N', id: 'N' },
        { label: 'M', id: 'M' },
        { label: ',', id: ',' },
        { label: '.', id: '.' },
        { label: '/', id: '/' },
        { label: 'Shift', id: 'RShift', width: 2.75 },
        { id: '_sp6', spacer: true, width: 1.25 },
        { label: '↑', id: '↑' }
      ]
    },
    // Row 5: 底行
    {
      keys: [
        { label: 'Ctrl', id: 'LCtrl', width: 1.25 },
        { label: 'Win', id: 'LWin', width: 1.25 },
        { label: 'Alt', id: 'LAlt', width: 1.25 },
        { label: '', id: 'Space', width: 6.25 },
        { label: 'Alt', id: 'RAlt', width: 1.25 },
        { label: 'Win', id: 'RWin', width: 1.25 },
        { label: 'Menu', id: 'Menu', width: 1.25 },
        { label: 'Ctrl', id: 'RCtrl', width: 1.25 },
        { id: '_sp7', spacer: true, width: 0.25 },
        { label: '←', id: '←' },
        { label: '↓', id: '↓' },
        { label: '→', id: '→' }
      ]
    }
  ];

  /* --------------------------------------------------
     键名映射表（快捷键文本 → 键盘ID）
     
     核心修复：完整覆盖所有中文方向键写法
  -------------------------------------------------- */
  const keyAliasToId = {
    // 修饰键 → 左侧
    'ctrl': 'LCtrl', 'control': 'LCtrl',
    'shift': 'LShift',
    'alt': 'LAlt',
    'win': 'LWin', 'meta': 'LWin',

    // 功能键
    'esc': 'Esc', 'escape': 'Esc',
    'enter': 'Enter', 'return': 'Enter',
    'tab': 'Tab',
    'backspace': 'Backspace',
    'delete': 'Del', 'del': 'Del',
    'capslock': 'CapsLock',
    'space': 'Space', '空格': 'Space', 'spacebar': 'Space', '空格键': 'Space',

    // 导航键
    'pageup': 'PageUp', 'pgup': 'PageUp',
    'pagedown': 'PageDown', 'pgdn': 'PageDown',
    'home': 'Home',
    'end': 'End',
    'insert': 'Insert', 'ins': 'Insert',
    'scrolllock': 'ScrollLock', 'scrlk': 'ScrollLock',
    'prtsc': 'PrtSc', 'printscreen': 'PrtSc',
    'pause': 'Pause',

    // ★★★ 方向键：所有可能的写法 ★★★
    '向上键': '↑', '向下键': '↓', '向左键': '←', '向右键': '→',
    '上键': '↑', '下键': '↓', '左键': '←', '右键': '→',
    '上': '↑', '下': '↓', '左': '←', '右': '→',
    '↑': '↑', '↓': '↓', '←': '←', '→': '→',
    'up': '↑', 'down': '↓', 'left': '←', 'right': '→',
    'arrowup': '↑', 'arrowdown': '↓', 'arrowleft': '←', 'arrowright': '→',

    // F键
    'f1': 'F1', 'f2': 'F2', 'f3': 'F3', 'f4': 'F4',
    'f5': 'F5', 'f6': 'F6', 'f7': 'F7', 'f8': 'F8',
    'f9': 'F9', 'f10': 'F10', 'f11': 'F11', 'f12': 'F12',

    // 符号 → 对应键
    '连字符': '-', '句号': '.',
    '`': '`', '~': '`',
    '!': '1', '@': '2', '#': '3', '$': '4', '%': '5','^': '6', '&': '7', '*': '8', '(': '9', ')': '0',
    '_': '-', '+': '=',
    '{': '[', '}': ']', '|': '\\',
    ':': ';', '"': "'",
    '<': ',', '>': '.', '?': '/'
  };

  /**
   * 将快捷键文本中的单个键名解析为键盘上的键ID
   * 例如: "Ctrl" → "LCtrl", "向下键" → "↓", "A" → "A"
   */
  function resolveKeyId(name) {
    const trimmed = name.trim();
    if (!trimmed) return '';

    const lower = trimmed.toLowerCase();

    // 1. 精确匹配别名表（先用原文匹配，再用小写匹配）
    if (keyAliasToId[trimmed]) return keyAliasToId[trimmed];
    if (keyAliasToId[lower]) return keyAliasToId[lower];

    // 2. 单字母→ 大写
    if (lower.length === 1 && lower >= 'a' && lower <= 'z') return lower.toUpperCase();

    // 3. 单数字
    if (lower.length === 1 && lower >= '0' && lower <= '9') return lower;

    // 4. 原样返回
    return trimmed;
  }

  /**
   * ★ 核心改进：智能拆分组合键
   *
   * 问题: "Alt+向下键" 经过简单的.split('+') 会被拆成 ['Alt', '向下键'] — 这是正确的
   * 但 "Ctrl+Shift+向右键" 或 "向下键" 单独出现时也需要正确处理
   * 
   * 还需处理连续中文如 "箭头键" 这种泛指（不高亮具体方向键）
   */
  function splitComboKeys(keyStr) {
    // 用+ 号拆分，但需要避免把键名中间拆开
    // 策略：先把已知的多字符键名替换为占位符，拆分后再还原
    const parts = [];
    const segments = keyStr.split('+');

    for (let seg of segments) {
      seg = seg.trim();
      if (seg) parts.push(seg);
    }

    return parts;
  }

  const KEY_UNIT = 48;
  const KEY_GAP = 4;

  function buildKeyboard() {
    const container = $('.keyboard-container');
    if (!container) return;
    container.innerHTML = '';

    keyboardLayout.forEach(row => {
      const rowDiv = document.createElement('div');
      rowDiv.className = 'keyboard-row';

      row.keys.forEach(key => {
        if (key.spacer) {
          const spacer = document.createElement('div');
          spacer.className = 'keyboard-spacer';
          spacer.style.width = (key.width * KEY_UNIT + Math.max(0, key.width - 1) * KEY_GAP) + 'px';
          rowDiv.appendChild(spacer);
          return;
        }

        const keyEl = document.createElement('div');
        keyEl.className = 'keyboard-key';
        keyEl.dataset.keyId = key.id;
        keyEl.textContent = key.label;

        const w = key.width || 1;
        if (w !== 1) {
          keyEl.style.width = (w * KEY_UNIT + (w - 1) * KEY_GAP) + 'px';
        }

        if (key.id === 'Space') {
          keyEl.classList.add('space-key');
        }

        rowDiv.appendChild(keyEl);
      });

      container.appendChild(rowDiv);
    });
  }

  function highlightKeys(keyStr) {
    clearHighlight();

    const parts = splitComboKeys(keyStr);
    const resolvedIds = parts.map(p => resolveKeyId(p)).filter(Boolean);

    const allKeys = $$('.keyboard-key');
    const matched = new Set();

    resolvedIds.forEach(targetId => {
      //遍历所有键盘按键，找第一个匹配的（左侧优先）
      for (const el of allKeys) {
        const elId = el.dataset.keyId;
        if (elId === targetId && !matched.has(el)) {
          el.classList.add('active');
          matched.add(el);
          break;
        }
      }
    });
  }

  function clearHighlight() {
    $$('.keyboard-key.active').forEach(el => el.classList.remove('active'));
  }

  /* --------------------------------------------------
     搜索
  -------------------------------------------------- */
  function initSearch() {
    const input = $('#search-input');
    if (!input) return;
    input.addEventListener('input', () => {
      const q = input.value.trim().toLowerCase();
      filterAndRender(q, getCurrentSubFilter());
    });
  }

  function getCurrentSubFilter() {
    const activeBtn = $('.sub-filter-btn.active');
    return activeBtn ? activeBtn.dataset.filter : 'all';
  }

  function filterAndRender(query, sub) {
    let list = shortcuts;
    if (sub && sub !== 'all') {
      list = list.filter(s => s.subcategory === sub);
    }
    if (query) {
      list = list.filter(s =>
        s.keys.toLowerCase().includes(query) ||
        s.desc.toLowerCase().includes(query) ||
        (s.subcategory && s.subcategory.toLowerCase().includes(query))
      );
    }
    renderCards(list);const noRes = $('.no-results');
    if (noRes) noRes.style.display = list.length === 0 ? 'block' : 'none';
  }

  /* --------------------------------------------------
     子分类筛选
  -------------------------------------------------- */
  function initSubFilter() {
    const btns = $$('.sub-filter-btn');
    if (!btns.length) return;
    btns.forEach(btn => {
      btn.addEventListener('click', () => {
        btns.forEach(b => b.classList.remove('active'));
        btn.classList.add('active');
        const input = $('#search-input');
        const q = input ? input.value.trim().toLowerCase() : '';
        filterAndRender(q, btn.dataset.filter);
      });
    });
  }

  /* --------------------------------------------------
     测验
  -------------------------------------------------- */
  function initQuiz() {
    const startBtn = $('#quiz-start-btn');
    const overlay = $('#quiz-overlay');
    const closeBtn = $('#quiz-close-btn');
    const submitBtn = $('#quiz-submit-btn');
    const nextBtn = $('#quiz-next-btn');
    const questionEl = $('#quiz-question');
    const descEl = $('#quiz-desc');
    const inputEl = $('#quiz-input');
    const feedbackEl = $('#quiz-feedback');
    const scoreEl = $('#quiz-score');

    if (!startBtn || !overlay || shortcuts.length === 0) return;

    let pool = [], current = null, score = 0, total = 0;

    function shuffle(arr) {
      const a = [...arr];
      for (let i = a.length - 1; i > 0; i--) {
        const j = Math.floor(Math.random() * (i + 1));
        [a[i], a[j]] = [a[j], a[i]];
      }
      return a;
    }

    function nextQuestion() {
      if (pool.length === 0) pool = shuffle(shortcuts);
      current = pool.pop();
      questionEl.textContent = `功能：${current.desc}`;
      descEl.textContent = '请输入对应的快捷键';
      inputEl.value = '';
      feedbackEl.textContent = '';
      feedbackEl.className = 'quiz-feedback';
      submitBtn.style.display = '';
      nextBtn.style.display = 'none';
      inputEl.focus();
    }

    function checkAnswer() {
      if (!current) return;
      const userAns = inputEl.value.trim().toLowerCase().replace(/\s/g, '');
      const correctAns = current.keys.toLowerCase().replace(/\s/g, '');
      total++;
      if (userAns === correctAns) {
        score++;
        feedbackEl.textContent = `✅ 正确！答案：${current.keys}`;
        feedbackEl.className = 'quiz-feedback correct';
      } else {
        feedbackEl.textContent = `❌ 错误。正确答案：${current.keys}`;
        feedbackEl.className = 'quiz-feedback wrong';
      }
      scoreEl.textContent = `得分：${score} / ${total}`;
      submitBtn.style.display = 'none';
      nextBtn.style.display = '';}

    startBtn.addEventListener('click', () => {
      overlay.classList.add('active');
      score = 0; total = 0;
      scoreEl.textContent = '得分：0 / 0';
      pool = shuffle(shortcuts);
      nextQuestion();
    });

    closeBtn.addEventListener('click', () => overlay.classList.remove('active'));
    submitBtn.addEventListener('click', checkAnswer);
    nextBtn.addEventListener('click', nextQuestion);

    inputEl.addEventListener('keydown', (e) => {
      if (e.key === 'Enter') {
        if (submitBtn.style.display !== 'none') checkAnswer();
        else nextQuestion();
      }
    });
  }

  /* --------------------------------------------------
     调整DOM顺序：键盘移到卡片上方
  -------------------------------------------------- */
  function reorderLayout() {
    const main = $('main.main-content');
    const keyboard = $('.keyboard-section');
    const cardsGrid = $('#shortcuts-container');
    if (!main || !keyboard || !cardsGrid) return;
    main.insertBefore(keyboard, cardsGrid);
  }

  /* --------------------------------------------------
     点击空白取消选中
  -------------------------------------------------- */
  function initClickOutside() {
    document.addEventListener('click', (e) => {
      if (!e.target.closest('.shortcut-card') && !e.target.closest('.keyboard-section')) {
        if (selectedCard) {
          selectedCard.classList.remove('selected');
          selectedCard = null;
          clearHighlight();
        }
      }
    });
  }

  /* --------------------------------------------------
     初始化
  -------------------------------------------------- */
  function init() {
    initTheme();
    initPrint();

    if (shortcuts.length > 0) {
      buildKeyboard();
      reorderLayout();
      renderCards(shortcuts);
      initSearch();
      initSubFilter();
      initQuiz();
      initClickOutside();
    }
  }

  if (document.readyState === 'loading') {
    document.addEventListener('DOMContentLoaded', init);
  } else {
    init();
  }

})();
