// ============================================================
//  Notion Project Template — script.js
//  Handles: date injection, open buttons, add subpage,
//  stat counter, editable fields, keyboard shortcuts
// ============================================================

(function () {
  'use strict';

  // ── Utilities ─────────────────────────────────────────────

  function formatDate(date) {
    return date.toLocaleDateString('en-US', {
      year: 'numeric', month: 'short', day: 'numeric'
    });
  }

  function getSubpageCount() {
    return document.querySelectorAll('#subpage-list .subpage-item').length;
  }

  function updateStatPages() {
    var el = document.getElementById('stat-pages');
    if (el) el.textContent = getSubpageCount();
  }

  // ── Init dates ────────────────────────────────────────────

  function initDates() {
    var now = new Date();
    var createdEl  = document.getElementById('created-date');
    var moveDateEl = document.getElementById('move-date');
    if (createdEl)  createdEl.textContent  = formatDate(now);
    if (moveDateEl) moveDateEl.textContent = formatDate(now);
  }

  // ── Open page button ──────────────────────────────────────

  function handleOpenBtn(btn) {
    btn.addEventListener('click', function (e) {
      e.stopPropagation(); // prevent row click from firing

      var item = btn.closest('.subpage-item');
      var url  = item && item.dataset.url;

      if (url) {
        window.open(url, '_blank', 'noopener');
      } else {
        // Visual feedback — flash success state
        var original = btn.innerHTML;
        btn.classList.add('success');
        btn.textContent = '✓';
        setTimeout(function () {
          btn.classList.remove('success');
          btn.innerHTML = original;
        }, 600);
      }
    });
  }

  function handleRowClick(item) {
    item.addEventListener('click', function () {
      var url = item.dataset.url;
      if (url) window.open(url, '_blank', 'noopener');
    });
  }

  // ── Wire all existing buttons ─────────────────────────────

  function wireButtons() {
    document.querySelectorAll('.open-btn').forEach(handleOpenBtn);
    document.querySelectorAll('.subpage-item').forEach(handleRowClick);
  }

  // ── Add sub-page ──────────────────────────────────────────

  function buildSubpageHTML(date) {
    return [
      '<div class="subpage-icon">📄</div>',
      '<div class="subpage-info">',
        '<div class="subpage-name" contenteditable="true" spellcheck="false">Untitled</div>',
        '<div class="subpage-meta">',
          '<span>' + formatDate(date) + '</span>',
        '</div>',
      '</div>',
      '<button class="open-btn">',
        'Open ',
        '<svg viewBox="0 0 10 10" fill="none" stroke="currentColor" stroke-width="1.5">',
          '<path d="M1 9L9 1M9 1H4M9 1V6"/>',
        '</svg>',
      '</button>',
    ].join('');
  }

  function addSubpage() {
    var list = document.getElementById('subpage-list');
    if (!list) return;

    var item = document.createElement('div');
    item.className = 'subpage-item new';
    item.innerHTML = buildSubpageHTML(new Date());

    list.appendChild(item);

    // Wire new button and row
    var btn = item.querySelector('.open-btn');
    if (btn) handleOpenBtn(btn);
    handleRowClick(item);

    // Focus and select name
    var nameEl = item.querySelector('.subpage-name');
    if (nameEl) {
      nameEl.focus();
      var range = document.createRange();
      range.selectNodeContents(nameEl);
      var sel = window.getSelection();
      sel.removeAllRanges();
      sel.addRange(range);
    }

    updateStatPages();
  }

  // Expose for onclick in HTML
  window.addSubpage = addSubpage;

  // ── Open page (legacy onclick support) ────────────────────

  window.openPage = function (item) {
    var url = item && item.dataset.url;
    if (url) {
      window.open(url, '_blank', 'noopener');
    } else {
      var btn = item && item.querySelector('.open-btn');
      if (btn) {
        var original = btn.innerHTML;
        btn.classList.add('success');
        btn.textContent = '✓';
        setTimeout(function () {
          btn.classList.remove('success');
          btn.innerHTML = original;
        }, 300);
      }
    }
  };

  // ── Keyboard shortcuts ────────────────────────────────────

  function initKeyboardShortcuts() {
    document.addEventListener('keydown', function (e) {
      // Cmd/Ctrl + Enter inside a subpage name → blur and confirm
      if ((e.metaKey || e.ctrlKey) && e.key === 'Enter') {
        var active = document.activeElement;
        if (active && active.classList.contains('subpage-name')) {
          active.blur();
        }
      }

      // Escape → blur any focused editable
      if (e.key === 'Escape') {
        var focused = document.activeElement;
        if (focused && focused.isContentEditable) focused.blur();
        if (focused && focused.tagName === 'INPUT') focused.blur();
      }
    });
  }

  // ── Editable title placeholder behaviour ──────────────────

  function initEditableTitle() {
    var title = document.querySelector('.page-title');
    if (!title) return;

    title.addEventListener('focus', function () {
      if (title.textContent.trim() === 'PROJECT NAME') {
        var range = document.createRange();
        range.selectNodeContents(title);
        var sel = window.getSelection();
        sel.removeAllRanges();
        sel.addRange(range);
      }
    });

    title.addEventListener('blur', function () {
      if (!title.textContent.trim()) {
        title.textContent = 'PROJECT NAME';
      }
    });
  }

  // ── Content block placeholder ─────────────────────────────

  function initContentBlock() {
    var block = document.querySelector('.content-block');
    if (!block) return;

    var placeholder = 'Add project notes, context, or any additional information here…';

    block.addEventListener('focus', function () {
      if (block.textContent.trim() === placeholder) {
        block.textContent = '';
        block.style.fontStyle = 'normal';
        block.style.color = 'var(--text-white)';
      }
    });

    block.addEventListener('blur', function () {
      if (!block.textContent.trim()) {
        block.textContent = placeholder;
        block.style.fontStyle = 'italic';
        block.style.color = '#666';
      }
    });
  }

  // ── Subpage name — prevent newline on Enter ───────────────

  function initSubpageNameBehavior() {
    document.addEventListener('keydown', function (e) {
      if (e.key === 'Enter') {
        var active = document.activeElement;
        if (active && active.classList.contains('subpage-name')) {
          e.preventDefault();
          active.blur();
        }
      }
    });
  }

  // ── To link a subpage to a Notion URL ────────────────────
  //
  //  Set data-url on the .subpage-item element:
  //
  //  <div class="subpage-item" data-url="https://notion.so/your-page-id">
  //
  //  The Open button and row click will both navigate to that URL.
  //
  //  Example — set all URLs programmatically:
  //
  //  setSubpageUrls([
  //    { index: 0, url: 'https://notion.so/abc123' },
  //    { index: 1, url: 'https://notion.so/def456' },
  //  ]);

  window.setSubpageUrls = function (mappings) {
    var items = document.querySelectorAll('#subpage-list .subpage-item');
    mappings.forEach(function (m) {
      if (items[m.index]) items[m.index].dataset.url = m.url;
    });
  };

  // ── Init ──────────────────────────────────────────────────

  document.addEventListener('DOMContentLoaded', function () {
    initDates();
    updateStatPages();
    wireButtons();
    initEditableTitle();
    initContentBlock();
    initSubpageNameBehavior();
    initKeyboardShortcuts();
  });

})();
