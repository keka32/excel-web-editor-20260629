(function () {
  'use strict';
  var KEY = 'excel-editor-theme';
  function get() { return localStorage.getItem(KEY); }
  function set(t) { localStorage.setItem(KEY, t); }
  function apply(t) {
    document.documentElement.setAttribute('data-theme', t);
    var btn = document.getElementById('btnTheme');
    if (btn) btn.textContent = t === 'dark' ? '☀️' : '🌙';
  }
  function detect() { return window.matchMedia('(prefers-color-scheme: dark)').matches ? 'dark' : 'light'; }
  function toggle() {
    var cur = get() || detect();
    var next = cur === 'dark' ? 'light' : 'dark';
    set(next); apply(next);
  }
  function init() {
    apply(get() || detect());
    window.matchMedia('(prefers-color-scheme: dark)').addEventListener('change', function (e) {
      if (!get()) apply(e.matches ? 'dark' : 'light');
    });
  }
  document.addEventListener('DOMContentLoaded', function () {
    var btn = document.getElementById('btnTheme');
    if (btn) btn.addEventListener('click', toggle);
    init();
  });
})();
