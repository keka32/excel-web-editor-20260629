// tabs.js — Multi-file tab manager
var TabManager = (function () {
  'use strict';

  function Tabs(opts) {
    opts = opts || {};
    this.storage = opts.storage;
    this.onSwitch = opts.onSwitch || function () {};
    this.onChange = opts.onChange || function () {};
    this.tabs = [];
    this.activeIdx = -1;
    this.nextId = 0;
  }

  Tabs.prototype.exists = function (name) {
    for (var i = 0; i < this.tabs.length; i++) {
      if (this.tabs[i].name === name) return i;
    }
    return -1;
  };

  Tabs.prototype.add = function (name, wb, sheet) {
    var existIdx = this.exists(name);
    if (existIdx >= 0) { this.switchTo(existIdx); return existIdx; }
    var tab = {
      id: 'tab_' + (this.nextId++),
      name: name,
      workbook: wb,
      sheetName: sheet || (wb && wb.SheetNames ? wb.SheetNames[0] : 'Sheet1'),
      dirty: false,
    };
    this.tabs.push(tab);
    this.activeIdx = this.tabs.length - 1;
    this.render();
    this.changed();
    return this.activeIdx;
  };

  Tabs.prototype.switchTo = function (idx) {
    if (idx < 0 || idx >= this.tabs.length || idx === this.activeIdx) return false;
    this._saveCurrent();
    this.activeIdx = idx;
    var t = this.tabs[idx];
    if (this.onSwitch) this.onSwitch(t);
    this.render();
    return true;
  };

  Tabs.prototype.close = function (idx) {
    if (this.tabs.length <= 1) return false;
    if (idx < 0 || idx >= this.tabs.length) return false;
    var name = this.tabs[idx].name;
    this.tabs.splice(idx, 1);
    if (this.storage) this.storage.remove(name);
    if (idx <= this.activeIdx) this.activeIdx = Math.max(0, this.activeIdx - 1);
    if (this.activeIdx >= this.tabs.length) this.activeIdx = 0;
    var t = this.tabs[this.activeIdx];
    if (this.onSwitch) this.onSwitch(t);
    this.render();
    this.changed();
  };

  Tabs.prototype.markDirty = function () {
    if (this.activeIdx >= 0 && this.activeIdx < this.tabs.length) {
      this.tabs[this.activeIdx].dirty = true;
    }
  };

  Tabs.prototype.getActive = function () {
    return this.activeIdx >= 0 ? this.tabs[this.activeIdx] : null;
  };

  Tabs.prototype._saveCurrent = function () {
    var t = this.getActive();
    if (!t || !this.storage) return;
    try {
      var data = XLSX.write(t.workbook, { bookType: 'xlsx', type: 'array' });
      this.storage.save(t.name, Array.from(new Uint8Array(data)));
      t.dirty = false;
    } catch (e) { console.warn('Tab save failed:', t.name, e); }
  };

  Tabs.prototype.saveCurrent = function () { this._saveCurrent(); };

  Tabs.prototype.saveAll = function () {
    for (var i = 0; i < this.tabs.length; i++) {
      var cur = this.activeIdx;
      this.activeIdx = i;
      this._saveCurrent();
      this.activeIdx = cur;
    }
    this.changed();
  };

  Tabs.prototype.loadAll = function () {
    if (!this.storage || !this.storage.available()) return [];
    var files = this.storage.listFiles();
    files.sort();
    var loaded = [];
    for (var i = 0; i < files.length; i++) {
      var raw = this.storage.load(files[i]);
      if (!raw) continue;
      try {
        var data = new Uint8Array(raw);
        var wb = XLSX.read(data, { type: 'array' });
        var st = wb.SheetNames[0];
        this.tabs.push({
          id: 'tab_' + (this.nextId++),
          name: files[i],
          workbook: wb,
          sheetName: st,
          dirty: false,
        });
        loaded.push(files[i]);
      } catch (e) { console.warn('Tab load failed:', files[i], e); }
    }
    if (this.tabs.length === 0) return [];
    this.activeIdx = 0;
    this.render();
    return loaded;
  };

  Tabs.prototype.exportActive = function () {
    var t = this.getActive();
    if (!t) return;
    var out = XLSX.write(t.workbook, { bookType: 'xlsx', type: 'array' });
    var blob = new Blob([out], { type: 'application/octet-stream' });
    var url = URL.createObjectURL(blob);
    var a = document.createElement('a');
    a.href = url;
    a.download = t.name.replace(/\.[^/.]+$/, '') + '_edited.xlsx';
    a.click();
    URL.revokeObjectURL(url);
    return a.download;
  };

  Tabs.prototype.render = function () {
    var el = document.getElementById('tabList');
    if (!el) return;
    el.innerHTML = '';
    for (var i = 0; i < this.tabs.length; i++) {
      var t = this.tabs[i];
      var span = document.createElement('span');
      span.className = 'tab-item' + (i === this.activeIdx ? ' active' : '');
      var nameSpan = document.createElement('span');
      nameSpan.className = 'tab-name';
      nameSpan.textContent = (t.dirty ? '● ' : '') + t.name;
      nameSpan.onclick = (function (idx) { return function () { self.switchTo(idx); }; })(i);
      span.appendChild(nameSpan);
      if (this.tabs.length > 1) {
        var closeBtn = document.createElement('button');
        closeBtn.className = 'tab-close';
        closeBtn.textContent = '×';
        closeBtn.onclick = (function (idx) { return function (e) { e.stopPropagation(); self.close(idx); }; })(i);
        span.appendChild(closeBtn);
      }
      el.appendChild(span);
    }
    var addBtn = document.createElement('span');
    addBtn.className = 'tab-item tab-add';
    addBtn.textContent = '＋';
    addBtn.title = 'Yeni';
    addBtn.onclick = function () { self.newEmpty(); };
    el.appendChild(addBtn);
  };

  Tabs.prototype.newEmpty = function () {
    var data = [];
    for (var r = 0; r < 50; r++) data.push(new Array(26).fill(''));
    var ws = XLSX.utils.aoa_to_sheet(data);
    var wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, 'Sheet1');
    var name = 'Yeni-' + (new Date().toLocaleTimeString().replace(/:/g, '-'));
    this.add(name, wb, 'Sheet1');
    this.changed();
  };

  Tabs.prototype.changed = function () {
    if (this.onChange) this.onChange(this.tabs.length);
  };

  var self;
  function create(opts) { self = new Tabs(opts); return self; }

  return { create: create };
})();
