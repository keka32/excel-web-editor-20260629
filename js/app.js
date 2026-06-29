// app.js — Excel Editor with multi-tab + mobile + manual sync
(function () {
  'use strict';

  var isMobile = window.matchMedia('(max-width: 768px)').matches;

  /* ──────────────────────────────
   *  DOM refs
   * ────────────────────────────── */
  var $ = function (id) { return document.getElementById(id); };
  var container = $('tableContainer');
  var fileNameEl = $('fileName');
  var cellInfoEl = $('cellInfo');
  var formulaInput = $('formulaInput');
  var fileInput = $('fileInput');
  var searchInput = $('searchInput');
  var statusText = $('statusText');
  var dropZone = $('dropZone');
  var editOverlay = $('editOverlay');
  var editInput = $('editInput');
  var editLabel = $('editLabel');

  /* ──────────────────────────────
   *  Toast
   * ────────────────────────────── */
  function toast(msg, type) {
    if (!msg) return;
    var tc = $('toastContainer');
    if (!tc) return;
    var el = document.createElement('div');
    el.className = 'toast toast-' + (type || 'info');
    el.textContent = msg;
    tc.appendChild(el);
    setTimeout(function () { el.remove(); }, 2500);
  }

  /* ──────────────────────────────
   *  Storage + Tabs
   * ────────────────────────────── */
  var storage = new AutoSaveStorage({
    onSave: function (id) {
      if (statusText) { statusText.textContent = 'Kaydedildi: ' + id; setTimeout(function () { statusText.textContent = 'Hazir'; }, 2000); }
    }
  });

  var tabs = TabManager.create({
    storage: storage,
    onSwitch: function (t) { loadTab(t); },
    onChange: function (n) {}
  });

  window.tabs = tabs;

  /* ──────────────────────────────
   *  Active tab state
   * ────────────────────────────── */
  var workbook, worksheet, sheetName, currentFileName;
  var ROWS = 50, COLS = 26;

  function loadTab(t) {
    workbook = t.workbook;
    sheetName = t.sheetName;
    worksheet = workbook.Sheets[sheetName];
    currentFileName = t.name;
    fileNameEl.textContent = currentFileName;
    cellInfoEl.textContent = 'Secili: A1';
    formulaInput.value = '';
    if (dropZone) dropZone.style.display = 'none';
    renderGrid();
  }

  /* ──────────────────────────────
   *  Sheet helpers
   * ────────────────────────────── */
  function colLabel(i) {
    var s = '';
    while (i >= 0) { s = String.fromCharCode(65 + (i % 26)) + s; i = Math.floor(i / 26) - 1; }
    return s;
  }
  function cellRef(r, c) { return colLabel(c) + (r + 1); }

  function getCell(r, c) {
    var ref = cellRef(r, c);
    if (!worksheet || !worksheet[ref]) return '';
    var cell = worksheet[ref];
    return cell.w !== undefined ? cell.w : (cell.v !== undefined ? String(cell.v) : '');
  }

  function setCell(r, c, val) {
    var ref = cellRef(r, c);
    if (!worksheet) return;
    if (val === '') { delete worksheet[ref]; return; }
    var num = parseFloat(val);
    if (!isNaN(num) && String(num) === val.trim()) XLSX.utils.sheet_add_aoa(worksheet, [[num]], { origin: ref });
    else XLSX.utils.sheet_add_aoa(worksheet, [[val]], { origin: ref });
  }

  /* ──────────────────────────────
   *  Render table
   * ────────────────────────────── */
  function renderGrid() {
    var searchVal = (searchInput && searchInput.value || '').toLowerCase();
    var table = document.createElement('table');
    table.className = 'table-grid';
    var thd = document.createElement('thead'), hRow = document.createElement('tr');
    var corner = document.createElement('th'); corner.textContent = ''; hRow.appendChild(corner);

    for (var c = 0; c < COLS; c++) {
      var th = document.createElement('th');
      th.textContent = colLabel(c);
      th.onclick = (function (ci) { return function () { sortColumn(ci); }; })(c);
      hRow.appendChild(th);
    }
    thd.appendChild(hRow); table.appendChild(thd);

    var tbd = document.createElement('tbody');
    for (var r = 0; r < ROWS; r++) {
      if (searchVal && !rowMatches(r, searchVal)) continue;
      var tr = document.createElement('tr');
      var rh = document.createElement('th'); rh.textContent = r + 1; tr.appendChild(rh);
      for (var cc = 0; cc < COLS; cc++) {
        var v = getCell(r, cc);
        var td = document.createElement('td');
        td.dataset.row = r; td.dataset.col = cc; td.textContent = v;

        if (isMobile) {
          td.contentEditable = false;
          td.addEventListener('click', onCellTap);
          td.addEventListener('dblclick', onCellDblTap);
        } else {
          td.contentEditable = true;
          td.addEventListener('focus', onCellFocusDesktop);
          td.addEventListener('input', onCellInputDesktop);
          td.addEventListener('keydown', onCellKeyDesktop);
        }
        tr.appendChild(td);
      }
      tbd.appendChild(tr);
    }
    table.appendChild(tbd);
    container.innerHTML = '';
    container.appendChild(table);
  }

  function rowMatches(r, q) {
    for (var cc = 0; cc < COLS; cc++) {
      if (String(getCell(r, cc) || '').toLowerCase().indexOf(q) !== -1) return true;
    }
    return false;
  }

  /* ──────────────────────────────
   *  DESKTOP: contentEditable cells
   * ────────────────────────────── */
  function onCellFocusDesktop(e) {
    var td = e.currentTarget;
    var r = +td.dataset.row, c = +td.dataset.col;
    cellInfoEl.textContent = 'Secili: ' + cellRef(r, c);
    formulaInput.value = getCell(r, c);
  }

  function onCellInputDesktop(e) {
    var td = e.currentTarget;
    var r = +td.dataset.row, c = +td.dataset.col;
    var val = td.textContent.trim();
    setCell(r, c, val);
    formulaInput.value = val;
    tabs.markDirty();
    autoSave();
  }

  function onCellKeyDesktop(e) {
    var td = e.currentTarget;
    var r = +td.dataset.row, c = +td.dataset.col;
    var nr = r, nc = c;
    switch (e.key) {
      case 'Enter': e.preventDefault(); nr = Math.min(r + 1, ROWS - 1); break;
      case 'Tab': e.preventDefault(); nc = e.shiftKey ? Math.max(c - 1, 0) : Math.min(c + 1, COLS - 1); break;
      case 'ArrowUp': e.preventDefault(); nr = Math.max(r - 1, 0); break;
      case 'ArrowDown': e.preventDefault(); nr = Math.min(r + 1, ROWS - 1); break;
      default: return;
    }
    var tds = container.querySelectorAll('td');
    for (var i = 0; i < tds.length; i++) {
      if (+tds[i].dataset.row === nr && +tds[i].dataset.col === nc) { tds[i].focus(); break; }
    }
  }

  /* ──────────────────────────────
   *  MOBILE: tap-select + overlay
   * ────────────────────────────── */
  var selectedCellTD = null;

  function onCellTap(e) {
    selectCell(e.currentTarget);
  }

  function onCellDblTap(e) {
    selectCell(e.currentTarget);
    openEditOverlay(e.currentTarget);
  }

  function selectCell(td) {
    if (selectedCellTD) selectedCellTD.classList.remove('cell-selected');
    selectedCellTD = td;
    td.classList.add('cell-selected');
    var r = +td.dataset.row, c = +td.dataset.col;
    cellInfoEl.textContent = 'Secili: ' + cellRef(r, c);
    formulaInput.value = getCell(r, c);
  }

  function openEditOverlay(td) {
    if (!editOverlay || !editInput) return;
    editLabel.textContent = cellRef(+td.dataset.row, +td.dataset.col);
    editInput.value = td.textContent.trim();
    editInput.dataset.row = td.dataset.row;
    editInput.dataset.col = td.dataset.col;
    editOverlay.classList.add('open');
    editInput.focus();
  }

  function closeEditOverlay(save) {
    if (!editOverlay) return;
    if (save && selectedCellTD) {
      var r = +editInput.dataset.row, c = +editInput.dataset.col;
      var val = editInput.value.trim();
      setCell(r, c, val);
      selectedCellTD.textContent = val;
      formulaInput.value = val;
      tabs.markDirty();
      autoSave();
    }
    editOverlay.classList.remove('open');
  }

  if ($('editDone')) $('editDone').addEventListener('click', function () { closeEditOverlay(true); });
  if ($('editCancel')) $('editCancel').addEventListener('click', function () { closeEditOverlay(false); });
  if ($('btnEdit')) $('btnEdit').addEventListener('click', function () { if (selectedCellTD) openEditOverlay(selectedCellTD); });

  // Mobilde ok tuslariyla gezinme
  document.addEventListener('keydown', function (e) {
    if (isMobile && selectedCellTD && !editOverlay.classList.contains('open')) {
      var r = +selectedCellTD.dataset.row, c = +selectedCellTD.dataset.col;
      var nr = r, nc = c;
      switch (e.key) {
        case 'ArrowUp': e.preventDefault(); nr = Math.max(r - 1, 0); break;
        case 'ArrowDown': e.preventDefault(); nr = Math.min(r + 1, ROWS - 1); break;
        case 'ArrowLeft': e.preventDefault(); nc = Math.max(c - 1, 0); break;
        case 'ArrowRight': e.preventDefault(); nc = Math.min(c + 1, COLS - 1); break;
        case 'Enter': e.preventDefault(); openEditOverlay(selectedCellTD); return;
        case 'Delete': case 'Backspace':
          e.preventDefault();
          setCell(r, c, '');
          selectedCellTD.textContent = '';
          tabs.markDirty(); autoSave();
          return;
        default: return;
      }
      var tds = container.querySelectorAll('td');
      for (var i = 0; i < tds.length; i++) {
        if (+tds[i].dataset.row === nr && +tds[i].dataset.col === nc) { selectCell(tds[i]); break; }
      }
    }
  });

  /* ──────────────────────────────
   *  Formula bar
   * ────────────────────────────── */
  formulaInput.addEventListener('keydown', function (e) {
    if (e.key !== 'Enter') return;
    e.preventDefault();
    var td = selectedCellTD || container.querySelector('td:focus');
    if (td) {
      var r = +td.dataset.row, c = +td.dataset.col;
      setCell(r, c, formulaInput.value);
      td.textContent = formulaInput.value;
      tabs.markDirty(); autoSave();
    }
  });

  /* ──────────────────────────────
   *  Sort
   * ────────────────────────────── */
  var sortCol = -1, sortAsc = true;
  function sortColumn(ci) {
    if (sortCol === ci) sortAsc = !sortAsc; else { sortCol = ci; sortAsc = true; }
    var ref = XLSX.utils.decode_range(worksheet['!ref'] || 'A1');
    var rows = [];
    for (var r = 0; r <= ref.e.r; r++) {
      var row = [];
      for (var c = 0; c <= ref.e.c; c++) row.push(getCell(r, c));
      rows.push(row);
    }
    rows.sort(function (a, b) {
      var va = a[ci] || '', vb = b[ci] || '';
      var na = Number(va), nb = Number(vb);
      if (!isNaN(na) && !isNaN(nb)) return sortAsc ? na - nb : nb - na;
      return sortAsc ? String(va).localeCompare(String(vb)) : String(vb).localeCompare(String(va));
    });
    for (var rr = 0; rr < rows.length; rr++)
      for (var cc = 0; cc <= ref.e.c; cc++)
        setCell(rr, cc, rows[rr][cc]);
    tabs.markDirty(); renderGrid(); autoSave();
  }

  /* ──────────────────────────────
   *  Search
   * ────────────────────────────── */
  if (searchInput) searchInput.addEventListener('input', function () { renderGrid(); });

  /* ──────────────────────────────
   *  AutoSave (localStorage only)
   * ────────────────────────────── */
  var saveTimer;
  function autoSave() {
    clearTimeout(saveTimer);
    saveTimer = setTimeout(function () {
      tabs.saveCurrent();
      var el = $('autoSaveStatus');
      if (el) { el.style.opacity = '1'; setTimeout(function () { el.style.opacity = '0.5'; }, 1000); }
    }, 800);
  }

  /* ──────────────────────────────
   *  File upload
   * ────────────────────────────── */
  window.handleFileUpload = function (file) {
    var reader = new FileReader();
    reader.onload = function (ev) {
      var data = new Uint8Array(ev.target.result);
      var wb = XLSX.read(data, { type: 'array' });
      var st = wb.SheetNames[0];
      var ws = wb.Sheets[st];
      var ref = XLSX.utils.decode_range(ws['!ref'] || 'A1');
      if (ref.e.r < ROWS - 1 || ref.e.c < COLS - 1) {
        ws['!ref'] = XLSX.utils.encode_range({ s: { r: 0, c: 0 }, e: { r: Math.max(ref.e.r, ROWS - 1), c: Math.max(ref.e.c, COLS - 1) } });
      }
      tabs.add(file.name, wb, st);
      loadTab(tabs.getActive());
      toast('Dosya yuklendi: ' + file.name, 'success');
    };
    reader.readAsArrayBuffer(file);
    fileInput.value = '';
  };

  fileInput.addEventListener('change', function (e) {
    var f = e.target.files[0];
    if (f) handleFileUpload(f);
  });

  if ($('btnOpen')) $('btnOpen').addEventListener('click', function () { fileInput.click(); });
  if ($('btnNew')) $('btnNew').addEventListener('click', function () { tabs.newEmpty(); loadTab(tabs.getActive()); toast('Yeni sekme', 'info'); });

  /* ──────────────────────────────
   *  Export
   * ────────────────────────────── */
  function doExport() {
    var name = tabs.exportActive();
    if (name) toast('Disa aktarildi: ' + name, 'success');
  }
  if ($('btnSave')) $('btnSave').addEventListener('click', doExport);
  if ($('btnExport')) $('btnExport').addEventListener('click', doExport);

  /* ──────────────────────────────
   *  Global keyboard shortcuts
   * ────────────────────────────── */
  document.addEventListener('keydown', function (e) {
    if (e.ctrlKey || e.metaKey) {
      if (e.key === 's') { e.preventDefault(); tabs.saveAll(); toast('Tum dosyalar kaydedildi', 'success'); }
      if (e.key === 'e') { e.preventDefault(); doExport(); }
      if (e.key === 'n') { e.preventDefault(); tabs.newEmpty(); loadTab(tabs.getActive()); }
    }
  });

  /* ──────────────────────────────
   *  Manual Cloud sync
   * ────────────────────────────── */
  var syncWatching = false;

  function pushAllToCloud() {
    if (!CloudSync.isReady()) return;
    var all = tabs.tabs;
    for (var i = 0; i < all.length; i++) {
      try {
        var data = XLSX.write(all[i].workbook, { bookType: 'xlsx', type: 'array' });
        CloudSync.saveFile(all[i].name, Array.from(new Uint8Array(data)));
      } catch (e) {}
    }
  }

  if ($('btnSync')) {
    $('btnSync').addEventListener('click', function () {
      if (!CloudSync.isReady()) {
        toast(CloudSync.lastError() || 'Senkronizasyon hazır değil', 'error');
        return;
      }
      toast('Senkronize ediliyor...', 'info');
      tabs.saveAll();
      pushAllToCloud();
      CloudSync.loadAllFiles().then(function (files) {
        var added = 0;
        files.forEach(function (f) {
          try {
            var data = new Uint8Array(f.data);
            var wb = XLSX.read(data, { type: 'array' });
            tabs.add(f.name, wb);
            added++;
          } catch (e) {}
        });
        loadTab(tabs.getActive());
        toast(added + ' dosya senkronize edildi', 'success');

        if (!syncWatching) {
          syncWatching = true;
          CloudSync.watchChanges(function (name, rawData) {
            if (!rawData) return;
            try {
              var data = new Uint8Array(rawData);
              var wb = XLSX.read(data, { type: 'array' });
              var existIdx = tabs.exists(name);
              if (existIdx >= 0) {
                tabs.tabs[existIdx].workbook = wb;
                tabs.tabs[existIdx].sheetName = wb.SheetNames[0];
                toast('Güncellendi: ' + name, 'info');
              } else {
                tabs.add(name, wb);
                toast('Yeni: ' + name, 'success');
              }
              if (tabs.activeIdx >= 0 && tabs.tabs[tabs.activeIdx].name === name) loadTab(tabs.getActive());
            } catch (e) {}
          });
        }
      });
    });
  }

  /* ──────────────────────────────
   *  Init
   * ────────────────────────────── */
  function init() {
    var loaded = tabs.loadAll();
    if (loaded.length === 0) tabs.newEmpty();
    loadTab(tabs.getActive());
    if (dropZone) dropZone.style.display = loaded.length ? 'none' : 'block';
  }

  if (typeof XLSX !== 'undefined') init();
  else document.addEventListener('DOMContentLoaded', init);
})();
