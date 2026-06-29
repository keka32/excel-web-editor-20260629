// sort-filter.js — table search & pagination
var TableFilter = (function () {
  'use strict';

  var PAGE_SIZE = 50;
  var currentSort = { col: -1, asc: true };

  function filterTable(data, query) {
    if (!query || !query.trim()) return data;
    var q = query.toLowerCase();
    return data.filter(function (row) {
      return row.some(function (cell) {
        return String(cell || '').toLowerCase().indexOf(q) !== -1;
      });
    });
  }

  function sortTable(data, colIdx) {
    if (currentSort.col === colIdx) currentSort.asc = !currentSort.asc;
    else { currentSort.col = colIdx; currentSort.asc = true; }
    return data.slice().sort(function (a, b) {
      var va = a[colIdx] || '', vb = b[colIdx] || '';
      var na = Number(va), nb = Number(vb);
      if (!isNaN(na) && !isNaN(nb)) return currentSort.asc ? na - nb : nb - na;
      return currentSort.asc ? String(va).localeCompare(String(vb)) : String(vb).localeCompare(String(va));
    });
  }

  function paginate(data, page) {
    var start = (page - 1) * PAGE_SIZE;
    return {
      items: data.slice(start, start + PAGE_SIZE),
      page: page,
      totalPages: Math.max(1, Math.ceil(data.length / PAGE_SIZE)),
      total: data.length,
    };
  }

  return { filter: filterTable, sort: sortTable, paginate: paginate, PAGE_SIZE: PAGE_SIZE };
})();
