class TableEditor {
  constructor(container, options = {}) {
    this.container = typeof container === 'string' ? document.querySelector(container) : container;
    if (!this.container) throw new Error('TableEditor: container not found');

    this.options = {
      columns: options.columns || 3,
      rows: options.rows || 3,
      headers: options.headers || null,
      data: options.data || null,
      minColWidth: options.minColWidth || 60,
      onChange: typeof options.onChange === 'function' ? options.onChange : null,
    };

    this.state = {
      data: [],
      headers: [],
      editingCell: null,
      colWidths: [],
      selectedCell: null,
    };

    this._init();
  }

  _init() {
    this._buildInitialData();
    this._render();
    this._attachEvents();
  }

  _buildInitialData() {
    const { columns, rows, headers, data } = this.options;

    if (data && Array.isArray(data)) {
      this.state.data = data.map(row => [...row]);
      this.state.headers = headers ? [...headers] : data[0] ? data[0].map((_, i) => `Sütun ${i + 1}`) : [];
    } else {
      this.state.headers = headers ? [...headers] : Array.from({ length: columns }, (_, i) => `Sütun ${i + 1}`);
      this.state.data = Array.from({ length: rows }, () => Array(this.state.headers.length).fill(''));
    }

    this.state.colWidths = this.state.headers.map(() => 150);
  }

  _render() {
    this.container.innerHTML = '';
    this.container.classList.add('table-editor-wrapper');

    this._renderToolbar();
    this._renderTable();
  }

  _renderToolbar() {
    const toolbar = document.createElement('div');
    toolbar.className = 'table-editor__toolbar';

    const addRowBtn = document.createElement('button');
    addRowBtn.className = 'table-editor__btn table-editor__btn--add-row';
    addRowBtn.type = 'button';
    addRowBtn.textContent = '+ Satır Ekle';
    addRowBtn.addEventListener('click', () => this.addRow());

    const addColBtn = document.createElement('button');
    addColBtn.className = 'table-editor__btn table-editor__btn--add-col';
    addColBtn.type = 'button';
    addColBtn.textContent = '+ Sütun Ekle';
    addColBtn.addEventListener('click', () => this.addColumn());

    const delRowBtn = document.createElement('button');
    delRowBtn.className = 'table-editor__btn table-editor__btn--del-row';
    delRowBtn.type = 'button';
    delRowBtn.textContent = '− Satır Sil';
    delRowBtn.addEventListener('click', () => this.deleteLastRow());

    const delColBtn = document.createElement('button');
    delColBtn.className = 'table-editor__btn table-editor__btn--del-col';
    delColBtn.type = 'button';
    delColBtn.textContent = '− Sütun Sil';
    delColBtn.addEventListener('click', () => this.deleteLastColumn());

    toolbar.appendChild(addRowBtn);
    toolbar.appendChild(addColBtn);
    toolbar.appendChild(delRowBtn);
    toolbar.appendChild(delColBtn);

    this.container.appendChild(toolbar);
    this.toolbar = toolbar;
  }

  _renderTable() {
    const tableWrap = document.createElement('div');
    tableWrap.className = 'table-editor__table-wrap';

    const table = document.createElement('table');
    table.className = 'table-editor__table';

    const thead = document.createElement('thead');
    const headerRow = document.createElement('tr');

    this.state.headers.forEach((header, colIdx) => {
      const th = document.createElement('th');
      th.className = 'table-editor__th';
      th.style.width = `${this.state.colWidths[colIdx]}px`;
      th.setAttribute('data-col', colIdx);

      const headerText = document.createElement('span');
      headerText.className = 'table-editor__header-text';
      headerText.textContent = header;
      th.appendChild(headerText);

      const resizeHandle = document.createElement('div');
      resizeHandle.className = 'table-editor__resize-handle';
      resizeHandle.addEventListener('mousedown', (e) => this._startResize(e, colIdx));
      th.appendChild(resizeHandle);

      headerRow.appendChild(th);
    });

    thead.appendChild(headerRow);
    table.appendChild(thead);

    const tbody = document.createElement('tbody');

    this.state.data.forEach((row, rowIdx) => {
      const tr = document.createElement('tr');
      tr.setAttribute('data-row', rowIdx);

      row.forEach((cell, colIdx) => {
        const td = document.createElement('td');
        td.className = 'table-editor__td';
        td.setAttribute('data-row', rowIdx);
        td.setAttribute('data-col', colIdx);
        td.setAttribute('tabindex', '0');
        td.style.width = `${this.state.colWidths[colIdx]}px`;
        td.textContent = cell;

        td.addEventListener('dblclick', (e) => this._startEdit(e, rowIdx, colIdx));
        td.addEventListener('keydown', (e) => this._handleCellKeydown(e, rowIdx, colIdx));
        td.addEventListener('focus', () => this._onCellFocus(rowIdx, colIdx));

        tr.appendChild(td);
      });

      tbody.appendChild(tr);
    });

    table.appendChild(tbody);
    tableWrap.appendChild(table);
    this.container.appendChild(tableWrap);

    this.table = table;
    this.tbody = tbody;
  }

  _attachEvents() {
    this._onDocMouseDown = (e) => {
      if (this.state.editingCell) {
        const td = e.target.closest('.table-editor__td');
        if (!td || td !== this.state.editingCell.td) {
          this._commitEdit();
        }
      }
    };
    document.addEventListener('mousedown', this._onDocMouseDown);
  }

  _startEdit(e, rowIdx, colIdx) {
    e.preventDefault();
    if (this.state.editingCell) {
      this._commitEdit();
    }

    const td = this.tbody.querySelector(`td[data-row="${rowIdx}"][data-col="${colIdx}"]`);
    if (!td) return;

    const currentValue = this.state.data[rowIdx][colIdx];

    const input = document.createElement('input');
    input.type = 'text';
    input.className = 'table-editor__input';
    input.value = currentValue;

    td.textContent = '';
    td.appendChild(input);
    td.classList.add('table-editor__td--editing');
    input.focus();
    input.select();

    input.addEventListener('keydown', (ev) => {
      if (ev.key === 'Enter') {
        ev.preventDefault();
        this._commitEdit();
        this._moveSelection(rowIdx + 1, colIdx);
      } else if (ev.key === 'Escape') {
        ev.preventDefault();
        this._cancelEdit();
      } else if (ev.key === 'Tab') {
        ev.preventDefault();
        this._commitEdit();
        if (ev.shiftKey) {
          this._moveSelection(rowIdx, colIdx - 1);
        } else {
          this._moveSelection(rowIdx, colIdx + 1);
        }
      }
    });

    this.state.editingCell = { td, input, rowIdx, colIdx, originalValue: currentValue };
  }

  _commitEdit() {
    const ec = this.state.editingCell;
    if (!ec) return;

    const newValue = ec.input.value;
    this.state.data[ec.rowIdx][ec.colIdx] = newValue;
    ec.td.textContent = newValue;
    ec.td.classList.remove('table-editor__td--editing');

    this.state.editingCell = null;

    if (newValue !== ec.originalValue) {
      this._notifyChange();
    }
  }

  _cancelEdit() {
    const ec = this.state.editingCell;
    if (!ec) return;

    ec.td.textContent = ec.originalValue;
    ec.td.classList.remove('table-editor__td--editing');
    this.state.editingCell = null;
  }

  _handleCellKeydown(e, rowIdx, colIdx) {
    if (this.state.editingCell) return;

    switch (e.key) {
      case 'Enter':
        e.preventDefault();
        this._startEdit(e, rowIdx, colIdx);
        break;
      case 'Tab':
        e.preventDefault();
        if (e.shiftKey) {
          this._moveSelection(rowIdx, colIdx - 1);
        } else {
          this._moveSelection(rowIdx, colIdx + 1);
        }
        break;
      case 'ArrowUp':
        e.preventDefault();
        this._moveSelection(rowIdx - 1, colIdx);
        break;
      case 'ArrowDown':
        e.preventDefault();
        this._moveSelection(rowIdx + 1, colIdx);
        break;
      case 'ArrowLeft':
        e.preventDefault();
        this._moveSelection(rowIdx, colIdx - 1);
        break;
      case 'ArrowRight':
        e.preventDefault();
        this._moveSelection(rowIdx, colIdx + 1);
        break;
      case 'Delete':
      case 'Backspace':
        e.preventDefault();
        this._clearCell(rowIdx, colIdx);
        break;
    }
  }

  _onCellFocus(rowIdx, colIdx) {
    this.state.selectedCell = { rowIdx, colIdx };
    this._clearSelectionHighlight();
    const td = this.tbody.querySelector(`td[data-row="${rowIdx}"][data-col="${colIdx}"]`);
    if (td) td.classList.add('table-editor__td--selected');
  }

  _clearSelectionHighlight() {
    const selected = this.tbody.querySelectorAll('.table-editor__td--selected');
    selected.forEach(el => el.classList.remove('table-editor__td--selected'));
  }

  _moveSelection(rowIdx, colIdx) {
    const maxRow = this.state.data.length - 1;
    const maxCol = this.state.headers.length - 1;

    if (rowIdx < 0) rowIdx = 0;
    if (rowIdx > maxRow) rowIdx = maxRow;
    if (colIdx < 0) {
      if (rowIdx > 0) { rowIdx--; colIdx = maxCol; }
      else { colIdx = 0; }
    }
    if (colIdx > maxCol) {
      if (rowIdx < maxRow) { rowIdx++; colIdx = 0; }
      else { colIdx = maxCol; }
    }

    const td = this.tbody.querySelector(`td[data-row="${rowIdx}"][data-col="${colIdx}"]`);
    if (td) td.focus();
  }

  _clearCell(rowIdx, colIdx) {
    if (this.state.data[rowIdx][colIdx] === '') return;
    this.state.data[rowIdx][colIdx] = '';
    const td = this.tbody.querySelector(`td[data-row="${rowIdx}"][data-col="${colIdx}"]`);
    if (td) td.textContent = '';
    this._notifyChange();
  }

  _startResize(e, colIdx) {
    e.preventDefault();
    e.stopPropagation();

    const startX = e.clientX;
    const startWidth = this.state.colWidths[colIdx];

    const onMouseMove = (ev) => {
      const diff = ev.clientX - startX;
      const newWidth = Math.max(this.options.minColWidth, startWidth + diff);
      this.state.colWidths[colIdx] = newWidth;

      const ths = this.table.querySelectorAll(`th[data-col="${colIdx}"]`);
      ths.forEach(th => { th.style.width = `${newWidth}px`; });

      const tds = this.tbody.querySelectorAll(`td[data-col="${colIdx}"]`);
      tds.forEach(td => { td.style.width = `${newWidth}px`; });
    };

    const onMouseUp = () => {
      document.removeEventListener('mousemove', onMouseMove);
      document.removeEventListener('mouseup', onMouseUp);
    };

    document.addEventListener('mousemove', onMouseMove);
    document.addEventListener('mouseup', onMouseUp);
  }

  addRow() {
    if (this.state.editingCell) this._commitEdit();
    this.state.data.push(Array(this.state.headers.length).fill(''));
    this._renderTable();
    this._notifyChange();
  }

  addColumn() {
    if (this.state.editingCell) this._commitEdit();
    const newColName = `Sütun ${this.state.headers.length + 1}`;
    this.state.headers.push(newColName);
    this.state.colWidths.push(150);
    this.state.data.forEach(row => row.push(''));
    this._renderTable();
    this._notifyChange();
  }

  deleteLastRow() {
    if (this.state.data.length <= 1) return;
    if (this.state.editingCell) this._cancelEdit();
    this.state.data.pop();
    this._renderTable();
    this._notifyChange();
  }

  deleteLastColumn() {
    if (this.state.headers.length <= 1) return;
    if (this.state.editingCell) this._cancelEdit();
    this.state.headers.pop();
    this.state.colWidths.pop();
    this.state.data.forEach(row => row.pop());
    this._renderTable();
    this._notifyChange();
  }

  deleteRow(rowIdx) {
    if (this.state.data.length <= 1) return;
    if (rowIdx < 0 || rowIdx >= this.state.data.length) return;
    if (this.state.editingCell && this.state.editingCell.rowIdx === rowIdx) this._cancelEdit();
    this.state.data.splice(rowIdx, 1);
    this._renderTable();
    this._notifyChange();
  }

  deleteColumn(colIdx) {
    if (this.state.headers.length <= 1) return;
    if (colIdx < 0 || colIdx >= this.state.headers.length) return;
    if (this.state.editingCell && this.state.editingCell.colIdx === colIdx) this._cancelEdit();
    this.state.headers.splice(colIdx, 1);
    this.state.colWidths.splice(colIdx, 1);
    this.state.data.forEach(row => row.splice(colIdx, 1));
    this._renderTable();
    this._notifyChange();
  }

  getData() {
    return {
      headers: [...this.state.headers],
      data: this.state.data.map(row => [...row]),
    };
  }

  setData(data, headers) {
    if (this.state.editingCell) this._cancelEdit();
    this.state.data = data.map(row => [...row]);
    if (headers) this.state.headers = [...headers];
    this._renderTable();
  }

  _notifyChange() {
    if (this.options.onChange) {
      this.options.onChange(this.getData());
    }
  }

  destroy() {
    if (this.state.editingCell) this._cancelEdit();
    document.removeEventListener('mousedown', this._onDocMouseDown);
    this.container.innerHTML = '';
  }
}

if (typeof module !== 'undefined' && module.exports) {
  module.exports = TableEditor;
}
