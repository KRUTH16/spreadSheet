import { Component, h, State, Prop, Listen } from '@stencil/core';
import { Sheet, Cell, CellStyle } from './types';
import { FormulaEngine } from './formula-engine';
import { CsvUtils } from './csv-utils';

type SpreadsheetSnapshot = {
  sheet: Sheet;
  selectedRow: number;
  selectedCol: number;
  columnWidths: number[];
  rowHeights: number[];
};

@Component({
  tag: 'app-spreadsheet',
  styleUrl: 'spreadsheet.css',
  shadow: true,
})
export class AppSpreadsheet {
  @State() sheet: Sheet = this.createInitialSheet(100, 26);
  @State() selectedRow: number = 0;
  @State() selectedCol: number = 0;
  @State() columnWidths: number[] = [];
  @State() rowHeights: number[] = [];
  @State() formulaDraft: string = '';
  @State() visibleStartRow: number = 0;
  @State() visibleEndRow: number = 40;
  @State() rangeStartRow: number = 0;
  @State() rangeEndRow: number = 0;
  @State() rangeStartCol: number = 0;
  @State() rangeEndCol: number = 0;
  @State() isFilling: boolean = false;
  @State() fillStartRow: number = 0;
  @State() fillStartCol: number = 0;

  
  private readonly DEFAULT_ROW_HEIGHT = 21;
  private lastChangedRow: number | null = null;
  private lastChangedCol: number | null = null;
  private lastChangedValue: string | null = null;

  @Listen('cellValueChanged')
  handleCellValueChanged(event: CustomEvent<{ row: number; col: number; value: string }>) {
    const { row, col, value } = event.detail;

    // 1️⃣ snapshot current state first
    this.pushHistory();

    this.lastChangedRow = row;
    this.lastChangedCol = col;
    this.lastChangedValue = value;

    // 2️⃣ then apply the change
    this.updateCellValue(row, col, value);
    this.recalculateFormulas();
  }

  @Listen('cellSelected')
  handleCellSelected(event: CustomEvent<{ row: number; col: number }>) {
    // this.selectCell(event.detail.row, event.detail.col);
    const { row, col } = event.detail;
    this.selectCell(row, col);

    const cell = this.sheet[row][col];
    this.formulaDraft = cell.formula ?? cell.value ?? '';

    console.log(`📌 cellSelected → now selected (${row},${col}) value="${cell.value}", formula="${cell.formula}"`);
  }



  @Listen('keydown', { target: 'window' })
  handleKeyDown(e: KeyboardEvent) {
    const row = this.selectedRow;
    const col = this.selectedCol;

    // --- Undo / Redo ---
    if (e.ctrlKey || e.metaKey) {
      if (e.key === 'z' || e.key === 'Z') {
        e.preventDefault();
        this.undo();
        return;
      }
      if (e.key === 'y' || e.key === 'Y') {
        e.preventDefault();
        this.redo();
        return;
      }
    }

    // --- Copy ---
    if (e.ctrlKey && e.key === 'c') {
      this.copyCell();
      return;
    }

    // --- Paste ---
    if (e.ctrlKey && e.key === 'v') {
      this.pasteCell();
      return;
    }

    // --- Delete Key ---
    if (e.key === 'Delete') {
      this.clearCellValue(row, col);
      return;
    }

    // --- Enter: Start editing ---
    if (e.key === 'Enter') {
      const active = document.querySelector('app-cell[selected="true"]') as HTMLElement;
      if (active) active.dispatchEvent(new MouseEvent('dblclick', { bubbles: true }));
      return;
    }

    // --- Arrow Keys for Navigation ---
    switch (e.key) {
      case 'ArrowUp':
        this.selectCell(Math.max(0, row - 1), col);
        break;
      case 'ArrowDown':
        this.selectCell(Math.min(this.sheet.length - 1, row + 1), col);
        break;
      case 'ArrowLeft':
        this.selectCell(row, Math.max(0, col - 1));
        break;
      case 'ArrowRight':
        this.selectCell(row, Math.min(this.sheet[0].length - 1, col + 1));
        break;
    }

    if (e.key === 'Home') {
      this.selectCell(row, 0);
      return;
    }

    if (e.key === 'End') {
      this.selectCell(row, this.sheet[0].length - 1);
      return;
    }

    if (e.ctrlKey && e.key === 'Home') {
      this.selectCell(0, 0);
      return;
    }

    const isPrintable = e.key.length === 1 && !e.ctrlKey && !e.metaKey && !e.altKey;
    if (isPrintable) {
      const active = document.querySelector('app-cell[selected="true"]') as HTMLElement;
      if (active) active.dispatchEvent(new MouseEvent('dblclick', { bubbles: true }));

      // Allow typing the actual key
      return;
    }

    // SHIFT + Arrow = expand range
    if (e.shiftKey) {
      switch (e.key) {
        case 'ArrowDown':
          this.rangeEndRow = Math.min(this.sheet.length - 1, this.rangeEndRow + 1);
          break;
        case 'ArrowUp':
          this.rangeEndRow = Math.max(0, this.rangeEndRow - 1);
          break;
        case 'ArrowRight':
          this.rangeEndCol = Math.min(this.sheet[0].length - 1, this.rangeEndCol + 1);
          break;
        case 'ArrowLeft':
          this.rangeEndCol = Math.max(0, this.rangeEndCol - 1);
          break;
      }
      return;
    }

    if (e.ctrlKey && e.key === 'c') {
      this.copyCellRange();
      return;
    }

    if (e.ctrlKey && e.key === 'v') {
      this.pasteCellRange();
      return;
    }
  }

  private textColorInput?: HTMLInputElement;
  private bgColorInput?: HTMLInputElement;


  private cloneSheet(sheet: Sheet): Sheet {
    return sheet.map(row =>
      row.map(cell => ({
        value: cell.value ?? '',
        formula: cell.formula,
        style: cell.style ? { ...cell.style } : {},
      })),
    );
  }

  private takeSnapshot(): SpreadsheetSnapshot {
    return {
      sheet: this.cloneSheet(this.sheet),
      selectedRow: this.selectedRow,
      selectedCol: this.selectedCol,
      columnWidths: [...this.columnWidths],
      rowHeights: [...this.rowHeights],
    };
  }

  private restoreSnapshot(snap: SpreadsheetSnapshot) {
    console.log('⏪ restoreSnapshot');

    

    this.sheet = this.cloneSheet(snap.sheet);
    this.selectedRow = snap.selectedRow;
    this.selectedCol = snap.selectedCol;

    this.columnWidths = [...snap.columnWidths];
    this.rowHeights = [...snap.rowHeights];


    this.sheet = [...this.sheet];

    this.formulaDraft = this.sheet[this.selectedRow]?.[this.selectedCol]?.formula ?? this.sheet[this.selectedRow]?.[this.selectedCol]?.value ?? '';

    console.log(`✔ Undo finished — selection restored to ${this.selectedRow},${this.selectedCol}`);
  }

  private history: SpreadsheetSnapshot[] = [];
  private future: SpreadsheetSnapshot[] = [];
  private readonly MAX_HISTORY = 50;



  private pushHistory() {
    //  this.recalculateFormulas();

    const snap = this.takeSnapshot();

    // 🚨 Compare with last snapshot to avoid duplicates
    const last = this.history[this.history.length - 1];
    if (last && JSON.stringify(last) === JSON.stringify(snap)) {
      console.log('🛑 No change detected → skipping history push');
      return;
    }

    this.history.push(snap);
    if (this.history.length > this.MAX_HISTORY) this.history.shift();
    this.future = [];

    if (this.lastChangedRow !== null && this.lastChangedCol !== null) {
      console.log(`📝 History pushed: cell=${this.getColumnName(this.lastChangedCol)}${this.lastChangedRow + 1} value="${this.lastChangedValue}"`);
    } else {
      console.log('📝 History pushed: (no cell info)');
    }
  }

  private onGridScroll(e: UIEvent) {
    const target = e.target as HTMLElement;
    const scrollTop = target.scrollTop;
    const viewportHeight = target.clientHeight;

    const baseRowHeight = this.rowHeights[0] || this.DEFAULT_ROW_HEIGHT;

    const buffer = 5; // extra rows above/below for smoothness

    const start = Math.floor(scrollTop / baseRowHeight) - buffer;
    const end = Math.ceil((scrollTop + viewportHeight) / baseRowHeight) + buffer;

    this.visibleStartRow = Math.max(0, start);
    this.visibleEndRow = Math.min(this.sheet.length - 1, end);
  }

  private exportCSV() {
    const csv = CsvUtils.sheetToCSV(this.sheet);

    const blob = new Blob([csv], { type: 'text/csv;charset=utf-8;' });
    const url = URL.createObjectURL(blob);

    const a = document.createElement('a');
    a.href = url;
    a.download = 'spreadsheet.csv';
    a.click();

    URL.revokeObjectURL(url);
  }

  private importInput?: HTMLInputElement;

  private handleImportCSV(e: Event) {
    const input = e.target as HTMLInputElement;
    const file = input.files?.[0];
    if (!file) return;

    const reader = new FileReader();

    reader.onload = () => {
      this.pushHistory(); // keep undo

      const text = String(reader.result || '');
      this.sheet = CsvUtils.csvToSheet(text);

      const rows = this.sheet.length;
      const cols = rows > 0 ? this.sheet[0].length : 0;

      this.columnWidths = Array(cols).fill(96);
      this.rowHeights = Array(rows).fill(21);

      this.recalculateFormulas();
      this.selectedRow = 0;
      this.selectedCol = 0;

      const cell = this.sheet[0]?.[0];
      this.formulaDraft = cell?.formula ?? cell?.value ?? '';
    };

    reader.readAsText(file);
    input.value = ''; // reset file input
  }

  private clipboardValue: string | string[][] = '';

  copyCell() {
    const cell = this.sheet[this.selectedRow][this.selectedCol];
    this.clipboardValue = cell.formula ?? cell.value;
    console.log('📋 Copied:', this.clipboardValue);
  }



  pasteCell() {
    // nothing copied → nothing to paste
    if (!this.clipboardValue) return;

    // 🚫 If clipboard is multi-cell, ignore
    if (Array.isArray(this.clipboardValue)) {
      console.warn('⚠️ Cannot paste multi-cell range with single-cell paste shortcut.');
      return;
    }

    // ✔️ Single cell paste
    const value = this.clipboardValue;
    const row = this.selectedRow;
    const col = this.selectedCol;

    console.log(`📌 Pasting "${value}" into ${this.getColumnName(col)}${row + 1}`);

    this.updateCellValue(row, col, value);
  }

  clearCellValue(row: number, col: number) {
    this.updateCellValue(row, col, '');
  }

  get selectionRange() {
    const r1 = Math.min(this.rangeStartRow, this.rangeEndRow);
    const r2 = Math.max(this.rangeStartRow, this.rangeEndRow);
    const c1 = Math.min(this.rangeStartCol, this.rangeEndCol);
    const c2 = Math.max(this.rangeStartCol, this.rangeEndCol);

    return { r1, r2, c1, c2 };
  }

  copyCellRange() {
    const { r1, r2, c1, c2 } = this.selectionRange;

    let buffer = [];

    for (let r = r1; r <= r2; r++) {
      const row = [];
      for (let c = c1; c <= c2; c++) {
        const cell = this.sheet[r][c];
        row.push(cell.formula ?? cell.value ?? '');
      }
      buffer.push(row);
    }

    this.clipboardValue = buffer;
    console.log('📋 Copied range:', buffer);
  }

  pasteCellRange() {
    if (!this.clipboardValue) return;

    if (!Array.isArray(this.clipboardValue)) return;

    const startR = this.selectedRow;
    const startC = this.selectedCol;

    const data = this.clipboardValue; // 2D array

    this.pushHistory();

    for (let r = 0; r < data.length; r++) {
      for (let c = 0; c < data[r].length; c++) {
        const targetR = startR + r;
        const targetC = startC + c;

        if (targetR >= this.sheet.length || targetC >= this.sheet[0].length) continue;

        this.updateCellValue(targetR, targetC, data[r][c]);
      }
    }

    this.recalculateFormulas();
  }

  startFill(e: MouseEvent, row: number, col: number) {
    e.preventDefault();
    this.isFilling = true;
    this.fillStartRow = row;
    this.fillStartCol = col;

    window.addEventListener('mousemove', this.performFill);
    window.addEventListener('mouseup', this.stopFill);
  }
  performFill = (e: MouseEvent) => {
    if (!this.isFilling) return;

    const targetRow = this.selectedRow;
    const targetCol = this.selectedCol;

    const value = this.sheet[this.fillStartRow][this.fillStartCol].value;

    const start = Math.min(this.fillStartRow, targetRow);
    const end = Math.max(this.fillStartRow, targetRow);

    for (let r = start; r <= end; r++) {
      this.updateCellValue(r, this.fillStartCol, value);
    }
  };

  stopFill = () => {
    this.isFilling = false;
    window.removeEventListener('mousemove', this.performFill);
    window.removeEventListener('mouseup', this.stopFill);
  };

  undo() {
    console.log('🔙 UNDO pressed');

    if (this.history.length === 0) {
      console.log('⚠️ Not enough history');
      return;
    }

    // Save current state for redo
    this.future.push(this.takeSnapshot());

    const prev = this.history.pop();
    if (prev) {
      this.restoreSnapshot(prev);

      const restoredCell = prev.sheet[prev.selectedRow]?.[prev.selectedCol];
      console.log(`✔ Undo complete: restored ${this.getColumnName(prev.selectedCol)}${prev.selectedRow + 1} = "${restoredCell?.value ?? ''}"`);
      // CHANGE END
    }
  }

  redo() {
    console.log('🔁 REDO pressed');

    if (this.future.length === 0) {
      console.log('⚠️ redo: future is empty');
      return;
    }

    // Save current for undo again
    this.history.push(this.takeSnapshot());

    const next = this.future.pop();
    if (next) {
      this.restoreSnapshot(next);

      const restoredCell = next.sheet[next.selectedRow]?.[next.selectedCol];
      console.log(`✔ Redo complete: restored ${this.getColumnName(next.selectedCol)}${next.selectedRow + 1} = "${restoredCell?.value ?? ''}"`);
      // CHANGE END
    }
  }

  startColumnResize(e: MouseEvent, colIndex: number) {
    this.pushHistory();
    e.preventDefault();
    const startX = e.clientX;
    const startWidth = this.columnWidths[colIndex] ?? 96;

    const onMouseMove = (ev: MouseEvent) => {
      const diff = ev.clientX - startX;
      this.columnWidths[colIndex] = Math.max(40, startWidth + diff); // min width 40
      this.columnWidths = [...this.columnWidths]; // trigger rerender
    };

    const onMouseUp = () => {
      this.pushHistory();
      window.removeEventListener('mousemove', onMouseMove);
      window.removeEventListener('mouseup', onMouseUp);
    };

    window.addEventListener('mousemove', onMouseMove);
    window.addEventListener('mouseup', onMouseUp);
  }

  startRowResize(e: MouseEvent, rowIndex: number) {
    this.pushHistory();
    e.preventDefault();
    const startY = e.clientY;
    const startHeight = this.rowHeights[rowIndex] ?? 21;

    const onMouseMove = (ev: MouseEvent) => {
      const diff = ev.clientY - startY;
      this.rowHeights[rowIndex] = Math.max(16, startHeight + diff);
      this.rowHeights = [...this.rowHeights];
    };

    const onMouseUp = () => {
      this.pushHistory();
      window.removeEventListener('mousemove', onMouseMove);
      window.removeEventListener('mouseup', onMouseUp);
      // this.pushHistory();
    };

    window.addEventListener('mousemove', onMouseMove);
    window.addEventListener('mouseup', onMouseUp);
  }

  createInitialSheet(rows: number, cols: number): Sheet {
    const sheet: Sheet = [];

    for (let r = 0; r < rows; r++) {
      const row: Cell[] = [];
      for (let c = 0; c < cols; c++) {
        row.push({ value: '', style: {} }); // empty cell
      }
      sheet.push(row);
    }

    requestAnimationFrame(() => {
      // avoid accessing during render
      if (!this.columnWidths.length) this.columnWidths = Array(cols).fill(96);
      if (!this.rowHeights.length) this.rowHeights = Array(rows).fill(21);
    });

    return sheet;
  }

  getColumnName(index: number): string {
    let name = '';
    while (index >= 0) {
      name = String.fromCharCode((index % 26) + 65) + name;
      index = Math.floor(index / 26) - 1;
    }
    return name;
  }

  selectCell(row: number, col: number) {
    this.selectedRow = row;
    this.selectedCol = col;
    this.rangeStartRow = row;
    this.rangeEndRow = row;
    this.rangeStartCol = col;
    this.rangeEndCol = col;
  }

  updateSelectionRange(row: number, col: number) {
    this.rangeEndRow = row;
    this.rangeEndCol = col;
  }



  updateCellValue(row: number, col: number, value: string) {
    const prevValue = this.sheet[row][col].value;
    if (prevValue === value) return;

    // 1️⃣ Take snapshot BEFORE change
    // this.pushHistory();

    console.log('🟡 Snapshot BEFORE update →', this.history[this.history.length - 1]?.sheet[row][col], '(Expected old value)', '   New incoming value =', value);

    // 2️⃣ Apply change
    const sheetCopy = this.cloneSheet(this.sheet);
    const cell = sheetCopy[row][col];
    cell.value = value;
    cell.formula = value.startsWith('=') ? value : undefined;

    // 3️⃣ Set final grid
    this.sheet = sheetCopy;

    this.formulaDraft = this.sheet[row][col]?.formula ?? this.sheet[row][col]?.value ?? '';
  }

  getCellAddress(): string {
    if (this.selectedRow == null || this.selectedCol == null) return '';
    return `${this.getColumnName(this.selectedCol)}${this.selectedRow + 1}`;
  }

  getCellFormulaOrValue(): string {
    return this.formulaDraft;
  }



  onFormulaInput(value: string) {
    this.formulaDraft = value;
  }

  onCellInput(row: number, col: number, value: string) {
    const sheetCopy = [...this.sheet];
    const cell = { ...sheetCopy[row][col] };

    cell.value = value;
    cell.formula = value.startsWith('=') ? value : undefined;

    sheetCopy[row][col] = cell;
    this.sheet = sheetCopy;

    this.recalculateFormulas();
  }

 

  applyFormula() {
    const row = this.selectedRow;
    const col = this.selectedCol;
    const value = this.formulaDraft ?? '';

    console.log('📥 applyFormula →', { row, col, value, beforeA1: this.sheet[0]?.[0]?.value ?? '' });

    // 1️⃣ Save previous state for undo

    this.pushHistory();
    console.log('📥 Formula Snapshot BEFORE update', this.history[this.history.length - 1]?.sheet[row][col]);

    // 2️⃣ Work on a cloned sheet
    const sheetCopy = this.cloneSheet(this.sheet);
    const cell = sheetCopy[row][col];

    // 3️⃣ Apply value / formula
    if (value.startsWith('=')) {
      cell.formula = value;
    } else {
      cell.value = value;
      cell.formula = undefined;
    }

  
    this.sheet = sheetCopy;
    this.recalculateFormulas();
    console.log('📐 After applyFormula+recalc → A1 =', this.sheet[0]?.[0]?.value ?? '');
  }

  recalculateFormulas() {
    // const sheetCopy = this.sheet;
    const engine = new FormulaEngine();
    const sheetCopy = this.cloneSheet(this.sheet);

    for (let r = 0; r < sheetCopy.length; r++) {
      for (let c = 0; c < sheetCopy[r].length; c++) {
        const cell = sheetCopy[r][c];
        if (cell.formula) {
          try {
            const result = engine.evaluate(cell.formula, sheetCopy);
            cell.value = result != null ? String(result) : '#ERROR';
          } catch (e) {
            cell.value = '#ERROR';
          }
        }
      }
    }

    this.sheet = sheetCopy;
    // this.sheet = this.sheet.map(r => [...r]);
    this.sheet = [...this.sheet];
  }

  applyCellStyle(styleKey: keyof CellStyle, value?: string) {
    this.pushHistory();
    const sheetCopy = this.cloneSheet(this.sheet);
    const oldCell = sheetCopy[this.selectedRow][this.selectedCol];

    const newStyle = { ...(oldCell.style || {}) };

    if (styleKey === 'bold' || styleKey === 'italic' || styleKey === 'underline') {
      newStyle[styleKey] = !newStyle[styleKey];
    } else if (styleKey == 'align') {
      newStyle.align = value as 'left' | 'center' | 'right';
    } else {
      newStyle[styleKey] = value;
    }

    sheetCopy[this.selectedRow][this.selectedCol] = {
      ...oldCell,
      style: newStyle,
    };

    this.sheet = sheetCopy;
    // this.pushHistory();
  }


  render() {
    const baseRowHeight = this.rowHeights[0] || this.DEFAULT_ROW_HEIGHT;

    const startRow = this.visibleStartRow;
    const endRow = Math.min(this.visibleEndRow, this.sheet.length - 1);

    const topSpacerHeight = startRow * baseRowHeight;
    const bottomSpacerHeight = (this.sheet.length - endRow - 1) * baseRowHeight;

    return (
      <div class="spreadsheet-container">
        <div class="header">
          <div class="toolbar">
            <button onClick={() => this.undo()}>↺</button>
            <button onClick={() => this.redo()}>↻</button>
            <button onClick={() => this.applyCellStyle('bold')}>𝗕</button>
            <button onClick={() => this.applyCellStyle('italic')}>𝘐</button>
            <button onClick={() => this.applyCellStyle('underline')}>U</button>

            <div class="color-picker">
              <button class="icon-btn" onClick={() => this.textColorInput?.click()}>
                <i class="icon-text-color"></i>
              </button>
              <input
                type="color"
                ref={el => (this.textColorInput = el as HTMLInputElement)}
                onInput={e => this.applyCellStyle('color', (e.target as HTMLInputElement).value)}
                style={{ display: 'none' }}
              />
            </div>

            <div class="color-picker">
              <button class="icon-btn" onClick={() => this.bgColorInput?.click()}>
                <i class="icon-fill-color"></i>
              </button>
              <input
                type="color"
                ref={el => (this.bgColorInput = el as HTMLInputElement)}
                onInput={e => this.applyCellStyle('bgColor', (e.target as HTMLInputElement).value)}
                style={{ display: 'none' }}
              />
            </div>

            <select class="align-dropdown" onInput={e => this.applyCellStyle('align', (e.target as HTMLSelectElement).value)}>
              <option value="" disabled selected>
                Align
              </option>
              <option value="left">left</option>
              <option value="center">center</option>
              <option value="right">right</option>
            </select>

  

            <button class="csv-btn" onClick={() => this.importInput?.click()}>
              📥 Import
            </button>
            <button class="csv-btn" onClick={() => this.exportCSV()}>
              📤 Export
            </button>

            <input type="file" accept=".csv,text/csv" ref={el => (this.importInput = el as HTMLInputElement)} style={{ display: 'none' }} onChange={e => this.handleImportCSV(e)} />
          </div>

          <div class="formula-bar">
            <span>{this.getCellAddress()}</span>
            <input
              value={this.getCellFormulaOrValue()}
              onInput={e => this.onFormulaInput((e.target as HTMLInputElement).value)}
              onKeyDown={e => e.key === 'Enter' && this.applyFormula()}
            />
          </div>
        </div>

        {/* 🔹 Scrollable viewport */}
        <div class="grid-viewport" onScroll={e => this.onGridScroll(e)}>
          <div class="sheet">
            {/* 🔹 Column Header Row */}
            <div class="row header-row">
              <div class="corner-cell"></div>
              {this.sheet[0].map((_, colIndex) => (
                <div class="column-header" style={{ width: `${this.columnWidths[colIndex]}px` }}>
                  {this.getColumnName(colIndex)}
                  <div class="resize-handle-column " onMouseDown={e => this.startColumnResize(e, colIndex)}></div>
                </div>
              ))}
            </div>

            {/* 🔹 Spacer before visible rows */}
            <div class="spacer-top" style={{ height: `${topSpacerHeight}px` }}></div>

            {/* 🔹 Only render visible rows */}
            {this.sheet.slice(startRow, endRow + 1).map((row, idx) => {
              const rIndex = startRow + idx;
              return (
                <div class="row">
                  <div class="row-header" style={{ height: `${this.rowHeights[rIndex]}px` }}>
                    {rIndex + 1}
                    <div class="resize-handle-row" onMouseDown={e => this.startRowResize(e, rIndex)}></div>
                  </div>

                  {row.map((cell, cIndex) => (
                    <app-cell
                      key={`${rIndex}-${cIndex}-${cell.value}`}
                      cell={cell}
                      row={rIndex}
                      col={cIndex}
                      isInRange={rIndex >= this.selectionRange.r1 && rIndex <= this.selectionRange.r2 && cIndex >= this.selectionRange.c1 && cIndex <= this.selectionRange.c2}
                      onCellSelected={event => this.selectCell(event.detail.row, event.detail.col)}
                      isSelected={this.selectedRow === rIndex && this.selectedCol === cIndex}
                      style={{
                        '--cell-width': `${this.columnWidths[cIndex]}px`,
                        '--cell-height': `${this.rowHeights[rIndex]}px`,
                      }}
                    />
                  ))}
                </div>
              );
            })}

            {/* 🔹 Spacer after visible rows */}
            <div class="spacer-bottom" style={{ height: `${bottomSpacerHeight}px` }}></div>
          </div>
        </div>
      </div>
    );
  }
}
