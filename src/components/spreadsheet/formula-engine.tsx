// import { Sheet } from './types';

// export class FormulaEngine {
//   constructor(private sheet: Sheet) {}

//   evaluate(formula: string): number | string {
//     if (!formula.startsWith('=')) return formula;

//     const expr = formula.substring(1).trim();

//     if (expr.startsWith('SUM')) return this.aggregate(expr, 'SUM');
//     if (expr.startsWith('AVG')) return this.aggregate(expr, 'AVG');
//     if (expr.startsWith('MIN')) return this.aggregate(expr, 'MIN');
//     if (expr.startsWith('MAX')) return this.aggregate(expr, 'MAX');

//     // If arithmetic formula like "=A1 + B1 * 2"
//     return this.calculateExpression(expr);
//   }

//   private aggregate(expr: string, type: 'SUM' | 'AVG' | 'MIN' | 'MAX'): number {
//     const range = expr.match(/\((.*)\)/)![1];
//     const values = this.getRangeValues(range);

//     console.log("Values in range:", values);


//     switch (type) {
//       case 'SUM': return values.reduce((a, b) => a + b, 0);
//       case 'AVG': return values.reduce((a, b) => a + b, 0) / values.length;
//       case 'MIN': return Math.min(...values);
//       case 'MAX': return Math.max(...values);
//     }
//   }

//   private getRangeValues(range: string): number[] {
//     const [start, end] = range.split(':');
//     const startPos = this.cellToIndices(start.trim());
//     const endPos = this.cellToIndices(end.trim());

//     let vals: number[] = [];
//     for (let r = startPos.row; r <= endPos.row; r++) {
//       for (let c = startPos.col; c <= endPos.col; c++) {
//         console.log("Range Evaluating:", range, "Start:", startPos, "End:", endPos);
//         console.log("Cell Value:", this.sheet[r][c].value);
//             let value = this.sheet[r][c].value;
//             if (!value) value = "0";
//             value = value.replace(/[^0-9.-]/g, '');  
//             const val = parseFloat(value) || 0;
//             vals.push(val);

//         // const val = parseFloat(this.sheet[r][c].value) || 0;
//         // vals.push(val);
//       }
//     }
//     return vals;
//   }

//   private cellToIndices(cell: string): { row: number; col: number } {
//     let col = 0;
//     let i = 0;

//     // Handle multi-letter columns (AA, AB...)
//     while (cell.charCodeAt(i) >= 65 && cell.charCodeAt(i) <= 90) {
//       col = col * 26 + (cell.charCodeAt(i) - 64);
//       i++;
//     }
//     col--;

//     const row = parseInt(cell.substring(i)) - 1;
//     return { row, col };
//   }

//   private calculateExpression(expr: string): number | string {
//     try {
//       const replacedExpr = expr.replace(/[A-Z]+[0-9]+/g, (match) => {
//         const { row, col } = this.cellToIndices(match);
//         let value = this.sheet[row][col].value;
//         if (!value) return '0';
//         value = value.replace(/[^0-9.-]/g, '');  // sanitize here too
//          return value || '0';

        
//         // return this.sheet[row][col].value || '0';
//       });

//       return Function(`return (${replacedExpr});`)();
//     } catch {
//       return '#ERROR';
//     }
//   }
// }


import { Sheet } from './types';

export class FormulaEngine {
  evaluate(formula: string, sheet: Sheet): number | string {
    if (!formula.startsWith('=')) return formula;

    const expr = formula.substring(1).trim();
    const upper = expr.toUpperCase();

    if (upper.startsWith('SUM')) return this.aggregate(upper, sheet, 'SUM');
    if (upper.startsWith('AVG')) return this.aggregate(upper, sheet, 'AVG');
    if (upper.startsWith('MIN')) return this.aggregate(upper, sheet, 'MIN');
    if (upper.startsWith('MAX')) return this.aggregate(upper, sheet, 'MAX');

    
    return this.calculateExpression(expr, sheet);
  }

  private aggregate(
    expr: string,
    sheet: Sheet,
    type: 'SUM' | 'AVG' | 'MIN' | 'MAX'
  ): number {
    const match = expr.match(/\((.*)\)/);
    if (!match) return 0;

    const range = match[1];
    const values = this.getRangeValues(sheet, range);
    console.log('Values in range:', values);

    if (!values.length) return 0;

    switch (type) {
      case 'SUM':
        return values.reduce((a, b) => a + b, 0);
      case 'AVG':
        return values.reduce((a, b) => a + b, 0) / values.length;
      case 'MIN':
        return Math.min(...values);
      case 'MAX':
        return Math.max(...values);
    }
  }

  private getRangeValues(sheet: Sheet, range: string): number[] {
    const [start, end] = range.split(':');
    const startPos = this.cellToIndices(start.trim());
    const endPos = this.cellToIndices(end.trim());

    const vals: number[] = [];

    for (let r = startPos.row; r <= endPos.row; r++) {
      for (let c = startPos.col; c <= endPos.col; c++) {
        console.log('Range Evaluating:', range, 'Start:', startPos, 'End:', endPos);

        const raw = sheet[r]?.[c]?.value ?? '';
        console.log('Cell Value:', raw);

        const cleaned = raw.toString().replace(/[^0-9.-]/g, '');
        const num = parseFloat(cleaned);

        vals.push(isNaN(num) ? 0 : num);
      }
    }

    return vals;
  }

  private cellToIndices(cell: string): { row: number; col: number } {
    let col = 0;
    let i = 0;

    const upper = cell.toUpperCase();

    while (
      upper.charCodeAt(i) >= 65 && // 'A'
      upper.charCodeAt(i) <= 90    // 'Z'
    ) {
      col = col * 26 + (upper.charCodeAt(i) - 64);
      i++;
    }
    col--;

    const row = parseInt(upper.substring(i), 10) - 1;
    return { row, col };
  }

  private calculateExpression(expr: string, sheet: Sheet): number | string {
    try {
      const replacedExpr = expr.replace(/[A-Z]+[0-9]+/gi, match => {
        const { row, col } = this.cellToIndices(match.toUpperCase());
        const raw = sheet[row]?.[col]?.value ?? '';
        const cleaned = raw.toString().replace(/[^0-9.-]/g, '');
        return cleaned || '0';
      });

      const result = Function(`return (${replacedExpr});`)();
      if (typeof result === 'number' && !isNaN(result)) return result;
      return '#ERROR';
    } catch {
      return '#ERROR';
    }
  }
}
