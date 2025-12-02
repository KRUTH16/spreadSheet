import { Sheet, Cell } from './types'; // Adjust path if needed

export class CsvUtils {
  static sheetToCSV(sheet: Sheet): string {
    return sheet
      .map(row =>
        row
          .map(cell => {
            const raw = cell.formula ?? cell.value ?? '';
            if (raw.includes('"') || raw.includes(',') || raw.includes('\n')) {
              const escaped = raw.replace(/"/g, '""');
              return `"${escaped}"`;
            }
            return raw;
          })
          .join(','),
      )
      .join('\n');
  }

  static csvToSheet(text: string): Sheet {
    const rows = text.split(/\r?\n/).filter(line => line.trim().length > 0);
    const sheet: Sheet = [];

    for (const rowText of rows) {
      const parts = CsvUtils.parseCSVLine(rowText);
      const cells: Cell[] = parts.map(part => {
        const trimmed = part.trim();
        return {
          value: trimmed,
          formula: trimmed.startsWith('=') ? trimmed : undefined,
        };
      });
      sheet.push(cells);
    }

    return sheet;
  }

  private static parseCSVLine(line: string): string[] {
    const result: string[] = [];
    let current = '';
    let inQuotes = false;

    for (let i = 0; i < line.length; i++) {
      const ch = line[i];
      if (inQuotes) {
        if (ch === '"') {
          if (line[i + 1] === '"') {
            current += '"';
            i++;
          } else {
            inQuotes = false;
          }
        } else {
          current += ch;
        }
      } else {
        if (ch === '"') {
          inQuotes = true;
        } else if (ch === ',') {
          result.push(current);
          current = '';
        } else {
          current += ch;
        }
      }
    }
    result.push(current);
    return result;
  }
}
