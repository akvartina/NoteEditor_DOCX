import { Component } from '@angular/core';
import { XlsxApiService, XlsxSheet } from '../services/xlsx-api.service';

@Component({
  selector: 'app-xlsx-editor',
  templateUrl: './xlsx-editor.component.html',
  styleUrls: ['./xlsx-editor.component.scss'],
})
export class XlsxEditorComponent {
  sheets: XlsxSheet[] = [];
  activeSheetIndex = 0;
  loading = false;
  error: string | null = null;

  constructor(private api: XlsxApiService) {}

  get activeSheet(): XlsxSheet | null {
    return this.sheets[this.activeSheetIndex] ?? null;
  }

  get cells(): string[][] {
    return this.activeSheet?.cells ?? [];
  }

  onFileSelected(ev: Event) {
    const input = ev.target as HTMLInputElement;
    const file = input.files?.[0];
    if (!file) return;

    this.loading = true;
    this.error = null;

    this.api.importXlsx(file).subscribe({
      next: (res) => {
        if (!res.ok) throw new Error('Import failed');
        this.sheets = res.sheets || [];
        this.activeSheetIndex = 0;
        if (this.sheets[0] && this.sheets[0].cells.length === 0) {
          this.sheets[0].cells = [['']];
        }
        this.loading = false;
      },
      error: (e) => {
        this.loading = false;
        this.error = e?.message || 'Import failed';
      }
    });

    // allow selecting same file again
    input.value = '';
  }

  updateCell(r: number, c: number, value: string) {
    const sheet = this.sheets[this.activeSheetIndex];
    if (!sheet) return;

    // Ensure row exists
    while (sheet.cells.length <= r) sheet.cells.push([]);

    // Ensure col exists
    while (sheet.cells[r].length <= c) sheet.cells[r].push('');

    sheet.cells[r][c] = value;
  }

  export() {
    if (!this.sheets.length) return;

    this.loading = true;
    this.error = null;

    this.api.exportXlsx(this.sheets).subscribe({
      next: (blob) => {
        this.loading = false;
        const url = URL.createObjectURL(blob);
        const a = document.createElement('a');
        a.href = url;
        a.download = 'export.xlsx';
        a.click();
        URL.revokeObjectURL(url);
      },
      error: (e) => {
        this.loading = false;
        this.error = e?.message || 'Export failed';
      }
    });
  }

  addRow() {
    const sheet = this.activeSheet;
    if (!sheet) return;
    const cols = Math.max(...sheet.cells.map(r => r.length), 0);
    sheet.cells.push(Array(cols).fill(''));
  }

  addCol() {
    const sheet = this.activeSheet;
    if (!sheet) return;
    const rows = sheet.cells.length;
    if (rows === 0) sheet.cells.push(['']);
    else sheet.cells = sheet.cells.map(r => [...r, '']);
  }

  setActiveSheet(i: number) {
    this.activeSheetIndex = i;
  }
}
