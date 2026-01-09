import { Injectable } from '@angular/core';
import { HttpClient } from '@angular/common/http';

export type XlsxSheet = { name: string; cells: string[][] };
export type ImportXlsxResponse = { ok: boolean; sheets: XlsxSheet[] };

@Injectable({ providedIn: 'root' })
export class XlsxApiService {
  // adjust if you use environment.ts
  private baseUrl = 'http://localhost:3000';

  constructor(private http: HttpClient) {}

  importXlsx(file: File) {
    const form = new FormData();
    form.append('file', file);
    return this.http.post<ImportXlsxResponse>(`${this.baseUrl}/import-xlsx`, form);
  }

  exportXlsx(sheets: XlsxSheet[]) {
    // backend returns a file (blob)
    return this.http.post(`${this.baseUrl}/export-xlsx`, { sheets }, { responseType: 'blob' });
  }
}
