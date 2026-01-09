import { Injectable } from '@angular/core';
import { HttpClient } from '@angular/common/http';

export type PptxSlide = { index?: number; text: string };
export type ImportPptxResponse = { ok: boolean; slides: PptxSlide[] };

@Injectable({ providedIn: 'root' })
export class PptxApiService {
  private baseUrl = 'http://localhost:3000';

  constructor(private http: HttpClient) {}

  importPptx(file: File) {
    const form = new FormData();
    form.append('file', file);
    return this.http.post<ImportPptxResponse>(`${this.baseUrl}/import-pptx`, form);
  }

  exportPptx(slides: PptxSlide[]) {
    return this.http.post(`${this.baseUrl}/export-pptx`, { slides }, { responseType: 'blob' });
  }
}
