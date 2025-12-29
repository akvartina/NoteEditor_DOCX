import { Injectable } from '@angular/core';
// @ts-ignore
import { renderAsync } from 'docx-preview';

@Injectable({
  providedIn: 'root'
})
export class DocxImportService {
  constructor() {}

  async convertDocxToHtml(file: File): Promise<string> {
    return new Promise((resolve, reject) => {
      const reader = new FileReader();
      reader.onload = async (e) => {
        try {
          const arrayBuffer = e.target?.result as ArrayBuffer;
          const container = document.createElement('div');
          await renderAsync(arrayBuffer, container);
          resolve(container.innerHTML);
        } catch (err) {
          reject(err);
        }
      };
      reader.onerror = (err) => reject(err);
      reader.readAsArrayBuffer(file);
    });
  }
}
