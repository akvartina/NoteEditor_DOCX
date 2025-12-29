import { Component, OnDestroy } from '@angular/core';
import { Editor } from 'ngx-editor';
import { HttpClient } from '@angular/common/http';

@Component({
  selector: 'app-note-editor',
  templateUrl: './note-editor.component.html',
  styleUrls: ['./note-editor.component.scss']
})
export class NoteEditorComponent implements OnDestroy {
  editor = new Editor();
  htmlContent = '';

  constructor(private http: HttpClient) {}

  // ========================
  // EXPORT TO DOCX
  // ========================
  exportToDocx() {
    this.http
      .post(
        'http://localhost:3000/export-docx',
        { html: this.htmlContent },
        { responseType: 'blob' }
      )
      .subscribe((blob) => {
        const url = window.URL.createObjectURL(blob);
        const a = document.createElement('a');
        a.href = url;
        a.download = 'note.docx';
        a.click();
        window.URL.revokeObjectURL(url);
      });
  }

  // ========================
  // IMPORT FROM DOCX
  // ========================
  importFromDocx(event: Event) {

    const input = event.target as HTMLInputElement;
    if (!input.files?.length) return;

    const file = input.files[0];
    const formData = new FormData();
    formData.append('file', file);

    this.http
      .post<{ html: string }>('http://localhost:3000/import-docx', formData)
      .subscribe({
        next: (res) => {
          this.htmlContent = res.html;
        },
        error: (err) => console.error('DOCX import failed:', err)
      });

    input.value = ''; // reset input

    /*
    const file = (event.target as HTMLInputElement).files?.[0];
    if (!file) return;

    const formData = new FormData();
    formData.append('file', file);

    this.http
      .post<{ html: string }>(
        'http://localhost:3000/import-docx',
        formData
      )
      .subscribe((res) => {
        this.htmlContent = this.postprocessImportedHtml(res.html);
      }); */
  }

  /*
  constructor(private http: HttpClient) {}

  exportToDocx() {
    this.http.post(
      'http://localhost:3000/export-docx',
      { html: this.htmlContent },
      { responseType: 'blob' }
    ).subscribe(blob => {
      const url = window.URL.createObjectURL(blob);
      const a = document.createElement('a');
      a.href = url;
      a.download = 'note.docx';
      a.click();
      window.URL.revokeObjectURL(url);
    });
  }

  importFromDocx(event: Event) {
    const input = event.target as HTMLInputElement;
    if (!input.files?.length) return;

    const formData = new FormData();
    formData.append('file', input.files[0]);

    this.http.post<{ html: string }>(
      'http://localhost:3000/import-docx',
      formData
    ).subscribe(res => {
      this.htmlContent = res.html;
    });
  }

   */

  ngOnDestroy() {
    this.editor.destroy();
  }
}
