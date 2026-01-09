import { Component } from '@angular/core';
import { PptxApiService, PptxSlide } from '../services/pptx-api.service';

@Component({
  selector: 'app-pptx-editor',
  templateUrl: './pptx-editor.component.html',
  styleUrls: ['./pptx-editor.component.scss'],
})
export class PptxEditorComponent {
  slides: PptxSlide[] = [];
  loading = false;
  error: string | null = null;

  constructor(private api: PptxApiService) {}

  onFileSelected(ev: Event) {
    const input = ev.target as HTMLInputElement;
    const file = input.files?.[0];
    if (!file) return;

    this.loading = true;
    this.error = null;

    this.api.importPptx(file).subscribe({
      next: (res) => {
        if (!res.ok) throw new Error('Import failed');
        this.slides = (res.slides || []).map((s, i) => ({
          index: s.index ?? (i + 1),
          text: s.text ?? '',
        }));

        // Make sure there is at least one slide to edit
        if (this.slides.length === 0) this.slides = [{ index: 1, text: '' }];

        this.loading = false;
      },
      error: (e) => {
        this.loading = false;
        this.error = e?.message || 'Import failed';
      },
    });

    // allow selecting same file again
    input.value = '';
  }

  updateSlideText(i: number, value: string) {
    if (!this.slides[i]) return;
    this.slides[i].text = value;
  }

  addSlide() {
    const nextIndex = this.slides.length + 1;
    this.slides.push({ index: nextIndex, text: '' });
  }

  deleteSlide(i: number) {
    this.slides.splice(i, 1);
    // re-number indices (purely cosmetic)
    this.slides = this.slides.map((s, idx) => ({ ...s, index: idx + 1 }));
    if (this.slides.length === 0) this.slides = [{ index: 1, text: '' }];
  }

  export() {
    if (!this.slides.length) return;

    this.loading = true;
    this.error = null;

    this.api.exportPptx(this.slides).subscribe({
      next: (blob) => {
        this.loading = false;
        const url = URL.createObjectURL(blob);
        const a = document.createElement('a');
        a.href = url;
        a.download = 'export.pptx';
        a.click();
        URL.revokeObjectURL(url);
      },
      error: (e) => {
        this.loading = false;
        this.error = e?.message || 'Export failed';
      },
    });
  }
}
