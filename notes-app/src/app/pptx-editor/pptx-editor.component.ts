import { Component, ElementRef, ViewChild, AfterViewInit, OnDestroy } from '@angular/core';
import Konva from 'konva';
import { PptxObjectsApiService } from '../services/pptx-objects-api.service';
import { PptxSlide, PptxObject, PptxTextObject, PptxImageObject } from '../models/pptx.model';

@Component({
  selector: 'app-pptx-editor',
  templateUrl: './pptx-editor.component.html',
  styleUrls: ['./pptx-editor.component.scss'],
})
export class PptxEditorComponent implements AfterViewInit, OnDestroy {
  @ViewChild('stageHost', { static: false }) stageHost?: ElementRef<HTMLDivElement>;

  deckId: string | null = null;
  slides: PptxSlide[] = [];
  activeSlideIndex = 0;

  loading = false;
  error: string | null = null;

  // PPTX wide defaults (inches)
  slideWIn = 13.333;
  slideHIn = 7.5;

  // render width in px
  viewportWpx = 1000;

  private stage?: Konva.Stage;
  private layer?: Konva.Layer;

  constructor(private api: PptxObjectsApiService) {}

  ngAfterViewInit(): void {
    // stage created after we have a host + slide to render
  }

  ngOnDestroy(): void {
    this.destroyStage();
  }

  get activeSlide(): PptxSlide | null {
    return this.slides[this.activeSlideIndex] ?? null;
  }

  get scalePxPerIn(): number {
    return this.viewportWpx / this.slideWIn;
  }

  get viewportHpx(): number {
    return Math.round(this.slideHIn * this.scalePxPerIn);
  }

  onFileSelected(ev: Event) {
    const input = ev.target as HTMLInputElement;
    const file = input.files?.[0];
    if (!file) return;

    this.loading = true;
    this.error = null;

    this.api.importPptx(file).subscribe({
      next: (res) => {
        if (!res.ok) throw new Error('Import failed');
        this.deckId = res.deckId;
        this.slides = res.slides || [];
        this.activeSlideIndex = 0;
        this.loading = false;

        // render first slide
        queueMicrotask(() => this.render());
      },
      error: (e) => {
        this.loading = false;
        this.error = e?.message || 'Import failed';
      }
    });

    input.value = '';
  }

  setActiveSlide(i: number) {
    this.activeSlideIndex = i;
    this.render();
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
      }
    });
  }

  // ---------------------------
  // Konva rendering + editing
  // ---------------------------

  private destroyStage() {
    try {
      this.stage?.destroy();
    } catch {}
    this.stage = undefined;
    this.layer = undefined;
  }

  render() {
    const host = this.stageHost?.nativeElement;
    const slide = this.activeSlide;
    if (!host || !slide) return;

    // rebuild stage for simplicity (MVP)
    this.destroyStage();
    host.innerHTML = '';

    this.stage = new Konva.Stage({
      container: host,
      width: this.viewportWpx,
      height: this.viewportHpx,
    });

    this.layer = new Konva.Layer();
    this.stage.add(this.layer);

    // background
    const bg = new Konva.Rect({
      x: 0, y: 0,
      width: this.viewportWpx,
      height: this.viewportHpx,
      fill: 'white',
      stroke: '#ddd',
      strokeWidth: 1
    });
    this.layer.add(bg);

    // render objects in z order
    const objects = (slide.objects || []).slice().sort((a,b) => (a.z ?? 0) - (b.z ?? 0));
    for (const obj of objects) {
      if (obj.type === 'text') this.addTextNode(slide, obj);
      if (obj.type === 'image') this.addImageNode(slide, obj);
    }

    this.layer.draw();
  }

  private addTextNode(slide: PptxSlide, obj: PptxTextObject) {
    const s = this.scalePxPerIn;

    const node = new Konva.Text({
      x: obj.x * s,
      y: obj.y * s,
      width: obj.w * s,
      height: obj.h * s,
      text: obj.text || '',
      fontSize: (obj.fontSize ?? 18) * (s / 96), // rough, ok for MVP
      draggable: true,
      fill: '#222',
    });

    node.on('dragend', () => {
      const pos = node.position();
      obj.x = pos.x / s;
      obj.y = pos.y / s;
    });

    node.on('dblclick', () => {
      const next = window.prompt('Edit text:', obj.text ?? '');
      if (next === null) return;
      obj.text = next;
      node.text(next);
      this.layer?.draw();
    });

    this.layer?.add(node);
  }

  private addImageNode(slide: PptxSlide, obj: PptxImageObject) {
    const s = this.scalePxPerIn;

    const imageObj = new window.Image();
    imageObj.crossOrigin = 'anonymous';
    imageObj.onload = () => {
      const node = new Konva.Image({
        x: obj.x * s,
        y: obj.y * s,
        width: obj.w * s,
        height: obj.h * s,
        image: imageObj,
        draggable: true,
      });

      node.on('dragend', () => {
        const pos = node.position();
        obj.x = pos.x / s;
        obj.y = pos.y / s;
      });

      this.layer?.add(node);
      this.layer?.draw();
    };

    imageObj.src = `http://localhost:3000${obj.src}`;
  }
}
