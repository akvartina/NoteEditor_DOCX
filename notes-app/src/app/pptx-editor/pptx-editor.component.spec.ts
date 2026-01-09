import { ComponentFixture, TestBed } from '@angular/core/testing';

import { PptxEditorComponent } from './pptx-editor.component';

describe('PptxEditorComponent', () => {
  let component: PptxEditorComponent;
  let fixture: ComponentFixture<PptxEditorComponent>;

  beforeEach(async () => {
    await TestBed.configureTestingModule({
      declarations: [ PptxEditorComponent ]
    })
    .compileComponents();

    fixture = TestBed.createComponent(PptxEditorComponent);
    component = fixture.componentInstance;
    fixture.detectChanges();
  });

  it('should create', () => {
    expect(component).toBeTruthy();
  });
});
