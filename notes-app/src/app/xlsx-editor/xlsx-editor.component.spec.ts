import { ComponentFixture, TestBed } from '@angular/core/testing';

import { XlsxEditorComponent } from './xlsx-editor.component';

describe('XlsxEditorComponent', () => {
  let component: XlsxEditorComponent;
  let fixture: ComponentFixture<XlsxEditorComponent>;

  beforeEach(async () => {
    await TestBed.configureTestingModule({
      declarations: [ XlsxEditorComponent ]
    })
    .compileComponents();

    fixture = TestBed.createComponent(XlsxEditorComponent);
    component = fixture.componentInstance;
    fixture.detectChanges();
  });

  it('should create', () => {
    expect(component).toBeTruthy();
  });
});
