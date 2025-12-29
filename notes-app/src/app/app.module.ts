import { NgModule } from '@angular/core';
import { BrowserModule } from '@angular/platform-browser';
import { FormsModule } from '@angular/forms';
import { HttpClientModule } from '@angular/common/http';
import { NgxEditorModule } from 'ngx-editor';

import { AppComponent } from './app.component';
import { NoteEditorComponent } from './note-editor/note-editor.component';

@NgModule({
  declarations: [
    AppComponent,
    NoteEditorComponent
  ],
  imports: [
    BrowserModule,
    FormsModule,
    HttpClientModule,
    NgxEditorModule
  ],
  bootstrap: [AppComponent]
})
export class AppModule {}
