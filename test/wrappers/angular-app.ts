// Angular wrapper test page (JIT): test/wrappers.mjs drives it through window.__wr.
import 'zone.js';
import '@angular/compiler';
import { ApplicationRef, ChangeDetectorRef, Component, ViewChild } from '@angular/core';
import { bootstrapApplication } from '@angular/platform-browser';
import { DocxEditorComponent } from '../../src/wrappers/angular/docx-editor.component.ts';

const log: string[] = [];

@Component({
  selector: 'odx-test-app',
  standalone: true,
  imports: [DocxEditorComponent],
  template: `@if (shown) {<odx-docx-editor [src]="src" [readOnly]="readOnly" fileName="form.docx" (opened)="log.push('opened ' + $event.fileName)" (ready)="log.push('ready')" />}`,
})
class TestApp {
  src: Uint8Array | null = null;
  readOnly = false;
  shown = true;
  log = log;
  @ViewChild(DocxEditorComponent) editor?: DocxEditorComponent;
}

bootstrapApplication(TestApp).then((ref) => {
  const app = ref.components[0].instance as TestApp;
  const appRef = ref.injector.get(ApplicationRef);
  // (Checked explicitly: from Angular 22 a component isn't refreshed on every tick unless marked, and these properties
  // are set from outside its template.)
  const cdr = ref.components[0].injector.get(ChangeDetectorRef);
  const run = (f: () => void) => { f(); cdr.markForCheck(); appRef.tick(); };
  (window as any).__wr = {
    log,
    open: (bytes: Uint8Array) => run(() => { app.src = bytes; }),
    readOnly: (on: boolean) => run(() => { app.readOnly = on; }),
    save: () => app.editor!.api!.save(),
    unmount: () => run(() => { app.shown = false; }),
  };
});
