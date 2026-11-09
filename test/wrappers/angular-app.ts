// Angular wrapper test page (JIT): test/wrappers.mjs drives it through window.__wr.
import 'zone.js';
import '@angular/compiler';
import { ApplicationRef, Component, ViewChild } from '@angular/core';
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
  const run = (f: () => void) => { f(); appRef.tick(); };
  (window as any).__wr = {
    log,
    open: (bytes: Uint8Array) => run(() => { app.src = bytes; }),
    readOnly: (on: boolean) => run(() => { app.readOnly = on; }),
    save: () => app.editor!.api!.save(),
    unmount: () => run(() => { app.shown = false; }),
  };
});
