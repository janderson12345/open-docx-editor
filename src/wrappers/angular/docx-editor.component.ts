// Angular wrapper: a standalone component, <odx-docx-editor [src]="file" (changed)="..." #editor />. TypeScript source,
// compiled by the Angular app that imports it (Angular 15+). The editor is mounted once into the component's element and
// kept: inputs that change after mounting are applied through the editor's API (readOnly, features, forms), a new `src`
// opens that document, and destroying the component destroys the editor.
import { AfterViewInit, Component, ElementRef, EventEmitter, Input, NgZone, OnChanges, OnDestroy, Output, SimpleChanges, ViewChild, inject } from '@angular/core';
import { mountDocxEditor } from '../../index.js';

/** The editor's API (mountDocxEditor's): save(), download(), getFieldValues(), view... */
export type DocxEditorApi = ReturnType<typeof mountDocxEditor>;

@Component({
  selector: 'odx-docx-editor',
  standalone: true,
  template: '<div #host></div>',
})
export class DocxEditorComponent implements AfterViewInit, OnChanges, OnDestroy {
  /** Bytes, File, Blob or Response to open; a new value opens it. */
  @Input() src: Uint8Array | ArrayBuffer | Blob | File | Response | null = null;
  @Input() fileName = 'document.docx';
  @Input() author = 'Author';
  @Input() readOnly = false;
  @Input() features: Record<string, boolean> | null = null;
  @Input() toolbar = true;
  @Input() forms = false;
  /** Replaces the download when the user saves. */
  @Input() save: ((bytes: Uint8Array, fileName: string) => void) | null = null;
  @Output() changed = new EventEmitter<unknown>();
  @Output() opened = new EventEmitter<{ timings: unknown; fileName: string }>();
  @Output() failed = new EventEmitter<unknown>();
  @Output() ready = new EventEmitter<DocxEditorApi>();
  @ViewChild('host', { static: true }) host!: ElementRef<HTMLDivElement>;

  /** The editor's API, once the view is ready. */
  api: DocxEditorApi | null = null;

  private zone = inject(NgZone);

  ngAfterViewInit(): void {
    // Outside Angular's zone: the editor's own events (typing, layout frames) don't run change detection; ours do.
    this.zone.runOutsideAngular(() => {
      this.api = mountDocxEditor(this.host.nativeElement, {
        fileName: this.fileName, author: this.author, readOnly: this.readOnly, features: this.features || {}, toolbar: this.toolbar, forms: this.forms,
        onChange: (view: unknown) => this.zone.run(() => this.changed.emit(view)),
        onSave: this.save ? (bytes: Uint8Array, name: string) => this.zone.run(() => this.save!(bytes, name)) : undefined,
        onOpened: (timings: unknown, name: string) => this.zone.run(() => this.opened.emit({ timings, fileName: name })),
        onError: (e: unknown) => this.zone.run(() => this.failed.emit(e)),
      });
      if (this.src != null) this.api.open(this.src, this.fileName).catch((e: unknown) => this.zone.run(() => this.failed.emit(e)));
      else this.api.newDocument(); // (no document given: a blank one, as Word's New)
    });
    this.ready.emit(this.api!);
  }

  ngOnChanges(changes: SimpleChanges): void {
    const api = this.api;
    if (!api) return; // (the first values are passed when the editor is mounted)
    if (changes['src'] && this.src != null) api.open(this.src, this.fileName).catch((e: unknown) => this.failed.emit(e));
    if (changes['readOnly']) api.setReadOnly(this.readOnly);
    if (changes['forms']) api.setForms(this.forms);
    if (changes['features'] && this.features) api.setFeatures(this.features);
  }

  ngOnDestroy(): void {
    this.api?.destroy();
    this.api = null;
  }
}
