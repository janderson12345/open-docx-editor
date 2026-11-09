/** The editor's pages as a PDF, drawn from its own page layout. `editor` is mountDocxEditor's API or an editor view. */
export function exportPdf(editor: unknown, options?: { title?: string; author?: string; localFonts?: boolean }): Promise<Uint8Array>;
