/** The editor's pages as a PDF, drawn from its own page layout. `editor` is mountDocxEditor's API or an editor view. */
export function exportPdf(editor: unknown, options?: {
  title?: string;
  author?: string;
  localFonts?: boolean;
  /** Where the Noto fonts for characters a run's font lacks load from (FALLBACK_CDN by default), or false for none. */
  fallbackFonts?: string | false;
}): Promise<Uint8Array>;

/** npm's @fontsource packages on jsDelivr, where the fallback fonts load from by default. */
export const FALLBACK_CDN: string;
