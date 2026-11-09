import type * as Y from 'yjs';
export interface CollabUser { name: string; color?: string }
export interface Collaboration {
  readonly ydoc: Y.Doc;
  readonly awareness: unknown;
  readonly user: CollabUser;
  readonly bytes: Uint8Array | null;
  /** Settles once the room has synced (the `ready` option). */
  readonly ready: Promise<unknown>;
  publish(bytes: Uint8Array | ArrayBuffer, doc: unknown): void;
  seed(doc: unknown): void;
  onDocument(fn: (bytes: Uint8Array) => void): () => void;
  plugins(): unknown[];
  remote(tr: unknown): boolean;
}
/** A room for mountDocxEditor's `collab` option: `ydoc` synced by your Yjs provider, its `awareness` for cursors. */
/** `ready`: a promise that settles once your provider has synced the room; the editor waits for it before opening the room's document or starting one. */
export function collaboration(ydoc: Y.Doc, options?: { awareness?: unknown; user?: CollabUser; ready?: Promise<unknown> }): Collaboration;
export { Y };
