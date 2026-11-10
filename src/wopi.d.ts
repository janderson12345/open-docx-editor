import type { DocxEditorApi, MountOptions } from './index';

/** What a WOPI host's CheckFileInfo says about the file and this user (the fields the editor reads, and any others). */
export interface WopiFileInfo {
  BaseFileName: string;
  Size?: number;
  Version?: string;
  UserFriendlyName?: string;
  UserCanWrite?: boolean;
  ReadOnly?: boolean;
  SupportsLocks?: boolean;
  SupportsUpdate?: boolean;
  [key: string]: unknown;
}

/** A WOPI failure: the host's HTTP status and, for a lock conflict, the lock the file is held under. */
export class WopiError extends Error {
  status: number;
  lock: string | null;
  reason: string | null;
}

export interface WopiDocument {
  info: WopiFileInfo;
  bytes: Uint8Array;
  /** The user may edit it and, where the host locks files, this editor holds the lock. */
  canWrite: boolean;
  readonly locked: boolean;
  /** Someone else's lock when the file was already locked ('' when the host didn't say whose), else null. */
  lockedBy: string | null;
  /** PutFile, under the lock. Resolves with the host's new version, if it gives one. */
  save(bytes: Uint8Array): Promise<string | null>;
  refreshLock(): Promise<void>;
  /** Takes the lock again after close() (a page come back from the back/forward cache). */
  reopen(): Promise<void>;
  /** Unlock. */
  close(options?: { keepalive?: boolean }): Promise<void>;
}

/** CheckFileInfo, GetFile and Lock for a WOPI host's file. `src` is its WOPI URL (https, or http on localhost). */
export function openWopi(src: string, accessToken: string, options?: { fetch?: typeof fetch; lock?: boolean; lockRefreshMs?: number; onLockError?: (error: WopiError) => void }): Promise<WopiDocument>;

export interface WopiEditorOptions extends MountOptions {
  src: string;
  accessToken: string;
  /** Saved this long after editing stops (ms; 2000). 0: only with Ctrl+S or save(). */
  autosaveMs?: number;
  onSaved?: (version: string | null) => void;
  onError?: (error: Error) => void;
  fetch?: typeof fetch;
  lockRefreshMs?: number;
}

export interface WopiEditor {
  editor: DocxEditorApi;
  wopi: WopiDocument;
  save(): Promise<string | null>;
  /** Saves what's unsaved, releases the lock and removes the editor. */
  close(): Promise<void>;
}

/** A WOPI host's file in an editor: locked while open, saved back as you go, read-only when it must be. */
export function mountWopiEditor(container: HTMLElement, options: WopiEditorOptions): Promise<WopiEditor>;
