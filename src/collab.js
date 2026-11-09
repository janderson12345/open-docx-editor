// Real-time co-editing over Yjs (MIT): import 'open-docx-editor/collab'. The document body is shared through
// y-prosemirror, and so is the .docx it came from (its styles, numbering, headers and footers, pictures), so every
// editor in the room builds the same pages. Bring your own Yjs provider (y-websocket, y-webrtc...) for the transport,
// and its awareness for others' cursors.
//
// What's shared: the body, including tracked changes; the comments (bindComments); headers, footers and notes as each
// is closed after editing, header and footer parts and pictures one editor adds (bindStories), and the document's own
// section settings (bindDoc). Styles and numbering definitions added after opening aren't: they're part of the .docx
// each editor opened.
import * as Y from 'yjs';
import { ySyncPlugin, ySyncPluginKey, yCursorPlugin, yUndoPlugin, undo, redo, prosemirrorToYXmlFragment } from 'y-prosemirror';

/**
 * A room for mountDocxEditor's `collab` option. `ydoc` is the Y.Doc your provider syncs; `awareness` its awareness
 * (for cursors); `user` { name, color } names this editor's cursor. `ready`: a promise that settles once the provider
 * has synced the room (y-websocket: `new Promise((r) => provider.once('synced', r))`). Until then the room looks empty,
 * and the editor waits for it before opening the room's document or, in a room that has none, starting a new one;
 * without it the editor decides at once, which a room that syncs later would see as a blank document published over its own.
 */
export function collaboration(ydoc, { awareness = null, user: who = { name: 'Author', color: '#2b6cdf' }, ready = null } = {}) {
  // (Cursor colours as six-digit hex, the only form y-prosemirror draws.)
  const user = { ...who, color: /^#[0-9a-f]{3}$/i.test(who.color || '') ? '#' + [...who.color.slice(1)].map((c) => c + c).join('') : who.color || '#2b6cdf' };
  const fragment = ydoc.getXmlFragment('odx-body');
  const files = ydoc.getMap('odx-files');
  const comments = ydoc.getMap('odx-comments'); // comment id -> { author, date, text, parentId, done... }
  const stories = ydoc.getMap('odx-stories'); // "story:<rId>" or "note:<kind>:<id>" -> its document's JSON
  const docAttrs = ydoc.getMap('odx-doc'); // the document node's attributes: sectPr
  if (awareness) awareness.setLocalStateField('user', user);
  return {
    ydoc,
    awareness,
    user,
    /** Settles once the room has synced (the `ready` option), at once without it. */
    ready: Promise.resolve(ready),
    /** The .docx the room is editing, or null before anyone has opened one. */
    get bytes() { return files.get('docx') || null; },
    /** Make `bytes` the room's document, its body `doc` (a ProseMirror document): everyone reopens it. */
    publish(bytes, doc) {
      ydoc.transact(() => {
        fragment.delete(0, fragment.length);
        comments.clear(); // (the new document's own come with its editor: bindComments)
        stories.clear(); // (and its headers, footers and notes are the file's)
        docAttrs.set('sectPr', doc.attrs.sectPr ?? null);
        prosemirrorToYXmlFragment(doc, fragment);
        files.set('docx', bytes instanceof Uint8Array ? bytes : new Uint8Array(bytes));
      }, 'odx-publish');
    },
    /** Give the room's body `doc` if it has none yet (the first editor in it). */
    seed(doc) {
      if (fragment.length) return;
      ydoc.transact(() => { prosemirrorToYXmlFragment(doc, fragment); docAttrs.set('sectPr', doc.attrs.sectPr ?? null); }, 'odx-seed');
    },
    /**
     * Shares the comments (their text, author, date, replies, resolved state) through the room: the room's replace
     * `session.comments`, or, in a room with none yet, this editor's become the room's. `onRemote()` runs when someone
     * else changes them. Returns { push(session), stop() }: push after this editor changes them.
     */
    bindComments(session, onRemote) {
      const plain = (c) => ({ id: c.id, author: c.author, initials: c.initials, date: c.date, text: c.text, paraId: c.paraId, parentId: c.parentId, done: !!c.done, xml: c.xml || null });
      // (In date order, then id: the same order in every editor, whatever order the map gives.)
      const all = () => [...comments.values()].map((c) => ({ ...c })).sort((a, b) => String(a.date).localeCompare(String(b.date)) || String(a.id).localeCompare(String(b.id)));
      if (!comments.size && session.comments.length) ydoc.transact(() => session.comments.forEach((c) => comments.set(String(c.id), plain(c))), 'odx-seed');
      else if (comments.size) session.comments = all();
      const watch = (e, tr) => {
        if (tr.local) return;
        session.comments = all();
        session.commentsDirty = true;
        onRemote();
      };
      comments.observe(watch);
      return {
        push(s) {
          ydoc.transact(() => {
            const ids = new Set(s.comments.map((c) => String(c.id)));
            for (const k of [...comments.keys()]) if (!ids.has(k)) comments.delete(k);
            for (const c of s.comments) {
              const p = plain(c);
              if (JSON.stringify(comments.get(String(c.id))) !== JSON.stringify(p)) comments.set(String(c.id), p);
            }
          }, 'odx-comments');
        },
        stop() { comments.unobserve(watch); },
      };
    },
    /**
     * Shares headers, footers and notes as they're closed after editing: `fromJSON` makes a story's ProseMirror
     * document from its JSON, `onRemote()` runs when someone else's arrive. Returns { story(rId, doc), note(kind, id,
     * doc), stop() }. Header and footer parts this editor adds (session.addStory) go to the room as they're made.
     */
    bindStories(session, { fromJSON, onRemote }) {
      // Relationship ids this editor makes from now on are its own (DocxPackage.addRel).
      session.pkg.idPrefix = `rIdc${ydoc.clientID}n`;
      const apply = (k, v) => {
        if (k.startsWith('part:')) session.addSharedStory(v);
        else if (k.startsWith('media:')) session.addSharedImage(v);
        else if (k.startsWith('story:')) {
          const rId = k.slice(6);
          if (session.pkg.relById(rId)) { session.story(rId); session.setStory(rId, fromJSON(v)); } // (story() first: it sets up the session's stories)
        } else {
          const [, kind, id] = k.split(':');
          const map = session.notes?.[kind];
          if (!map) return;
          const note = map.get(id);
          if (note) { note.doc = fromJSON(v); note.dirty = true; }
          else map.set(id, { id, type: 'normal', xml: null, doc: fromJSON(v), dirty: true });
        }
      };
      // (Parts first: a header's text needs its part.)
      const first = (k) => (k.startsWith('part:') || k.startsWith('media:') ? 1 : 0);
      const order = (keys) => [...keys].sort((x, y) => first(y) - first(x));
      for (const k of order(stories.keys())) apply(k, stories.get(k)); // (the room's, as they were when this editor joined)
      const watch = (e, tr) => {
        if (tr.local) return;
        for (const k of order(e.keysChanged)) if (stories.has(k)) apply(k, stories.get(k));
        onRemote();
      };
      stories.observe(watch);
      // Header and footer parts this editor adds (Different First Page, unlinking a section's) go to the room too.
      session.onAddStory = (part) => stories.set(`part:${part.rId}`, part);
      // And pictures inserted here: their files, under the same rIds, so everyone's saved .docx has them.
      session.onAddImage = (pic) => stories.set(`media:${pic.rId}`, pic);
      return {
        story: (rId, doc) => stories.set(`story:${rId}`, doc.toJSON()),
        note: (kind, id, doc) => stories.set(`note:${kind}:${id}`, doc.toJSON()),
        stop() { stories.unobserve(watch); session.onAddStory = null; session.onAddImage = null; },
      };
    },
    /**
     * Shares the document node's own attributes (its last section's w:sectPr: page size, margins, columns, header and
     * footer references), which the shared body (its content) doesn't carry. The room's are applied to `view` now and
     * as they change, in transactions marked 'odx-remote'. Returns { push(attrs), stop() }: push after a local change.
     */
    bindDoc(view) {
      const take = () => {
        const v = docAttrs.get('sectPr');
        if (v === undefined || v === view.state.doc.attrs.sectPr) return;
        view.dispatch(view.state.tr.setDocAttribute('sectPr', v).setMeta('odx-remote', true).setMeta('addToHistory', false));
      };
      take();
      // (After the shared body has taken the same update: a view changed before it would write its old body back to the
      // room, as y-prosemirror sends a view's whole document when it changes. Not into a view since destroyed.)
      const watch = (e, tr) => { if (!tr.local) setTimeout(() => { if (!view.isDestroyed) take(); }); };
      docAttrs.observe(watch);
      return {
        push(attrs) { if (docAttrs.get('sectPr') !== attrs.sectPr) docAttrs.set('sectPr', attrs.sectPr); },
        stop() { docAttrs.unobserve(watch); },
      };
    },
    /** Calls `fn(bytes)` when someone else opens a document in the room; returns a function to stop. */
    onDocument(fn) {
      const watch = (e, tr) => { if (!tr.local && e.keysChanged.has('docx')) fn(files.get('docx')); };
      files.observe(watch);
      return () => files.unobserve(watch);
    },
    /** The editor plugins: the shared body, others' cursors, and an undo of this editor's own edits only. */
    plugins() {
      return [ySyncPlugin(fragment), ...(awareness ? [yCursorPlugin(awareness)] : []), yUndoPlugin()];
    },
    /** Whether a transaction brings someone else's edits (not tracked again as this editor's). */
    remote: (tr) => !!tr.getMeta(ySyncPluginKey)?.isChangeOrigin,
    undo,
    redo,
  };
}

export { Y };
