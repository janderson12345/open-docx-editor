// Restrict Editing in the editor, as Word enforces it: a document protected as read-only, for comments only or for
// tracked changes only, and "Limit formatting to a selection of styles". Edits the protection doesn't allow are
// refused (and said, as Word's status bar says it); the parts left editable for everyone or for this user (w:permStart
// .. w:permEnd) stay editable and are shaded, as Word highlights them. Forms protection is fields.js's forms mode
// (with its unprotected sections, openSectionAt below). The file side is ../ooxml/protect.js.
import { Plugin, PluginKey } from 'prosemirror-state';
import { Decoration, DecorationSet } from 'prosemirror-view';
import { AddMarkStep, RemoveMarkStep, ReplaceStep, ReplaceAroundStep, AttrStep } from 'prosemirror-transform';
import { schema } from './schema.js';
import { announce } from './live.js';
import { trackKey } from './track.js';
import { sectionEnd } from './sections.js';

export const protectKey = new PluginKey('odx-protect');

/** The protection the editor enforces: { edit, formatting, locked, editors } or null. */
export const protectionOf = (state) => protectKey.getState(state)?.cfg || null;

/** Comments can be added (not in a read-only document). */
export const commentsAllowed = (state) => protectionOf(state)?.edit !== 'readOnly';

/** Tracked changes can be accepted or rejected (not when the document is protected for tracked changes). */
export const resolvingAllowed = (state) => protectionOf(state)?.edit !== 'trackedChanges';

const permAttr = (xml, name) => new RegExp(`\\bw:${name}="([^"]*)"`).exec(xml)?.[1] ?? null;
const isMarker = (n, kinds) => (n.type === schema.nodes.opaque_inline || n.type === schema.nodes.opaque_block) && kinds.test(n.attrs.kind);

/**
 * The ranges this user may edit: w:permStart .. w:permEnd for everyone (w:edGrp="everyone") or for one of `editors`
 * (w:ed: a name, an e-mail address, DOMAIN\name).
 */
export function editableRanges(doc, editors = []) {
  const who = new Set(editors.filter(Boolean).flatMap((e) => { const s = String(e).toLowerCase(); return [s, s.split('\\').pop()]; }));
  const starts = new Map();
  const out = [];
  doc.descendants((n, pos) => {
    if (!isMarker(n, /^perm(Start|End)$/)) return true;
    const id = permAttr(n.attrs.xml, 'id');
    if (n.attrs.kind === 'permStart') {
      const grp = (permAttr(n.attrs.xml, 'edGrp') || '').toLowerCase();
      const ed = (permAttr(n.attrs.xml, 'ed') || '').toLowerCase();
      if (grp === 'everyone' || (ed && (who.has(ed) || who.has(ed.split('\\').pop())))) starts.set(id, pos + n.nodeSize);
    } else if (starts.has(id)) {
      out.push({ from: starts.get(id), to: pos });
      starts.delete(id);
    }
    return false;
  });
  return out;
}

/** True when a protected form leaves the section at `pos` open (its w:formProt w:val="false"). */
export function openSectionAt(doc, pos) {
  let xml = null;
  // (The section's w:sectPr: on the first section-ending paragraph from `pos` on, else the document's last.)
  doc.forEach((child, offset) => {
    if (xml != null || offset + child.nodeSize <= pos || child.type !== schema.nodes.paragraph) return;
    xml = sectionEnd(child);
  });
  return /<w:formProt\b[^>]*w:val="(?:0|false|off)"/.test(xml ?? doc.attrs.sectPr ?? '');
}

const FORMAT = new Set(['font', 'bold', 'italic', 'strike', 'color', 'size', 'highlight', 'underline', 'vertAlign', 'rextra']);
const COMMENT = /^comment(RangeStart|RangeEnd|Reference)$/;
const onlyMarkers = (fragment) => { let ok = true; fragment.forEach((n) => { if (!isMarker(n, COMMENT)) ok = false; }); return ok; };

let lastSaid = 0;
function refuse(msg) {
  if (Date.now() - lastSaid > 1500) { lastSaid = Date.now(); announce(msg); }
  return false;
}

const MESSAGE = {
  readOnly: 'This document is protected: it can only be read. The highlighted parts can be changed.',
  comments: 'This document is protected: only comments can be added. The highlighted parts can be changed.',
};

/** Whether step `step` (on `doc`) is a change the protection allows; else the reason. */
function stepAllowed(cfg, step, doc, ranges) {
  const from = step.from ?? step.pos;
  const to = step.to ?? (step.pos != null ? step.pos + 1 : null);
  if (cfg.formatting) {
    const locked = (id) => id != null && cfg.locked.has(id);
    if (step instanceof AddMarkStep || step instanceof RemoveMarkStep) {
      const m = step.mark;
      if (FORMAT.has(m.type.name) || (m.type === schema.marks.cstyle && locked(m.attrs.val))) return 'That formatting is locked in this document.';
    } else if (step instanceof AttrStep) {
      if (step.attr !== 'styleId' || locked(step.value)) return 'That formatting is locked in this document.';
    } else if (step instanceof ReplaceAroundStep) {
      const old = doc.nodeAt(step.from);
      const now = step.slice.content.firstChild;
      if (old && now && old.type === now.type) {
        const changed = Object.keys(now.attrs).filter((k) => now.attrs[k] !== old.attrs[k]);
        if (changed.some((k) => k !== 'styleId') || (changed.includes('styleId') && locked(now.attrs.styleId))) return 'That formatting is locked in this document.';
      }
    }
  }
  if (cfg.edit !== 'readOnly' && cfg.edit !== 'comments') return null;
  if (from != null && ranges.some((r) => from >= r.from && to <= r.to)) return null;
  if (cfg.edit === 'comments' && step instanceof ReplaceStep && onlyMarkers(doc.slice(step.from, step.to).content) && onlyMarkers(step.slice.content)) return null;
  return MESSAGE[cfg.edit];
}

/** The plugin; protection is set with setProtection. */
export function protectPlugin() {
  return new Plugin({
    key: protectKey,
    state: {
      init: () => ({ cfg: null, ranges: [], decos: null }),
      apply(tr, st, _old, state) {
        const m = tr.getMeta(protectKey);
        if (!m && !(tr.docChanged && st.cfg)) return st;
        const cfg = m ? m.cfg : st.cfg;
        if (!cfg || (cfg.edit !== 'readOnly' && cfg.edit !== 'comments')) return { cfg, ranges: [], decos: null };
        const ranges = editableRanges(state.doc, cfg.editors);
        const decos = ranges.length ? DecorationSet.create(state.doc, ranges.filter((r) => r.to > r.from).map((r) => Decoration.inline(r.from, r.to, { class: 'odx-perm' }))) : null;
        return { cfg, ranges, decos };
      },
    },
    filterTransaction(tr, state) {
      const cfg = protectionOf(state);
      if (!cfg || tr.getMeta(protectKey)) return true;
      // Track Changes stays on in a document protected for tracked changes, and its changes stay as they are.
      if (cfg.edit === 'trackedChanges' && tr.getMeta(trackKey)?.on === false) return refuse('Track Changes stays on: this document is protected for tracked changes.');
      if (cfg.edit === 'trackedChanges' && tr.getMeta('odx-resolve')) return refuse('Changes can\'t be accepted or rejected: this document is protected for tracked changes.');
      // (Others' edits arriving in a shared document, what follows from an edit already allowed, and the fields' results
      // as pages change, which Word updates in a protected document too.)
      // (Bold or a font picked for the next characters typed is direct formatting too.)
      if (cfg.formatting && tr.storedMarksSet && !tr.docChanged) {
        const fmt = (marks) => (marks || []).filter((m) => FORMAT.has(m.type.name));
        const was = fmt(state.storedMarks || state.selection.$from.marks()), now = fmt(tr.storedMarks);
        if (was.length !== now.length || now.some((m) => !m.isInSet(was))) return refuse('That formatting is locked in this document.');
      }
      if (!tr.docChanged || tr.getMeta('odx-remote') || tr.getMeta('y-sync$') || tr.getMeta('appendedTransaction') || tr.getMeta('odx-fields')) return true;
      const { ranges } = protectKey.getState(state);
      for (let i = 0; i < tr.steps.length; i++) {
        const why = stepAllowed(cfg, tr.steps[i], tr.docs[i], i ? editableRanges(tr.docs[i], cfg.editors) : ranges);
        if (why) return refuse(why);
      }
      return true;
    },
    props: {
      decorations: (state) => protectKey.getState(state)?.decos || null,
      attributes: (state) => { const c = protectionOf(state); return c?.edit ? { 'data-odx-protect': c.edit } : {}; },
    },
  });
}

/**
 * Enforces `cfg` ({ edit: 'readOnly' | 'comments' | 'trackedChanges' | null, formatting, locked: Set of style ids
 * and lower-case names, editors: this user's names }), or nothing with null. Tracked-changes protection turns
 * Track Changes on.
 */
export function setProtection(view, cfg) {
  const tr = view.state.tr.setMeta(protectKey, { cfg: cfg && (cfg.edit || cfg.formatting) ? cfg : null }).setMeta('addToHistory', false);
  if (cfg?.edit === 'trackedChanges' && !trackKey.getState(view.state)?.on) tr.setMeta(trackKey, { on: true });
  view.dispatch(tr);
}
