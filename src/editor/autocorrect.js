// AutoCorrect as you type, Word's everyday set: straight quotes to curly ones, (c) (r) (tm) to © ® ™, three dots to an
// ellipsis, --> <-- ==> to arrows, a double hyphen between words to an em dash and a spaced hyphen to an en dash,
// 1/2 1/4 3/4 to fractions, a sentence's first word and a lone "i" capitalised, and common misspellings. Each
// correction is its own undo step after the character that made it, so Ctrl+Z gives back exactly what was typed, as in
// Word. (Not when co-editing: the shared undo there takes the character and its correction together.) Off with
// features: { autocorrect: false }.
import { Plugin } from 'prosemirror-state';
import { closeHistory } from 'prosemirror-history';
import { announce } from './live.js';

const SYMBOLS = [['(c)', '©'], ['(r)', '®'], ['(tm)', '™']];
const ARROWS = [['-->', '→'], ['<--', '←'], ['==>', '⇒'], ['<==', '⇐']];
const FRACTIONS = { '1/2': '½', '1/4': '¼', '3/4': '¾' };
// Common misspellings (lower case; a capitalised one is corrected with its capital).
const TYPOS = {
  teh: 'the', adn: 'and', taht: 'that', hte: 'the', recieve: 'receive', recieved: 'received', seperate: 'separate',
  definately: 'definitely', occured: 'occurred', occurence: 'occurrence', untill: 'until', wich: 'which',
  becuase: 'because', thier: 'their', alot: 'a lot', tommorow: 'tomorrow', tomorow: 'tomorrow', accomodate: 'accommodate',
  beleive: 'believe', freind: 'friend', goverment: 'government', wierd: 'weird', acheive: 'achieve',
  neccessary: 'necessary', existance: 'existence', independant: 'independent', publically: 'publicly',
  recomend: 'recommend', calender: 'calendar', enviroment: 'environment', foward: 'forward', begining: 'beginning',
  dont: 'don’t', didnt: 'didn’t', doesnt: 'doesn’t', isnt: 'isn’t', wasnt: 'wasn’t', couldnt: 'couldn’t',
  shouldnt: 'shouldn’t', wouldnt: 'wouldn’t', youre: 'you’re', theyre: 'they’re', thats: 'that’s',
};
// Words ending in a full stop that don't end a sentence.
const ABBREVIATIONS = new Set(['e.g', 'i.e', 'etc', 'vs', 'mr', 'mrs', 'ms', 'dr', 'prof', 'inc', 'ltd', 'co', 'no', 'fig', 'approx', 'dept', 'st', 'jan', 'feb', 'mar', 'apr', 'jun', 'jul', 'aug', 'sep', 'sept', 'oct', 'nov', 'dec']);
const OPENS = /[\s([{“‘—–/-]$/;

/**
 * The correction for `before` (the paragraph's text up to the caret, the character just typed at its end): { from, to,
 * text } as offsets into `before`, or null. Exported for the tests.
 */
export function correction(before) {
  const typed = before.slice(-1);
  const prior = before.slice(0, -1);
  // Quotes: an opening one at the start, after a space or an opening bracket or dash; else a closing one (an apostrophe).
  if (typed === '"' || typed === "'") {
    const open = !prior || OPENS.test(prior);
    return { from: before.length - 1, to: before.length, text: typed === '"' ? (open ? '“' : '”') : open ? '‘' : '’' };
  }
  if (typed === ')') for (const [k, v] of SYMBOLS) if (before.toLowerCase().endsWith(k)) return { from: before.length - k.length, to: before.length, text: v };
  if (typed === '.' && before.endsWith('...')) return { from: before.length - 3, to: before.length, text: '…' };
  if (typed === '>' || typed === '-' || typed === '=') for (const [k, v] of ARROWS) if (before.endsWith(k)) return { from: before.length - k.length, to: before.length, text: v };
  if (/[\p{L}\p{N}]/u.test(typed)) {
    // A double hyphen between words: an em dash; a hyphen with a space each side: an en dash. (When the next word starts.)
    if (/[\p{L}\p{N}]--$/u.test(prior)) return { from: prior.length - 2, to: prior.length, text: '—' };
    const spaced = /\S (--?) $/u.exec(prior); // (one hyphen or two, a space each side: an en dash, as in Word)
    if (spaced) return { from: prior.length - 1 - spaced[1].length, to: prior.length - 1, text: '–' };
    return null;
  }
  // A word just ended (a space or punctuation typed after it).
  if (!/[\s.,;:!?)]/.test(typed)) return null;
  const m = /(^|[\s(“‘"'])([^\s(“‘"']+)$/u.exec(prior);
  if (!m) return null;
  const word = m[2];
  const at = prior.length - word.length;
  // (Own entries only: "constructor" or "toString" mustn't find the object's built-in ones.)
  const own = (table, k) => (Object.hasOwn(table, k) ? table[k] : null);
  if (typed === ' ' && own(FRACTIONS, word)) return { from: at, to: prior.length, text: FRACTIONS[word] };
  // A sentence's first word: at the paragraph's start, or after a full stop, question or exclamation mark and a space
  // (not after an abbreviation or an initial).
  const head = prior.slice(0, at);
  const sentence = !head.trim() || (/[.!?]["”’)]?\s+$/u.test(head) && !ABBREVIATIONS.has((/([\p{L}.]+)\.["”’)]?\s+$/u.exec(head)?.[1] || '').toLowerCase()) && !/\b\p{L}\.\s+$/u.test(head));
  const lower = word.toLowerCase();
  const fixed = /^\p{L}+$/u.test(word) && own(TYPOS, lower);
  if (fixed) return { from: at, to: prior.length, text: word[0] !== lower[0] || sentence ? fixed[0].toUpperCase() + fixed.slice(1) : fixed };
  // A lone "i" (a word on its own, not "i.e." or a list's "i)").
  if (word === 'i' && /\s/.test(typed)) return { from: at, to: prior.length, text: 'I' };
  // Only words all in lower case are capitalised (not "iPhone", "e.g." or a web address).
  if (sentence && /^\p{Ll}[\p{Ll}'’-]*$/u.test(word)) return { from: at, to: at + 1, text: word[0].toUpperCase() };
  return null;
}

export const autocorrectPlugin = new Plugin({
  props: {
    handleTextInput(view, from, to, text) {
      if (view.odx?.features?.autocorrect === false || text.length !== 1 || view.composing) return false;
      const $from = view.state.doc.resolve(from);
      if (!$from.parent.isTextblock || $from.parent.type.spec.code) return false;
      const before = $from.parent.textBetween(0, $from.parentOffset, null, '￼') + text;
      const fix = correction(before);
      if (!fix) return false;
      // The character as typed, then the correction as a step of its own, worked out again where the character landed
      // (with Track Changes on, typing over a selection keeps the selection as a deletion, before the new character).
      view.dispatch(view.state.tr.insertText(text, from, to));
      const $at = view.state.selection.$from;
      const again = correction($at.parent.textBetween(0, $at.parentOffset, null, '￼'));
      if (!again) return true;
      const start = $at.start();
      const tr = closeHistory(view.state.tr).insertText(again.text, start + again.from, start + again.to);
      view.dispatch(tr.setMeta('odx-autocorrect', true));
      // A changed word or a symbol is said to screen-reader users (curly quotes and dashes aren't: they'd be noise).
      if (/\p{L}{2}|[©®™…→←⇒⇐½¼¾]/u.test(again.text)) announce(`AutoCorrect: ${again.text}`);
      return true;
    },
  },
});
