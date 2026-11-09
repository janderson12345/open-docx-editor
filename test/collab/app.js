// Co-editing test page: test/collab.mjs opens two of these on one room (?port=&room=&name=).
import * as Y from 'yjs';
import { WebsocketProvider } from 'y-websocket';
import { mountDocxEditor } from '../../src/index.js';
import { collaboration } from '../../src/collab.js';
import * as comments from '../../src/editor/comments.js';

import * as hf from '../../src/editor/hfedit.js';
import * as notes from '../../src/editor/notes.js';

window.cm = comments; // (the test adds, answers, resolves and deletes comments through these)
window.hf = hf; // (and edits a header and adds a footnote through these)
window.notes = notes;

const q = new URLSearchParams(location.search);
const ydoc = new Y.Doc();
const provider = new WebsocketProvider(`ws://localhost:${q.get('port')}`, q.get('room'), ydoc);
provider.on('sync', (synced) => {
  if (!synced || window.odx) return;
  // Mounted once the room's state has arrived, as collab needs.
  const collab = collaboration(ydoc, { awareness: provider.awareness, user: { name: q.get('name'), color: q.get('name') === 'A' ? '#c00' : '#06c' } });
  window.odx = mountDocxEditor(document.getElementById('app'), { collab, author: q.get('name') });
});
