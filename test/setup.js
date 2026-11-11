import { DOMParser, XMLSerializer } from '@xmldom/xmldom';
import { setXmlImpl } from '../src/ooxml/xml.js';

// Strict parser: any XML error fails the test instead of being recovered from.
class StrictParser {
  parseFromString(text, type) {
    // (HTML is read as a browser reads it, recovering from errors: imported content, src/ooxml/altchunk.js.)
    if (type === 'text/html') return new DOMParser({ onError: () => {} }).parseFromString(text, type);
    return new DOMParser({ onError: (level, msg) => { if (level !== 'warning') throw new Error('XML ' + level + ': ' + msg); } }).parseFromString(text, type);
  }
}
setXmlImpl({ DOMParser: StrictParser, XMLSerializer });
