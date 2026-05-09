// jsdom bootstrap so @tiptap/core helpers (generateJSON / generateHTML)
// can run in Node. Tiptap requires a DOM environment; jsdom gives us
// one. Importing this file for side effects sets up the globals.

import { JSDOM } from 'jsdom';

const dom = new JSDOM('<!DOCTYPE html><html><body></body></html>', {
  url: 'http://localhost/',
});

globalThis.window = dom.window;
globalThis.document = dom.window.document;
globalThis.DOMParser = dom.window.DOMParser;
globalThis.XMLSerializer = dom.window.XMLSerializer;
globalThis.Node = dom.window.Node;
globalThis.Element = dom.window.Element;
globalThis.HTMLElement = dom.window.HTMLElement;
globalThis.Text = dom.window.Text;
globalThis.NodeFilter = dom.window.NodeFilter;
