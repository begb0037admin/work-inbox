#!/usr/bin/env node
'use strict';

/* Extract the shipped visibility rule, rather than carrying a second implementation. */
const fs = require('fs');
const path = require('path');
const source = fs.readFileSync(path.join(__dirname, '..', 'js', 'app.js'), 'utf8');

function extract(pattern) {
  const match = source.match(pattern);
  if (!match) throw new Error(`Could not find ${pattern}`);
  const start = match.index;
  let at = source.indexOf('{', start), depth = 0;
  for (; at < source.length; at++) {
    if (source[at] === '{') depth++;
    if (source[at] === '}' && --depth === 0) return source.slice(start, at + 1);
  }
  throw new Error(`Could not brace-match ${pattern}`);
}

const helpers = [
  /function _tickStorageKey\(id\)\{/, /function _priGetLegacyTitleKey\(p\)\{/, /function _priGetKey\(p\)\{/, /function _priCardVisible\(p,ticks,showingDone\)\{/
].map(extract).join('\n');
const priCardVisible = new Function(`let currentKey='test';\n${helpers}\nreturn _priCardVisible;`)();

function card(id) { return { id, title: id }; }
function expect(name, actual, wanted) {
  if (actual !== wanted) throw new Error(`${name}: expected ${wanted}; got ${actual}`);
  console.log(`PASS: ${name}`);
}

const plain = card('plain');
const handled = card('handled');
const deleted = card('deleted');
const ticks = { id_handled: true, del_id_deleted: true };

expect('deleted cards are invisible', priCardVisible(deleted, ticks, false), false);
expect('handled cards are hidden when Show done is off', priCardVisible(handled, ticks, false), false);
expect('handled cards are visible when Show done is on', priCardVisible(handled, ticks, true), true);
expect('plain cards are visible', priCardVisible(plain, ticks, false), true);
expect('section count excludes handled and deleted cards', [plain, handled, deleted].filter(p => priCardVisible(p, ticks, false)).length, 1);
expect('section count includes handled but never deleted cards when Show done is on', [plain, handled, deleted].filter(p => priCardVisible(p, ticks, true)).length, 2);
console.log('All section count checks passed.');
