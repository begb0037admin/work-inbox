#!/usr/bin/env node
'use strict';

const fs = require('fs');
const path = require('path');
const source = fs.readFileSync(path.join(__dirname, '..', 'js', 'app.js'), 'utf8');

function extract(name) {
  const start = source.indexOf(`function ${name}(`);
  if (start < 0) throw new Error(`Missing ${name}`);
  let at = source.indexOf('{', start), depth = 0;
  for (; at < source.length; at += 1) {
    if (source[at] === '{') depth += 1;
    if (source[at] === '}' && --depth === 0) return source.slice(start, at + 1);
  }
  throw new Error(`Could not parse ${name}`);
}

const wiLinkDestinations = new Function('wiLinkDocument', `${extract('wiLinkDestinations')}; return wiLinkDestinations;`)({
  links: [
    { status: 'active', wiKeys: ['id_t1'], trackerId: 'trk_one' },
    { status: 'active', wiKeys: ['id_t1'], trackerId: 'trk_two' },
    { status: 'proposed', wiKeys: ['id_t1'], trackerId: 'trk_ignored' },
  ],
});
const destinations = wiLinkDestinations('id_t1');
if (destinations.length !== 2 || !destinations[0].url.includes('tracker.lelitte.co.uk/#trk_one')) throw new Error('active Tracker destinations were not rendered from the link map');
if (!source.includes("grid.append(button)")) throw new Error('Work Inbox Tracker icon was not placed in the allowed third row');
if (!source.includes("function goToWorkInboxCard(key)") || !source.includes('applySecCollapse(sec,false)')) throw new Error('Work Inbox deep-link arrival does not unfold its section');
if (!source.includes("destinations.length===1") || !source.includes("className='dashboard-link-picker'")) throw new Error('multi-link picker path is missing');
console.log('Work Inbox link icon, deep-link and picker checks passed.');
