#!/usr/bin/env node
'use strict';

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

const calEl = { innerHTML: '' };
const teamsEl = { innerHTML: '' };
const documentStub = { getElementById: id => id === 'calPanel' ? calEl : id === 'teamsPanel' ? teamsEl : null };
const helpers = [
  /function escapeHtml\(text\)\{/,
  /function _connectorStatusInfo\(data,key\)\{/,
  /function _connectorFreshnessNote\(data,key\)\{/,
  /function _connectorUnavailableNote\(data,key\)\{/,
  /function renderTeamsPanel\(data\)\{/,
  /function renderCalPanel\(data\)\{/,
].map(extract).join('\n');
const render = new Function('document', `${helpers}; return {renderCalPanel, renderTeamsPanel};`)(documentStub);

const data = {
  connector_status: {
    calendar: { status: 'carried_forward', as_of: '2026-09-21T11:25:26Z' },
    teams: { status: 'carried_forward', as_of: '2026-09-21T11:25:26Z' },
  },
  calToday: [{ time: '09:00', title: 'Daily catch-up' }],
  calTomorrow: [{ time: '11:00', title: 'Planning' }],
  calDay2: [], calDay3: [],
  teams: [{ channel: 'HR Systems', sender: 'Alex', time: '09:30', preview: 'Status update' }],
};

render.renderCalPanel(data);
if (!calEl.innerHTML.includes('Calendar from Mon 21 Sep')) throw new Error('calendar carry-forward label was not rendered');
if (!calEl.innerHTML.includes('Daily catch-up') || !calEl.innerHTML.includes('Planning')) throw new Error('carried calendar items were not rendered');

render.renderTeamsPanel(data);
if (!teamsEl.innerHTML.includes('Teams from Mon 21 Sep')) throw new Error('Teams carry-forward label was not rendered');
if (!teamsEl.innerHTML.includes('Status update')) throw new Error('Teams data was not rendered');

console.log('Dashboard carry-forward render checks passed.');
