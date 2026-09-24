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
  /function _calendarBucketsForDate\(data,now\)\{/,
  /function renderTeamsPanel\(data\)\{/,
  /function renderCalPanel\(data\)\{/,
].map(extract).join('\n');
const RealDate = Date;
function ThursdayDate(...args) {
  return args.length ? new RealDate(...args) : new RealDate('2026-09-24T10:00:00');
}
const render = new Function('document', 'Date', `${helpers}; return {renderCalPanel, renderTeamsPanel};`)(documentStub, ThursdayDate);

const data = {
  connector_status: {
    calendar: { status: 'carried_forward', as_of: '2026-09-21T11:25:26Z' },
    teams: { status: 'carried_forward', as_of: '2026-09-21T11:25:26Z' },
  },
  calToday: [{ time: '09:00', title: 'Stale Monday meeting' }],
  calTomorrow: [{ time: '11:00', title: 'Stale Tuesday meeting' }],
  calDay2: [], calDay3: [],
  calFull: [
    { date: '2026-09-21', items: [{ time: '09:00', title: 'Monday meeting' }] },
    { date: '2026-09-24', items: [{ time: '10:00', title: 'Thursday meeting' }] },
    { date: '2026-09-25', items: [{ time: '11:00', title: 'Friday meeting' }] },
    { date: '2026-09-28', items: [{ time: '14:00', title: 'Next Monday meeting' }] },
    { date: '2026-09-29', items: [{ time: '15:00', title: 'Next Tuesday meeting' }] },
  ],
  teams: [{ channel: 'HR Systems', sender: 'Alex', time: '09:30', preview: 'Status update' }],
};

render.renderCalPanel(data);
if (!calEl.innerHTML.includes('Calendar from Mon 21 Sep')) throw new Error('calendar carry-forward label was not rendered');
if (!calEl.innerHTML.includes('Thursday meeting')) throw new Error('Thursday event was not rendered under the carried calendar');
if (!calEl.innerHTML.includes('Friday meeting') || !calEl.innerHTML.includes('Next Monday meeting') || !calEl.innerHTML.includes('Next Tuesday meeting')) throw new Error('future carried calendar items were not projected');
if (calEl.innerHTML.includes('main-cal-title">Monday meeting</div>') || calEl.innerHTML.includes('main-cal-title">Stale Monday meeting</div>')) throw new Error('Monday events leaked into the carried Thursday calendar');

render.renderTeamsPanel(data);
if (!teamsEl.innerHTML.includes('Teams from Mon 21 Sep')) throw new Error('Teams carry-forward label was not rendered');
if (!teamsEl.innerHTML.includes('Status update')) throw new Error('Teams data was not rendered');

console.log('Dashboard carry-forward render checks passed.');
