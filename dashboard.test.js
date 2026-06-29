import test from 'node:test';
import assert from 'node:assert/strict';
import { readFileSync } from 'node:fs';

import {
  buildAdsApiUrl,
  computeStats,
  extractSheetPayload,
  formatAdsFromSheetRows,
  highlightText,
  resolveInitialSource,
  resolveSheetSource,
  shouldCollapseSidebar,
} from './dashboard.js';
import {
  buildSheetUrl,
  resolveProxySource,
} from './functions/api/ads.js';

test('buildAdsApiUrl points the browser to the controlled proxy endpoint', () => {
  const url = buildAdsApiUrl(resolveSheetSource('CG'));

  assert.equal(url, '/api/ads?source=CG');
});

test('resolveSheetSource matches public source codes case-insensitively without exposing sheet IDs', () => {
  assert.deepEqual(resolveSheetSource('cg'), {
    code: 'CG',
    label: 'CG FB SPY',
  });
  assert.deepEqual(resolveSheetSource('CH'), {
    code: 'CH',
    label: 'CH FB SPY',
  });
  assert.deepEqual(resolveSheetSource('kenny'), {
    code: 'KENNY',
    label: 'KENNY FB SPY',
  });
});

test('resolveSheetSource returns null for unknown source codes', () => {
  assert.equal(resolveSheetSource('unknown'), null);
});

test('resolveInitialSource starts with no active sheet source', () => {
  assert.equal(resolveInitialSource(), null);
});

test('dashboard code does not expose Google Sheet IDs or direct Google Sheet URLs', () => {
  const script = readFileSync(new URL('./dashboard.js', import.meta.url), 'utf8');

  assert.doesNotMatch(script, /docs\.google\.com\/spreadsheets/);
  assert.doesNotMatch(script, /[a-zA-Z0-9_-]{30,}/);
});

test('worker proxy resolves sheet sources from server-side environment only', () => {
  const env = {
    ADSPY_SOURCES_JSON: JSON.stringify({
      CG: {
        label: 'CG FB SPY',
        sheetId: 'server-only-sheet-id',
        sheetName: 'Sheet1',
      },
    }),
  };

  assert.deepEqual(resolveProxySource('cg', env), {
    code: 'CG',
    label: 'CG FB SPY',
    sheetId: 'server-only-sheet-id',
    sheetName: 'Sheet1',
  });
  assert.equal(resolveProxySource('missing', env), null);
});

test('worker proxy builds the Google Sheet URL server-side', () => {
  const source = {
    sheetId: 'server-only-sheet-id',
    sheetName: 'Sheet 1',
  };

  assert.equal(
    buildSheetUrl(source),
    'https://docs.google.com/spreadsheets/d/server-only-sheet-id/gviz/tq?tqx=out:json&sheet=Sheet%201',
  );
});

test('extractSheetPayload parses Google Visualization responses', () => {
  const wrapped = 'google.visualization.Query.setResponse({"table":{"rows":[{"c":[{"v":"123"}]}]}});';
  const payload = extractSheetPayload(wrapped);

  assert.equal(payload.table.rows[0].c[0].v, '123');
});

test('extractSheetPayload parses plain JSON responses', () => {
  const payload = extractSheetPayload('{"table":{"rows":[{"c":[{"v":"abc"}]}]}}');

  assert.equal(payload.table.rows[0].c[0].v, 'abc');
});

test('highlightText escapes dangerous HTML before adding highlight markup', () => {
  const html = highlightText('<img src=x onerror=alert(1)>promo', 'promo');

  assert.match(html, /&lt;img/);
  assert.doesNotMatch(html, /<img/);
  assert.match(html, /<span class="highlight">promo<\/span>/i);
});

test('formatAdsFromSheetRows prefers final media URL when present', () => {
  const columns = [
    { label: 'ad_archive_id' },
    { label: 'type' },
    { label: 'original_media_url' },
    { label: 'start_date' },
    { label: 'page_name' },
    { label: 'original_feed' },
    { label: 'platforms' },
    { label: 'cta_text' },
    { label: 'original_url' },
    { label: 'status' },
    { label: 'final_media_url' },
  ];
  const rows = [
    {
      c: [
        { v: 'ad-1' },
        { v: 'video' },
        { v: 'https://origin.example/video.mp4' },
        { v: '2026-03-24' },
        { v: 'Brand A' },
        { v: 'feed' },
        { v: 'Facebook, Instagram' },
        { v: 'Book now' },
        { v: 'https://brand.example/landing' },
        { v: 'Done' },
        { v: 'https://final.example/video.mp4' },
      ],
    },
  ];

  const ads = formatAdsFromSheetRows(rows, columns);

  assert.equal(ads[0].media_url, 'https://final.example/video.mp4');
  assert.equal(ads[0].page_name, 'Brand A');
  assert.equal(ads[0].platforms, 'Facebook, Instagram');
  assert.equal(ads[0].cta_text, 'Book now');
  assert.equal(ads[0].original_url, 'https://brand.example/landing');
});

test('computeStats groups CTA distribution and falls back to No CTA', () => {
  const stats = computeStats([
    { page_name: 'Brand A', type: 'image', cta_text: 'Learn more' },
    { page_name: 'Brand A', type: 'video', cta_text: 'Book now' },
    { page_name: 'Brand B', type: 'video', cta_text: '' },
    { page_name: 'Brand C', type: 'image', cta_text: 'Learn more' },
  ]);

  assert.equal(stats.totalAds, 4);
  assert.equal(stats.totalBrands, 3);
  assert.deepEqual(stats.ctas, {
    'Learn more': 2,
    'Book now': 1,
    'No CTA': 1,
  });
});

test('index places the chart KPI block before the creative feed list', () => {
  const html = readFileSync(new URL('./index.html', import.meta.url), 'utf8');

  const chartIndex = html.indexOf('<section class="charts-grid">');
  const feedIndex = html.indexOf('<div id="adsContainer" class="ads-grid"></div>');

  assert.notEqual(chartIndex, -1);
  assert.notEqual(feedIndex, -1);
  assert.ok(chartIndex < feedIndex);
});

test('index and static deployment files opt out of public indexing', () => {
  const html = readFileSync(new URL('./index.html', import.meta.url), 'utf8');
  const headers = readFileSync(new URL('./_headers', import.meta.url), 'utf8');
  const robots = readFileSync(new URL('./robots.txt', import.meta.url), 'utf8');

  assert.match(html, /<meta name="robots" content="noindex, nofollow, noarchive, nosnippet">/);
  assert.match(headers, /X-Robots-Tag: noindex, nofollow, noarchive, nosnippet/);
  assert.match(robots, /Disallow: \//);
});

test('worker routes API requests and keeps static responses out of search indexes', () => {
  const worker = readFileSync(new URL('./_worker.js', import.meta.url), 'utf8');

  assert.match(worker, /pathname === '\/api\/ads'/);
  assert.match(worker, /env\.ASSETS\.fetch/);
  assert.match(worker, /X-Robots-Tag/);
});

test('shouldCollapseSidebar only collapses on desktop when stored state is collapsed', () => {
  assert.equal(shouldCollapseSidebar('collapsed', 1440), true);
  assert.equal(shouldCollapseSidebar('expanded', 1440), false);
  assert.equal(shouldCollapseSidebar('collapsed', 1024), false);
});

test('index keeps only the refresh status UI from the top control strip', () => {
  const html = readFileSync(new URL('./index.html', import.meta.url), 'utf8');

  assert.doesNotMatch(html, /id="brandInput"/);
  assert.doesNotMatch(html, /id="goBtn"/);
  assert.doesNotMatch(html, /id="currentBrandDisplay"/);
  assert.match(html, /id="lastUpdate"/);
});

test('index uses the Buuluu light visual system', () => {
  const html = readFileSync(new URL('./index.html', import.meta.url), 'utf8');

  assert.match(html, /Noto\+Sans\+HK/);
  assert.match(html, /color-scheme:\s*light/);
  assert.match(html, /--bg:\s*#ffffff/);
});

test('index includes the sidebar source switcher controls', () => {
  const html = readFileSync(new URL('./index.html', import.meta.url), 'utf8');

  assert.match(html, /id="sourceInput"/);
  assert.match(html, /id="sourceApplyBtn"/);
  assert.match(html, /id="sourceStatus"/);
});

test('index empty state asks the user to enter a source code first', () => {
  const html = readFileSync(new URL('./index.html', import.meta.url), 'utf8');

  assert.match(html, /請先輸入資料源代號/);
});

test('index keeps the source switcher container visible on first render', () => {
  const html = readFileSync(new URL('./index.html', import.meta.url), 'utf8');

  assert.doesNotMatch(html, /id="mainContent" hidden/);
  assert.match(html, /id="feedPanel"/);
});
