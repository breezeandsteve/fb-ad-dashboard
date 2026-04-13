import test from 'node:test';
import assert from 'node:assert/strict';
import { readFileSync } from 'node:fs';

import {
  buildSheetUrl,
  computeStats,
  extractSheetPayload,
  formatAdsFromSheetRows,
  highlightText,
  resolveSheetSource,
  shouldCollapseSidebar,
} from './dashboard.js';

test('buildSheetUrl points to the configured Google Sheet gviz endpoint', () => {
  const url = buildSheetUrl(resolveSheetSource('CG'));

  assert.equal(
    url,
    'https://docs.google.com/spreadsheets/d/1NzSHaQe6puchCA1B-tU2-4VLR1_gHlOQCiCuV9DIltk/gviz/tq?tqx=out:json&sheet=Sheet1',
  );
});

test('resolveSheetSource matches hardcoded source codes case-insensitively', () => {
  assert.equal(resolveSheetSource('cg')?.sheetId, '1NzSHaQe6puchCA1B-tU2-4VLR1_gHlOQCiCuV9DIltk');
  assert.equal(resolveSheetSource('CH')?.sheetId, '1_Ni_mQ4xVJRZ86q5y75KTgtFBhxi45tHM2YeDYGf2dA');
});

test('resolveSheetSource returns null for unknown source codes', () => {
  assert.equal(resolveSheetSource('unknown'), null);
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
