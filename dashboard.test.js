import test from 'node:test';
import assert from 'node:assert/strict';

import {
  buildSheetUrl,
  extractSheetPayload,
  formatAdsFromSheetRows,
  highlightText,
} from './dashboard.js';

test('buildSheetUrl points to the configured Google Sheet gviz endpoint', () => {
  const url = buildSheetUrl();

  assert.equal(
    url,
    'https://docs.google.com/spreadsheets/d/1NzSHaQe6puchCA1B-tU2-4VLR1_gHlOQCiCuV9DIltk/gviz/tq?tqx=out:json&sheet=Sheet1',
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
