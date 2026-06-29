const DEFAULT_SHEET_NAME = 'Sheet1';

const API_HEADERS = {
  'Cache-Control': 'no-store',
  'Referrer-Policy': 'strict-origin-when-cross-origin',
  'X-Content-Type-Options': 'nosniff',
  'X-Robots-Tag': 'noindex, nofollow, noarchive, nosnippet',
};

function normalizeSourceCode(value) {
  return String(value ?? '').trim().toUpperCase();
}

function parseSources(env = {}) {
  if (!env.ADSPY_SOURCES_JSON) {
    return {};
  }

  try {
    return JSON.parse(env.ADSPY_SOURCES_JSON);
  } catch {
    return {};
  }
}

export function resolveProxySource(code, env = {}) {
  const sourceCode = normalizeSourceCode(code);
  const sources = parseSources(env);
  const source = sources[sourceCode];

  if (!source?.sheetId) {
    return null;
  }

  return {
    code: sourceCode,
    label: source.label || sourceCode,
    sheetId: String(source.sheetId),
    sheetName: source.sheetName || DEFAULT_SHEET_NAME,
  };
}

export function buildSheetUrl(source) {
  return `https://docs.google.com/spreadsheets/d/${source.sheetId}/gviz/tq?tqx=out:json&sheet=${encodeURIComponent(source.sheetName)}`;
}

function jsonResponse(payload, status = 200) {
  return Response.json(payload, {
    status,
    headers: API_HEADERS,
  });
}

export async function onRequestGet({ request, env }) {
  const url = new URL(request.url);
  const source = resolveProxySource(url.searchParams.get('source'), env);

  if (!source) {
    return jsonResponse({ error: '找不到代號' }, 404);
  }

  let sheetResponse;
  try {
    sheetResponse = await fetch(buildSheetUrl(source));
  } catch {
    return jsonResponse({ error: '資料來源暫時無法連線' }, 502);
  }

  const body = await sheetResponse.text();

  if (!sheetResponse.ok) {
    return jsonResponse({ error: '資料來源暫時無法讀取' }, 502);
  }

  return new Response(body, {
    status: 200,
    headers: {
      ...API_HEADERS,
      'Content-Type': 'application/json; charset=utf-8',
    },
  });
}
