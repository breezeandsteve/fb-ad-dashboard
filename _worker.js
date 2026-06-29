import { onRequestGet } from './functions/api/ads.js';

const STATIC_HEADERS = {
  'Permissions-Policy': 'camera=(), microphone=(), geolocation=()',
  'Referrer-Policy': 'strict-origin-when-cross-origin',
  'X-Content-Type-Options': 'nosniff',
  'X-Robots-Tag': 'noindex, nofollow, noarchive, nosnippet',
};

function withStaticHeaders(response) {
  const headers = new Headers(response.headers);

  Object.entries(STATIC_HEADERS).forEach(([key, value]) => {
    headers.set(key, value);
  });

  return new Response(response.body, {
    status: response.status,
    statusText: response.statusText,
    headers,
  });
}

export default {
  async fetch(request, env, ctx) {
    const url = new URL(request.url);

    if (url.pathname === '/api/ads') {
      if (request.method !== 'GET') {
        return Response.json(
          { error: 'Method not allowed' },
          {
            status: 405,
            headers: STATIC_HEADERS,
          },
        );
      }

      return onRequestGet({
        request,
        env,
        params: {},
        data: {},
        waitUntil: ctx.waitUntil.bind(ctx),
      });
    }

    return withStaticHeaders(await env.ASSETS.fetch(request));
  },
};
