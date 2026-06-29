# AGENTS.md

## Product Positioning

Except for the public `buuluu.com` marketing website, Buuluu client tools are private client delivery tools. They are not public SEO products and should not expose public landing pages.

## Security Rules

- Do not add public SEO pages, marketing landing pages, sitemap entries, or indexable public content for client tools.
- Keep client tools out of search indexes with `noindex`, `robots.txt`, and deployment headers where possible.
- Require login or client-level access control before exposing client data.
- Keep customer data isolated so one client cannot access another client's data source or dashboard.
- Keep session and token behavior explicit. Do not put access tokens in URLs, and do not persist long-lived tokens in `localStorage`.
- Do not let AdSpy read public Google Sheets directly from the browser in the target architecture. Move data access behind a controlled API or Worker proxy that maps the authenticated client to allowed data sources.
- Do not commit live secrets in workflow exports. Keep sanitized exports separate from local/import-only secret-filled artifacts.

## Current AdSpy Gap

The current static AdSpy frontend still contains hardcoded Google Sheet IDs and reads Google Sheets directly from the browser. Treat this as a temporary implementation until a Worker proxy and access-control layer are added.
