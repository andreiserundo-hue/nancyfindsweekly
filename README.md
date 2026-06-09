# Nancy Finds Weekly

Live paid-acquisition dashboard for Nancy Finds — pulls **Glued (Meta Ads)** + **Shopify** in real time, with **Meta Marketing API failover**. Tabbed UI (Overview, Scorecard, Campaigns, Landing Pages, Creatives, Products, Countries, Strategy), date + time-of-day filters, strict 🟢🟡🔴 performance tagging, click-to-expand creatives, and a one-click weekly **PPTX export**.

## Setup
```bash
npm install
cp secrets.example.json secrets.json   # then fill in your tokens
# (optional) paste a Meta ads_read token into meta_token.txt for failover
node server.js                          # → http://localhost:4321
```

## Data sources & conventions
- **Glued (primary)** via MCP (`app.glued.me/mcp`) — campaigns, landing pages, creatives + images, hourly, demographics.
- **Meta Marketing API (failover)** — `meta_ads.js`, token in `meta_token.txt` (gitignored).
- **Shopify Admin API** — orders → revenue (Total Sales, net of refunds), products, countries.
- **Glued date windows are end-exclusive** (`00:00 → next-day 00:00`); the server handles this automatically.
- **Tagging:** 🟢 above Best · 🟡 Good–Best · 🔴 below Good. ROAS floor 1.5, target 2.0.

## Files
- `server.js` — local API + static host (port 4321)
- `index.html` — the dashboard SPA
- `weekly_refresh.js` — standalone generator (static dashboard + data dump)
- `build_live_deck.js` / `build_deck.js` — PPTX generators
- `meta_ads.js` — direct Meta connector (failover)
- `config.js` / `secrets.json` — credential loading

## Security
Secrets live in `secrets.json` and `meta_token.txt` — both **gitignored**. Never commit them.
