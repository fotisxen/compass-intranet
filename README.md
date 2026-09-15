# compass-intranet

## Summary

SharePoint Framework (SPFx) solution for the Compass company intranet. Includes:

- **Compass Home** — a web part with the intranet home page content: upcoming events, a live fleet map, newsroom, welcome/promotion spotlights, work anniversaries, and public holidays.
- **CompassChromeApplicationCustomizer** — a site-wide extension that injects the header (with hover mega-menus and an AI assistant entry point) and footer into every page of the site. The assistant chat calls the tenant's existing AskStarBulk backend (`chatResourceUri`/`chatQueryUrl` in `sharepoint/assets/elements.xml`) via SPFx's built-in `AadHttpClientFactory`, authenticated as the signed-in user — no API key involved on our side. This requires the `webApiPermissionRequests` entry in `config/package-solution.json` to be approved once by a tenant admin under SharePoint Admin Center → Advanced → API access after the first deployment. If `chatResourceUri`/`chatQueryUrl` are left blank, `ChatWidget` falls back to `chatApiUrl` (a custom proxy, if one is configured); if that's blank too, it shows an "assistant not configured" message rather than a fake reply. The footer's stock ticker fetches real SBLK/S&P 500 quotes from `stockApiUrl` (same `elements.xml`, defaults to the company's `GetStockPrice` endpoint on the fleet API's Function App); the app never shows placeholder data anywhere — if a real fetch fails or a data source isn't configured, the affected widget (news, events, holidays, stock ticker, chat) simply renders empty rather than fabricated content.
- **api/** — an Azure Function backend that proxies fleet vessel positions to the company's fleet-tracking provider, keeping its API key server-side. Configure `FLEET_API_URL` and `FLEET_API_KEY` in the function app's settings (see `api/local.settings.json.template` for local development).

## Used SharePoint Framework Version

![version](https://img.shields.io/badge/version-1.23.2-green.svg)

## Applies to

- [SharePoint Framework](https://aka.ms/spfx)
- [Microsoft 365 tenant](https://docs.microsoft.com/sharepoint/dev/spfx/set-up-your-developer-tenant)

## Prerequisites

- Node.js >= 22.14.0 < 23.0.0
- A SharePoint Online site to deploy to (hosted workbench needs a tenant; the standalone preview does not)

## Getting started

```bash
npm install
npm start
```

`npm start` runs `heft start --clean` and serves the local dev bundle. To see it rendered, open the hosted workbench against your tenant:

```
https://<your-tenant>.sharepoint.com/_layouts/15/workbench.aspx?debugManifestsFile=https://localhost:4321/temp/build/manifests.js&debug=true&noredir=true
```

### Standalone preview (no tenant required)

```bash
npm run preview
```

Bundles the real component source with esbuild and serves it at `http://localhost:5600` — useful for checking layout/styling changes without a SharePoint tenant.

## Building and deploying

```bash
npm run build
```

Produces `solution/compass-intranet.sppkg`. Upload it to a SharePoint App Catalog, deploy it, then install the app on the target site — this makes the "Compass Home" web part available to add to a page, and activates the site-wide header/footer automatically.

## Solution

| Solution           | Author(s) |
| ------------------ | --------- |
| compass-intranet   |           |

## Disclaimer

**THIS CODE IS PROVIDED _AS IS_ WITHOUT WARRANTY OF ANY KIND, EITHER EXPRESS OR IMPLIED, INCLUDING ANY IMPLIED WARRANTIES OF FITNESS FOR A PARTICULAR PURPOSE, MERCHANTABILITY, OR NON-INFRINGEMENT.**

## References

- [Getting started with SharePoint Framework](https://docs.microsoft.com/sharepoint/dev/spfx/set-up-your-developer-tenant)
- [Use Microsoft Graph in your solution](https://docs.microsoft.com/sharepoint/dev/spfx/web-parts/get-started/using-microsoft-graph-apis)
- [Publish SharePoint Framework applications to the Marketplace](https://docs.microsoft.com/sharepoint/dev/spfx/publish-to-marketplace-overview)
- [Heft Documentation](https://heft.rushstack.io/)
