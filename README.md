# Dynamics Desk

A lightweight Electron desktop client for Microsoft Dynamics 365 Field Service — bookings, schedule board, work orders, parts requests, accounts, and contacts in a fast native window.

> This is a sanitized public mirror. All organization-specific values (URLs, record IDs, app IDs, account/territory names, locations, credentials) have been replaced with `YOUR-…` placeholders.

## Setup

```bash
npm install
npm start
```

On first launch, enter your Dynamics 365 org URL and sign in with your Microsoft account. The app then prompts you to add your team members (used for the schedule board, bookings filter, and team view).

### Configure for your own org

- **Model-driven app ID** — replace `YOUR-MODEL-DRIVEN-APP-ID` in `renderer/workorder.js` and `renderer/contact.js` (found in your Dynamics app's URL).
- **SharePoint PTO calendar** (optional, for time-off bars on the schedule board) — set `SP_SITE` and `SP_TIMEOFF_LIST` in `renderer/app.js` to your SharePoint site and calendar list. Columns expected: `Title` (person), `EventDate`, `EndDate`.
- **Quick "Travel Home" work order** (optional) — fill in the `TRAVEL_WO` GUIDs in `renderer/app.js` (service account, incident type, priority, price list, service territory) plus the booking status / setup-metadata GUIDs, all pulled from a sample work order in your org.
- **Ship-to locations** — the parts-request form's `Ship to Location` options in `renderer/workorder.html` and the `SHIP_TO_LABELS` map in `renderer/workorder.js` use placeholder labels/values; replace with your org's option-set.

### Auto-update (optional)

`renderer` ships with a self-updater. To use it, set `owner`/`repo`/`token` in `update-config.json` (a read-only, fine-grained GitHub token scoped to a private releases repo), then cut releases with `npm run release`. Leave the placeholders as-is to disable it.

## Features

- Bookings list with search (customer, work order, resource, status), status/substatus/team-member filters, and reset
- Weekly schedule board with a timeline view, hover previews, utilization %, auto-scroll to today, and time-off bars (from a SharePoint PTO calendar)
- Work order window: status, sub-status, resource, notes, service tasks, parts requests, assigned engineers, add-engineer / transfer, and a timeline tab
- Parts requests: build a list then submit as real Dynamics records that appear on the work order timeline
- One-click "Travel Home" internal work order (creates the work order and books it)
- Standalone contact and team-member detail windows with live status (Free / Scheduled / Traveling / In Progress)
- Accounts and contacts directories with server-side search and a clickable account detail modal
- Native spellcheck across all text fields (right-click suggestions + custom dictionary)
- Fast child windows: work order / contact / team windows share the main window's warm Dynamics session over IPC instead of each loading the full app shell
- Configurable team roster (add/remove members in-app)
- Custom light-blue/grey theme with Windows-style window controls
