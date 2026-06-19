# Dynamics Desk

A lightweight Electron desktop client for Microsoft Dynamics 365 Field Service — bookings, schedule board, accounts, and contacts in a fast native window.

## Setup

```bash
npm install
npm start
```

On first launch, enter your Dynamics 365 org URL and sign in with your Microsoft account.

Before running, configure these for your own org:
- `DEFAULT_BOOKINGS_RESOURCE` and `SCHEDULE_RESOURCES` in `renderer/app.js` — your resource name and your team's resource/technician names
- `APP_ID` in `renderer/workorder.js` and `renderer/contact.js` — your Dynamics model-driven app ID (found in that app's URL)

## Features

- Bookings list with search, status/substatus filters, team-member filter, and a reset-filters button
- Weekly schedule board with a timeline view, hover previews, and auto-scroll to today
- Work order detail window with status, sub-status, resource, notes, tasks, products, and assigned-engineers list
- Standalone contact and team-member detail windows showing contact info and current status (Free / Scheduled / Traveling / In Progress)
- Accounts and contacts directories with server-side search and a clickable account detail modal
- Custom light-blue/grey theme with Windows-style window controls
