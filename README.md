# Dynamics Desk

A lightweight Electron desktop client for Microsoft Dynamics 365 Field Service — bookings, schedule board, accounts, and contacts in a fast native window.

## Setup

```bash
npm install
npm start
```

On first launch, enter your Dynamics 365 org URL and sign in with your Microsoft account.

Before running, set `MY_RESOURCE_NAME` and `SCHEDULE_RESOURCES` in `renderer/app.js` to match the resource/technician names in your own org.

## Features

- Bookings list with search and filters
- Weekly schedule board with drag-free timeline view and hover previews
- Work order detail view with status, resource, and notes editing
- Accounts and contacts directories
