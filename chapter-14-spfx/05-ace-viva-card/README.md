# 05 — Adaptive Card Extension: Viva Card

An **Adaptive Card Extension (ACE)** for Viva Connections that shows a count of
open tickets on the dashboard card and a list of tickets in the QuickView.

Covers the book's chapter 14 "Adaptive Card Extensions (ACE) in detail".

## Structure

```
src/adaptiveCardExtensions/ticketsAce/
  TicketsAceAdaptiveCardExtension.ts   # onInit: state (demo tickets) + register views
  cardView/CardView.ts                 # BasicCardView: count of open tickets + "View" button
  quickView/QuickView.ts               # data() returns the ticket list
  quickView/template/QuickViewTemplate.json   # Adaptive Card JSON ($data binding)
  TicketsAcePropertyPane.ts            # title field
```

## How it works

- `onInit` seeds `this.state.tickets` with demo data and registers the
  `CardView` (compact dashboard card) and `QuickView` (expanded panel).
- The card shows `${openCount} open tickets`; tapping it opens the QuickView.
- The QuickView template binds `$data` to the tickets array with a `ColumnSet`
  per ticket (title + status).

> Wire `loadTickets()` to Microsoft Graph (`/me/planner/tasks`, scope
> `Tasks.Read` — already declared in `package-solution.json`) to make it a live
> feed instead of demo data.

## Build

```bash
nvm use 18
npm install
gulp build
gulp bundle --ship
gulp package-solution --ship     # sharepoint/solution/ace-viva-card.sppkg
```

## Deploy

1. Upload the `.sppkg` to the App Catalog and enable it.
2. Approve **Microsoft Graph → Tasks.Read** in API Access (for a live feed).
3. Add the card to a **Viva Connections** dashboard and set **Card title**.