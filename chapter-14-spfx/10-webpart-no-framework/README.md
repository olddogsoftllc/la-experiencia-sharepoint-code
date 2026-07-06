# 10 — Web Part: No Framework

A **vanilla TypeScript Web Part** with no React and no UI framework. Renders
straight into `this.domElement`, wires a button listener, and — critically —
removes that listener in `onDispose` (no framework means *you* own cleanup).

Covers the book's cap14 "Web Part sin React (No Framework)" and demonstrates the
key differences from the React path: no virtual DOM, manual listener teardown,
and CSS Modules still applying.

## What it shows

- `render()` builds HTML from `pageContext` (site title, user display name) and
  the Property Pane `heading`.
- A **Recargar** button re-renders the Web Part on click.
- `onDispose()` removes the click listener — forgetting this is the classic
  no-framework memory leak.
- Property Pane with a single `PropertyPaneTextField` for the heading.
- `onThemeChanged` + CSS custom properties (`--bodyText`, etc.) for theming,
  including a `.dark` class for inverted themes.

## Build

```bash
nvm use 18
npm install
gulp build
gulp bundle --ship
gulp package-solution --ship     # sharepoint/solution/no-framework-webpart.sppkg
```

## Deploy

1. Upload the `.sppkg` to the App Catalog and enable it.
2. Add the Web Part to a page and edit the **Heading** in the Property Pane.

## Properties

| Property | Default | Purpose |
|---|---|---|
| `heading` | `No Framework Web Part` | Title shown at the top of the Web Part |