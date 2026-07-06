# 09 — SPFx Library: Shared code

An **SPFx Library** deployed to the App Catalog that other solutions consume as
a dependency. Loads **once** per page regardless of how many Web Parts use it,
so shared code counts once toward the bundle size.

Covers the book's chapter 14 "SPFx Library: shared code across solutions".

## Public surface (`src/index.ts`)

```typescript
import {
  classifyHttpStatus,    // 401/403/404/429/5xx → coarse category for UI branching
  truncate,               // text → "…"-suffixed truncation
  formatRelativeDate      // ISO date → "3 days ago" style string
} from 'spfx-library-shared';
```

Pure TypeScript utilities — no React / Graph dependency, so the library stays
small. Add React components or a shared `GraphService` here when your solutions
need them (the library then declares those deps).

## Build

```bash
nvm use 18
npm install
gulp build
gulp bundle --ship
gulp package-solution --ship     # sharepoint/solution/spfx-library-shared.sppkg
```

## Deploy & consume

1. Upload the `.sppkg` to the App Catalog and enable it (tenant-wide).
2. In the consuming solution: `npm install ../09-spfx-library-shared --save`
   (or reference it by its library id once published to the App Catalog).
3. `import { classifyHttpStatus } from 'spfx-library-shared';`

> Treat the library as a **versioned** dependency: bump its `version` and
> re-deploy consumers when the public API changes, or pages will load a stale
> cached version.