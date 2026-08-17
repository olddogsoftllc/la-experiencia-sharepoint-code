# Graph Report - .  (2026-08-17)

## Corpus Check
- 383 files · ~88,683 words
- Verdict: corpus is large enough that graph structure adds value.

## Summary
- 76 nodes · 67 edges · 28 communities (17 shown, 11 thin omitted)
- Extraction: 100% EXTRACTED · 0% INFERRED · 0% AMBIGUOUS
- Token cost: 0 input · 0 output

## Community Hubs (Navigation)
- Community 0
- Community 1
- Community 2
- Community 3
- Community 4
- Community 5
- Community 6
- Community 7
- Community 8
- Community 9
- Community 10
- Community 12
- Community 13
- Community 14
- Community 15
- Community 16

## God Nodes (most connected - your core abstractions)
1. `getGraphClient()` - 4 edges
2. `getPriorityClass()` - 2 edges
3. `buildBadgeText()` - 2 edges
4. `shouldShowExport()` - 2 edges
5. `buildNotifyMessage()` - 2 edges
6. `buildExportMessage()` - 2 edges
7. `DEFAULT_FOOTER_TEXT` - 2 edges
8. `resolveFooterText()` - 2 edges
9. `footerBarStyle()` - 2 edges
10. `countOpenTickets()` - 2 edges

## Surprising Connections (you probably didn't know these)
- None detected - all connections are within the same source files.

## Import Cycles
- None detected.

## Communities (28 total, 11 thin omitted)

### Community 0 - "Community 0"
Cohesion: 0.21
Nodes (10): getGraphClient(), IGraphDrive, IGraphDriveItem, IGraphSite, IGraphSiteCollection, adapter, bot, botFrameworkAuthentication (+2 more)

### Community 1 - "Community 1"
Cohesion: 0.53
Nodes (4): DEFAULT_FOOTER_TEXT, footerBarStyle(), IFooterBarStyle, resolveFooterText()

### Community 2 - "Community 2"
Cohesion: 0.60
Nodes (3): buildBadgeText(), getPriorityClass(), PriorityClass

### Community 3 - "Community 3"
Cohesion: 0.70
Nodes (3): buildExportMessage(), buildNotifyMessage(), shouldShowExport()

### Community 5 - "Community 5"
Cohesion: 0.60
Nodes (3): hostRunningSuffix(), resolveThemeName(), ThemeName

## Knowledge Gaps
- **18 isolated node(s):** `styles`, `styles`, `PriorityClass`, `IFooterBarStyle`, `styles` (+13 more)
  These have ≤1 connection - possible missing edges or undocumented components.
- **11 thin communities (<3 nodes) omitted from report** — run `graphify query` to explore isolated nodes.

## Suggested Questions
_Questions this graph is uniquely positioned to answer:_

- **What connects `styles`, `styles`, `PriorityClass` to the rest of the system?**
  _18 weakly-connected nodes found - possible documentation gaps or missing edges._