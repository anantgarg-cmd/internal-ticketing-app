# Shadowfax world-class UI — live visual QA

## UX rationale

The redesign uses progressive disclosure for dependent client fields, a persistent review panel to reduce memory load, responsive tables/cards to preserve scanability without horizontal scrolling, and a right-side detail drawer to retain list context. Status combines text, shape, and colour; loading uses skeletons for list structure and a delayed reference-counted overlay for explicit work. All visual enhancements reuse the initial payload and existing APIs.

## DEV/live checklist

A deployed Google Apps Script environment was not available in this container. Every item below is therefore **NOT RUN** and should be exercised against DEV before production validation.

1. **NOT RUN** — Initial loading state
2. **NOT RUN** — Active SALES navigation
3. **NOT RUN** — Active POC navigation
4. **NOT RUN** — Active ADMIN navigation
5. **NOT RUN** — Expanded sidebar
6. **NOT RUN** — Collapsed sidebar and persisted preference
7. **NOT RUN** — Raise form desktop
8. **NOT RUN** — Raise form mobile
9. **NOT RUN** — Client-type transitions
10. **NOT RUN** — Dynamic-field transitions
11. **NOT RUN** — File selection, drag/drop, metadata, and clear
12. **NOT RUN** — Draft save, restore, and clear
13. **NOT RUN** — Duplicate modal
14. **NOT RUN** — Ticket creation success
15. **NOT RUN** — Sales Tickets desktop
16. **NOT RUN** — Sales Tickets mobile
17. **NOT RUN** — Debounced search and clear
18. **NOT RUN** — Filters and active state
19. **NOT RUN** — Empty states
20. **NOT RUN** — Pagination
21. **NOT RUN** — Ticket drawer and focus return
22. **NOT RUN** — Long description
23. **NOT RUN** — Many timeline events
24. **NOT RUN** — Multiple SLA cycles
25. **NOT RUN** — Attachment link
26. **NOT RUN** — Investigate action
27. **NOT RUN** — Resolve action
28. **NOT RUN** — Reopen action
29. **NOT RUN** — Work Queue and current-view metrics
30. **NOT RUN** — SLA Dashboard
31. **NOT RUN** — Loader on slow request
32. **NOT RUN** — Loader after failure
33. **NOT RUN** — Keyboard-only navigation
34. **NOT RUN** — 390 px viewport
35. **NOT RUN** — 768 px viewport
36. **NOT RUN** — 1440 px viewport
37. **NOT RUN** — Slow network simulation
38. **NOT RUN** — Unexpected horizontal overflow
39. **NOT RUN** — Logo load without layout shift and fallback
40. **NOT RUN** — Normal flows without erroneous red toast
