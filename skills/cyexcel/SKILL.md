---
name: cyexcel
description: Integrate, configure, or extend the cyexcel browser spreadsheet component, including workbook tabs, formulas, validation, context menus, and XLSX export. Use when a task names cyexcel or an existing project imports cyexcel.
license: MIT
---

# cyexcel

Implement the requested spreadsheet behavior using the public cyexcel API and the host project's existing framework and lifecycle conventions.

## Choose the integration

- `createSheet()` creates one sheet with a toolbar, formula bar, and context menu enabled by default.
- `createWorkbook()` adds workbook tabs and enables the bundled formula engine by default.
- `CyExcelCore` provides rendering and editing without preset widgets. Its context menu and formula engine are opt-in.
- All three require a browser DOM and a container with measurable dimensions. Create them after mounting and call `dispose()` before unmounting. Avoid constructing them during server rendering or recreating them on each render.

Read [integration.md](references/integration.md) for setup, framework lifecycle examples, snapshot import, and installation constraints. Read [features.md](references/features.md) for formulas, workbook management, validation, clipboard, and export. Read [extensions.md](references/extensions.md) when customizing styles, menus, plugins, or performance-sensitive behavior.

## API rules

- Coordinates and range endpoints are zero-based; range endpoints are inclusive. Formula text uses one-based A1 references.
- Import arrays with `matrixToWorkbookSnapshot()` or `rowsToWorkbookSnapshot()`. Data is copied once; subsequent edits do not synchronize with the source array.
- Use public setters and command interfaces for edits. Do not mutate snapshot records or internal stores to bypass history or rendering updates.
- `sheet.toSnapshot()` serializes a sheet's backing model. A `WorkbookHandle` is a separate UI wrapper: do not invent `workbook.toSnapshot()` or promise whole-workbook persistence from the active sheet alone.
- Follow the installed version's exported types if they differ from these references. Package imports use `cyexcel` or its published subpaths; `@cyexcel/*` aliases are specific to this repository.

## Current scope

The package in this repository is currently private with version `0.1.0`. Check availability before recommending registry installation; installing this Skill does not install the component. Do not change release flags or publish a package as part of ordinary integration.

Formula support is limited to same-sheet references and twelve functions. Cross-sheet references, formula relocation on copy/fill, and XLSX import are unavailable. Formula-containing sheets restrict structural operations and sorting. Report these limits when relevant instead of generating calls for unsupported features.

## Verify the result

Typecheck the host integration and verify creation, editing, and disposal in a browser. Exercise the requested behavior, including undo when applicable. For source changes in the cyexcel repository, run relevant tests and `pnpm typecheck`; use `pnpm docs:build` for documentation-site changes. Describe the checks actually performed and any unverified browser behavior.
