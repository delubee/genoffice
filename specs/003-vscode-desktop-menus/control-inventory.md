# Desktop control source inventory

Run `node scripts/inventory-desktop-menus.mjs` from the workspace root to refresh
[control-inventory.json](control-inventory.json). Run with `--check` to detect
source drift. The artifact records source hashes, exact source locations, action
expressions, labels, selection constraints, menu options and command dispatches.
It covers SheetsRibbon/ExcelShell, the Function Library, Docs Ribbon and its
ribbon tab modules, plus directly named context-menu modules in their directories.

The initial snapshot contained 16 modules and 1,103 evidence entries. The generator now also follows the extracted shared SortDialog, GoalSeekDialog, RemoveDuplicatesDialog and AdvancedFilterDialog as explicit inventory sources; its current counts are printed on regeneration. These are
syntax locations, including repeated dispatches and callback forwarding; they
must not be reported as 1,103 distinct commands or implemented plugin controls.
The generator reports dynamic dispatch/option entries requiring expansion. JSX callback forwarding
can also be dynamic even when that counter does not include it.

Every entry still requires a human review linking it to the plugin handler,
authoritative operation, persistence path and acceptance evidence. Dynamic menu
providers, nested imported controls, dialogs, panes and keyboard/focus states
must be expanded before the complete ledger task can be checked off. This source
inventory is a starting point, not proof of complete menu parity.

Existing reviewed group ledgers:

- [Sheets Insert](sheets-insert-ledger.md)
- [Sheets Page Layout](sheets-page-layout-ledger.md)
- [Sheets Formulas](sheets-formulas-ledger.md)
- [Sheets Data](sheets-data-ledger.md)

The complete Docs control/variant review and Sheets Home, Review, View, contextual
menu and dynamically generated variant reviews remain open. A matching command
name or shared component does not establish a working adapter or saved result.
