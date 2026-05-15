Changelog

All notable changes to the Karza Deep-Consolidation Engine will be documented in this file.
[v2026.005] - Infinite Deep Scan & Unregistered Entity Support
🚀 Added

    Infinite Deep Scanning (UsedRange Bypass): The engine no longer relies on Excel's notoriously unreliable UsedRange.Rows.Count to determine the end of a month's data block. It now safely executes an infinite downward scan (capped at an extreme 100,000 rows) and breaks solely on a mathematical 50-row blank streak. This guarantees no internal or standard transactions are missed, regardless of formatting inconsistencies.

    Unregistered/B2C Capture: Added a safety net for rows containing valid Taxable/Invoice values but missing a PAN (e.g., Unregistered B2B or B2C transactions). These are now assigned to an "UNREGISTERED" PAN bucket, ensuring they are fully captured and appropriately grouped in the matrix output instead of being completely skipped.

    Safe Type-Casting Logic: Data extraction cells are now wrapped in try/catch statements. If the Karza report contains a text string like NA, -, or #DIV/0! in a numeric column, the script will silently default to 0 rather than crashing the loop and dropping the row entirely.

🐛 Fixed

    Data Exclusion on Blank PANs: Fixed a critical loop termination bug where the engine would abort processing an entire month block if it encountered a row with a blank PAN (such as a manually inserted empty row or an unregistered customer).

    Global Name Overwrites: Hardened the $panToNameMap dictionary so that it explicitly prevents overwriting legitimate party names with the "UNREGISTERED" tag during the name backfilling process.
