### Changelog Version 2026.003

🚀 Added

    Smart Entity Grouping (PAN-Based): The engine now reads the PAN directly from the Entity Profile sheet (Cell B5/B6) rather than relying on filenames.

    Proprietorship vs. Corporate Logic: Implemented intelligent grouping based on the 4th character of the PAN. Proprietorships (4th character 'P') are grouped by PAN + Trade Name, while Companies ('C'/'F') are strictly grouped by PAN to prevent fragmentation caused by trade name typos across different state reports.

    Multi-GSTIN Resolution: Automatically detects if an entity has multiple GST registrations within the exact same state. It splits these into separate columns using the last 3 digits of the GSTIN as a suffix (e.g., 29-Karnataka-1Z7 vs 29-Karnataka-2Z9).

    Unlimited Row Scanning: Removed hardcoded 150-row limits. The script now uses UsedRange.Rows.Count to dynamically scan an unlimited number of transactions per month, ensuring no data is dropped for high-turnover entities.

    Extended Historical Support: Expanded column scanning range up to 400 columns to support parsing Karza reports containing over 3 years of historical data.

    Live Dashboard UI: Introduced a custom, flicker-free terminal UI with a progress bar that uses cursor manipulation to cleanly update statuses on a single line.

    File-Lock Failsafe: Added an automatic detection and prompt system that warns the user if the target CONSOLIDATED Excel file is currently open, pausing the script until the file is closed rather than crashing.

🔄 Changed

    Single-File Polyglot Architecture: Combined the .ps1 PowerShell script and the execution wrapper into a single .bat file for seamless sharing and "double-click" execution.

    Sheet Nomenclature: Optimized the naming of the output sheets to clearly indicate Inter-company transactions while strictly adhering to Excel's 31-character limit:

        Tax. Value - Internal Sales

        Inv. Value - Internal Sales

        Tax. Value - Internal Purchases

        Inv. Value - Internal Purchases

    Hierarchy Visualization: Updated the state-level drill-down indicator from a Unicode branch to >> to ensure cross-system compatibility.

🐛 Fixed

    Excel Date-Serial Bug: Fixed an issue where Excel would automatically convert parsed month strings (e.g., "Mar-25") into serial numbers (e.g., "46106"). Month headers are now strictly forced into text format (@).

    Revenue Netting Zeros: Fixed a logical mapping error where Measure-Object was dropping values if internal transactions were missing in a given month. Gross and Net revenues now calculate perfectly.

    PowerShell 5.1 Compatibility: Refactored modern operators (like ??) and variable enclosures to ensure the script runs natively on older, out-of-the-box Windows Enterprise environments without requiring framework updates.

    Terminal Text Wrapping: Fixed a UI bug where extremely long Karza filenames would wrap to a second line and break the terminal progress bar's cursor alignment. Filenames are now dynamically truncated in the UI display.
