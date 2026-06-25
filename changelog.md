[v2026.08] - Invariant Parsing & Case-Insensitive Thread Synchronization
🐛 Fixed

- Data Exclusion & Disk Overwrite Bug: Patched a severe system thread collision where entity strings possessing identical spelling but different casing boundaries (e.g., "ABC TRADERS" vs "Abc Traders") bypassed memory grouping and subsequently forced destructive disk-level overwrites.

🚀 Added

 - Invariant Text Transformation (Global Uppercase): Shifted string standardisation directly to the extraction tier. The engine immediately normalises counterparties to uppercase arrays before assigning internal dictionary keys.

 - Corporate Suffix Normalization Layer: Integrated boundary-locked Regular Expressions to intelligently sanitize and uniform variable corporate designations (PRIVATE LIMITED, (P) LTD, LTD., etc.) into a consistent data format (PVT LTD and LTD) without manipulating internal brand naming structures.

https://github.com/pratikvoice-web/Karza-Consolidation-Tool/releases/tag/v2026.beta.008
