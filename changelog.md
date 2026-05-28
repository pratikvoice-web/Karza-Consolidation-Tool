Changelog

All notable changes to the Karza Deep-Consolidation Engine will be documented in this file.
[v2026.006] - GSTR1 Fallback Protocol & Dynamic Formatting
🚀 Added

    GSTR1 Fallback System: Implemented intelligent cross-checking logic during the revenue extraction phase. If an entity has not filed GSTR-3B for a specific month (resulting in a 0 or blank in the GSTR1 vs 3B sheet), the engine will automatically parse the adjacent GSTR1 columns. If valid numbers exist in GSTR1, they will be utilized to prevent mathematically false negative netting when internal transactions are subtracted.

    Dynamic Visual Formatting: When the GSTR1 Fallback is triggered for a specific month and state, the engine will automatically format the corresponding Gross and Net values in the generated output sheets with Italics and a Light Yellow background color. This ensures complete transparency for anyone auditing the final file.

    Glossary Update: Added a new legend entry to the Audit_Glossary sheet explicitly explaining the Light Yellow/Italic formatting applied to the GSTR1 Fallback values.
