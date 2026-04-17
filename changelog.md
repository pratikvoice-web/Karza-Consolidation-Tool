Changelog

All notable changes to the Karza Deep-Consolidation Engine will be documented in this file.
[v2026.004] - Matrix & Header Precision Update
🚀 Added

    Dual Matrix Generation: The engine now automatically generates four detailed matrices instead of two. It produces distinct counterparty breakdowns for both Taxable Value and Invoice Value for all Customers and Suppliers.

        New sheets: Detailed_Customer_Taxable, Detailed_Customer_Invoice, Detailed_Supplier_Taxable, Detailed_Supplier_Invoice.

🔄 Changed

    Standardized Reporting Headers: Renamed the data blocks within the four main revenue sheets to strictly adhere to requested financial terminologies. Replaced generic generic references like "Gross - Supplier Taxable" with highly specific equivalents:

        Gross Revenue - Taxable Value

        Internal Purchases/Sales - Taxable Value

        Net Revenue - Taxable Value

        Gross Revenue - Invoice Value

        Internal Purchases/Sales - Invoice Value

        Net Revenue - Invoice Value

    Progress Bar Synchronization: Updated the dynamic step counter to accurately reflect the 9 total generation steps (4 Revenue + 4 Matrices + 1 Glossary) while maintaining the flicker-free single-line UI.

🐛 Fixed

    Data Variable Unification: Resolved a potential data-loss conflict by unifying customer and supplier array collections into a single $matrixData array, preventing overwriting during the extraction phase and ensuring flawless matrix building for both sides of the ledger.
