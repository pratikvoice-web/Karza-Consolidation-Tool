[v2026.07] - Financial Year Outlining & Volumetric Sorting
🚀 Added

    Dynamic Financial Year Groupings: Built a chronologic parser supporting standard Indian FY bounds (April to March). Outputs now explicitly group months into overarching Financial Year boundaries (e.g., FY24-25).

    Multi-Axis Expandable Subledgers: Utilized the FY logic to automatically configure native Excel row and column outlines (ws.Outline.Group()). Summary sheets now collapse months into FY rows, while detail sheets dynamically span and collapse FY columns with explicit "FY Total" safety pillars.

    Volumetric Counterparty Sorting: Integrated deep LINQ array evaluations during the subledger compilation layer. Suppliers and Customers are now automatically ranked and written to the matrix strictly by total transaction volume (largest to smallest) for immediate executive visibility.

    LZ4 Native Compression: Explicitly requested the .NET compiler flag <EnableCompressionInSingleFile> to shrink the total payload size of the final GUI .exe dramatically.
