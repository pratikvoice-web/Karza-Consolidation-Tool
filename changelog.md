🚀 Added

    Graphical User Interface (GUI): Completely retired the terminal command-line interface in favor of a modern, standalone Windows WPF desktop application.

    Dynamic Execution Pathways: Implemented native Windows Folder Pickers, allowing users to dynamically select Source and Destination directories without modifying code or moving files.

    Asynchronous Progress Matrix: Engine processes have been decoupled from the UI using background worker threads (Task.Run). Features dual synchronized progress bars tracking Extraction and Compilation phases simultaneously without freezing the application window.

    Live Audit Terminal: Added a built-in scrolling terminal console to the application window to display real-time transaction validations, error catches, and runtime metrics.

    The "Black Box" Logger: Implemented a global exception failsafe (KARZA_CRASH_LOG.txt). If the application encounters an OS-level block or initialization failure, it will now trap the error, write a detailed stack trace to the execution folder, and notify the user rather than silently vanishing.

🛠️ Fixed & Optimized

    Corporate AV "Silent Kill" Bypass: Resolved a critical issue where Windows Defender and enterprise security policies were silently terminating the application upon launch. Native UI rendering libraries (wpfgfx_cor3.dll) are now configured to extract safely to the local disk rather than loading suspiciously into active RAM.

    XAML Compilation Hardening: Sanitized UI layout properties to ensure strict backward compatibility with standard Windows Presentation Foundation rendering engines.

    Domain Model Stabilization: Re-linked structural data classes (FileMetadata, SummaryRecord, etc.) directly to the core application namespace, ensuring flawless single-file compilation.

⏭️ Planned for Next Release

    Enable LZ4 Native AOT Compression to reduce overall executable file size.
