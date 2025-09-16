# autoprint-and-archive

Automated PDF printing and archiving tool for Windows. Monitors your downloads folder and automatically prints and organizes files based on regex patterns.

## Features

-   Automatic printing to configured printer with print/prompt/skip options
-   Smart file organization based on regex patterns with variable extraction
-   Duplicate file detection to prevent reprocessing
-   Configurable monitoring and archiving

## Common Use Cases

-   **Invoice Management**: Automatically print and archive downloaded invoices by date/vendor
-   **Receipt Organization**: Sort receipts into monthly folders and print for bookkeeping
-   **Document Processing**: Organize contracts, forms, or reports by type and date
-   **Business Workflow**: Streamline document handling for accounting or administrative tasks

## Installation

1. Download latest release executable from [Releases](https://github.com/ptmrio/autoprint-and-archive/releases)
2. Create `config.yaml` in same directory as executable
3. Run the executable

## Configuration

Create `config.yaml`:

```yaml
default_printer: "Your Printer Name"
language: "en" # Interface language: "en" for English, "de" for German
dedupe_ttl_seconds: 60 # Prevent duplicate processing within time window (optional, default: 30)

patterns:
    - pattern: "Invoice-(?P<year>\\d{4})-(?P<month>\\d{2})-\\d+\\.pdf$"
      destination: "C:\\Archive\\Invoices\\{year}\\{month}"
      print: true # true=always print, false=never print, prompt=ask each time
      printer: "HP LaserJet" # Optional: override default printer for this pattern
```

## Build

```bash
pyinstaller --noconsole --onefile autoprint-and-archive.py
```
