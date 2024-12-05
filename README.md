# autoprint-and-archive

Automated PDF printing and archiving tool for Windows. Monitors downloads folder and automatically prints + archives files based on regex patterns.

## Features

-   Automatic printing to configured printer
-   File organization based on regex patterns
-   Log

## Installation

1. Download latest release executable from [Releases](https://github.com/ptmrio/autoprint-and-archive/releases)
2. Create `config.yaml` in same directory as executable
3. Run the executable

## Configuration

Create `config.yaml`:

```yaml
default_printer: "Your Printer Name"

patterns:
  - pattern: "RE-(?P<year>\\d{4})(?P<month>\\d{2})\\d{2}-\\d+\\.pdf$"
    destination: "C:\\Archive\\{year}\\{month}"
    print: true
```
