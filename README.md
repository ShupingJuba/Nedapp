# NedapStockExporter

C# port of the Nedap RFID Stock API → Excel exporter.  
Produces a self-contained Windows EXE (no .NET runtime required on the target machine).

---

## Prerequisites (build machine only)

- [.NET 8 SDK](https://dotnet.microsoft.com/download/dotnet/8.0)

---

## Build

### Windows
```bat
build.bat
```
Output: `publish\win-x64\NedapStockExporter.exe`

### Linux / macOS
```bash
chmod +x build.sh
./build.sh          # Windows EXE
./build.sh linux-x64
./build.sh osx-x64
```

---

## Usage

```
NedapStockExporter.exe [options]

Options:
  --locations <uris>       Comma-separated Nedap store location URIs
                           e.g. http://nedapretail.com/loc/store-548225
  --locations-file <path>  CSV file with a Location column (or one URI per line)
  --after <timestamp>      ISO-8601 start timestamp  e.g. 2026-03-21T10:51:42.129Z
  --days-back <n>          Days to look back when --after is not set (default: 90)
  --disposition <uri>      Nedap disposition URI
                           (default: http://nedapretail.com/disp/received_order)
  --output <file>          Output Excel filename (default: nedap_reserved_stock.xlsx)
  --token <token>          Nedap API token
  --help, -h               Show help
```

### API token

Prefer the environment variable so the token stays out of shell history:

**Windows**
```bat
set NEDAP_API_TOKEN=your-token-here
NedapStockExporter.exe --locations http://nedapretail.com/loc/store-548225
```

**Linux / macOS**
```bash
export NEDAP_API_TOKEN="your-token-here"
./NedapStockExporter --locations http://nedapretail.com/loc/store-548225
```

### Examples

Retrieve last 90 days for two stores:
```
NedapStockExporter.exe --locations "http://nedapretail.com/loc/store-1,http://nedapretail.com/loc/store-2"
```

Use a CSV file of locations and custom date range:
```
NedapStockExporter.exe --locations-file stores.csv --days-back 30 --output stock_march.xlsx
```

---

## Output

The generated Excel file contains two sheets:

| Sheet | Contents |
|-------|----------|
| **SGTINs** | One row per SGTIN/EPC across all stores |
| **Stocks Summary** | One row per (location, disposition, timestamp) group |

Both sheets have frozen header rows, styled headers, and auto-filter enabled.

---

## NuGet packages used

| Package | Purpose |
|---------|---------|
| [ClosedXML](https://github.com/ClosedXML/ClosedXML) | Excel generation |
| CsvHelper | (included for future CSV work; not currently active) |
