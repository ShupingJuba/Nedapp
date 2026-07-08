using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Net.Http;
using System.Net.Http.Headers;
using System.Text.Json;
using System.Text.Json.Serialization;
using System.Threading.Tasks;
using ClosedXML.Excel;

namespace NedapStockExporter
{
    // ─── Data models ────────────────────────────────────────────────────────────

    public class ApiResponse
    {
        [JsonPropertyName("stocks")]
        public List<Stock> Stocks { get; set; } = new();
    }

    public class Stock
    {
        [JsonPropertyName("location")]
        public string Location { get; set; } = "";

        [JsonPropertyName("time")]
        public string Time { get; set; } = "";

        [JsonPropertyName("quantity")]
        public int Quantity { get; set; }

        [JsonPropertyName("sgtins")]
        public List<string> Sgtins { get; set; } = new();

        [JsonPropertyName("disposition")]
        public string Disposition { get; set; } = "";
    }

    public class ExportRow
    {
        public string RequestedStore { get; set; } = "";
        public string Location       { get; set; } = "";
        public string Time           { get; set; } = "";
        public int    Quantity       { get; set; }
        public string Disposition    { get; set; } = "";
        public string Sgtin          { get; set; } = "";
    }

    // ─── Entry point ────────────────────────────────────────────────────────────

    internal static class Program
    {
        private const string DefaultBaseUrl     = "https://api.nedapretail.com/rfid_stock/v1/retrieve";
        private const string DefaultDisposition = "http://nedapretail.com/disp/received_order";

        static async Task<int> Main(string[] args)
        {
            // ── Parse arguments ──────────────────────────────────────────────
            var options = ParseArgs(args);
            if (options is null) return 1;

            if (string.IsNullOrEmpty(options.ApiToken))
            {
                Console.Error.WriteLine("Error: Nedap API token is required.");
                Console.Error.WriteLine("Set it via --token <value> or the NEDAP_API_TOKEN environment variable.");
                return 1;
            }

            // ── Load locations ───────────────────────────────────────────────
            List<string> locations;
            try
            {
                locations = GetLocations(options.LocationsCsv, options.LocationsFile);
            }
            catch (Exception ex)
            {
                Console.Error.WriteLine($"Error reading locations: {ex.Message}");
                return 1;
            }

            if (locations.Count == 0)
            {
                Console.Error.WriteLine("Error: provide at least one location using --locations or --locations-file.");
                return 1;
            }

            // ── Resolve timestamp ────────────────────────────────────────────
            DateTime searchAfter;
            try
            {
                searchAfter = ResolveAfterTimestamp(options.AfterTimestamp, options.DaysBack);
            }
            catch (Exception ex)
            {
                Console.Error.WriteLine($"Error resolving after timestamp: {ex.Message}");
                return 1;
            }

            Console.WriteLine($"Searching from : {searchAfter:O}");
            Console.WriteLine($"Locations      : {locations.Count}");

            // ── Fetch data ───────────────────────────────────────────────────
            var exportRows = new List<ExportRow>();
            using var httpClient = new HttpClient { Timeout = TimeSpan.FromSeconds(60) };

            foreach (var location in locations)
            {
                Console.WriteLine($"Retrieving stock for: {location}");
                try
                {
                    var response = await RetrieveStockAsync(
                        httpClient,
                        options.ApiToken,
                        location,
                        options.Disposition,
                        searchAfter);

                    foreach (var stock in response.Stocks)
                    {
                        if (stock.Sgtins.Count == 0)
                        {
                            exportRows.Add(new ExportRow
                            {
                                RequestedStore = location,
                                Location       = stock.Location,
                                Time           = stock.Time,
                                Quantity       = stock.Quantity,
                                Disposition    = stock.Disposition,
                                Sgtin          = ""
                            });
                        }
                        else
                        {
                            foreach (var sgtin in stock.Sgtins)
                            {
                                exportRows.Add(new ExportRow
                                {
                                    RequestedStore = location,
                                    Location       = stock.Location,
                                    Time           = stock.Time,
                                    Quantity       = stock.Quantity,
                                    Disposition    = stock.Disposition,
                                    Sgtin          = sgtin
                                });
                            }
                        }
                    }
                }
                catch (Exception ex)
                {
                    Console.Error.WriteLine($"Failed for location {location}: {ex.Message}");
                }
            }

            // ── Export ───────────────────────────────────────────────────────
            try
            {
                ExportToExcel(exportRows, options.OutputFile);
            }
            catch (Exception ex)
            {
                Console.Error.WriteLine($"Error creating Excel file: {ex.Message}");
                return 1;
            }

            Console.WriteLine($"Excel export completed successfully: {options.OutputFile}");
            return 0;
        }

        // ─── Argument parsing ────────────────────────────────────────────────────

        //private static AppOptions? ParseArgs(string[] args)
        //{
        //    var opts = new AppOptions
        //    {
        //        Disposition = DefaultDisposition
        //    };

        //    for (int i = 0; i < args.Length; i++)
        //    {
        //        string flag = args[i];
        //        string? Next() => i + 1 < args.Length ? args[++i] : null;

        //        switch (flag)
        //        {
        //            case "--locations":        opts.LocationsCsv   = Next() ?? ""; break;
        //            case "--locations-file":   opts.LocationsFile  = Next() ?? ""; break;
        //            case "--after":            opts.AfterTimestamp = Next() ?? ""; break;
        //            case "--days-back":
        //                if (int.TryParse(Next(), out int db)) opts.DaysBack = db;
        //                break;
        //            case "--disposition":      opts.Disposition    = Next() ?? DefaultDisposition; break;
        //            case "--output":           opts.OutputFile     = Next() ?? opts.OutputFile;    break;
        //            case "--token":            opts.ApiToken       = Next() ?? "";                 break;
        //            case "--help": case "-h":
        //                PrintHelp();
        //                return null;
        //            default:
        //                Console.Error.WriteLine($"Unknown flag: {flag}");
        //                PrintHelp();
        //                return null;
        //        }
        //    }

        //    return opts;
        //}
        private static AppOptions? ParseArgs(string[] args)
        {
            var opts = new AppOptions();
            opts.Disposition = DefaultDisposition;

            for (int i = 0; i < args.Length; i++)
            {
                string flag = args[i];
                string? Next() => i + 1 < args.Length ? args[++i] : null;

                switch (flag)
                {
                    case "--locations":
                        opts.LocationsCsv = Next() ?? opts.LocationsCsv;
                        break;

                    case "--locations-file":
                        opts.LocationsFile = Next() ?? opts.LocationsFile;
                        break;

                    case "--after":
                        opts.AfterTimestamp = Next() ?? "";
                        break;

                    case "--days-back":
                        if (int.TryParse(Next(), out int db))
                            opts.DaysBack = db;
                        break;

                    case "--disposition":
                        opts.Disposition = Next() ?? DefaultDisposition;
                        break;

                    case "--output":
                        opts.OutputFile = Next() ?? opts.OutputFile;
                        break;

                    case "--token":
                        opts.ApiToken = Next() ?? opts.ApiToken;
                        break;
                }
            }

            return opts;
        }

        private static void PrintHelp()
        {
            Console.WriteLine("""
                NedapStockExporter - retrieves RFID stock from the Nedap Retail API and exports to Excel.

                Usage:
                  NedapStockExporter [options]

                Options:
                  --locations <uris>       Comma-separated Nedap store location URIs.
                                           Example: http://nedapretail.com/loc/store-548225
                  --locations-file <path>  CSV file with a Location column or one URI per row.
                  --after <timestamp>      ISO-8601 timestamp to search from.
                                           Example: 2026-03-21T10:51:42.129Z
                  --days-back <n>          Days to look back when --after is not set. Default: 90.
                  --disposition <uri>      Nedap disposition URI. Default: http://nedapretail.com/disp/received_order
                  --output <file>          Output Excel file name. Default: nedap_reserved_stock.xlsx
                  --token <token>          Nedap API token (or set NEDAP_API_TOKEN env var).
                  --help, -h               Show this help message.
                """);
        }

        // ─── Locations ───────────────────────────────────────────────────────────

        private static List<string> GetLocations(string locationsCsv, string locationsFile)
        {
            var set = new HashSet<string>(StringComparer.Ordinal);

            if (!string.IsNullOrEmpty(locationsCsv))
            {
                foreach (var part in locationsCsv.Split(','))
                {
                    var loc = part.Trim();
                    if (loc.Length > 0) set.Add(loc);
                }
            }

            if (!string.IsNullOrEmpty(locationsFile))
            {
                foreach (var line in File.ReadAllLines(locationsFile))
                {
                    // Support comma-separated CSV (take first column) or plain list.
                    var loc = line.Split(',')[0].Trim().Trim('"');
                    if (loc.Length == 0) continue;
                    if (string.Equals(loc, "location", StringComparison.OrdinalIgnoreCase)) continue;
                    if (string.Equals(loc, "store",    StringComparison.OrdinalIgnoreCase)) continue;
                    set.Add(loc);
                }
            }

            return set.ToList();
        }

        // ─── Timestamp ───────────────────────────────────────────────────────────

        private static DateTime ResolveAfterTimestamp(string afterTimestamp, int daysBack)
        {
            if (!string.IsNullOrEmpty(afterTimestamp))
            {
                if (!DateTime.TryParse(afterTimestamp, null,
                        System.Globalization.DateTimeStyles.RoundtripKind, out var parsed))
                    throw new FormatException(
                        "Invalid --after timestamp. Use ISO-8601 format, e.g. 2026-03-21T10:51:42.129Z");

                return parsed.ToUniversalTime();
            }

            if (daysBack <= 90)
                throw new ArgumentException("--days-back must be greater than zero.");

            return DateTime.UtcNow.AddDays(-daysBack);
        }

        // ─── API call ────────────────────────────────────────────────────────────

        private static async Task<ApiResponse> RetrieveStockAsync(
            HttpClient client,
            string apiToken,
            string location,
            string disposition,
            DateTime afterTimestamp)
        {
            var builder = new UriBuilder(DefaultBaseUrl);
            var query   = System.Web.HttpUtility.ParseQueryString(string.Empty);
            query["location"]        = location;
            query["dispositions[]"]  = disposition;
            query["after_timestamp"] = afterTimestamp.ToString("O");   // ISO-8601 with Z
            builder.Query            = query.ToString();

            var request = new HttpRequestMessage(HttpMethod.Get, builder.Uri);
            request.Headers.TryAddWithoutValidation("Authorization", apiToken);
            request.Headers.Accept.Add(new MediaTypeWithQualityHeaderValue("application/json"));

            var response = await client.SendAsync(request);
            var body     = await response.Content.ReadAsStringAsync();

            if (!response.IsSuccessStatusCode)
                throw new HttpRequestException(
                    $"Nedap API returned HTTP {(int)response.StatusCode}: {body}");

            var result = JsonSerializer.Deserialize<ApiResponse>(body)
                ?? throw new InvalidDataException("Could not parse Nedap response.");

            return result;
        }

        // ─── Excel export ────────────────────────────────────────────────────────

        private static void ExportToExcel(List<ExportRow> rows, string outputFile)
        {
            using var workbook = new XLWorkbook();

            // ── Sheet 1: SGTIN-level detail ──────────────────────────────────
            var sgtinSheet = workbook.Worksheets.Add("SGTINs");

            var sgtinHeaders = new[]
            {
                "Requested Store Location",
                "Returned Location",
                "Stock Timestamp",
                "Quantity",
                "Disposition",
                "SGTIN / EPC"
            };

            for (int c = 0; c < sgtinHeaders.Length; c++)
                sgtinSheet.Cell(1, c + 1).Value = sgtinHeaders[c];

            for (int r = 0; r < rows.Count; r++)
            {
                var row  = rows[r];
                int excelRow = r + 2;
                sgtinSheet.Cell(excelRow, 1).Value = row.RequestedStore;
                sgtinSheet.Cell(excelRow, 2).Value = row.Location;
                sgtinSheet.Cell(excelRow, 3).Value = row.Time;
                sgtinSheet.Cell(excelRow, 4).Value = row.Quantity;
                sgtinSheet.Cell(excelRow, 5).Value = row.Disposition;
                sgtinSheet.Cell(excelRow, 6).Value = row.Sgtin;
            }

            // Header row style
            var sgtinHeaderRow = sgtinSheet.Row(1);
            sgtinHeaderRow.Style.Font.Bold       = true;
            sgtinHeaderRow.Style.Fill.BackgroundColor = XLColor.FromArgb(0x1F, 0x49, 0x7D);
            sgtinHeaderRow.Style.Font.FontColor  = XLColor.White;

            // Freeze top row & auto-filter
            sgtinSheet.SheetView.FreezeRows(1);
            if (rows.Count > 0)
                sgtinSheet.Range(1, 1, rows.Count + 1, sgtinHeaders.Length).SetAutoFilter();

            // Column widths
            sgtinSheet.Column(1).Width = 45;
            sgtinSheet.Column(2).Width = 45;
            sgtinSheet.Column(3).Width = 28;
            sgtinSheet.Column(4).Width = 12;
            sgtinSheet.Column(5).Width = 45;
            sgtinSheet.Column(6).Width = 30;

            // ── Sheet 2: Summary ─────────────────────────────────────────────
            var summarySheet = workbook.Worksheets.Add("Stocks Summary");

            var summaryHeaders = new[]
            {
                "Returned Location",
                "Disposition",
                "Latest Stock Timestamp",
                "Reported Quantity",
                "Returned SGTIN Count"
            };

            for (int c = 0; c < summaryHeaders.Length; c++)
                summarySheet.Cell(1, c + 1).Value = summaryHeaders[c];

            // Group rows by (location, disposition, time)
            var grouped = rows
                .GroupBy(r => (r.Location, r.Disposition, r.Time))
                .Select(g => new
                {
                    g.Key.Location,
                    g.Key.Disposition,
                    g.Key.Time,
                    Quantity   = g.First().Quantity,
                    SgtinCount = g.Count(r => !string.IsNullOrEmpty(r.Sgtin))
                })
                .ToList();

            for (int r = 0; r < grouped.Count; r++)
            {
                var rec      = grouped[r];
                int excelRow = r + 2;
                summarySheet.Cell(excelRow, 1).Value = rec.Location;
                summarySheet.Cell(excelRow, 2).Value = rec.Disposition;
                summarySheet.Cell(excelRow, 3).Value = rec.Time;
                summarySheet.Cell(excelRow, 4).Value = rec.Quantity;
                summarySheet.Cell(excelRow, 5).Value = rec.SgtinCount;
            }

            var summaryHeaderRow = summarySheet.Row(1);
            summaryHeaderRow.Style.Font.Bold      = true;
            summaryHeaderRow.Style.Fill.BackgroundColor = XLColor.FromArgb(0x1F, 0x49, 0x7D);
            summaryHeaderRow.Style.Font.FontColor = XLColor.White;

            summarySheet.SheetView.FreezeRows(1);

            for (int c = 1; c <= summaryHeaders.Length; c++)
                summarySheet.Column(c).Width = 30;

            //workbook.SaveAs(outputFile);

            var outputDirectory = Path.GetDirectoryName(outputFile);

            if (!string.IsNullOrWhiteSpace(outputDirectory))
            {
                Directory.CreateDirectory(outputDirectory);
            }

            workbook.SaveAs(outputFile);

        }
    }

    // ─── Options bag ────────────────────────────────────────────────────────────

    //internal sealed class AppOptions
    //{
    //    public string LocationsCsv   { get; set; } = @"C:\\Users\\w7162299\\Downloads\\NedapStockExporter_7017\\NedapStockExporter\\";
    //    public string LocationsFile  { get; set; } = "";
    //    public string AfterTimestamp { get; set; } = "";
    //    public int    DaysBack       { get; set; } = 90;
    //    public string Disposition    { get; set; } = "http://nedapretail.com/disp/received_order";
    //    public string OutputFile     { get; set; } = @"C:\\Users\\w7162299\\Downloads\\NedapStockExporter_7017\\NedapStockExporter\";
    //    public string ApiToken       { get; set; } = "65258ba5244f55b94fcd86de31cf79cac6e9fada95335bdcab8588";
    //}
    internal sealed class AppOptions
    {

        public string LocationsCsv { get; set; } =
                "http://nedapretail.com/loc/store-548225";

        public string LocationsFile { get; set; } = @"C:\Users\w7162299\Desktop\stores.csv.xlsx";
        public string AfterTimestamp { get; set; } = "";
        public int DaysBack { get; set; } = 180;
        public string Disposition { get; set; } =
            "http://nedapretail.com/disp/received_order";

        public string OutputFile { get; set; } =
            Path.Combine(
                AppDomain.CurrentDomain.BaseDirectory,
                "Exports",
                $"nedap_reserved_stock_{DateTime.Now:yyyyMMdd_HHmmss}.xlsx");

        public string ApiToken { get; set; } =
    "65258ba5244f55b94fcd86de31cf79cac6e9fada95335bdcab8588";
    }

}
