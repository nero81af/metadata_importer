using Playnite.SDK;
using Playnite.SDK.Models;
using Playnite.SDK.Plugins;
using System;
using System.Collections.Generic;
using System.IO;
using System.IO.Compression;
using System.Linq;
using System.Xml.Linq;

namespace ExcelMetadataImporter
{
    public class ExcelMetadataImporter : GenericPlugin
    {
        private static readonly ILogger logger = LogManager.GetLogger();

        // Every plugin must have a unique GUID. Generate your own and paste it below.
        public override Guid Id { get; } = Guid.Parse("837f7763-5d5b-4006-a993-2657ca63b549");

        // Fixed path to your Excel file. Leave it empty to show the file picker every time.
        private const string DefaultExcelPath = @""; // e.g. @"D:\games\collection.xlsx"

        // Names of the metadata columns, looked up in the games table header (case-insensitive).
        // The game NAME column does not need to be configured: it is detected automatically as
        // the first used cell of the header row ("Game Digital" / "Game Physical").
        private const string ColPublisher = "Publisher";
        private const string ColDeveloper = "Developer";
        private const string ColGenre = "Genre";
        private const string ColRegion = "Region";

        // Maps the sheet's region values to cleaner names for Playnite. Values not listed here
        // are used as-is. Empty the map if you want to keep the original sheet values.
        private static readonly Dictionary<string, string> RegionAliases = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase)
        {
            { "Pal", "PAL" },
            { "Usa", "USA" }
        };

        private const string ColRarity = "Rarity";

        // UserScore assigned based on the number of asterisks (index = asterisk count; index 0 is unused).
        // 10 levels, from 1 to 10 asterisks -> from 10 to 100, in steps of 10.
        private static readonly int[] RarityScores = { 0, 10, 20, 30, 40, 50, 60, 70, 80, 90, 100 };

        // Rows that are not games and must be ignored (labels / repeated headers).
        private static readonly HashSet<string> SkipNames = new HashSet<string>(StringComparer.OrdinalIgnoreCase)
        {
            "Wishlist Digital", "Wishlist Physical", "Game Digital", "Game Physical"
        };

        public ExcelMetadataImporter(IPlayniteAPI api) : base(api)
        {
        }

        public override IEnumerable<GameMenuItem> GetGameMenuItems(GetGameMenuItemsArgs args)
        {
            yield return new GameMenuItem
            {
                Description = "Import metadata from Excel (matching row)",
                MenuSection = "Excel Importer",
                Action = a => ImportFromExcel(a.Games)
            };
        }

        private void ImportFromExcel(List<Game> games)
        {
            if (games == null || games.Count == 0)
            {
                return;
            }

            // 1. Locate the Excel file
            var path = DefaultExcelPath;
            if (string.IsNullOrEmpty(path) || !File.Exists(path))
            {
                path = PlayniteApi.Dialogs.SelectFile("Excel|*.xlsx");
            }
            if (string.IsNullOrEmpty(path) || !File.Exists(path))
            {
                return; // user cancelled
            }

            // 2. Load every sheet separately (sheet name -> rows keyed by game name)
            List<KeyValuePair<string, Dictionary<string, ExcelRow>>> sheets;
            try
            {
                sheets = LoadAllSheets(path);
            }
            catch (Exception ex)
            {
                logger.Error(ex, "Error reading the Excel file.");
                PlayniteApi.Dialogs.ShowErrorMessage(
                    "Could not read the Excel file:\n" + ex.Message, "Excel Importer");
                return;
            }

            if (sheets.Count == 0)
            {
                PlayniteApi.Dialogs.ShowMessage("No sheet with a games table was found.", "Excel Importer");
                return;
            }

            // 2b. Let the user choose which sheet to import from
            const string allSheetsLabel = "(All sheets)";
            var options = new List<GenericItemOption>
            {
                new GenericItemOption(allSheetsLabel, "Import from every sheet (merged)")
            };
            options.AddRange(sheets.Select(s => new GenericItemOption(s.Key, $"{s.Value.Count} games")));

            var chosen = PlayniteApi.Dialogs.ChooseItemWithSearch(
                options,
                query => string.IsNullOrEmpty(query)
                    ? options
                    : options.Where(o => o.Name.IndexOf(query, StringComparison.OrdinalIgnoreCase) >= 0).ToList(),
                null,
                "Choose the Excel sheet to import from");

            if (chosen == null)
            {
                return; // user cancelled
            }

            // Build the lookup table from the chosen sheet (or merge all)
            Dictionary<string, ExcelRow> table;
            if (chosen.Name == allSheetsLabel)
            {
                table = new Dictionary<string, ExcelRow>();
                foreach (var sheet in sheets)
                {
                    foreach (var kv in sheet.Value)
                    {
                        table[kv.Key] = kv.Value; // on duplicates the last sheet wins
                    }
                }
            }
            else
            {
                table = sheets.First(s => s.Key == chosen.Name).Value;
            }

            // 3. For each selected game: find the row and update the fields
            int updated = 0;
            var notFound = new List<string>();

            foreach (var game in games)
            {
                var key = Normalize(game.Name);
                if (!table.TryGetValue(key, out var row))
                {
                    notFound.Add(game.Name);
                    continue;
                }

                bool changed = false;

                if (!string.IsNullOrWhiteSpace(row.Publisher))
                {
                    game.PublisherIds = SplitCompanies(row.Publisher)
                        .Select(GetOrCreateCompanyId).ToList();
                    changed = true;
                }

                if (!string.IsNullOrWhiteSpace(row.Developer))
                {
                    game.DeveloperIds = SplitCompanies(row.Developer)
                        .Select(GetOrCreateCompanyId).ToList();
                    changed = true;
                }

                if (!string.IsNullOrWhiteSpace(row.Genre))
                {
                    game.GenreIds = SplitGenres(row.Genre)
                        .Select(GetOrCreateGenreId).ToList();
                    changed = true;
                }

                if (!string.IsNullOrWhiteSpace(row.Region))
                {
                    game.RegionIds = SplitRegions(row.Region)
                        .Select(GetOrCreateRegionId).ToList();
                    changed = true;
                }

                if (!string.IsNullOrWhiteSpace(row.Rarity))
                {
                    int stars = row.Rarity.Count(c => c == '*');
                    if (stars >= 1 && stars < RarityScores.Length)
                    {
                        game.UserScore = RarityScores[stars];
                        changed = true;
                    }
                }

                if (changed)
                {
                    PlayniteApi.Database.Games.Update(game);
                    updated++;
                }
            }

            // 4. Summary
            var msg = $"Games updated: {updated} of {games.Count}.";
            if (notFound.Count > 0)
            {
                msg += "\n\nNo matching row found for:\n - " + string.Join("\n - ", notFound.Take(20));
                if (notFound.Count > 20)
                {
                    msg += $"\n ... and {notFound.Count - 20} more.";
                }
            }
            PlayniteApi.Dialogs.ShowMessage(msg, "Excel Importer");
        }

        // ============================================================
        //  Reading .xlsx WITHOUT external dependencies.
        //  An .xlsx is a zip of XML files: we open it with
        //  System.IO.Compression and read the XML with System.Xml.Linq,
        //  both included in the .NET Framework.
        // ============================================================

        private List<KeyValuePair<string, Dictionary<string, ExcelRow>>> LoadAllSheets(string path)
        {
            var result = new List<KeyValuePair<string, Dictionary<string, ExcelRow>>>();

            using (var fs = File.OpenRead(path))
            using (var zip = new ZipArchive(fs, ZipArchiveMode.Read))
            {
                var shared = ReadSharedStrings(zip);

                foreach (var sheet in ReadSheetTargets(zip))
                {
                    var entry = zip.GetEntry("xl/" + sheet.Value);
                    if (entry == null)
                    {
                        continue;
                    }

                    var rows = ReadSheetRows(entry, shared);

                    // Find the games table header: the row that contains
                    // Developer, Publisher and Genre at the same time.
                    int headerIdx = -1, devCol = 0, pubCol = 0, genCol = 0, regCol = 0, rarCol = 0, nameCol = 0;
                    for (int i = 0; i < rows.Count; i++)
                    {
                        int d = FindCol(rows[i], ColDeveloper);
                        int p = FindCol(rows[i], ColPublisher);
                        int g = FindCol(rows[i], ColGenre);
                        if (d > 0 && p > 0 && g > 0)
                        {
                            headerIdx = i;
                            devCol = d; pubCol = p; genCol = g;
                            regCol = FindCol(rows[i], ColRegion); // 0 if the column is missing
                            rarCol = FindCol(rows[i], ColRarity); // 0 if the column is missing
                            nameCol = rows[i].Keys.Min(); // first used column = game name
                            break;
                        }
                    }

                    if (headerIdx < 0)
                    {
                        continue; // sheet without a games table
                    }

                    var dict = new Dictionary<string, ExcelRow>();
                    for (int i = headerIdx + 1; i < rows.Count; i++)
                    {
                        var gameName = Get(rows[i], nameCol);
                        if (string.IsNullOrEmpty(gameName) || SkipNames.Contains(gameName))
                        {
                            continue;
                        }

                        var key = Normalize(gameName);
                        if (string.IsNullOrEmpty(key))
                        {
                            continue;
                        }

                        dict[key] = new ExcelRow
                        {
                            Publisher = Get(rows[i], pubCol),
                            Developer = Get(rows[i], devCol),
                            Genre = Get(rows[i], genCol),
                            Region = regCol > 0 ? Get(rows[i], regCol) : string.Empty,
                            Rarity = rarCol > 0 ? Get(rows[i], rarCol) : string.Empty
                        };
                    }

                    if (dict.Count > 0)
                    {
                        result.Add(new KeyValuePair<string, Dictionary<string, ExcelRow>>(sheet.Key, dict));
                    }
                }
            }

            return result;
        }

        // Shared strings table (xl/sharedStrings.xml)
        private static List<string> ReadSharedStrings(ZipArchive zip)
        {
            var list = new List<string>();
            var entry = zip.GetEntry("xl/sharedStrings.xml");
            if (entry == null)
            {
                return list;
            }
            using (var s = entry.Open())
            {
                var doc = XDocument.Load(s);
                foreach (var si in doc.Root.Elements().Where(e => e.Name.LocalName == "si"))
                {
                    // Concatenate all <t> elements (handles rich text <r><t> too)
                    var text = string.Concat(si.Descendants()
                        .Where(e => e.Name.LocalName == "t")
                        .Select(e => e.Value));
                    list.Add(text);
                }
            }
            return list;
        }

        // Ordered list of sheets: name -> file path (e.g. "worksheets/sheet1.xml")
        private static List<KeyValuePair<string, string>> ReadSheetTargets(ZipArchive zip)
        {
            var result = new List<KeyValuePair<string, string>>();
            var wbEntry = zip.GetEntry("xl/workbook.xml");
            var relEntry = zip.GetEntry("xl/_rels/workbook.xml.rels");
            if (wbEntry == null || relEntry == null)
            {
                return result;
            }

            // r:id -> target
            var relMap = new Dictionary<string, string>();
            using (var s = relEntry.Open())
            {
                var doc = XDocument.Load(s);
                foreach (var rel in doc.Root.Elements().Where(e => e.Name.LocalName == "Relationship"))
                {
                    var id = (string)rel.Attribute("Id");
                    var target = (string)rel.Attribute("Target");
                    if (id != null && target != null)
                    {
                        // Normalize any absolute paths such as "/xl/worksheets/..."
                        target = target.Replace("/xl/", "").TrimStart('/');
                        relMap[id] = target;
                    }
                }
            }

            using (var s = wbEntry.Open())
            {
                var doc = XDocument.Load(s);
                var sheets = doc.Root.Elements().FirstOrDefault(e => e.Name.LocalName == "sheets");
                if (sheets == null)
                {
                    return result;
                }
                foreach (var sheet in sheets.Elements().Where(e => e.Name.LocalName == "sheet"))
                {
                    var name = (string)sheet.Attribute("name");
                    var ridAttr = sheet.Attributes().FirstOrDefault(a => a.Name.LocalName == "id");
                    if (name != null && ridAttr != null && relMap.TryGetValue(ridAttr.Value, out var target))
                    {
                        result.Add(new KeyValuePair<string, string>(name, target));
                    }
                }
            }
            return result;
        }

        // Reads the rows of a sheet: for each row, a dictionary column(1-based) -> text
        private static List<Dictionary<int, string>> ReadSheetRows(ZipArchiveEntry entry, List<string> shared)
        {
            var rows = new List<Dictionary<int, string>>();
            using (var s = entry.Open())
            {
                var doc = XDocument.Load(s);
                var sheetData = doc.Root.Elements().FirstOrDefault(e => e.Name.LocalName == "sheetData");
                if (sheetData == null)
                {
                    return rows;
                }

                foreach (var row in sheetData.Elements().Where(e => e.Name.LocalName == "row"))
                {
                    var cells = new Dictionary<int, string>();
                    foreach (var c in row.Elements().Where(e => e.Name.LocalName == "c"))
                    {
                        int col = ColumnFromRef((string)c.Attribute("r"));
                        if (col <= 0)
                        {
                            continue;
                        }

                        var t = (string)c.Attribute("t");
                        string value;

                        if (t == "s") // shared string
                        {
                            var v = c.Elements().FirstOrDefault(e => e.Name.LocalName == "v");
                            value = (v != null && int.TryParse(v.Value, out int idx) && idx >= 0 && idx < shared.Count)
                                ? shared[idx] : "";
                        }
                        else if (t == "inlineStr") // inline string
                        {
                            value = string.Concat(c.Descendants()
                                .Where(e => e.Name.LocalName == "t")
                                .Select(e => e.Value));
                        }
                        else // number or formula string
                        {
                            var v = c.Elements().FirstOrDefault(e => e.Name.LocalName == "v");
                            value = v != null ? v.Value : "";
                        }

                        if (!string.IsNullOrEmpty(value))
                        {
                            cells[col] = value;
                        }
                    }
                    rows.Add(cells);
                }
            }
            return rows;
        }

        // "C5" -> 3 ; "AB12" -> 28
        private static int ColumnFromRef(string cellRef)
        {
            if (string.IsNullOrEmpty(cellRef))
            {
                return -1;
            }
            int col = 0;
            foreach (char ch in cellRef)
            {
                if (ch >= 'A' && ch <= 'Z') col = col * 26 + (ch - 'A' + 1);
                else if (ch >= 'a' && ch <= 'z') col = col * 26 + (ch - 'a' + 1);
                else break;
            }
            return col;
        }

        private static int FindCol(Dictionary<int, string> row, string header)
        {
            foreach (var kv in row)
            {
                if (string.Equals(kv.Value.Trim(), header, StringComparison.OrdinalIgnoreCase))
                {
                    return kv.Key;
                }
            }
            return 0;
        }

        private static string Get(Dictionary<int, string> row, int col)
        {
            return row.TryGetValue(col, out var v) ? v.Trim() : string.Empty;
        }

        // --- Helpers: get-or-create to avoid duplicates in the database ---

        private Guid GetOrCreateCompanyId(string name)
        {
            var existing = PlayniteApi.Database.Companies
                .FirstOrDefault(c => string.Equals(c.Name, name, StringComparison.OrdinalIgnoreCase));
            if (existing != null)
            {
                return existing.Id;
            }
            var company = new Company(name);
            PlayniteApi.Database.Companies.Add(company);
            return company.Id;
        }

        private Guid GetOrCreateGenreId(string name)
        {
            var existing = PlayniteApi.Database.Genres
                .FirstOrDefault(g => string.Equals(g.Name, name, StringComparison.OrdinalIgnoreCase));
            if (existing != null)
            {
                return existing.Id;
            }
            var genre = new Genre(name);
            PlayniteApi.Database.Genres.Add(genre);
            return genre.Id;
        }

        private Guid GetOrCreateRegionId(string name)
        {
            // Apply the alias if any (e.g. "Jap" -> "Japan")
            if (RegionAliases.TryGetValue(name, out var mapped))
            {
                name = mapped;
            }

            var existing = PlayniteApi.Database.Regions
                .FirstOrDefault(r => string.Equals(r.Name, name, StringComparison.OrdinalIgnoreCase));
            if (existing != null)
            {
                return existing.Id;
            }
            var region = new Region(name);
            PlayniteApi.Database.Regions.Add(region);
            return region.Id;
        }

        // --- Helpers: normalization and splitting ---

        private static string Normalize(string s)
        {
            if (string.IsNullOrWhiteSpace(s))
            {
                return string.Empty;
            }

            s = s.Trim().ToLowerInvariant();

            // Curly apostrophes and quotes -> straight ones (e.g. Demon's Souls)
            s = s.Replace('\u2019', '\'').Replace('\u2018', '\'')
                 .Replace('\u201C', '"').Replace('\u201D', '"');

            // Trademark symbols
            s = s.Replace("\u2122", "").Replace("\u00AE", "").Replace("\u00A9", "");

            while (s.Contains("  "))
            {
                s = s.Replace("  ", " ");
            }

            return s.Trim();
        }

        // Multiple publishers/developers separated by ';'
        private static IEnumerable<string> SplitCompanies(string value)
        {
            return value.Split(';')
                .Select(x => x.Trim())
                .Where(x => !string.IsNullOrEmpty(x));
        }

        // Multiple genres separated by ',' or ';'
        private static IEnumerable<string> SplitGenres(string value)
        {
            return value.Split(',', ';')
                .Select(x => x.Trim())
                .Where(x => !string.IsNullOrEmpty(x));
        }

        // Multiple regions separated by ',' or ';'
        private static IEnumerable<string> SplitRegions(string value)
        {
            return value.Split(',', ';')
                .Select(x => x.Trim())
                .Where(x => !string.IsNullOrEmpty(x));
        }

        private class ExcelRow
        {
            public string Publisher { get; set; }
            public string Developer { get; set; }
            public string Genre { get; set; }
            public string Region { get; set; }
            public string Rarity { get; set; }
        }
    }
}