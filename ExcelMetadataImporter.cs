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

        // Ogni plugin deve avere un GUID univoco. Generane uno tuo e incollalo qui sotto.
        public override Guid Id { get; } = Guid.Parse("837f7763-5d5b-4006-a993-2657ca63b549");

        // Percorso fisso del tuo Excel. Lascialo vuoto per far comparire ogni volta
        // la finestra di selezione del file.
        private const string DefaultExcelPath = @""; // es. @"D:\giochi\collection.xlsx"

        // Nomi delle colonne dei metadati, cercati nell'intestazione della tabella giochi
        // (case-insensitive). La colonna del NOME gioco viene rilevata automaticamente come
        // prima colonna usata della riga di intestazione ("Game Digital" / "Game Physical").
        private const string ColPublisher = "Publisher";
        private const string ColDeveloper = "Developer";
        private const string ColGenre = "Genre";

        // Righe che non sono giochi e vanno ignorate (etichette/intestazioni ripetute).
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
                Description = "Importa metadati da Excel (riga corrispondente)",
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

            // 1. Trova il file Excel
            var path = DefaultExcelPath;
            if (string.IsNullOrEmpty(path) || !File.Exists(path))
            {
                path = PlayniteApi.Dialogs.SelectFile("Excel|*.xlsx");
            }
            if (string.IsNullOrEmpty(path) || !File.Exists(path))
            {
                return; // utente ha annullato
            }

            // 2. Carica l'Excel in un dizionario (nome gioco normalizzato -> riga)
            Dictionary<string, ExcelRow> table;
            try
            {
                table = LoadExcel(path);
            }
            catch (Exception ex)
            {
                logger.Error(ex, "Errore nella lettura dell'Excel.");
                PlayniteApi.Dialogs.ShowErrorMessage(
                    "Impossibile leggere il file Excel:\n" + ex.Message, "Excel Importer");
                return;
            }

            // 3. Per ogni gioco selezionato: cerca la riga e aggiorna i campi
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

                if (changed)
                {
                    PlayniteApi.Database.Games.Update(game);
                    updated++;
                }
            }

            // 4. Riepilogo
            var msg = $"Giochi aggiornati: {updated} su {games.Count}.";
            if (notFound.Count > 0)
            {
                msg += "\n\nNessuna riga trovata per:\n - " + string.Join("\n - ", notFound.Take(20));
                if (notFound.Count > 20)
                {
                    msg += $"\n ... e altri {notFound.Count - 20}.";
                }
            }
            PlayniteApi.Dialogs.ShowMessage(msg, "Excel Importer");
        }

        // ============================================================
        //  Lettura .xlsx SENZA dipendenze esterne.
        //  Un .xlsx e' uno zip di file XML: lo apriamo con
        //  System.IO.Compression e leggiamo l'XML con System.Xml.Linq,
        //  entrambi inclusi nel .NET Framework.
        // ============================================================

        private Dictionary<string, ExcelRow> LoadExcel(string path)
        {
            var result = new Dictionary<string, ExcelRow>();

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

                    // Trova l'intestazione della tabella giochi: la riga che contiene
                    // contemporaneamente Developer, Publisher e Genre.
                    int headerIdx = -1, devCol = 0, pubCol = 0, genCol = 0, nameCol = 0;
                    for (int i = 0; i < rows.Count; i++)
                    {
                        int d = FindCol(rows[i], ColDeveloper);
                        int p = FindCol(rows[i], ColPublisher);
                        int g = FindCol(rows[i], ColGenre);
                        if (d > 0 && p > 0 && g > 0)
                        {
                            headerIdx = i;
                            devCol = d; pubCol = p; genCol = g;
                            nameCol = rows[i].Keys.Min(); // prima colonna usata = nome gioco
                            break;
                        }
                    }

                    if (headerIdx < 0)
                    {
                        continue; // foglio senza tabella giochi
                    }

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

                        result[key] = new ExcelRow
                        {
                            Publisher = Get(rows[i], pubCol),
                            Developer = Get(rows[i], devCol),
                            Genre = Get(rows[i], genCol)
                        };
                    }
                }
            }

            return result;
        }

        // Tabella delle stringhe condivise (xl/sharedStrings.xml)
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
                    // Concatena tutti i <t> (gestisce anche il rich text <r><t>)
                    var text = string.Concat(si.Descendants()
                        .Where(e => e.Name.LocalName == "t")
                        .Select(e => e.Value));
                    list.Add(text);
                }
            }
            return list;
        }

        // Elenco fogli in ordine: nome -> percorso file (es. "worksheets/sheet1.xml")
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
                        // Normalizza eventuali percorsi assoluti tipo "/xl/worksheets/..."
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

        // Legge le righe di un foglio: per ogni riga, dizionario colonna(1-based) -> testo
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

                        if (t == "s") // stringa condivisa
                        {
                            var v = c.Elements().FirstOrDefault(e => e.Name.LocalName == "v");
                            value = (v != null && int.TryParse(v.Value, out int idx) && idx >= 0 && idx < shared.Count)
                                ? shared[idx] : "";
                        }
                        else if (t == "inlineStr") // stringa inline
                        {
                            value = string.Concat(c.Descendants()
                                .Where(e => e.Name.LocalName == "t")
                                .Select(e => e.Value));
                        }
                        else // numero o stringa formula
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

        // --- Helper: get-or-create per evitare duplicati nel database ---

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

        // --- Helper: normalizzazione e split ---

        private static string Normalize(string s)
        {
            if (string.IsNullOrWhiteSpace(s))
            {
                return string.Empty;
            }

            s = s.Trim().ToLowerInvariant();

            // Apostrofi e virgolette "ricurvi" -> dritti (es. Demon's Souls)
            s = s.Replace('\u2019', '\'').Replace('\u2018', '\'')
                 .Replace('\u201C', '"').Replace('\u201D', '"');

            // Simboli di marchio
            s = s.Replace("\u2122", "").Replace("\u00AE", "").Replace("\u00A9", "");

            while (s.Contains("  "))
            {
                s = s.Replace("  ", " ");
            }

            return s.Trim();
        }

        // Publisher/Developer multipli separati da ';'
        private static IEnumerable<string> SplitCompanies(string value)
        {
            return value.Split(';')
                .Select(x => x.Trim())
                .Where(x => !string.IsNullOrEmpty(x));
        }

        // Generi multipli separati da ',' o ';'
        private static IEnumerable<string> SplitGenres(string value)
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
        }
    }
}