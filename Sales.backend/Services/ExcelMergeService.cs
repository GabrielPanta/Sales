using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.IO.Compression;
using System.Linq;
using System.Security;
using System.Text;
using System.Text.RegularExpressions;
using System.Web;
using System.Xml.Linq;

namespace Sales.backend.Services
{
    public class ExcelMergeService
    {
        private static readonly string[] KeyCandidates = { "documento", "dni", "cedula", "identificacion", "id", "codigo", "legajo", "empleado", "trabajador" };
        private static readonly string[] ActivityCandidates = { "actividad", "labor", "labores", "estado", "situacion", "observacion" };

        public byte[] Merge(HttpPostedFileBase trabajadores, HttpPostedFileBase ultimoDiaLaborado, HttpPostedFileBase marcaciones)
        {
            var workerRows = ReadRows(trabajadores.InputStream);
            var lastDayRows = ReadRows(ultimoDiaLaborado.InputStream);
            var punchRows = ReadRows(marcaciones.InputStream);

            if (!workerRows.Any()) throw new InvalidOperationException("El archivo de trabajadores no contiene registros.");

            var workerKey = FindColumn(workerRows, KeyCandidates) ?? throw new InvalidOperationException("No se encontro una columna de identificacion en trabajadores (Documento, DNI, Cedula, Id, Codigo, etc.).");
            var lastDayKey = FindColumn(lastDayRows, KeyCandidates) ?? workerKey;
            var punchKey = FindColumn(punchRows, KeyCandidates) ?? workerKey;
            var activityColumn = FindColumn(lastDayRows, ActivityCandidates);

            var lastDayByKey = lastDayRows
                .Where(r => r.ContainsKey(lastDayKey) && !string.IsNullOrWhiteSpace(r[lastDayKey]))
                .GroupBy(r => NormalizeKey(r[lastDayKey]))
                .ToDictionary(g => g.Key, g => g.First());
            var punchedKeys = new HashSet<string>(punchRows
                .Where(r => r.ContainsKey(punchKey) && !string.IsNullOrWhiteSpace(r[punchKey]))
                .Select(r => NormalizeKey(r[punchKey])));

            var headers = new List<string>();
            AddHeaders(headers, workerRows.SelectMany(r => r.Keys), "Trabajador");
            AddHeaders(headers, lastDayRows.SelectMany(r => r.Keys), "UltimoDia");
            headers.Add("Estado");

            var outputRows = new List<IList<string>> { headers };
            foreach (var worker in workerRows)
            {
                var key = worker.ContainsKey(workerKey) ? NormalizeKey(worker[workerKey]) : string.Empty;
                lastDayByKey.TryGetValue(key, out var lastDay);
                var hasPunches = !string.IsNullOrEmpty(key) && punchedKeys.Contains(key);
                var activity = lastDay != null && activityColumn != null && lastDay.ContainsKey(activityColumn) ? lastDay[activityColumn] : string.Empty;
                var row = new List<string>();
                AppendValues(row, headers, worker, "Trabajador");
                AppendValues(row, headers, lastDay, "UltimoDia");
                row.Add(hasPunches ? "ACTIVO" : (string.IsNullOrWhiteSpace(activity) ? "SIN MARCACIONES" : activity));
                outputRows.Add(row);
            }

            return CreateWorkbook(outputRows);
        }

        private static List<Dictionary<string, string>> ReadRows(Stream stream)
        {
            using (var archive = new ZipArchive(stream, ZipArchiveMode.Read, true))
            {
                var sharedStrings = ReadSharedStrings(archive);
                var sheetEntry = archive.GetEntry("xl/worksheets/sheet1.xml") ?? archive.Entries.FirstOrDefault(e => e.FullName.StartsWith("xl/worksheets/sheet", StringComparison.OrdinalIgnoreCase));
                if (sheetEntry == null) return new List<Dictionary<string, string>>();
                var rows = ReadSheetRows(sheetEntry, sharedStrings);
                if (!rows.Any()) return new List<Dictionary<string, string>>();
                var headers = rows.First().Select((h, i) => string.IsNullOrWhiteSpace(h) ? "Columna" + (i + 1) : h.Trim()).ToList();
                return rows.Skip(1).Where(r => r.Any(v => !string.IsNullOrWhiteSpace(v))).Select(r =>
                {
                    var dict = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
                    for (var i = 0; i < headers.Count; i++) dict[headers[i]] = i < r.Count ? r[i] : string.Empty;
                    return dict;
                }).ToList();
            }
        }

        private static List<string> ReadSharedStrings(ZipArchive archive)
        {
            var entry = archive.GetEntry("xl/sharedStrings.xml");
            if (entry == null) return new List<string>();
            using (var reader = new StreamReader(entry.Open()))
            {
                XNamespace ns = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";
                return XDocument.Parse(reader.ReadToEnd()).Descendants(ns + "si").Select(si => string.Concat(si.Descendants(ns + "t").Select(t => t.Value))).ToList();
            }
        }

        private static List<List<string>> ReadSheetRows(ZipArchiveEntry sheetEntry, IList<string> sharedStrings)
        {
            XNamespace ns = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";
            using (var reader = new StreamReader(sheetEntry.Open()))
            {
                return XDocument.Parse(reader.ReadToEnd()).Descendants(ns + "row").Select(row =>
                {
                    var values = new List<string>();
                    foreach (var cell in row.Elements(ns + "c"))
                    {
                        var index = ColumnIndex((string)cell.Attribute("r"));
                        while (values.Count < index) values.Add(string.Empty);
                        var raw = (string)cell.Element(ns + "v") ?? (string)cell.Element(ns + "is")?.Element(ns + "t") ?? string.Empty;
                        values.Add((string)cell.Attribute("t") == "s" && int.TryParse(raw, out var sstIndex) && sstIndex < sharedStrings.Count ? sharedStrings[sstIndex] : raw);
                    }
                    return values;
                }).ToList();
            }
        }

        private static int ColumnIndex(string cellReference)
        {
            var letters = new string((cellReference ?? "A").TakeWhile(char.IsLetter).ToArray()).ToUpperInvariant();
            var sum = 0;
            foreach (var c in letters) sum = sum * 26 + c - 'A' + 1;
            return Math.Max(sum - 1, 0);
        }

        private static string FindColumn(IEnumerable<Dictionary<string, string>> rows, IEnumerable<string> candidates)
        {
            var headers = rows.SelectMany(r => r.Keys).Distinct(StringComparer.OrdinalIgnoreCase).ToList();
            return headers.FirstOrDefault(h => candidates.Any(c => NormalizeHeader(h).Contains(c)));
        }

        private static string NormalizeHeader(string value)
        {
            return Regex.Replace((value ?? string.Empty).Normalize(NormalizationForm.FormD), "[^a-zA-Z0-9]", string.Empty).ToLowerInvariant();
        }

        private static string NormalizeKey(string value)
        {
            return (value ?? string.Empty).Trim().ToUpperInvariant();
        }

        private static void AddHeaders(ICollection<string> target, IEnumerable<string> source, string prefix)
        {
            foreach (var header in source.Distinct(StringComparer.OrdinalIgnoreCase))
            {
                var name = prefix + " - " + header;
                if (!target.Contains(name)) target.Add(name);
            }
        }

        private static void AppendValues(ICollection<string> row, IEnumerable<string> headers, IDictionary<string, string> source, string prefix)
        {
            foreach (var header in headers.Where(h => h.StartsWith(prefix + " - ", StringComparison.Ordinal)))
            {
                var original = header.Substring(prefix.Length + 3);
                row.Add(source != null && source.ContainsKey(original) ? source[original] : string.Empty);
            }
        }

        private static byte[] CreateWorkbook(IList<IList<string>> rows)
        {
            using (var memory = new MemoryStream())
            {
                using (var archive = new ZipArchive(memory, ZipArchiveMode.Create, true))
                {
                    AddEntry(archive, "[Content_Types].xml", "<?xml version=\"1.0\" encoding=\"UTF-8\"?><Types xmlns=\"http://schemas.openxmlformats.org/package/2006/content-types\"><Default Extension=\"rels\" ContentType=\"application/vnd.openxmlformats-package.relationships+xml\"/><Default Extension=\"xml\" ContentType=\"application/xml\"/><Override PartName=\"/xl/workbook.xml\" ContentType=\"application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml\"/><Override PartName=\"/xl/worksheets/sheet1.xml\" ContentType=\"application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml\"/></Types>");
                    AddEntry(archive, "_rels/.rels", "<?xml version=\"1.0\" encoding=\"UTF-8\"?><Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\"><Relationship Id=\"rId1\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument\" Target=\"xl/workbook.xml\"/></Relationships>");
                    AddEntry(archive, "xl/_rels/workbook.xml.rels", "<?xml version=\"1.0\" encoding=\"UTF-8\"?><Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\"><Relationship Id=\"rId1\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet\" Target=\"worksheets/sheet1.xml\"/></Relationships>");
                    AddEntry(archive, "xl/workbook.xml", "<?xml version=\"1.0\" encoding=\"UTF-8\"?><workbook xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\" xmlns:r=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\"><sheets><sheet name=\"Resultado\" sheetId=\"1\" r:id=\"rId1\"/></sheets></workbook>");
                    AddEntry(archive, "xl/worksheets/sheet1.xml", BuildSheetXml(rows));
                }

                return memory.ToArray();
            }
        }

        private static string BuildSheetXml(IList<IList<string>> rows)
        {
            var xml = new StringBuilder("<?xml version=\"1.0\" encoding=\"UTF-8\"?><worksheet xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\"><sheetData>");
            for (var r = 0; r < rows.Count; r++)
            {
                xml.Append("<row r=\"").Append(r + 1).Append("\">");
                for (var c = 0; c < rows[r].Count; c++)
                {
                    xml.Append("<c r=\"").Append(ColumnName(c + 1)).Append(r + 1).Append("\" t=\"inlineStr\"><is><t>").Append(SecurityElement.Escape(rows[r][c] ?? string.Empty)).Append("</t></is></c>");
                }
                xml.Append("</row>");
            }
            return xml.Append("</sheetData></worksheet>").ToString();
        }

        private static string ColumnName(int index)
        {
            var name = string.Empty;
            while (index > 0)
            {
                var modulo = (index - 1) % 26;
                name = Convert.ToChar('A' + modulo) + name;
                index = (index - modulo) / 26;
            }
            return name;
        }

        private static void AddEntry(ZipArchive archive, string path, string content)
        {
            var entry = archive.CreateEntry(path);
            using (var writer = new StreamWriter(entry.Open(), new UTF8Encoding(false))) writer.Write(content);
        }
    }
}
