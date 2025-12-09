using ExcelParsing.Model;
using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO.Compression;
using System.Linq;
using System.Reflection;
using System.Text;
using System.Threading.Tasks;
using System.Xml;

namespace ExcelParsing.Core
{
    public static class ExcelReader
    {
        /// <summary>
        /// Reads an Excel file and maps it to a list of entities
        /// </summary>
        public static List<T> ReadToEntities<T>(string filePath, int sheetIndex = 0, bool hasHeader = true, ColumnMapping mapping = null)
            where T : class, new()
        {
            if (!File.Exists(filePath))
                throw new FileNotFoundException($"Excel file not found: {filePath}");

            using (var stream = File.OpenRead(filePath))
            {
                return ReadToEntities<T>(stream, sheetIndex, hasHeader, mapping);
            }
        }

        /// <summary>
        /// Reads an Excel stream and maps it to a list of entities
        /// </summary>
        public static List<T> ReadToEntities<T>(Stream stream, int sheetIndex = 0, bool hasHeader = true, ColumnMapping mapping = null)
            where T : class, new()
        {
            var result = new List<T>();

            try
            {
                using (var archive = new ZipArchive(stream, ZipArchiveMode.Read, leaveOpen: false))
                {
                    var sharedStrings = LoadSharedStrings(archive);
                    var sheets = GetSheetList(archive);

                    if (sheetIndex >= sheets.Count)
                        return result;

                    var sheetPath = sheets[sheetIndex];
                    var sheetEntry = archive.GetEntry(sheetPath);

                    if (sheetEntry == null)
                        return result;

                    using (var sheetStream = sheetEntry.Open())
                    {
                        var doc = new XmlDocument();
                        doc.Load(sheetStream);

                        var rows = doc.GetElementsByTagName("row");

                        if (rows.Count == 0)
                            return result;

                        var properties = GetEntityProperties<T>();
                        var columnHeaders = new List<string>();
                        var propertyMapping = new Dictionary<int, PropertyInfo>(); // Column index to property
                        int startRow = 0;

                        if (hasHeader && rows.Count > 0)
                        {
                            var headerRow = rows[0];
                            var cells = headerRow.SelectNodes(".//*[local-name()='c']");

                            for (int i = 0; i < cells.Count; i++)
                            {
                                try
                                {
                                    var cellNode = cells[i];
                                    var cellValue = ExtractCellValue(cellNode, sharedStrings);
                                    columnHeaders.Add(cellValue ?? $"Column{i}");
                                }
                                catch
                                {
                                    columnHeaders.Add($"Column{i}");
                                }
                            }
                            startRow = 1;
                        }
                        else
                        {
                            if (rows.Count > 0)
                            {
                                var firstRow = rows[0];
                                var cells = firstRow.SelectNodes(".//*[local-name()='c']");
                                for (int i = 0; i < cells.Count; i++)
                                {
                                    columnHeaders.Add($"Column{i}");
                                }
                            }
                        }

                        for (int colIndex = 0; colIndex < columnHeaders.Count; colIndex++)
                        {
                            var columnName = columnHeaders[colIndex];
                            PropertyInfo property = null;

                            if (mapping != null && mapping.Map.ContainsKey(columnName))
                            {
                                var propertyName = mapping.Map[columnName];
                                property = properties.FirstOrDefault(p =>
                                    string.Equals(p.Name, propertyName, StringComparison.OrdinalIgnoreCase));
                            }

                            if (property == null)
                            {
                                property = properties.FirstOrDefault(p =>
                                    string.Equals(p.Name, columnName, StringComparison.OrdinalIgnoreCase));
                            }

                            if (property != null)
                            {
                                propertyMapping[colIndex] = property;
                            }
                        }

                        for (int rowIndex = startRow; rowIndex < rows.Count; rowIndex++)
                        {
                            try
                            {
                                var row = rows[rowIndex];
                                var entity = new T();
                                bool hasData = false;

                                var cells = row.SelectNodes(".//*[local-name()='c']");

                                for (int colIndex = 0; colIndex < cells.Count; colIndex++)
                                {
                                    try
                                    {
                                        if (!propertyMapping.ContainsKey(colIndex))
                                            continue;

                                        var cellNode = cells[colIndex];
                                        var cellValue = ExtractCellValue(cellNode, sharedStrings);

                                        if (cellValue == null)
                                            continue;

                                        var property = propertyMapping[colIndex];
                                        var convertedValue = ConvertValue(cellValue, property.PropertyType);

                                        if (convertedValue != null)
                                        {
                                            property.SetValue(entity, convertedValue);
                                            hasData = true;
                                        }
                                    }
                                    catch
                                    {
                                        continue;
                                    }
                                }

                                if (hasData)
                                {
                                    result.Add(entity);
                                }
                            }
                            catch
                            {
                                continue;
                            }
                        }
                    }
                }
            }
            catch (Exception ex)
            {

                System.Diagnostics.Debug.WriteLine($"Excel reading error: {ex.Message}");
                return result;
            }

            return result;
        }

        /// <summary>
        /// Loads shared strings from the workbook
        /// </summary>
        private static Dictionary<int, string> LoadSharedStrings(ZipArchive archive)
        {
            var strings = new Dictionary<int, string>();
            var entry = archive.GetEntry("xl/sharedStrings.xml");

            if (entry == null)
                return strings;

            try
            {
                using (var stream = entry.Open())
                {
                    var doc = new XmlDocument();
                    doc.Load(stream);

                    var siNodes = doc.GetElementsByTagName("si");
                    for (int i = 0; i < siNodes.Count; i++)
                    {
                        var siNode = siNodes[i];
                        var tNode = siNode.SelectSingleNode(".//*[local-name()='t']");
                        strings[i] = tNode?.InnerText ?? string.Empty;
                    }
                }
            }
            catch
            {

            }

            return strings;
        }

        /// <summary>
        /// Gets the list of sheet paths
        /// </summary>
        private static List<string> GetSheetList(ZipArchive archive)
        {
            var sheets = new List<string>();

            try
            {
                var workbookEntry = archive.GetEntry("xl/workbook.xml");
                if (workbookEntry == null)
                    return sheets;

                using (var stream = workbookEntry.Open())
                {
                    var doc = new XmlDocument();
                    doc.Load(stream);

                    var sheetNodes = doc.GetElementsByTagName("sheet");

                    var relsEntry = archive.GetEntry("xl/_rels/workbook.xml.rels");
                    var rels = new Dictionary<string, string>();

                    if (relsEntry != null)
                    {
                        using (var relsStream = relsEntry.Open())
                        {
                            var relsDoc = new XmlDocument();
                            relsDoc.Load(relsStream);
                            var relNodes = relsDoc.GetElementsByTagName("Relationship");

                            for (int i = 0; i < relNodes.Count; i++)
                            {
                                var rel = relNodes[i];
                                var id = rel.Attributes["Id"]?.Value;
                                var target = rel.Attributes["Target"]?.Value;

                                if (!string.IsNullOrEmpty(id) && !string.IsNullOrEmpty(target))
                                {
                                    rels[id] = target;
                                }
                            }
                        }
                    }

                    for (int i = 0; i < sheetNodes.Count; i++)
                    {
                        var sheetNode = sheetNodes[i];
                        var rId = sheetNode.Attributes["id"]?.Value;

                        if (string.IsNullOrEmpty(rId))
                        {
                            var rIdAttr = sheetNode.Attributes["id", "http://schemas.openxmlformats.org/officeDocument/2006/relationships"];
                            rId = rIdAttr?.Value;
                        }

                        if (!string.IsNullOrEmpty(rId) && rels.ContainsKey(rId))
                        {
                            var target = rels[rId];
                            if (!target.StartsWith("xl/"))
                                target = "xl/" + target;

                            sheets.Add(target);
                        }
                    }
                }
            }
            catch
            {

            }

            return sheets;
        }

        /// <summary>
        /// Extracts the value from a cell node
        /// </summary>
        private static string ExtractCellValue(XmlNode cellNode, Dictionary<int, string> sharedStrings)
        {
            var typeAttr = cellNode.Attributes["t"]?.Value ?? string.Empty;
            var vNode = cellNode.SelectSingleNode(".//*[local-name()='v']");
            var value = vNode?.InnerText ?? string.Empty;

            if (string.IsNullOrEmpty(value))
                return null;

            try
            {
                switch (typeAttr)
                {
                    case "s":
                        if (int.TryParse(value, out int stringIndex) && sharedStrings.ContainsKey(stringIndex))
                            return sharedStrings[stringIndex];
                        return value;

                    case "b":
                        return value == "1" ? "TRUE" : "FALSE";

                    case "e":
                        return null;

                    case "d":
                        return value;

                    default:
                        return value;
                }
            }
            catch
            {
                return value;
            }
        }

        /// <summary>
        /// Gets writable properties for an entity type
        /// </summary>
        private static PropertyInfo[] GetEntityProperties<T>() where T : class
        {
            return typeof(T)
                .GetProperties(BindingFlags.Public | BindingFlags.IgnoreCase | BindingFlags.Instance)
                .Where(p => p.CanWrite)
                .ToArray();
        }

        /// <summary>
        /// Converts a cell value to the target property type
        /// </summary>
        private static object ConvertValue(object cellValue, Type targetType)
        {
            if (cellValue == null)
                return null;

            var stringValue = cellValue.ToString().Trim();

            if (string.IsNullOrEmpty(stringValue))
                return null;

            if (targetType.IsAssignableFrom(cellValue.GetType()))
                return cellValue;

            var underlyingType = Nullable.GetUnderlyingType(targetType) ?? targetType;

            try
            {
                if (underlyingType == typeof(string))
                    return stringValue;

                if (underlyingType == typeof(int))
                {
                    if (int.TryParse(stringValue, NumberStyles.Any, CultureInfo.InvariantCulture, out int intValue))
                        return intValue;
                    if (double.TryParse(stringValue, NumberStyles.Any, CultureInfo.InvariantCulture, out double doubleVal))
                        return (int)doubleVal;
                    return null;
                }

                if (underlyingType == typeof(long))
                {
                    if (long.TryParse(stringValue, NumberStyles.Any, CultureInfo.InvariantCulture, out long longValue))
                        return longValue;
                    if (double.TryParse(stringValue, NumberStyles.Any, CultureInfo.InvariantCulture, out double doubleVal))
                        return (long)doubleVal;
                    return null;
                }

                if (underlyingType == typeof(short))
                {
                    if (short.TryParse(stringValue, NumberStyles.Any, CultureInfo.InvariantCulture, out short shortValue))
                        return shortValue;
                    if (double.TryParse(stringValue, NumberStyles.Any, CultureInfo.InvariantCulture, out double doubleVal))
                        return (short)doubleVal;
                    return null;
                }

                if (underlyingType == typeof(byte))
                {
                    if (byte.TryParse(stringValue, NumberStyles.Any, CultureInfo.InvariantCulture, out byte byteValue))
                        return byteValue;
                    if (double.TryParse(stringValue, NumberStyles.Any, CultureInfo.InvariantCulture, out double doubleVal))
                        return (byte)doubleVal;
                    return null;
                }

                if (underlyingType == typeof(decimal))
                {
                    if (decimal.TryParse(stringValue, NumberStyles.Any, CultureInfo.InvariantCulture, out decimal decimalValue))
                        return decimalValue;
                    return null;
                }

                if (underlyingType == typeof(double))
                {
                    if (double.TryParse(stringValue, NumberStyles.Any, CultureInfo.InvariantCulture, out double doubleValue))
                        return doubleValue;
                    return null;
                }

                if (underlyingType == typeof(float))
                {
                    if (float.TryParse(stringValue, NumberStyles.Any, CultureInfo.InvariantCulture, out float floatValue))
                        return floatValue;
                    return null;
                }

                if (underlyingType == typeof(bool))
                {
                    if (bool.TryParse(stringValue, out bool boolValue))
                        return boolValue;
                    if (stringValue == "1" || stringValue.Equals("TRUE", StringComparison.OrdinalIgnoreCase) || stringValue.Equals("YES", StringComparison.OrdinalIgnoreCase))
                        return true;
                    if (stringValue == "0" || stringValue.Equals("FALSE", StringComparison.OrdinalIgnoreCase) || stringValue.Equals("NO", StringComparison.OrdinalIgnoreCase))
                        return false;
                    return null;
                }

                if (underlyingType == typeof(DateTime))
                {
                    if (double.TryParse(stringValue, NumberStyles.Any, CultureInfo.InvariantCulture, out double excelDate))
                    {
                        if (excelDate >= 0 && excelDate <= 2958465) // Valid OLE date range
                        {
                            try
                            {
                                return DateTime.FromOADate(excelDate);
                            }
                            catch
                            {

                            }
                        }
                    }

                    if (DateTime.TryParse(stringValue, CultureInfo.InvariantCulture, DateTimeStyles.None, out DateTime dateValue))
                        return dateValue;

                    return null;
                }

                if (underlyingType == typeof(Guid))
                {
                    if (Guid.TryParse(stringValue, out Guid guidValue))
                        return guidValue;
                    return null;
                }

                return null;
            }
            catch
            {
                return null;
            }
        }
    }
}
