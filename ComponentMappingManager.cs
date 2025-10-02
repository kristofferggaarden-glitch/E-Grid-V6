using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text.Json;
using System.Windows;

namespace WpfEGridApp
{
    public class ComponentMapping
    {
        public string ExcelReference { get; set; }
        public int GridRow { get; set; }
        public int GridColumn { get; set; }
        public bool DefaultToBottom { get; set; }
    }

    public class ComponentMappingManager
    {
        private Dictionary<string, ComponentMapping> _mappings;
        private readonly string _mappingFileName;
        private readonly string _bulkMappingFileName;
        private readonly MainWindow _mainWindow;

        private readonly List<BulkRangeMapping> _bulkRangeMappings = new List<BulkRangeMapping>();

        public ComponentMappingManager(MainWindow mainWindow, string excelFileName)
        {
            _mainWindow = mainWindow;
            _mappings = new Dictionary<string, ComponentMapping>(StringComparer.OrdinalIgnoreCase);

            var fileNameWithoutExt = Path.GetFileNameWithoutExtension(excelFileName);
            _mappingFileName = $"{fileNameWithoutExt}_ComponentMapping.json";
            _bulkMappingFileName = $"{fileNameWithoutExt}_BulkMapping.json";

            LoadMappings();
            LoadBulkMappings();
        }

        // Hjelpemetode: Sjekk om cleanRef matches mappingKey smart
        private bool IsSmartMatch(string cleanRef, string mappingKey)
        {
            // Eksakt match først
            if (cleanRef.Equals(mappingKey, StringComparison.OrdinalIgnoreCase))
                return true;

            // Prefix match: mappingKey må ende med ":" og cleanRef må starte med den
            if (mappingKey.EndsWith(":") && cleanRef.StartsWith(mappingKey, StringComparison.OrdinalIgnoreCase))
                return true;

            // Contains match: MEN kun hvis cleanRef inneholder mappingKey som HELT ord/segment
            // F.eks: "J03-F3:11" skal IKKE matche "J03-F3:1" 
            // Men "J01-X1:1" skal matche "J01-X1" (hvis mappet som prefix)

            // Unngå delstring-problem: Sjekk om det er et word boundary etter match
            int index = cleanRef.IndexOf(mappingKey, StringComparison.OrdinalIgnoreCase);
            if (index >= 0)
            {
                int endIndex = index + mappingKey.Length;
                // Match er OK hvis:
                // - Det er slutten av strengen, ELLER
                // - Neste tegn er ikke et siffer (unngår "F3:1" matcher "F3:11")
                if (endIndex >= cleanRef.Length || !char.IsDigit(cleanRef[endIndex]))
                    return true;
            }

            return false;
        }

        public bool HasMapping(string excelReference)
        {
            var cleanRef = excelReference.Trim();

            foreach (var mappingKey in _mappings.Keys)
            {
                if (IsSmartMatch(cleanRef, mappingKey))
                    return true;
            }

            var bulkMapping = GetBulkRangeMappingForReference(cleanRef);
            if (bulkMapping != null)
                return true;

            return false;
        }

        public ComponentMapping GetMapping(string excelReference)
        {
            if (string.IsNullOrWhiteSpace(excelReference))
                return null;

            var cleanRef = excelReference.Trim();

            // 1. EKSAKT match først (høyest prioritet)
            if (_mappings.TryGetValue(cleanRef, out var mapping))
            {
                return new ComponentMapping
                {
                    ExcelReference = mapping.ExcelReference,
                    GridRow = mapping.GridRow,
                    GridColumn = mapping.GridColumn,
                    DefaultToBottom = mapping.DefaultToBottom
                };
            }

            // 2. Smart match logikk
            ComponentMapping bestMatch = null;
            int bestMatchLength = 0;

            foreach (var kvp in _mappings)
            {
                var mappingKey = kvp.Key;
                var mappingValue = kvp.Value;

                if (IsSmartMatch(cleanRef, mappingKey) && mappingKey.Length > bestMatchLength)
                {
                    bestMatch = new ComponentMapping
                    {
                        ExcelReference = mappingValue.ExcelReference,
                        GridRow = mappingValue.GridRow,
                        GridColumn = mappingValue.GridColumn,
                        DefaultToBottom = mappingValue.DefaultToBottom
                    };
                    bestMatchLength = mappingKey.Length;
                }
            }

            return bestMatch;
        }

        public void AddMapping(string excelReference, int gridRow, int gridCol, bool defaultToBottom = false)
        {
            var cleanRef = excelReference.Trim();

            bool hasStar = cleanRef.EndsWith("*");
            string baseRef = hasStar ? cleanRef.TrimEnd('*') : cleanRef;
            string starRef = hasStar ? cleanRef : cleanRef + "*";

            _mappings[cleanRef] = new ComponentMapping
            {
                ExcelReference = cleanRef,
                GridRow = gridRow,
                GridColumn = gridCol,
                DefaultToBottom = defaultToBottom
            };

            if (hasStar)
            {
                _mappings[baseRef] = new ComponentMapping
                {
                    ExcelReference = baseRef,
                    GridRow = gridRow,
                    GridColumn = gridCol,
                    DefaultToBottom = !defaultToBottom
                };
            }
            else
            {
                _mappings[starRef] = new ComponentMapping
                {
                    ExcelReference = starRef,
                    GridRow = gridRow,
                    GridColumn = gridCol,
                    DefaultToBottom = !defaultToBottom
                };
            }

            SaveMappings();
        }

        public class CellCoord
        {
            public int Row { get; set; }
            public int Col { get; set; }
        }

        public class BulkRangeMapping
        {
            public string Prefix { get; set; }
            public int StartIndex { get; set; }
            public int EndIndex { get; set; }
            public List<CellCoord> Cells { get; set; } = new List<CellCoord>();
            public bool SelectedIsTop { get; set; }
        }

        public void AddBulkRangeMapping(string prefix, int startIndex, int endIndex, List<(int Row, int Col)> selectedCells, bool selectedIsTop)
        {
            if (string.IsNullOrWhiteSpace(prefix)) throw new ArgumentException(nameof(prefix));
            if (endIndex < startIndex) throw new ArgumentException("endIndex < startIndex");
            if (selectedCells == null || selectedCells.Count == 0) throw new ArgumentException("selectedCells");

            _bulkRangeMappings.RemoveAll(b =>
                b.Prefix.Equals(prefix, StringComparison.OrdinalIgnoreCase) &&
                b.StartIndex == startIndex &&
                b.EndIndex == endIndex);

            var bulk = new BulkRangeMapping
            {
                Prefix = prefix,
                StartIndex = startIndex,
                EndIndex = endIndex,
                Cells = selectedCells
                    .Distinct()
                    .Select(c => new CellCoord { Row = c.Row, Col = c.Col })
                    .ToList(),
                SelectedIsTop = selectedIsTop
            };

            _bulkRangeMappings.Add(bulk);

            for (int i = startIndex; i <= endIndex; i++)
            {
                string baseReference = $"{prefix}:{i}";
                string starReference = $"{prefix}:{i}*";

                var firstCell = selectedCells.First();

                _mappings[baseReference] = new ComponentMapping
                {
                    ExcelReference = baseReference,
                    GridRow = firstCell.Row,
                    GridColumn = firstCell.Col,
                    DefaultToBottom = !selectedIsTop
                };

                int oppositeRow = selectedIsTop ? firstCell.Row + 1 : firstCell.Row - 1;

                _mappings[starReference] = new ComponentMapping
                {
                    ExcelReference = starReference,
                    GridRow = oppositeRow,
                    GridColumn = firstCell.Col,
                    DefaultToBottom = selectedIsTop
                };
            }

            SaveBulkMappings();
            SaveMappings();
        }

        public void RemoveBulkRangeMapping(string prefix, int startIndex, int endIndex)
        {
            _bulkRangeMappings.RemoveAll(b =>
                b.Prefix.Equals(prefix, StringComparison.OrdinalIgnoreCase) &&
                b.StartIndex == startIndex &&
                b.EndIndex == endIndex);

            for (int i = startIndex; i <= endIndex; i++)
            {
                string baseRef = $"{prefix}:{i}";
                string starRef = $"{prefix}:{i}*";
                _mappings.Remove(baseRef);
                _mappings.Remove(starRef);
            }

            SaveBulkMappings();
            SaveMappings();
        }

        public List<BulkRangeMapping> GetAllBulkRanges() => _bulkRangeMappings.ToList();

        public List<ComponentMapping> GetAllMappingsIncludingBulk()
        {
            var allMappings = new List<ComponentMapping>();

            var bulkReferences = new HashSet<string>();
            foreach (var bulk in _bulkRangeMappings)
            {
                for (int i = bulk.StartIndex; i <= bulk.EndIndex; i++)
                {
                    bulkReferences.Add($"{bulk.Prefix}:{i}");
                    bulkReferences.Add($"{bulk.Prefix}:{i}*");
                }
            }

            foreach (var mapping in _mappings.Values)
            {
                if (!bulkReferences.Contains(mapping.ExcelReference))
                {
                    allMappings.Add(mapping);
                }
            }

            foreach (var bulk in _bulkRangeMappings)
            {
                var displayMapping = new ComponentMapping
                {
                    ExcelReference = $"{bulk.Prefix}:{bulk.StartIndex}-{bulk.EndIndex}",
                    GridRow = -99,
                    GridColumn = bulk.Cells.Count,
                    DefaultToBottom = !bulk.SelectedIsTop
                };
                allMappings.Add(displayMapping);
            }

            return allMappings;
        }

        public BulkRangeMapping GetBulkRangeMappingForReference(string excelReference)
        {
            if (string.IsNullOrWhiteSpace(excelReference)) return null;
            var cleaned = excelReference.Trim().TrimEnd('*');

            var lastColonIndex = cleaned.LastIndexOf(':');
            if (lastColonIndex == -1) return null;

            var prefix = cleaned.Substring(0, lastColonIndex);
            var numPart = cleaned.Substring(lastColonIndex + 1);

            if (!int.TryParse(numPart, out var index)) return null;

            return _bulkRangeMappings.FirstOrDefault(b =>
                b.Prefix.Equals(prefix, StringComparison.OrdinalIgnoreCase) &&
                index >= b.StartIndex && index <= b.EndIndex);
        }

        public bool IsReferenceInBulkRange(string reference)
        {
            return GetBulkRangeMappingForReference(reference) != null;
        }

        private void SaveBulkMappings()
        {
            try
            {
                var json = JsonSerializer.Serialize(_bulkRangeMappings, new JsonSerializerOptions
                {
                    WriteIndented = true
                });
                File.WriteAllText(_bulkMappingFileName, json);
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Kunne ikke lagre bulk mappings: {ex.Message}", "Feil",
                              MessageBoxButton.OK, MessageBoxImage.Warning);
            }
        }

        private void LoadBulkMappings()
        {
            try
            {
                if (File.Exists(_bulkMappingFileName))
                {
                    var json = File.ReadAllText(_bulkMappingFileName);
                    var list = JsonSerializer.Deserialize<List<BulkRangeMapping>>(json);
                    _bulkRangeMappings.Clear();
                    if (list != null)
                        _bulkRangeMappings.AddRange(list);
                }
            }
            catch
            {
                _bulkRangeMappings.Clear();
            }
        }

        public void RemoveMapping(string excelReference)
        {
            var cleanRef = excelReference.Trim();

            if (cleanRef.Contains("-"))
            {
                var lastColonIndex = cleanRef.LastIndexOf(':');
                if (lastColonIndex != -1)
                {
                    var prefix = cleanRef.Substring(0, lastColonIndex);
                    var rangePart = cleanRef.Substring(lastColonIndex + 1);
                    var rangeParts = rangePart.Split('-');

                    if (rangeParts.Length == 2 &&
                        int.TryParse(rangeParts[0], out var start) &&
                        int.TryParse(rangeParts[1], out var end))
                    {
                        RemoveBulkRangeMapping(prefix, start, end);
                        return;
                    }
                }
            }

            bool hasStar = cleanRef.EndsWith("*");
            string baseRef = hasStar ? cleanRef.TrimEnd('*') : cleanRef;
            string starRef = hasStar ? cleanRef : cleanRef + "*";

            _mappings.Remove(cleanRef);
            _mappings.Remove(hasStar ? baseRef : starRef);

            SaveMappings();
        }

        public bool IsPositionMapped(int row, int col, bool isBottomSide)
        {
            return _mappings.Values.Any(m => m.GridRow == row && m.GridColumn == col && m.DefaultToBottom == isBottomSide);
        }

        public bool IsReferenceMapped(string reference)
        {
            if (string.IsNullOrWhiteSpace(reference))
                return false;

            var cleanRef = reference.Trim();

            foreach (var mappingKey in _mappings.Keys)
            {
                if (IsSmartMatch(cleanRef, mappingKey))
                    return true;
            }

            return GetBulkRangeMappingForReference(cleanRef) != null;
        }

        public List<ComponentMapping> GetAllMappings()
        {
            return _mappings.Values.ToList();
        }

        private void SaveMappings()
        {
            try
            {
                var json = JsonSerializer.Serialize(_mappings, new JsonSerializerOptions
                {
                    WriteIndented = true
                });
                File.WriteAllText(_mappingFileName, json);
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Kunne ikke lagre mappings: {ex.Message}", "Feil",
                              MessageBoxButton.OK, MessageBoxImage.Warning);
            }
        }

        private void LoadMappings()
        {
            try
            {
                if (File.Exists(_mappingFileName))
                {
                    var json = File.ReadAllText(_mappingFileName);
                    _mappings = JsonSerializer.Deserialize<Dictionary<string, ComponentMapping>>(json)
                               ?? new Dictionary<string, ComponentMapping>();
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Kunne ikke laste mappings: {ex.Message}", "Feil",
                              MessageBoxButton.OK, MessageBoxImage.Warning);
                _mappings = new Dictionary<string, ComponentMapping>();
            }
        }
    }
}