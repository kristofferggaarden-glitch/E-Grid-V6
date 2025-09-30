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

        public bool HasMapping(string excelReference)
        {
            var cleanRef = excelReference.Trim();

            if (_mappings.ContainsKey(cleanRef))
                return true;

            foreach (var mappingKey in _mappings.Keys)
            {
                if (mappingKey.EndsWith(":") && cleanRef.StartsWith(mappingKey))
                    return true;

                if (cleanRef.Contains(mappingKey, StringComparison.OrdinalIgnoreCase))
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

            ComponentMapping bestMatch = null;
            int bestMatchLength = 0;

            foreach (var kvp in _mappings)
            {
                var mappingKey = kvp.Key;
                var mappingValue = kvp.Value;

                if (mappingKey.EndsWith(":") && cleanRef.StartsWith(mappingKey, StringComparison.OrdinalIgnoreCase))
                {
                    if (mappingKey.Length > bestMatchLength)
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
                else if (cleanRef.Contains(mappingKey, StringComparison.OrdinalIgnoreCase))
                {
                    if (mappingKey.Length > bestMatchLength)
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

        /// <summary>
        /// FIKSET: Legger til bulk range mapping og oppretter automatisk individuelle mappings
        /// for hver referanse (base og starred) til ALLE valgte celler.
        /// </summary>
        public void AddBulkRangeMapping(string prefix, int startIndex, int endIndex, List<(int Row, int Col)> selectedCells, bool selectedIsTop)
        {
            if (string.IsNullOrWhiteSpace(prefix)) throw new ArgumentException(nameof(prefix));
            if (endIndex < startIndex) throw new ArgumentException("endIndex < startIndex");
            if (selectedCells == null || selectedCells.Count == 0) throw new ArgumentException("selectedCells");

            // Fjern eksisterende bulk mapping for samme range
            _bulkRangeMappings.RemoveAll(b =>
                b.Prefix.Equals(prefix, StringComparison.OrdinalIgnoreCase) &&
                b.StartIndex == startIndex &&
                b.EndIndex == endIndex);

            // Opprett bulk range mapping
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

            // VIKTIG: Opprett individuelle mappings for hver referanse i bulk range
            // Dette gjør at GetMapping() kan finne dem!
            for (int i = startIndex; i <= endIndex; i++)
            {
                string baseReference = $"{prefix}:{i}";
                string starReference = $"{prefix}:{i}*";

                // Map base reference til alle valgte celler
                // Vi bruker første celle som "hovedmapping", men lagrer alle i bulk
                var firstCell = selectedCells.First();

                // Base reference (uten stjerne) mappes til valgt side
                _mappings[baseReference] = new ComponentMapping
                {
                    ExcelReference = baseReference,
                    GridRow = firstCell.Row,
                    GridColumn = firstCell.Col,
                    DefaultToBottom = !selectedIsTop // If T selected, DefaultToBottom=false
                };

                // Starred reference mappes til motsatt side
                // Hvis selectedIsTop (T-celler), så mappes starred til B-siden (row+1)
                int oppositeRow = selectedIsTop ? firstCell.Row + 1 : firstCell.Row - 1;

                _mappings[starReference] = new ComponentMapping
                {
                    ExcelReference = starReference,
                    GridRow = oppositeRow,
                    GridColumn = firstCell.Col,
                    DefaultToBottom = selectedIsTop // Opposite of base
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

            // Fjern også alle individuelle mappings for dette bulk range
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

            // Legg til vanlige mappings, men filtrer bort de som tilhører bulk ranges
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

            // Legg til bulk ranges som egne display items
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
            var parts = cleaned.Split(':');
            if (parts.Length != 2) return null;
            var prefix = parts[0];
            var numPart = parts[1];
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

            // Sjekk om dette er en bulk range referanse (format: X2:21-100)
            if (cleanRef.Contains("-"))
            {
                var parts = cleanRef.Split(':');
                if (parts.Length == 2)
                {
                    var prefix = parts[0];
                    var rangeParts = parts[1].Split('-');
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

            if (_mappings.ContainsKey(cleanRef) ||
                _mappings.Keys.Any(key => key.EndsWith(":") && cleanRef.StartsWith(key, StringComparison.OrdinalIgnoreCase)) ||
                _mappings.Keys.Any(key => cleanRef.Contains(key, StringComparison.OrdinalIgnoreCase)))
            {
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