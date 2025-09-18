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
        public string ExcelReference { get; set; } // F1, X2:, K3, etc.
        public int GridRow { get; set; }
        public int GridColumn { get; set; }
        public bool DefaultToBottom { get; set; } // Default side for this mapping
    }

    public class ComponentMappingManager
    {
        private Dictionary<string, ComponentMapping> _mappings;
        private readonly string _mappingFileName;
        private readonly string _bulkMappingFileName;
        private readonly MainWindow _mainWindow;

        // Liste over bulk-range mappings (f.eks. X2:21-100).  Hver
        // BulkRangeMapping inneholder både start/slutt-indeks og en liste
        // over grid-koordinater som brukeren valgte i bulk-modus.  Dette gjør
        // at vi senere kan beregne lengste avstand mellom venstre og høyre
        // terminalblokk i en ledning ved å prøve alle kombinasjoner av
        // mappede celler.
        private readonly List<BulkRangeMapping> _bulkRangeMappings = new List<BulkRangeMapping>();

        public ComponentMappingManager(MainWindow mainWindow, string excelFileName)
        {
            _mainWindow = mainWindow;
            _mappings = new Dictionary<string, ComponentMapping>(StringComparer.OrdinalIgnoreCase);

            // Lag unike filnavn for lagring basert på Excel-filnavnet
            var fileNameWithoutExt = Path.GetFileNameWithoutExtension(excelFileName);
            _mappingFileName = $"{fileNameWithoutExt}_ComponentMapping.json";
            _bulkMappingFileName = $"{fileNameWithoutExt}_BulkMapping.json";

            LoadMappings();
            LoadBulkMappings();
        }

        public bool HasMapping(string excelReference)
        {
            var cleanRef = excelReference.Trim();

            // Check exact match first
            if (_mappings.ContainsKey(cleanRef))
                return true;

            // Check prefix match for terminal blocks (X20:41 should match X20:)
            foreach (var mappingKey in _mappings.Keys)
            {
                if (mappingKey.EndsWith(":") && cleanRef.StartsWith(mappingKey))
                    return true;

                // Also check if cleanRef contains mappingKey
                if (cleanRef.Contains(mappingKey, StringComparison.OrdinalIgnoreCase))
                    return true;
            }

            // Also check bulk mappings
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

            // Try exact match first
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

            // Try prefix match for terminal blocks and other components
            ComponentMapping bestMatch = null;
            int bestMatchLength = 0;

            foreach (var kvp in _mappings)
            {
                var mappingKey = kvp.Key;
                var mappingValue = kvp.Value;

                // Terminal blocks - exact prefix match (X20: should match X20:41)
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
                // Component match (A2 should match E01-A2-X4:10)
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
            _mappings[cleanRef] = new ComponentMapping
            {
                ExcelReference = cleanRef,
                GridRow = gridRow,
                GridColumn = gridCol,
                DefaultToBottom = defaultToBottom
            };
            SaveMappings();
        }

        // ===================================================================
        //                BULK RANGE MAPPING SUPPORT
        //
        // Bulk-range mappings gjør det mulig å mappe et intervall av
        // terminalblokk-referanser (f.eks. X2:21-100) til flere grid-celler i
        // ett steg.  Når to ledninger i Excel begge refererer til referanser
        // innenfor et bulk-intervall, beregner programmet automatisk lengste
        // avstand mellom alle kombinasjoner av cellene i de to intervallene.

        /// <summary>
        /// Intern struktur som representerer en grid-koordinat valgt under
        /// bulk-mapping.  Rad og kolonne refererer til globale koordinater i
        /// gridet (samme som brukes av Cell i MainWindow).
        /// </summary>
        public class CellCoord
        {
            public int Row { get; set; }
            public int Col { get; set; }
        }

        /// <summary>
        /// Representerer et intervall (prefix:start-end) som er mappet til en
        /// liste av grid-celler.  Prefix er f.eks. "X2", StartIndex og
        /// EndIndex angir tallene (21 og 100 i X2:21-100), og Cells
        /// inneholder en unik liste av koordinater valgt av brukeren.
        /// </summary>
        public class BulkRangeMapping
        {
            public string Prefix { get; set; }
            public int StartIndex { get; set; }
            public int EndIndex { get; set; }
            public List<CellCoord> Cells { get; set; } = new List<CellCoord>();

            /// <summary>
            /// Angir om de valgte cellene representerer oversiden (T) eller undersiden (B) av
            /// rekkeklemmene. True betyr overside (T), false betyr underside (B).  Dette brukes
            /// for å kunne håndtere stjernemerkede referanser som skal mappes på motsatt side.
            /// </summary>
            public bool SelectedIsTop { get; set; }
        }

        /// <summary>
        /// Legger til eller erstatter et bulk-range i samlingen.  Hvis et
        /// eksisterende range med samme prefix/start/end finnes, erstattes det.
        /// Argumentet selectedCells er en liste av (Row, Col)-par hentet fra
        /// MainWindow.  Koordinatene lagres som CellCoord-objekter.  Listen
        /// dedupliseres automatisk.  Kaster hvis argumenter er ugyldige.
        /// </summary>
        public void AddBulkRangeMapping(string prefix, int startIndex, int endIndex, List<(int Row, int Col)> selectedCells, bool selectedIsTop)
        {
            if (string.IsNullOrWhiteSpace(prefix)) throw new ArgumentException(nameof(prefix));
            if (endIndex < startIndex) throw new ArgumentException("endIndex < startIndex");
            if (selectedCells == null || selectedCells.Count == 0) throw new ArgumentException("selectedCells");

            // Fjern eksisterende range med samme identifikator
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
            SaveBulkMappings();
        }

        /// <summary>
        /// Fjerner et bulk-range basert på prefix og indeksområde.
        /// </summary>
        public void RemoveBulkRangeMapping(string prefix, int startIndex, int endIndex)
        {
            _bulkRangeMappings.RemoveAll(b =>
                b.Prefix.Equals(prefix, StringComparison.OrdinalIgnoreCase) &&
                b.StartIndex == startIndex &&
                b.EndIndex == endIndex);
            SaveBulkMappings();
        }

        /// <summary>
        /// Henter alle bulk-ranger for visning eller debugging.
        /// </summary>
        public List<BulkRangeMapping> GetAllBulkRanges() => _bulkRangeMappings.ToList();

        /// <summary>
        /// Get all mappings including bulk range mappings for display
        /// </summary>
        public List<ComponentMapping> GetAllMappingsIncludingBulk()
        {
            var allMappings = new List<ComponentMapping>();

            // Add regular mappings
            allMappings.AddRange(_mappings.Values);

            // Add bulk mappings as virtual mappings for display
            foreach (var bulk in _bulkRangeMappings)
            {
                // Create a display mapping for the bulk range
                var displayMapping = new ComponentMapping
                {
                    ExcelReference = $"{bulk.Prefix}:{bulk.StartIndex}-{bulk.EndIndex}",
                    GridRow = -99, // Special marker for bulk
                    GridColumn = bulk.Cells.Count, // Number of cells selected
                    DefaultToBottom = !bulk.SelectedIsTop
                };
                allMappings.Add(displayMapping);
            }

            return allMappings;
        }

        /// <summary>
        /// Returnerer bulk-range som en gitt referanse tilhører.  Tar høyde
        /// for at referansen kan inneholde stjerne (*), men ser bort fra
        /// den ved sammenligning.  Returnerer null hvis referansen ikke
        /// ligger i et bulk-intervall.
        /// </summary>
        public BulkRangeMapping GetBulkRangeMappingForReference(string excelReference)
        {
            if (string.IsNullOrWhiteSpace(excelReference)) return null;
            var cleaned = excelReference.Trim();
            var parts = cleaned.Split(':');
            if (parts.Length != 2) return null;
            var prefix = parts[0];
            var numPart = parts[1].TrimEnd('*');
            if (!int.TryParse(numPart, out var index)) return null;
            return _bulkRangeMappings.FirstOrDefault(b =>
                b.Prefix.Equals(prefix, StringComparison.OrdinalIgnoreCase) &&
                index >= b.StartIndex && index <= b.EndIndex);
        }

        /// <summary>
        /// Serialiserer listen av bulk-range mappings til fil.  Kalles etter
        /// tillegg, fjerning eller oppdatering av bulk-ranger.  Filen
        /// navngis basert på Excel-filen som ble brukt til å opprette
        /// denne mapping manageren.
        /// </summary>
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
            catch
            {
                // Ignorer skrivefeil; bruker vil bli varslet ved neste
                // lasting hvis filen mangler eller er korrupt
            }
        }

        /// <summary>
        /// Leser bulk-range mappings fra fil.  Kalles i konstruktøren.
        /// Hvis filen ikke finnes eller er korrupt, startes med tom liste.
        /// </summary>
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

        // ===================================================================

        public void RemoveMapping(string excelReference)
        {
            var cleanRef = excelReference.Trim();

            // Check if it's a bulk mapping reference
            if (cleanRef.Contains("-"))
            {
                // Parse bulk mapping format "X2:21-100"
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

            // Regular mapping removal
            if (_mappings.Remove(cleanRef))
            {
                SaveMappings();
            }
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

            // Check regular mappings
            if (_mappings.ContainsKey(cleanRef) ||
                _mappings.Keys.Any(key => key.EndsWith(":") && cleanRef.StartsWith(key, StringComparison.OrdinalIgnoreCase)) ||
                _mappings.Keys.Any(key => cleanRef.Contains(key, StringComparison.OrdinalIgnoreCase)))
            {
                return true;
            }

            // Check bulk mappings
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