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
        private MainWindow _mainWindow;

        public ComponentMappingManager(MainWindow mainWindow, string excelFileName)
        {
            _mainWindow = mainWindow;
            _mappings = new Dictionary<string, ComponentMapping>();

            // Create unique filename based on Excel file
            var fileNameWithoutExt = Path.GetFileNameWithoutExtension(excelFileName);
            _mappingFileName = $"{fileNameWithoutExt}_ComponentMapping.json";

            LoadMappings();
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

        public void RemoveMapping(string excelReference)
        {
            var cleanRef = excelReference.Trim();
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
            return _mappings.ContainsKey(cleanRef) ||
                   _mappings.Keys.Any(key => key.EndsWith(":") && cleanRef.StartsWith(key, StringComparison.OrdinalIgnoreCase)) ||
                   _mappings.Keys.Any(key => cleanRef.Contains(key, StringComparison.OrdinalIgnoreCase));
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