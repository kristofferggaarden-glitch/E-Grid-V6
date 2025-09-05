using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.IO;
using System.Linq;
using System.Text.Json;
using System.Text.RegularExpressions;
using System.Windows;
using System.Windows.Controls;
using Microsoft.Win32;
using Excel = Microsoft.Office.Interop.Excel;

namespace WpfEGridApp
{
    public partial class ComponentMappingWindow : Window
    {
        private MainWindow _mainWindow;
        private ComponentMappingManager _mappingManager;
        public ObservableCollection<MappingDisplayItem> MappingDisplayItems { get; set; }
        public ObservableCollection<UnmappedReferenceItem> UnmappedReferences { get; set; }

        private Queue<string> _specificMappingQueue;
        private bool _isInSpecificMappingMode = false;
        private string _originalSearchText = "";

        public ComponentMappingWindow(MainWindow mainWindow, ComponentMappingManager mappingManager)
        {
            InitializeComponent();
            _mainWindow = mainWindow;
            _mappingManager = mappingManager;

            MappingDisplayItems = new ObservableCollection<MappingDisplayItem>();
            UnmappedReferences = new ObservableCollection<UnmappedReferenceItem>();

            MappingsList.ItemsSource = MappingDisplayItems;
            UnmappedReferencesList.ItemsSource = UnmappedReferences;

            LoadExistingMappings();
            UpdateStatistics();
            UpdateProgress();
            UpdateStatus("Component mapping vindu åpnet");
        }

        private void LoadExistingMappings()
        {
            MappingDisplayItems.Clear();
            var mappings = _mappingManager.GetAllMappings();

            var filteredMappings = mappings;

            // Apply search filter if search text exists
            if (!string.IsNullOrWhiteSpace(SearchMappingsBox?.Text))
            {
                var searchText = SearchMappingsBox.Text.ToLower();
                filteredMappings = mappings.Where(m =>
                    m.ExcelReference.ToLower().Contains(searchText)).ToList();
            }

            foreach (var mapping in filteredMappings)
            {
                string position;
                if (mapping.GridRow == -1)
                {
                    position = $"Motor {mapping.GridColumn - 1000}";
                }
                else if (mapping.GridRow == -2)
                {
                    position = $"Door {mapping.GridColumn - 1000}";
                }
                else
                {
                    position = $"({mapping.GridRow},{mapping.GridColumn})";
                    if (mapping.DefaultToBottom)
                        position += " (B)";
                    else
                        position += " (T)";
                }

                MappingDisplayItems.Add(new MappingDisplayItem
                {
                    ExcelReference = mapping.ExcelReference,
                    GridPosition = position
                });
            }

            UpdateStatistics();
        }

        private void UpdateStatistics()
        {
            var allMappings = _mappingManager.GetAllMappings();
            var componentMappings = allMappings.Where(m => m.GridRow >= 0).Count();
            var specialMappings = allMappings.Where(m => m.GridRow < 0).Count();

            TotalMappingsCount.Text = allMappings.Count.ToString();
            ComponentMappingsCount.Text = componentMappings.ToString();
            SpecialMappingsCount.Text = specialMappings.ToString();
        }

        private void UpdateProgress()
        {
            try
            {
                if (_mainWindow.worksheet == null)
                {
                    MappingProgressBar.Value = 0;
                    ProgressPercentageText.Text = "0%";
                    return;
                }

                var uniqueCells = GetUniqueExcelCells();
                var mappedCount = uniqueCells.Count(cell => _mappingManager.HasMapping(cell));
                var totalCount = uniqueCells.Count;

                if (totalCount > 0)
                {
                    var percentage = (double)mappedCount / totalCount * 100;
                    MappingProgressBar.Value = percentage;
                    ProgressPercentageText.Text = $"{percentage:F0}%";
                }
                else
                {
                    MappingProgressBar.Value = 0;
                    ProgressPercentageText.Text = "0%";
                }
            }
            catch (Exception)
            {
                MappingProgressBar.Value = 0;
                ProgressPercentageText.Text = "0%";
            }
        }

        private void SearchMappings_TextChanged(object sender, TextChangedEventArgs e)
        {
            LoadExistingMappings();
        }

        private void ClearAllMappings_Click(object sender, RoutedEventArgs e)
        {
            var result = MessageBox.Show(
                "Er du sikker på at du vil slette ALLE mappings?\n\nDenne handlingen kan ikke angres.",
                "Bekreft sletting av alle mappings",
                MessageBoxButton.YesNo,
                MessageBoxImage.Warning);

            if (result == MessageBoxResult.Yes)
            {
                var mappings = _mappingManager.GetAllMappings().ToList();
                foreach (var mapping in mappings)
                {
                    _mappingManager.RemoveMapping(mapping.ExcelReference);
                }
                LoadExistingMappings();
                UpdateProgress();
                UpdateStatus("Alle mappings slettet");
            }
        }

        private void StartSpecificMapping_Click(object sender, RoutedEventArgs e)
        {
            if (_mainWindow.worksheet == null)
            {
                UpdateStatus("Ingen Excel-fil åpen");
                return;
            }

            try
            {
                var uniqueCells = GetUniqueExcelCells();
                var unmappedCells = FilterUnmappedCells(uniqueCells);

                if (unmappedCells.Count == 0)
                {
                    MessageBox.Show("Alle celler er allerede mappet!", "Ingen celler å mappe",
                                   MessageBoxButton.OK, MessageBoxImage.Information);
                    return;
                }

                _specificMappingQueue = new Queue<string>(unmappedCells);
                _isInSpecificMappingMode = true;

                ProcessNextSpecificMapping();

                UpdateStatus($"Starter spesifikk mapping av {unmappedCells.Count} celler.");
            }
            catch (Exception ex)
            {
                UpdateStatus($"Feil ved oppstart av spesifikk mapping: {ex.Message}");
            }
        }

        private List<string> GetUniqueExcelCells()
        {
            var uniqueCells = new HashSet<string>();
            var usedRange = _mainWindow.worksheet.UsedRange;
            if (usedRange == null) return new List<string>();

            var lastRow = usedRange.Rows.Count;

            for (int row = 2; row <= lastRow; row++)
            {
                try
                {
                    var cellB = (_mainWindow.worksheet.Cells[row, 2] as Excel.Range)?.Value?.ToString() ?? "";
                    var cellC = (_mainWindow.worksheet.Cells[row, 3] as Excel.Range)?.Value?.ToString() ?? "";

                    if (!string.IsNullOrWhiteSpace(cellB)) uniqueCells.Add(cellB.Trim());
                    if (!string.IsNullOrWhiteSpace(cellC)) uniqueCells.Add(cellC.Trim());
                }
                catch { }
            }

            return uniqueCells.ToList();
        }

        private List<string> FilterUnmappedCells(List<string> cells)
        {
            var unmappedCells = new List<string>();
            var existingMappings = _mappingManager.GetAllMappings();

            foreach (var cell in cells)
            {
                bool isMapped = false;
                foreach (var mapping in existingMappings)
                {
                    if (cell.Contains(mapping.ExcelReference, StringComparison.OrdinalIgnoreCase))
                    {
                        isMapped = true;
                        break;
                    }
                }

                if (!isMapped)
                {
                    unmappedCells.Add(cell);
                }
            }

            return unmappedCells;
        }

        private void ProcessNextSpecificMapping()
        {
            if (!_isInSpecificMappingMode || _specificMappingQueue == null || _specificMappingQueue.Count == 0)
            {
                EndSpecificMappingMode();
                return;
            }

            var currentCell = _specificMappingQueue.Dequeue();

            // Show in MainWindow
            _mainWindow.StartInteractiveMapping(currentCell, "", OnSpecificMappingCompleted);

            UpdateStatus($"Mapper '{currentCell}' - {_specificMappingQueue.Count} gjenstår");
            this.WindowState = WindowState.Minimized;
        }

        private void OnSpecificMappingCompleted(string reference, string ignore)
        {
            this.WindowState = WindowState.Normal;
            this.Activate();
            LoadExistingMappings();
            UpdateProgress();

            // Continue with next mapping immediately
            System.Threading.Tasks.Task.Delay(100).ContinueWith(_ =>
            {
                this.Dispatcher.Invoke(() => ProcessNextSpecificMapping());
            });
        }

        private void EndSpecificMappingMode()
        {
            _isInSpecificMappingMode = false;
            _specificMappingQueue = null;
            _mainWindow.EndMappingMode();
            UpdateStatus("Spesifikk mapping fullført");
            UpdateProgress();

            this.WindowState = WindowState.Normal;
            this.Activate();
        }

        protected override void OnKeyDown(System.Windows.Input.KeyEventArgs e)
        {
            if (e.Key == System.Windows.Input.Key.Escape && _isInSpecificMappingMode)
            {
                EndSpecificMappingMode();
                e.Handled = true;
            }
            base.OnKeyDown(e);
        }

        private void StartInteractiveMapping_Click(object sender, RoutedEventArgs e)
        {
            if (string.IsNullOrWhiteSpace(NewExcelReference.Text))
            {
                UpdateStatus("Skriv inn Excel-referanse først");
                return;
            }

            var reference = NewExcelReference.Text.Trim();

            _mainWindow.StartInteractiveMapping(reference, "", OnMappingCompleted);
            UpdateStatus($"Interaktiv mapping startet for {reference}. Klikk på mellomrad i hovedvinduet.");
            this.WindowState = WindowState.Minimized;
        }

        private void OnMappingCompleted(string reference, string ignore)
        {
            LoadExistingMappings();
            UpdateProgress();
            NewExcelReference.Clear();
            this.WindowState = WindowState.Normal;
            this.Activate();
            UpdateStatus($"Mapping fullført for {reference}");
        }

        private void FindUnmappedReferences_Click(object sender, RoutedEventArgs e)
        {
            UnmappedReferences.Clear();

            try
            {
                if (_mainWindow.worksheet == null)
                {
                    UpdateStatus("Ingen Excel-fil åpen");
                    return;
                }

                var foundReferences = new HashSet<string>();
                var usedRange = _mainWindow.worksheet.UsedRange;
                var lastRow = usedRange?.Rows?.Count ?? 100;

                for (int row = 2; row <= lastRow; row++)
                {
                    var cellB = (_mainWindow.worksheet.Cells[row, 2] as Excel.Range)?.Value?.ToString() ?? "";
                    var cellC = (_mainWindow.worksheet.Cells[row, 3] as Excel.Range)?.Value?.ToString() ?? "";

                    ExtractBaseReferences(cellB, foundReferences);
                    ExtractBaseReferences(cellC, foundReferences);
                }

                foreach (var reference in foundReferences.OrderBy(r => r))
                {
                    if (!_mappingManager.HasMapping(reference))
                    {
                        UnmappedReferences.Add(new UnmappedReferenceItem { Reference = reference, IsSelected = false });
                    }
                }

                UpdateStatus($"Fant {UnmappedReferences.Count} umappede base-referanser av totalt {foundReferences.Count}");
            }
            catch (Exception ex)
            {
                UpdateStatus($"Feil ved søking: {ex.Message}");
                MessageBox.Show($"Feil ved søking etter referanser: {ex.Message}", "Feil",
                               MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        private void ExtractBaseReferences(string cellValue, HashSet<string> references)
        {
            if (string.IsNullOrWhiteSpace(cellValue)) return;

            // Extract prefix (everything before the last component reference)
            // For "E01-A1-X3:1" we want "E01-A1"
            // For "A01-K3" we want "A01-K3"
            // For "J01-K4" we want "J01-K4"

            var allMatches = new List<string>();

            // Look for patterns like "XXX-YY" where XXX is prefix and YY is component
            var prefixMatches = Regex.Matches(cellValue, @"([A-Z]\d+-[A-Z]\d+)");
            foreach (Match match in prefixMatches)
            {
                allMatches.Add(match.Groups[1].Value);
            }

            // If no prefix pattern found, try terminal block patterns
            if (allMatches.Count == 0)
            {
                // Terminal block matches (X20:41 should match X20:)
                var terminalMatches = Regex.Matches(cellValue, @"[A-Z]\d+:\d+");
                foreach (Match match in terminalMatches)
                {
                    var fullRef = match.Value;
                    var baseRef = fullRef.Substring(0, fullRef.IndexOf(':') + 1);
                    allMatches.Add(baseRef);
                }

                // Simple matches (X20, F5, etc.)
                var simpleMatches = Regex.Matches(cellValue, @"[A-Z]\d+(?!:)");
                foreach (Match match in simpleMatches)
                {
                    var startPos = match.Index;
                    var endPos = startPos + match.Length;

                    if (endPos < cellValue.Length && cellValue[endPos] == ':')
                        continue;

                    allMatches.Add(match.Value);
                }
            }

            foreach (var reference in allMatches.Distinct())
            {
                references.Add(reference);
            }
        }

        private void StartSelectedMapping_Click(object sender, RoutedEventArgs e)
        {
            var selectedRefs = UnmappedReferences.Where(u => u.IsSelected).Select(u => u.Reference).ToList();
            if (selectedRefs.Count == 0)
            {
                UpdateStatus("Ingen referanser valgt");
                return;
            }

            _specificMappingQueue = new Queue<string>(selectedRefs);
            _isInSpecificMappingMode = true;

            ProcessNextSpecificMapping();

            UpdateStatus($"Starter mapping av {selectedRefs.Count} valgte referanser");
        }

        private void DeleteMapping_Click(object sender, RoutedEventArgs e)
        {
            if (sender is Button button && button.Tag is string excelReference)
            {
                var result = MessageBox.Show(
                    $"Slett mapping for {excelReference}?",
                    "Bekreft sletting",
                    MessageBoxButton.YesNo,
                    MessageBoxImage.Question);

                if (result == MessageBoxResult.Yes)
                {
                    _mappingManager.RemoveMapping(excelReference);
                    LoadExistingMappings();
                    UpdateProgress();
                    UpdateStatus($"Slettet mapping for {excelReference}");
                }
            }
        }

        private void ProcessAllWithMappings_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                ProcessingStatusText.Text = "Prosesserer...";
                ProcessingStatusText.Foreground = System.Windows.Media.Brushes.Orange;

                var processor = new ExcelConnectionProcessor(_mainWindow, _mappingManager);
                var processedCount = processor.ProcessAllConnections();

                ProcessingStatusText.Text = "Prosessering fullført";
                ProcessingStatusText.Foreground = System.Windows.Media.Brushes.LightGreen;
                LastProcessingResultText.Text = $"Prosesserte {processedCount} ledninger";

                UpdateStatus($"Prosesserte {processedCount} ledninger");
                MessageBox.Show($"Ferdig! Prosesserte {processedCount} ledninger automatisk.",
                               "Automatisk prosessering fullført",
                               MessageBoxButton.OK, MessageBoxImage.Information);

                _mainWindow.UpdateExcelDisplayText();
            }
            catch (Exception ex)
            {
                ProcessingStatusText.Text = "Prosessering feilet";
                ProcessingStatusText.Foreground = System.Windows.Media.Brushes.Red;
                LastProcessingResultText.Text = $"Feil: {ex.Message}";

                UpdateStatus($"Feil under prosessering: {ex.Message}");
                MessageBox.Show($"Feil under automatisk prosessering: {ex.Message}", "Feil",
                               MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        private void TestProcessing_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                ProcessingStatusText.Text = "Kjører test...";
                ProcessingStatusText.Foreground = System.Windows.Media.Brushes.Orange;

                var processor = new ExcelConnectionProcessor(_mainWindow, _mappingManager);
                var result = processor.TestProcessing();

                ProcessingStatusText.Text = "Test fullført";
                ProcessingStatusText.Foreground = System.Windows.Media.Brushes.LightGreen;

                var message = $"Test resultat:\n\n{result}";
                MessageBox.Show(message, "Test prosessering", MessageBoxButton.OK, MessageBoxImage.Information);
            }
            catch (Exception ex)
            {
                ProcessingStatusText.Text = "Test feilet";
                ProcessingStatusText.Foreground = System.Windows.Media.Brushes.Red;
                LastProcessingResultText.Text = $"Feil: {ex.Message}";

                UpdateStatus($"Feil under test: {ex.Message}");
                MessageBox.Show($"Feil under test prosessering: {ex.Message}", "Feil",
                               MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        private void ExportMappings_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                var saveFileDialog = new SaveFileDialog
                {
                    Filter = "JSON Files (*.json)|*.json|All Files (*.*)|*.*",
                    Title = "Eksporter Mappings"
                };

                if (saveFileDialog.ShowDialog() == true)
                {
                    var mappings = _mappingManager.GetAllMappings();
                    var json = JsonSerializer.Serialize(mappings, new JsonSerializerOptions
                    {
                        WriteIndented = true
                    });
                    File.WriteAllText(saveFileDialog.FileName, json);
                    UpdateStatus($"Eksporterte {mappings.Count} mappings til {Path.GetFileName(saveFileDialog.FileName)}");
                }
            }
            catch (Exception ex)
            {
                UpdateStatus($"Feil ved eksport: {ex.Message}");
                MessageBox.Show($"Feil ved eksport av mappings: {ex.Message}", "Feil",
                               MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        private void ImportMappings_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                var openFileDialog = new OpenFileDialog
                {
                    Filter = "JSON Files (*.json)|*.json|All Files (*.*)|*.*",
                    Title = "Importer Mappings"
                };

                if (openFileDialog.ShowDialog() == true)
                {
                    var json = File.ReadAllText(openFileDialog.FileName);
                    var mappings = JsonSerializer.Deserialize<List<ComponentMapping>>(json);

                    if (mappings != null)
                    {
                        var result = MessageBox.Show(
                            $"Dette vil erstatte alle eksisterende mappings med {mappings.Count} mappings fra filen.\n\nFortsette?",
                            "Bekreft import",
                            MessageBoxButton.YesNo,
                            MessageBoxImage.Question);

                        if (result == MessageBoxResult.Yes)
                        {
                            // Clear existing mappings
                            var existingMappings = _mappingManager.GetAllMappings().ToList();
                            foreach (var mapping in existingMappings)
                            {
                                _mappingManager.RemoveMapping(mapping.ExcelReference);
                            }

                            // Add imported mappings
                            foreach (var mapping in mappings)
                            {
                                _mappingManager.AddMapping(mapping.ExcelReference, mapping.GridRow,
                                                         mapping.GridColumn, mapping.DefaultToBottom);
                            }

                            LoadExistingMappings();
                            UpdateProgress();
                            UpdateStatus($"Importerte {mappings.Count} mappings fra {Path.GetFileName(openFileDialog.FileName)}");
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                UpdateStatus($"Feil ved import: {ex.Message}");
                MessageBox.Show($"Feil ved import av mappings: {ex.Message}", "Feil",
                               MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        private void Close_Click(object sender, RoutedEventArgs e)
        {
            this.Close();
        }

        private void UpdateStatus(string message)
        {
            StatusText.Text = message;
            TimestampText.Text = DateTime.Now.ToString("HH:mm:ss");
        }
    }

    public class MappingDisplayItem
    {
        public string ExcelReference { get; set; }
        public string GridPosition { get; set; }
    }

    public class UnmappedReferenceItem
    {
        public string Reference { get; set; }
        public bool IsSelected { get; set; }
    }
}