using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.IO;
using System.Linq;
using System.Text.Json;
using System.Text.RegularExpressions;
using System.Windows;
using System.Windows.Controls;
using Excel = Microsoft.Office.Interop.Excel;

namespace WpfEGridApp
{
    public partial class ComponentMappingWindow : Window
    {
        private MainWindow _mainWindow;
        private ComponentMappingManager _mappingManager;
        public ObservableCollection<MappingDisplayItem> MappingDisplayItems { get; set; }
        public ObservableCollection<UnmappedReferenceItem> UnmappedReferences { get; set; }

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
            UpdateStatus("Component mapping vindu åpnet");
        }

        private void LoadExistingMappings()
        {
            MappingDisplayItems.Clear();
            var mappings = _mappingManager.GetAllMappings();

            foreach (var mapping in mappings)
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

            UpdateStatus($"Lastet {mappings.Count} eksisterende mappings");
        }

        private void ClearAllMappings_Click(object sender, RoutedEventArgs e)
        {
            var result = MessageBox.Show(
                "Er du sikker på at du vil slette ALLE mappings?",
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
                UpdateStatus("Alle mappings slettet");
            }
        }

        private Queue<string> _specificMappingQueue;
        private bool _isInSpecificMappingMode = false;

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

                UpdateStatus($"Starter spesifikk mapping av {unmappedCells.Count} celler. Trykk ESC for å avbryte.");
            }
            catch (Exception ex)
            {
                UpdateStatus($"Feil ved oppstart av spesifikk mapping: {ex.Message}");
            }
        }

        private void CancelSpecificMapping_Click(object sender, RoutedEventArgs e)
        {
            EndSpecificMappingMode();
            UpdateStatus("Spesifikk mapping avbrutt");
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
                MessageBox.Show("Spesifikk mapping fullført!", "Ferdig", MessageBoxButton.OK, MessageBoxImage.Information);
                return;
            }

            var currentCell = _specificMappingQueue.Dequeue();

            // Show in MainWindow instead of MessageBox
            _mainWindow.StartInteractiveMapping(currentCell, "", OnSpecificMappingCompleted);

            UpdateStatus($"Mapper '{currentCell}' - {_specificMappingQueue.Count} gjenstår");
            this.WindowState = WindowState.Minimized;
        }

        private void OnSpecificMappingCompleted(string reference, string ignore)
        {
            this.WindowState = WindowState.Normal;
            this.Activate();
            LoadExistingMappings();

            // Short pause and continue with next mapping
            System.Threading.Tasks.Task.Delay(500).ContinueWith(_ =>
            {
                this.Dispatcher.Invoke(() => ProcessNextSpecificMapping());
            });
        }

        private void EndSpecificMappingMode()
        {
            _isInSpecificMappingMode = false;
            _specificMappingQueue = null;
            _mainWindow.EndMappingMode();
            UpdateStatus("Spesifikk mapping fullført/avbrutt");

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

            var allMatches = new List<string>();

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
                    UpdateStatus($"Slettet mapping for {excelReference}");
                }
            }
        }

        private void ProcessAllWithMappings_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                var processor = new ExcelConnectionProcessor(_mainWindow, _mappingManager);
                var processedCount = processor.ProcessAllConnections();

                UpdateStatus($"Prosesserte {processedCount} ledninger");
                MessageBox.Show($"Ferdig! Prosesserte {processedCount} ledninger automatisk.",
                               "Automatisk prosessering fullført",
                               MessageBoxButton.OK, MessageBoxImage.Information);

                _mainWindow.UpdateExcelDisplayText();
            }
            catch (Exception ex)
            {
                UpdateStatus($"Feil under prosessering: {ex.Message}");
                MessageBox.Show($"Feil under automatisk prosessering: {ex.Message}", "Feil",
                               MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        private void TestProcessing_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                var processor = new ExcelConnectionProcessor(_mainWindow, _mappingManager);
                var result = processor.TestProcessing();

                var message = $"Test resultat:\n\n{result}";
                MessageBox.Show(message, "Test prosessering", MessageBoxButton.OK, MessageBoxImage.Information);
            }
            catch (Exception ex)
            {
                UpdateStatus($"Feil under test: {ex.Message}");
                MessageBox.Show($"Feil under test prosessering: {ex.Message}", "Feil",
                               MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        private void Close_Click(object sender, RoutedEventArgs e)
        {
            this.Close();
        }

        private void UpdateStatus(string message)
        {
            StatusText.Text = $"{DateTime.Now:HH:mm:ss} - {message}";
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
