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

        private Queue<string> _specificMappingQueue;
        private bool _isInSpecificMappingMode = false;
        private bool _waitingForUserMapping = false;

        private Queue<string> _rangeMappingQueue;
        private bool _isInRangeMappingMode = false;

        private bool _isBulkMappingMode = false;
        private string _bulkPrefix;
        private int _bulkStart;
        private int _bulkEnd;

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

        public void LoadExistingMappings()
        {
            MappingDisplayItems.Clear();
            var mappings = _mappingManager.GetAllMappingsIncludingBulk();

            foreach (var mapping in mappings)
            {
                string position;
                if (mapping.GridRow == -99)
                {
                    position = $"BULK: {mapping.GridColumn} celler";
                    if (mapping.DefaultToBottom) position += " (B)";
                    else position += " (T)";
                }
                else if (mapping.GridRow == -1)
                    position = $"Motor {mapping.GridColumn - 1000}";
                else if (mapping.GridRow == -2)
                    position = $"Door {mapping.GridColumn - 1000}";
                else
                {
                    position = $"({mapping.GridRow},{mapping.GridColumn})";
                    if (mapping.DefaultToBottom) position += " (B)";
                    else position += " (T)";
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
                "Er du sikker på at du vil slette ALLE mappings (inkludert bulk mappings)?",
                "Bekreft sletting av alle mappings",
                MessageBoxButton.YesNo,
                MessageBoxImage.Warning);

            if (result == MessageBoxResult.Yes)
            {
                var mappings = _mappingManager.GetAllMappings().ToList();
                foreach (var mapping in mappings)
                    _mappingManager.RemoveMapping(mapping.ExcelReference);

                var bulkMappings = _mappingManager.GetAllBulkRanges().ToList();
                foreach (var bulk in bulkMappings)
                    _mappingManager.RemoveBulkRangeMapping(bulk.Prefix, bulk.StartIndex, bulk.EndIndex);

                LoadExistingMappings();
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
                var uniqueCells = GetUniqueExcelCellsOrdered();
                var unmappedCells = FilterUnmappedCellsExcludingBulk(uniqueCells);

                if (unmappedCells.Count == 0)
                {
                    MessageBox.Show("Alle celler er allerede mappet!", "Ingen celler å mappe",
                                   MessageBoxButton.OK, MessageBoxImage.Information);
                    return;
                }

                _specificMappingQueue = new Queue<string>(unmappedCells);
                _isInSpecificMappingMode = true;
                _waitingForUserMapping = false;

                _mainWindow.StartSequentialMappingMode();
                ProcessNextSpecificMapping();
                this.WindowState = WindowState.Minimized;
                UpdateStatus($"Starter spesifikk mapping av {unmappedCells.Count} celler (ekskluderer bulk-mappede referanser).");
            }
            catch (Exception ex)
            {
                UpdateStatus($"Feil ved oppstart av spesifikk mapping: {ex.Message}");
            }
        }

        private List<string> GetUniqueExcelCellsOrdered()
        {
            var allCells = new HashSet<string>();
            var usedRange = _mainWindow.worksheet.UsedRange;
            if (usedRange == null) return new List<string>();

            var lastRow = usedRange.Rows.Count;

            for (int row = 2; row <= lastRow; row++)
            {
                try
                {
                    var cellB = (_mainWindow.worksheet.Cells[row, 2] as Excel.Range)?.Value?.ToString() ?? "";
                    var cellC = (_mainWindow.worksheet.Cells[row, 3] as Excel.Range)?.Value?.ToString() ?? "";

                    if (!string.IsNullOrWhiteSpace(cellB)) allCells.Add(cellB.Trim());
                    if (!string.IsNullOrWhiteSpace(cellC)) allCells.Add(cellC.Trim());
                }
                catch { }
            }

            return allCells.OrderBy(x => x).ToList();
        }

        private List<string> FilterUnmappedCellsExcludingBulk(List<string> cells)
        {
            var unmappedCells = new List<string>();
            foreach (var cell in cells)
            {
                if (_mappingManager.IsReferenceMapped(cell))
                    continue;

                if (_mappingManager.IsReferenceInBulkRange(cell))
                    continue;

                unmappedCells.Add(cell);
            }
            return unmappedCells;
        }

        private void ProcessNextSpecificMapping()
        {
            if (!_isInSpecificMappingMode || _specificMappingQueue == null)
            {
                EndSpecificMappingMode();
                return;
            }

            if (_waitingForUserMapping)
                return;

            while (_specificMappingQueue.Count > 0)
            {
                var currentCell = _specificMappingQueue.Peek();

                if (_mappingManager.IsReferenceMapped(currentCell) ||
                    _mappingManager.IsReferenceInBulkRange(currentCell))
                {
                    _specificMappingQueue.Dequeue();
                    continue;
                }

                _specificMappingQueue.Dequeue();
                _mainWindow.StartInteractiveMapping(currentCell, "", OnSpecificMappingCompleted);
                UpdateStatus($"Mapper '{currentCell}' - {_specificMappingQueue.Count} gjenstår");
                return;
            }
            EndSpecificMappingMode();
        }

        private void OnSpecificMappingCompleted(string reference, string ignore)
        {
            _waitingForUserMapping = false;
            LoadExistingMappings();

            System.Threading.Tasks.Task.Delay(100).ContinueWith(_ =>
            {
                this.Dispatcher.Invoke(() => ProcessNextSpecificMapping());
            });
        }

        public void OnSequentialMappingFinished()
        {
            EndSpecificMappingMode();
        }

        public void PutReferenceBackInQueue(string reference)
        {
            if (_specificMappingQueue != null && _isInSpecificMappingMode)
            {
                var newQueue = new Queue<string>();
                newQueue.Enqueue(reference);

                while (_specificMappingQueue.Count > 0)
                    newQueue.Enqueue(_specificMappingQueue.Dequeue());

                _specificMappingQueue = newQueue;
                _waitingForUserMapping = false;

                UpdateStatus($"Angret mapping for '{reference}' - klar for ny mapping. ({_specificMappingQueue.Count - 1} gjenstår etter denne)");
            }
        }

        private void EndSpecificMappingMode()
        {
            _isInSpecificMappingMode = false;
            _specificMappingQueue = null;
            _waitingForUserMapping = false;
            _mainWindow.EndMappingMode();
            _mainWindow.EndSequentialMappingMode();

            this.WindowState = WindowState.Normal;
            this.Activate();
            LoadExistingMappings();
            UpdateStatus("Spesifikk mapping fullført");
        }

        private void StartRangeMapping_Click(object sender, RoutedEventArgs e)
        {
            var text = NewExcelReference.Text?.Trim();
            if (string.IsNullOrWhiteSpace(text))
            {
                UpdateStatus("Skriv inn prefiks eller intervall først");
                return;
            }

            var rangeRegex = new Regex(@"^([A-Za-z]+\d+):(\d+)(?:-(\d+))?$");
            var match = rangeRegex.Match(text);
            if (!match.Success)
            {
                MessageBox.Show("Ugyldig format. Bruk f.eks. X2:21-100 eller X2:21", "Feil format",
                                MessageBoxButton.OK, MessageBoxImage.Warning);
                return;
            }

            var prefix = match.Groups[1].Value;
            var startStr = match.Groups[2].Value;
            var endStr = match.Groups[3].Success ? match.Groups[3].Value : null;

            if (!int.TryParse(startStr, out var startNum))
            {
                MessageBox.Show("Startnummeret kunne ikke tolkes", "Feil", MessageBoxButton.OK, MessageBoxImage.Warning);
                return;
            }

            int endNum = string.IsNullOrEmpty(endStr) ? startNum :
                        (int.TryParse(endStr, out endNum) ? endNum : startNum);

            if (endNum < startNum)
            {
                MessageBox.Show("Sluttnummeret må være større eller lik startnummeret", "Feil", MessageBoxButton.OK, MessageBoxImage.Warning);
                return;
            }

            _isBulkMappingMode = true;
            _bulkPrefix = prefix;
            _bulkStart = startNum;
            _bulkEnd = endNum;
            FinishBulkMappingButton.Visibility = Visibility.Visible;

            UpdateStatus($"Bulk mapping: velg grid-celler for {prefix}:{startNum}-{endNum}. Klikk 'Ferdig bulk mapping' når ferdig.");

            MessageBox.Show(
                $"Bulk mapping startet for {prefix}:{startNum}-{endNum}.\n\n" +
                "1. Gå tilbake til hovedvinduet\n" +
                "2. Klikk på T-celler ELLER B-celler (1 eller flere)\n" +
                "3. Klikk 'Ferdig bulk mapping' i hovedvinduet\n\n" +
                $"Dette vil mappe alle {endNum - startNum + 1} referanser til de valgte cellene.",
                "Bulk Mapping", MessageBoxButton.OK, MessageBoxImage.Information);

            _mainWindow.StartBulkMappingSelection(prefix, startNum, endNum, OnBulkMappingCompleted);
            this.WindowState = WindowState.Minimized;
        }

        private void FinishBulkMappingButton_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                _mainWindow?.FinishBulkMappingSelection();
            }
            catch (Exception ex)
            {
                UpdateStatus($"Feil ved fullføring av bulk mapping: {ex.Message}");
            }
        }

        private void OnBulkMappingCompleted(string prefix, int startNumber, int endNumber, List<(int Row, int Col)> cells)
        {
            this.Dispatcher.Invoke(() =>
            {
                try
                {
                    if (cells == null || cells.Count == 0)
                    {
                        MessageBox.Show("Ingen celler ble valgt. Bulk mapping avbrutt.", "Ingen valg",
                                       MessageBoxButton.OK, MessageBoxImage.Warning);
                        FinishBulkMappingButton.Visibility = Visibility.Collapsed;
                        _isBulkMappingMode = false;
                        NewExcelReference.Clear();
                        this.WindowState = WindowState.Normal;
                        this.Activate();
                        UpdateStatus("Bulk mapping avbrutt - ingen celler valgt");
                        return;
                    }

                    bool selectedIsTop = _mainWindow.BulkMappingSelectedIsTop;

                    // VIKTIG: Kall AddBulkRangeMapping med de valgte cellene
                    _mappingManager.AddBulkRangeMapping(prefix, startNumber, endNumber, cells, selectedIsTop);

                    FinishBulkMappingButton.Visibility = Visibility.Collapsed;
                    _isBulkMappingMode = false;
                    NewExcelReference.Clear();

                    // Reload mappings for å vise de nye
                    LoadExistingMappings();

                    int totalReferences = (endNumber - startNumber + 1) * 2;
                    string message = $"Bulk mapping fullført!\n\n" +
                        $"Range: {prefix}:{startNumber}-{endNumber}\n" +
                        $"Celler valgt: {cells.Count}\n" +
                        $"Side: {(selectedIsTop ? "T (oversiden)" : "B (undersiden)")}\n\n" +
                        $"Totalt {totalReferences} referanser mappet:\n" +
                        $"- {endNumber - startNumber + 1} base referanser (eks: {prefix}:{startNumber})\n" +
                        $"- {endNumber - startNumber + 1} stjerne referanser (eks: {prefix}:{startNumber}*)";

                    MessageBox.Show(message, "Bulk Mapping Fullført", MessageBoxButton.OK, MessageBoxImage.Information);

                    this.WindowState = WindowState.Normal;
                    this.Activate();
                    UpdateStatus($"Bulk mapping lagret: {prefix}:{startNumber}-{endNumber} ({cells.Count} celler, {totalReferences} ref)");
                }
                catch (Exception ex)
                {
                    UpdateStatus($"Feil under lagring av bulk mapping: {ex.Message}");
                    MessageBox.Show($"Feil under lagring av bulk mapping: {ex.Message}", "Feil", MessageBoxButton.OK, MessageBoxImage.Error);
                }
            });
        }

        protected override void OnKeyDown(System.Windows.Input.KeyEventArgs e)
        {
            if (e.Key == System.Windows.Input.Key.Escape)
            {
                if (_isInSpecificMappingMode)
                {
                    EndSpecificMappingMode();
                    e.Handled = true;
                }
                else if (_isInRangeMappingMode)
                {
                    EndRangeMapping();
                    e.Handled = true;
                }
            }
            base.OnKeyDown(e);
        }

        private void ProcessNextRangeMapping()
        {
            if (!_isInRangeMappingMode || _rangeMappingQueue == null || _rangeMappingQueue.Count == 0)
            {
                EndRangeMapping();
                return;
            }
            var currentRef = _rangeMappingQueue.Dequeue();
            _mainWindow.StartInteractiveMapping(currentRef, "", OnRangeMappingCompleted);
            UpdateStatus($"Mapper '{currentRef}' - {_rangeMappingQueue.Count} gjenstår");
        }

        private void OnRangeMappingCompleted(string reference, string ignore)
        {
            LoadExistingMappings();
            System.Threading.Tasks.Task.Delay(100).ContinueWith(_ =>
            {
                this.Dispatcher.Invoke(() => ProcessNextRangeMapping());
            });
        }

        private void EndRangeMapping()
        {
            _isInRangeMappingMode = false;
            _rangeMappingQueue = null;
            _mainWindow.EndMappingMode();
            this.WindowState = WindowState.Normal;
            this.Activate();
            NewExcelReference.Clear();
            UpdateStatus("Rekke-mapping fullført");
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
                    if (_mappingManager.IsReferenceInBulkRange(reference))
                        continue;

                    if (!_mappingManager.HasMapping(reference))
                        UnmappedReferences.Add(new UnmappedReferenceItem { Reference = reference, IsSelected = false });
                }

                UpdateStatus($"Fant {UnmappedReferences.Count} umappede base-referanser (ekskluderer bulk ranges)");
            }
            catch (Exception ex)
            {
                UpdateStatus($"Feil ved søking: {ex.Message}");
                MessageBox.Show($"Feil ved søking etter referanser: {ex.Message}", "Feil", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        private void ExtractBaseReferences(string cellValue, HashSet<string> references)
        {
            if (string.IsNullOrWhiteSpace(cellValue)) return;

            var allMatches = new List<string>();
            var prefixMatches = Regex.Matches(cellValue, @"([A-Z]\d+-[A-Z]\d+)");
            foreach (Match match in prefixMatches)
                allMatches.Add(match.Groups[1].Value);

            if (allMatches.Count == 0)
            {
                var terminalMatches = Regex.Matches(cellValue, @"[A-Z]\d+:\d+");
                foreach (Match match in terminalMatches)
                {
                    var fullRef = match.Value;
                    var baseRef = fullRef.Substring(0, fullRef.IndexOf(':') + 1);
                    allMatches.Add(baseRef);
                }

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
                references.Add(reference);
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
            _waitingForUserMapping = false;

            _mainWindow.StartSequentialMappingMode();
            ProcessNextSpecificMapping();
            this.WindowState = WindowState.Minimized;
            UpdateStatus($"Starter mapping av {selectedRefs.Count} valgte referanser");
        }

        private void DeleteMapping_Click(object sender, RoutedEventArgs e)
        {
            if (sender is Button button && button.Tag is string excelReference)
            {
                var result = MessageBox.Show($"Slett mapping for {excelReference}?", "Bekreft sletting",
                    MessageBoxButton.YesNo, MessageBoxImage.Question);

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
                               "Automatisk prosessering fullført", MessageBoxButton.OK, MessageBoxImage.Information);

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
                MessageBox.Show($"Test resultat:\n\n{result}", "Test prosessering", MessageBoxButton.OK, MessageBoxImage.Information);
            }
            catch (Exception ex)
            {
                UpdateStatus($"Feil under test: {ex.Message}");
                MessageBox.Show($"Feil under test prosessering: {ex.Message}", "Feil", MessageBoxButton.OK, MessageBoxImage.Error);
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