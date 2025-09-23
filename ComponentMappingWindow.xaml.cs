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

        public void LoadExistingMappings()
        {
            MappingDisplayItems.Clear();

            // Get all mappings including bulk mappings
            var mappings = _mappingManager.GetAllMappingsIncludingBulk();

            foreach (var mapping in mappings)
            {
                string position;

                // Check if it's a bulk mapping (special marker -99)
                if (mapping.GridRow == -99)
                {
                    position = $"BULK: {mapping.GridColumn} celler";
                    if (mapping.DefaultToBottom)
                        position += " (B)";
                    else
                        position += " (T)";
                }
                else if (mapping.GridRow == -1)
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
                "Er du sikker på at du vil slette ALLE mappings (inkludert bulk mappings)?",
                "Bekreft sletting av alle mappings",
                MessageBoxButton.YesNo,
                MessageBoxImage.Warning);

            if (result == MessageBoxResult.Yes)
            {
                // Clear regular mappings
                var mappings = _mappingManager.GetAllMappings().ToList();
                foreach (var mapping in mappings)
                {
                    _mappingManager.RemoveMapping(mapping.ExcelReference);
                }

                // Clear bulk mappings
                var bulkMappings = _mappingManager.GetAllBulkRanges().ToList();
                foreach (var bulk in bulkMappings)
                {
                    _mappingManager.RemoveBulkRangeMapping(bulk.Prefix, bulk.StartIndex, bulk.EndIndex);
                }

                LoadExistingMappings();
                UpdateStatus("Alle mappings slettet");
            }
        }

        private Queue<string> _specificMappingQueue;
        private bool _isInSpecificMappingMode = false;

        // Kø for rekke-mapping (bulk mapping av terminalblokker)
        private Queue<string> _rangeMappingQueue;
        private bool _isInRangeMappingMode = false;

        // ==== Bulk mapping mode state ====
        // Brukes når brukeren vil mappe et helt intervall av rekkeklemmer til
        // flere grid-celler i ett steg.  Når aktiv vil RangeMapping-knappen
        // trigge StartBulkMappingSelection i MainWindow, FinishBulkMappingButton
        // vises, og OnBulkMappingCompleted kalles når brukeren er ferdig.
        private bool _isBulkMappingMode = false;
        private string _bulkPrefix;
        private int _bulkStart;
        private int _bulkEnd;

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
                var unmappedCells = FilterUnmappedCells(uniqueCells);

                if (unmappedCells.Count == 0)
                {
                    MessageBox.Show("Alle celler er allerede mappet!", "Ingen celler å mappe",
                                   MessageBoxButton.OK, MessageBoxImage.Information);
                    return;
                }

                _specificMappingQueue = new Queue<string>(unmappedCells);
                _isInSpecificMappingMode = true;

                // Notify MainWindow that we're starting sequential mapping mode
                _mainWindow.StartSequentialMappingMode();

                ProcessNextSpecificMapping();

                // Minimize this window during sequential mapping
                this.WindowState = WindowState.Minimized;

                UpdateStatus($"Starter spesifikk mapping av {unmappedCells.Count} celler.");
            }
            catch (Exception ex)
            {
                UpdateStatus($"Feil ved oppstart av spesifikk mapping: {ex.Message}");
            }
        }

        private List<string> GetUniqueExcelCellsOrdered()
        {
            var columnBCells = new HashSet<string>();
            var columnCCells = new HashSet<string>();
            var usedRange = _mainWindow.worksheet.UsedRange;
            if (usedRange == null) return new List<string>();

            var lastRow = usedRange.Rows.Count;

            // Collect cells from column B and C separately
            for (int row = 2; row <= lastRow; row++)
            {
                try
                {
                    var cellB = (_mainWindow.worksheet.Cells[row, 2] as Excel.Range)?.Value?.ToString() ?? "";
                    var cellC = (_mainWindow.worksheet.Cells[row, 3] as Excel.Range)?.Value?.ToString() ?? "";

                    if (!string.IsNullOrWhiteSpace(cellB)) columnBCells.Add(cellB.Trim());
                    if (!string.IsNullOrWhiteSpace(cellC)) columnCCells.Add(cellC.Trim());
                }
                catch { }
            }

            // Return with column B cells first, then column C cells
            var result = new List<string>();
            result.AddRange(columnBCells.OrderBy(x => x));
            result.AddRange(columnCCells.Except(columnBCells).OrderBy(x => x)); // Avoid duplicates

            return result;
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

            foreach (var cell in cells)
            {
                if (!_mappingManager.IsReferenceMapped(cell))
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
        }

        private void OnSpecificMappingCompleted(string reference, string ignore)
        {
            // DO NOT restore and activate window after each mapping
            // Just update the mappings list
            LoadExistingMappings();

            // Continue with next mapping immediately
            System.Threading.Tasks.Task.Delay(100).ContinueWith(_ =>
            {
                this.Dispatcher.Invoke(() => ProcessNextSpecificMapping());
            });
        }

        /// <summary>
        /// Called when sequential mapping mode is finished (user clicked "Ferdig")
        /// </summary>
        public void OnSequentialMappingFinished()
        {
            EndSpecificMappingMode();
            this.WindowState = WindowState.Normal;
            this.Activate();
            LoadExistingMappings();
        }

        /// <summary>
        /// Puts a reference back at the front of the sequential mapping queue
        /// </summary>
        public void PutReferenceBackInQueue(string reference)
        {
            if (_specificMappingQueue != null && _isInSpecificMappingMode)
            {
                // Create a new queue with the undone reference at the front
                var newQueue = new Queue<string>();
                newQueue.Enqueue(reference);

                // Add all remaining items from the original queue
                while (_specificMappingQueue.Count > 0)
                {
                    newQueue.Enqueue(_specificMappingQueue.Dequeue());
                }

                _specificMappingQueue = newQueue;

                // Continue with the next mapping (which is now the undone reference)
                ProcessNextSpecificMapping();
            }
        }

        /// <summary>
        /// Starter bulk-mapping av et intervall av terminalblokker (f.eks. "X2:21-100").
        /// I stedet for å mappe én og én referanse sekvensielt, åpnes hovedvinduet i
        /// multi-valgmodus der brukeren kan klikke på flere grid-celler som representerer
        /// den fysiske skinnen. Når brukeren fullfører valget via "Ferdig bulk mapping",
        /// lagres én eneste BulkRangeMapping med alle de valgte cellene.
        /// </summary>
        private void StartRangeMapping_Click(object sender, RoutedEventArgs e)
        {
            var text = NewExcelReference.Text?.Trim();
            if (string.IsNullOrWhiteSpace(text))
            {
                UpdateStatus("Skriv inn prefiks eller intervall først");
                return;
            }

            // Gjenkjenn formatet PREFIX:START-END eller PREFIX:NUMMER
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

            int endNum;
            if (string.IsNullOrEmpty(endStr))
            {
                endNum = startNum;
            }
            else if (!int.TryParse(endStr, out endNum))
            {
                MessageBox.Show("Sluttnummeret kunne ikke tolkes", "Feil", MessageBoxButton.OK, MessageBoxImage.Warning);
                return;
            }

            if (endNum < startNum)
            {
                MessageBox.Show("Sluttnummeret må være større eller lik startnummeret", "Feil", MessageBoxButton.OK, MessageBoxImage.Warning);
                return;
            }

            // Start bulk mapping modus
            _isBulkMappingMode = true;
            _bulkPrefix = prefix;
            _bulkStart = startNum;
            _bulkEnd = endNum;
            FinishBulkMappingButton.Visibility = Visibility.Visible;

            UpdateStatus($"Bulk mapping: velg grid-celler for {prefix}:{startNum}-{endNum}. Klikk på 'Ferdig bulk mapping' når du er ferdig.");

            // Inform the user via a dialog so they understand that the bulk mapping
            // selection happens in the main window.  This message appears
            // before the window is minimized so the user sees it.
            MessageBox.Show(
                "Bulk mapping startet. Gå tilbake til hovedvinduet og klikk på en eller flere T/B-celler for å velge posisjoner.\n\n" +
                "Når du er ferdig, klikk på 'Ferdig bulk mapping' i hovedvinduet for å fullføre.",
                "Bulk mapping",
                MessageBoxButton.OK,
                MessageBoxImage.Information);

            // Start bulk selection i MainWindow
            _mainWindow.StartBulkMappingSelection(prefix, startNum, endNum, OnBulkMappingCompleted);

            // Minimer dette vinduet slik at brukeren kan klikke på gridet
            this.WindowState = WindowState.Minimized;
        }

        /// <summary>
        /// Behandler neste referanse i rekke-mapping modus. Kalles etter at en mapping er fullført.
        /// </summary>
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
            // Oppdater liste og gjenoppta state (men ikke fullfør enda)
            LoadExistingMappings();
            // Fortsett med neste mapping via dispatcher etter en liten delay
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

        /// <summary>
        /// Fullfører bulk-mapping når brukeren klikker på "Ferdig bulk mapping"-knappen.
        /// Denne metoden instruerer MainWindow om å avslutte bulkselection og
        /// trigger deretter vår callback via OnBulkMappingCompleted.
        /// </summary>
        private void FinishBulkMappingButton_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                // Signaliser til MainWindow at brukeren er ferdig med å velge celler
                _mainWindow?.FinishBulkMappingSelection();
            }
            catch (Exception ex)
            {
                UpdateStatus($"Feil ved fullføring av bulk mapping: {ex.Message}");
            }
        }

        /// <summary>
        /// Callback som mottar resultatet av en bulk-mapping fra MainWindow.  Her
        /// konverteres de valgte cellene til en liste av CellCoord og lagres i
        /// mappingManager.  Etterpå oppdateres UI og status og bulk-modus
        /// avsluttes.
        /// </summary>
        private void OnBulkMappingCompleted(string prefix, int startNumber, int endNumber, List<(int Row, int Col)> cells)
        {
            // Kjør på UI-tråden for å sikre at vi kan oppdatere kontroller
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

                    // Get side information from MainWindow
                    bool selectedIsTop = _mainWindow.BulkMappingSelectedIsTop;

                    // Lagre bulk mapping til fil via manager
                    _mappingManager.AddBulkRangeMapping(prefix, startNumber, endNumber, cells, selectedIsTop);

                    // Oppdater UI
                    FinishBulkMappingButton.Visibility = Visibility.Collapsed;
                    _isBulkMappingMode = false;
                    NewExcelReference.Clear();

                    // Ensure the mappings list is updated to show the new bulk mapping
                    LoadExistingMappings();

                    this.WindowState = WindowState.Normal;
                    this.Activate();
                    UpdateStatus($"Bulk mapping lagret: {prefix}:{startNumber}-{endNumber} ({cells.Count} celler)");
                }
                catch (Exception ex)
                {
                    UpdateStatus($"Feil under lagring av bulk mapping: {ex.Message}");
                    MessageBox.Show($"Feil under lagring av bulk mapping: {ex.Message}", "Feil", MessageBoxButton.OK, MessageBoxImage.Error);
                }
            });
        }

        private void EndSpecificMappingMode()
        {
            _isInSpecificMappingMode = false;
            _specificMappingQueue = null;
            _mainWindow.EndMappingMode();
            _mainWindow.EndSequentialMappingMode();
            UpdateStatus("Spesifikk mapping fullført");
        }

        protected override void OnKeyDown(System.Windows.Input.KeyEventArgs e)
        {
            if (e.Key == System.Windows.Input.Key.Escape)
            {
                // Avbryt spesifikk mapping eller rekke-mapping med Esc
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

            // Notify MainWindow that we're starting sequential mapping mode
            _mainWindow.StartSequentialMappingMode();

            ProcessNextSpecificMapping();

            // Minimize this window during sequential mapping
            this.WindowState = WindowState.Minimized;

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