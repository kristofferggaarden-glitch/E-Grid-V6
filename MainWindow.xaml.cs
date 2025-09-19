using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Input;
using System.Windows.Media;
using Microsoft.Win32;
using Excel = Microsoft.Office.Interop.Excel;

namespace WpfEGridApp
{
    public partial class MainWindow : Window, INotifyPropertyChanged
    {
        private int _sections = 5;
        private int _rows = 7;
        private int _cols = 4;
        private string _selectedExcelFile;
        private string _excelDisplayText;
        private SpecialPoint _lockedPointA;
        private Button _lockedButton;
        private int _currentExcelRow = 2;
        private ComponentMappingManager _componentMappingManager;
        private string _currentMappingReference = "";
        private bool _isInMappingMode = false;
        private Action<string, string> _mappingCompletedCallback;

        // Sequential mapping mode state
        private bool _isInSequentialMappingMode = false;
        private Stack<string> _sequentialMappingUndoStack = new Stack<string>();
        private ComponentMappingWindow _componentMappingWindow;

        // Indicates whether the application is in cell removal mode.  When true
        // regular cell clicks will remove the cell from the grid instead of
        // initiating a measurement.  The RemoveCellsButton toggles this flag.
        private bool _isRemovingCells = false;

        // ==== Bulk mapping mode state ====
        // When true the user is selecting multiple grid cells to map a range of
        // terminal blocks (rekkeklemmer).  Clicks on grid cells will toggle
        // membership in the selection and no measurement or normal mapping
        // processing will occur.  When finished the selection is passed back
        // to the ComponentMappingWindow via a callback.
        private bool _isBulkMappingSelectionMode = false;
        private string _bulkMappingPrefix;
        private int _bulkMappingStart;
        private int _bulkMappingEnd;
        private List<(int Row, int Col)> _bulkSelectedCells;
        private Action<string, int, int, List<(int Row, int Col)>> _bulkMappingCompletedCallback;

        // Flag to remember whether the current bulk mapping selection is on top (T) or bottom (B).
        // null means not yet decided.  True means T (overside), false means B (underside).
        private bool? _bulkMappingSelectedIsTop;

        /// <summary>
        /// Angir hvilken side (oversiden/T eller undersiden/B) som er valgt i gjeldende
        /// bulk-mapping.  Denne verdien settes når brukeren klikker på den første
        /// T- eller B-knappen i bulk mapping-modus.  Den brukes av
        /// ComponentMappingWindow når bulk mapping fullføres.
        /// </summary>
        public bool BulkMappingSelectedIsTop { get; private set; }

        public int Sections
        {
            get => _sections;
            set { _sections = value; OnPropertyChanged(nameof(Sections)); }
        }

        public int Rows
        {
            get => _rows;
            set { _rows = value; OnPropertyChanged(nameof(Rows)); }
        }

        public int Cols
        {
            get => _cols;
            set { _cols = value; OnPropertyChanged(nameof(Cols)); }
        }

        public string SelectedExcelFile
        {
            get => _selectedExcelFile;
            set { _selectedExcelFile = value; OnPropertyChanged(nameof(SelectedExcelFile)); }
        }

        public string ExcelDisplayText
        {
            get => _excelDisplayText;
            set { _excelDisplayText = value; OnPropertyChanged(nameof(ExcelDisplayText)); }
        }

        private readonly Dictionary<(int globalRow, int globalCol), Cell> allCells = new();
        private readonly Dictionary<(int globalRow, int globalCol), Cell> mappingCells = new();
        private readonly List<SpecialPoint> doorPoints = new();
        private readonly List<SpecialPoint> motorPoints = new();
        private object startPoint;
        private object endPoint;
        private Excel.Application excelApp;
        private Excel.Workbook workbook;
        public Excel.Worksheet worksheet;

        public MainWindow()
        {
            InitializeComponent();
            DataContext = this;
            UpdateExcelDisplayText();
            BuildAllSections();
        }

        private void BuildAllSections()
        {
            MainPanel.Children.Clear();
            allCells.Clear();
            mappingCells.Clear();
            doorPoints.Clear();
            motorPoints.Clear();
            ResultText.Text = "";
            startPoint = null;
            endPoint = null;

            for (int s = 0; s < Sections; s++)
            {
                var sectionPanel = new StackPanel
                {
                    Margin = new Thickness(15),
                    Orientation = Orientation.Vertical,
                    Background = Brushes.Transparent
                };

                // Door button
                var doorBtn = new Button
                {
                    Content = $"Door {s + 1}",
                    Width = 200,
                    Height = 40,
                    Style = (Style)FindResource("OrangeButtonStyle"),
                    HorizontalAlignment = HorizontalAlignment.Left,
                    ToolTip = $"Select door {s + 1} as a path point",
                    Tag = s
                };
                doorBtn.Click += (sender, e) => HandlePointClick(sender, doorPoints);
                sectionPanel.Children.Add(doorBtn);

                var doorLockBtn = new Button
                {
                    Content = $"Lock Door {s + 1}",
                    Width = 200,
                    Height = 40,
                    Style = (Style)FindResource("RoundedButtonStyle"),
                    HorizontalAlignment = HorizontalAlignment.Left,
                    ToolTip = $"Lock or unlock door {s + 1} as point A",
                    Tag = s
                };
                doorLockBtn.Click += (sender, e) => LockPointA_Click(sender, doorPoints);
                sectionPanel.Children.Add(doorLockBtn);

                // Grid
                var grid = new Grid
                {
                    Background = Brushes.Transparent,
                    Margin = new Thickness(0, 10, 0, 10)
                };
                sectionPanel.Children.Add(grid);

                // Motor button  
                var motorBtn = new Button
                {
                    Content = $"Motor {s + 1}",
                    Width = 200,
                    Height = 40,
                    Style = (Style)FindResource("OrangeButtonStyle"),
                    HorizontalAlignment = HorizontalAlignment.Left,
                    ToolTip = $"Select motor {s + 1} as a path point",
                    Tag = s
                };
                motorBtn.Click += (sender, e) => HandlePointClick(sender, motorPoints);
                sectionPanel.Children.Add(motorBtn);

                var motorLockBtn = new Button
                {
                    Content = $"Lock Motor {s + 1}",
                    Width = 200,
                    Height = 40,
                    Style = (Style)FindResource("RoundedButtonStyle"),
                    HorizontalAlignment = HorizontalAlignment.Left,
                    ToolTip = $"Lock or unlock motor {s + 1} as point A",
                    Tag = s
                };
                motorLockBtn.Click += (sender, e) => LockPointA_Click(sender, motorPoints);
                sectionPanel.Children.Add(motorLockBtn);

                // Setup grid structure
                grid.RowDefinitions.Clear();
                grid.ColumnDefinitions.Clear();
                for (int r = 0; r < Rows; r++)
                    grid.RowDefinitions.Add(new RowDefinition { Height = GridLength.Auto });
                for (int c = 0; c < Cols; c++)
                    grid.ColumnDefinitions.Add(new ColumnDefinition { Width = GridLength.Auto });

                // Add cells to grid
                for (int row = 0; row < Rows; row++)
                {
                    if (row % 2 == 0) // Even rows - regular cells
                    {
                        for (int col = 0; col < Cols; col++)
                            AddCell(grid, row, col, s);
                    }
                    else // Odd rows - mapping cells
                    {
                        AddCell(grid, row, 0, s); // First column regular cell
                        for (int col = 1; col < Cols; col++)
                            AddMappingCell(grid, row, col, s);
                    }
                }

                // Add special points
                doorPoints.Add(new SpecialPoint
                {
                    SectionIndex = s,
                    Type = SpecialPointType.Door,
                    Button = doorBtn,
                    GlobalRow = 0,
                    GlobalCol = s * Cols // Fixed: Direct section * cols
                });

                motorPoints.Add(new SpecialPoint
                {
                    SectionIndex = s,
                    Type = SpecialPointType.Motor,
                    Button = motorBtn,
                    GlobalRow = Rows - 1,
                    GlobalCol = s * Cols + (Cols - 1) // Fixed: Section * cols + last col in section
                });

                MainPanel.Children.Add(sectionPanel);
            }

            // Restore locked point if exists
            if (_lockedPointA != null)
            {
                var points = _lockedPointA.Type == SpecialPointType.Door ? doorPoints : motorPoints;
                var point = points.FirstOrDefault(p => p.SectionIndex == _lockedPointA.SectionIndex);
                if (point != null)
                {
                    startPoint = point;
                    point.Button.Background = new SolidColorBrush(Color.FromRgb(0, 120, 212));
                    ResultText.Text = $"Locked point A: {point.Type} {point.SectionIndex + 1}";
                    foreach (var panel in MainPanel.Children.OfType<StackPanel>())
                    {
                        foreach (var btn in panel.Children.OfType<Button>().Where(b => b.Tag != null && (int)b.Tag == point.SectionIndex && b.Content.ToString().StartsWith("Lock")))
                        {
                            if ((btn.Content.ToString().Contains("Door") && point.Type == SpecialPointType.Door) ||
                                (btn.Content.ToString().Contains("Motor") && point.Type == SpecialPointType.Motor))
                            {
                                _lockedButton = btn;
                                btn.Content = $"Unlock {point.Type} {point.SectionIndex + 1}";
                                break;
                            }
                        }
                    }
                }
                else
                {
                    _lockedPointA = null;
                    _lockedButton = null;
                }
            }
        }

        private void AddCell(Grid grid, int localRow, int localCol, int sectionIndex)
        {
            var btn = new Button
            {
                Width = 50,
                Height = 50,
                Margin = new Thickness(4, 4, 4, 4),
                Background = new SolidColorBrush(Color.FromRgb(74, 90, 91)),
                ToolTip = $"Cell ({localRow}, {localCol}) in section {sectionIndex + 1}",
                Style = (Style)FindResource("RoundedButtonStyle")
            };
            btn.Click += CellClick;
            Grid.SetRow(btn, localRow);
            Grid.SetColumn(btn, localCol);
            grid.Children.Add(btn);

            // Fixed global coordinates
            int globalRow = localRow;
            int globalCol = sectionIndex * Cols + localCol;
            allCells[(globalRow, globalCol)] = new Cell(globalRow, globalCol, btn);
        }

        private void AddMappingCell(Grid grid, int localRow, int localCol, int sectionIndex)
        {
            var btnTop = new Button
            {
                Width = 50,
                Height = 25,
                Margin = new Thickness(2, 2, 2, 2),
                Background = new SolidColorBrush(Color.FromRgb(80, 90, 100)),
                ToolTip = $"Mapping cell TOP ({localRow}, {localCol}) in section {sectionIndex + 1}",
                Style = (Style)FindResource("MappingButtonStyle"),
                Content = "T"
            };
            btnTop.Click += MappingCell_Click;

            var btnBottom = new Button
            {
                Width = 50,
                Height = 25,
                Margin = new Thickness(2, 2, 2, 2),
                Background = new SolidColorBrush(Color.FromRgb(60, 70, 80)),
                ToolTip = $"Mapping cell BOTTOM ({localRow}, {localCol}) in section {sectionIndex + 1}",
                Style = (Style)FindResource("MappingButtonStyle"),
                Content = "B"
            };
            btnBottom.Click += MappingCell_Click;

            var stackPanel = new StackPanel
            {
                Orientation = Orientation.Vertical,
                VerticalAlignment = VerticalAlignment.Center,
                HorizontalAlignment = HorizontalAlignment.Center
            };

            stackPanel.Children.Add(btnTop);
            stackPanel.Children.Add(btnBottom);

            Grid.SetRow(stackPanel, localRow);
            Grid.SetColumn(stackPanel, localCol);
            grid.Children.Add(stackPanel);

            // Fixed mapping cell coordinates
            int globalRow = localRow;
            int globalCol = sectionIndex * Cols + localCol;

            // Store mapping cells with proper keys
            mappingCells[(globalRow, globalCol)] = new Cell(globalRow, globalCol, btnTop);  // Top side
            mappingCells[(-globalRow - 1, globalCol)] = new Cell(-globalRow - 1, globalCol, btnBottom); // Bottom side (negative)
        }

        /// <summary>
        /// Starts sequential mapping mode - shows undo and finish buttons
        /// </summary>
        public void StartSequentialMappingMode()
        {
            _isInSequentialMappingMode = true;
            _sequentialMappingUndoStack.Clear();
            SequentialMappingControls.Visibility = Visibility.Visible;
            UpdateUndoButtonState();
        }

        /// <summary>
        /// Ends sequential mapping mode - hides undo and finish buttons
        /// </summary>
        public void EndSequentialMappingMode()
        {
            _isInSequentialMappingMode = false;
            _sequentialMappingUndoStack.Clear();
            SequentialMappingControls.Visibility = Visibility.Collapsed;
        }

        /// <summary>
        /// Updates the undo button enabled state based on stack contents
        /// </summary>
        private void UpdateUndoButtonState()
        {
            if (UndoMappingButton != null)
            {
                UndoMappingButton.IsEnabled = _sequentialMappingUndoStack.Count > 0;
            }
        }

        /// <summary>
        /// Records a mapping in the undo stack
        /// </summary>
        private void RecordMappingForUndo(string excelReference)
        {
            if (_isInSequentialMappingMode)
            {
                _sequentialMappingUndoStack.Push(excelReference);
                UpdateUndoButtonState();
            }
        }

        /// <summary>
        /// Handles undo button click - removes the last mapping and puts it back in queue
        /// </summary>
        private void UndoMapping_Click(object sender, RoutedEventArgs e)
        {
            if (!_isInSequentialMappingMode || _sequentialMappingUndoStack.Count == 0)
                return;

            var lastMapping = _sequentialMappingUndoStack.Pop();

            // Remove the mapping from storage
            if (_componentMappingManager != null)
            {
                _componentMappingManager.RemoveMapping(lastMapping);

                // Update the component mapping window if it exists and is loaded
                if (_componentMappingWindow != null)
                {
                    try
                    {
                        _componentMappingWindow.LoadExistingMappings();
                    }
                    catch
                    {
                        // If the window is closed or disposed, ignore the error
                    }
                }
            }

            // Put the reference back at the front of the queue to be mapped again
            if (_componentMappingWindow != null)
            {
                _componentMappingWindow.PutReferenceBackInQueue(lastMapping);
            }

            UpdateUndoButtonState();
            ResultText.Text = $"Angret mapping for: {lastMapping} - klar for ny mapping";
        }

        /// <summary>
        /// Handles finish button click - ends sequential mapping and opens component mapping window
        /// </summary>
        private void FinishMapping_Click(object sender, RoutedEventArgs e)
        {
            EndSequentialMappingMode();
            EndMappingMode();

            // Open component mapping window
            if (_componentMappingWindow != null)
            {
                _componentMappingWindow.OnSequentialMappingFinished();
            }
            else
            {
                OpenComponentMapping_Click(sender, e);
            }
        }

        private void MappingCell_Click(object sender, RoutedEventArgs e)
        {
            // During bulk mapping selection we want mapping cell clicks to behave like
            // normal cell clicks so that multiple mapping cells (T/B) can be selected.
            // Redirect the event to the CellClick handler which contains bulk
            // selection logic.  Note that CellClick checks _isBulkMappingSelectionMode
            // and will perform appropriate selection highlighting and side locking.
            if (_isBulkMappingSelectionMode)
            {
                CellClick(sender, e);
                return;
            }

            var btn = sender as Button;

            if (_isInMappingMode)
            {
                Cell cell = null;
                bool isBottomSide = false;

                // Find which mapping cell was clicked
                foreach (var kvp in mappingCells)
                {
                    if (kvp.Value.ButtonRef == btn)
                    {
                        cell = kvp.Value;
                        isBottomSide = kvp.Key.Item1 < 0; // Negative row = bottom side
                        break;
                    }
                }

                if (cell != null && !string.IsNullOrEmpty(_currentMappingReference))
                {
                    int actualRow = isBottomSide ? -(cell.Row + 1) : cell.Row;

                    _componentMappingManager?.AddMapping(_currentMappingReference, actualRow, cell.Col, isBottomSide);

                    // Record for undo
                    RecordMappingForUndo(_currentMappingReference);

                    _mappingCompletedCallback?.Invoke(_currentMappingReference, "");
                    EndMappingMode();
                }
            }
        }

        // Fortsetter i del 2...
        // Del 2 av MainWindow.xaml.cs - Versjon 2.0

        public void EndMappingMode()
        {
            _isInMappingMode = false;
            _currentMappingReference = "";
            _mappingCompletedCallback = null;

            // Hide mapping indicator and cancel button (but not sequential mapping controls)
            MappingIndicator.Visibility = Visibility.Collapsed;
            if (!_isInSequentialMappingMode)
            {
                CancelMappingButton.Visibility = Visibility.Collapsed;
            }

            ResetCellColors();
            foreach (var cell in allCells.Values)
            {
                cell.ButtonRef.IsEnabled = true;
            }

            // Reset Door and Motor buttons to original color
            foreach (var dp in doorPoints)
                dp.Button.Background = (Brush)FindResource("OrangeButtonBrush");
            foreach (var mp in motorPoints)
                mp.Button.Background = (Brush)FindResource("OrangeButtonBrush");

            if (!_isInSequentialMappingMode)
            {
                ResultText.Text = "";
            }
        }

        private void ResetCellColors()
        {
            foreach (var cell in allCells.Values)
            {
                cell.ButtonRef.Background = new SolidColorBrush(Color.FromRgb(74, 90, 91));
            }
            foreach (var cell in mappingCells.Values)
            {
                if (cell.ButtonRef.Content?.ToString() == "T")
                    cell.ButtonRef.Background = new SolidColorBrush(Color.FromRgb(80, 90, 100));
                else if (cell.ButtonRef.Content?.ToString() == "B")
                    cell.ButtonRef.Background = new SolidColorBrush(Color.FromRgb(60, 70, 80));
            }
        }

        public bool HasHorizontalNeighbor(int row, int col)
        {
            return allCells.ContainsKey((row, col - 1)) || allCells.ContainsKey((row, col + 1));
        }

        private void CellClick(object sender, RoutedEventArgs e)
        {
            var btn = sender as Button;

            // If we are in bulk mapping selection mode, treat clicks as selection toggles.
            if (_isBulkMappingSelectionMode)
            {
                if (btn != null)
                {
                    // Only allow mapping cells (T/B) to be selected
                    var mappingEntry = mappingCells.FirstOrDefault(kvp => kvp.Value.ButtonRef == btn);
                    if (!mappingEntry.Equals(default(KeyValuePair<(int, int), Cell>)))
                    {
                        // Determine whether user clicked on a top (T) or bottom (B) mapping cell
                        bool clickedIsTop = mappingEntry.Key.Item1 >= 0;
                        // Set side if not set, else check consistency
                        if (_bulkMappingSelectedIsTop == null)
                        {
                            _bulkMappingSelectedIsTop = clickedIsTop;
                            BulkMappingSelectedIsTop = clickedIsTop;
                        }
                        else if (_bulkMappingSelectedIsTop != clickedIsTop)
                        {
                            // Disallow mixing T and B cells in a single bulk mapping
                            MessageBox.Show("Du kan kun velge enten T eller B celler under bulk mapping.", "Ugyldig valg", MessageBoxButton.OK, MessageBoxImage.Warning);
                            return;
                        }
                        // Convert mapping key to actual row index
                        int actualRow = clickedIsTop ? mappingEntry.Key.Item1 : -(mappingEntry.Key.Item1 + 1);
                        int column = mappingEntry.Key.Item2;
                        bool alreadySelected = _bulkSelectedCells?.Any(c => c.Row == actualRow && c.Col == column) ?? false;
                        if (alreadySelected)
                        {
                            // Remove from selection
                            _bulkSelectedCells?.RemoveAll(c => c.Row == actualRow && c.Col == column);
                            // Reset visual on both T and B buttons for this cell
                            // Reset colour on clicked button
                            btn.Background = new SolidColorBrush(Color.FromRgb(80, 120, 140));
                            // Highlight counterpart (the other side) if exists
                            var topKey = (actualRow, column);
                            var bottomKey = (-actualRow - 1, column);
                            if (mappingCells.TryGetValue(topKey, out var topCell))
                                topCell.ButtonRef.Background = new SolidColorBrush(Color.FromRgb(80, 120, 140));
                            if (mappingCells.TryGetValue(bottomKey, out var bottomCell))
                                bottomCell.ButtonRef.Background = new SolidColorBrush(Color.FromRgb(80, 120, 140));
                        }
                        else
                        {
                            // Add to selection
                            _bulkSelectedCells?.Add((actualRow, column));
                            // Highlight selected colour on the clicked button and its counterpart
                            var selectedColour = new SolidColorBrush(Color.FromRgb(0, 178, 148));
                            btn.Background = selectedColour;
                            var topKey = (actualRow, column);
                            var bottomKey = (-actualRow - 1, column);
                            if (mappingCells.TryGetValue(topKey, out var topCell))
                                topCell.ButtonRef.Background = selectedColour;
                            if (mappingCells.TryGetValue(bottomKey, out var bottomCell))
                                bottomCell.ButtonRef.Background = selectedColour;
                        }
                    }
                    // Ignore clicks on regular cells during bulk mapping selection
                }
                return;
            }

            // If we are in removal mode, remove the clicked cell from the grid and
            // internal cell dictionary.  Skip any measurement processing.
            if (_isRemovingCells)
            {
                if (btn != null)
                {
                    // Locate the corresponding cell in the allCells dictionary
                    var cellEntry = allCells.FirstOrDefault(kvp => kvp.Value.ButtonRef == btn);
                    if (!cellEntry.Equals(default(KeyValuePair<(int globalRow, int globalCol), Cell>)))
                    {
                        // Instead of removing the Button from the Grid (which causes
                        // the row to collapse), disable it and hide it from view.  A
                        // hidden element still occupies layout space, preserving the
                        // visual structure of the grid.  We then remove the cell
                        // from the allCells dictionary so that path finding no
                        // longer considers it.
                        btn.IsEnabled = false;
                        btn.Visibility = Visibility.Hidden;
                        allCells.Remove(cellEntry.Key);
                    }
                }
                return;
            }
            if (_lockedPointA != null)
            {
                if (endPoint != null)
                    ResetSelection();
                endPoint = btn;
                btn.Background = new SolidColorBrush(Color.FromRgb(232, 17, 35));
                ProcessPath();
            }
            else
            {
                if (startPoint != null && endPoint != null)
                    ResetSelection();
                if (startPoint == null)
                {
                    startPoint = btn;
                    btn.Background = new SolidColorBrush(Color.FromRgb(0, 120, 212));
                    ResultText.Text = "";
                }
                else if (endPoint == null)
                {
                    endPoint = btn;
                    btn.Background = new SolidColorBrush(Color.FromRgb(232, 17, 35));
                    ProcessPath();
                }
            }
        }

        private void HandlePointClick(object sender, List<SpecialPoint> points)
        {
            var btn = sender as Button;
            if (btn.Tag == null) return;
            int sectionIndex = (int)btn.Tag;
            var special = points.FirstOrDefault(p => p.SectionIndex == sectionIndex);
            if (special == null) return;

            // Check if we're in mapping mode
            if (_isInMappingMode && !string.IsNullOrEmpty(_currentMappingReference))
            {
                // Map to special point (Door or Motor)
                int specialRow = special.Type == SpecialPointType.Door ? -2 : -1;
                int specialCol = 1000 + sectionIndex; // Special encoding for Door/Motor

                _componentMappingManager?.AddMapping(_currentMappingReference, specialRow, specialCol, false);

                // Record for undo
                RecordMappingForUndo(_currentMappingReference);

                _mappingCompletedCallback?.Invoke(_currentMappingReference, "");
                EndMappingMode();
                return;
            }

            // Normal click handling for manual measurement
            if (_lockedPointA != null)
            {
                if (endPoint != null)
                    ResetSelection();
                endPoint = special;
                special.Button.Background = new SolidColorBrush(Color.FromRgb(232, 17, 35));
                ProcessPath();
            }
            else
            {
                if (startPoint != null && endPoint != null)
                    ResetSelection();
                if (startPoint == null)
                {
                    startPoint = special;
                    special.Button.Background = new SolidColorBrush(Color.FromRgb(0, 120, 212));
                    ResultText.Text = "";
                }
                else if (endPoint == null)
                {
                    endPoint = special;
                    special.Button.Background = new SolidColorBrush(Color.FromRgb(232, 17, 35));
                    ProcessPath();
                }
            }
        }

        private void LockPointA_Click(object sender, List<SpecialPoint> points)
        {
            var btn = sender as Button;
            if (btn.Tag == null) return;
            int sectionIndex = (int)btn.Tag;
            var special = points.FirstOrDefault(p => p.SectionIndex == sectionIndex);
            if (special == null) return;

            if (_lockedPointA == special)
            {
                _lockedPointA = null;
                _lockedButton = null;
                btn.Content = $"Lock {special.Type} {sectionIndex + 1}";
                ResetSelection();
            }
            else
            {
                if (_lockedPointA != null && _lockedButton != null)
                {
                    _lockedButton.Content = $"Lock {_lockedPointA.Type} {_lockedPointA.SectionIndex + 1}";
                    _lockedPointA.Button.Background = (Brush)FindResource("OrangeButtonBrush");
                }
                _lockedPointA = special;
                _lockedButton = btn;
                ResetSelection();
                startPoint = special;
                special.Button.Background = new SolidColorBrush(Color.FromRgb(0, 120, 212));
                btn.Content = $"Unlock {special.Type} {sectionIndex + 1}";
                ResultText.Text = $"Locked point A: {special.Type} {sectionIndex + 1}";
            }
        }

        private void ProcessPath()
        {
            Cell startCell = null;
            Cell endCell = null;

            // Ekstra distanse for å justere start- og sluttdistanser. Vi
            // ønsker at motor/dør som startpunkt skal ha samme bidrag som når de
            // er sluttpunkter (500 for motor, 1000 for dør), og at vanlige
            // celler bidrar med 200 uavhengig av om de er start eller slutt.
            // PathFinder.CalculateDistance legger til 100 for startpunkt og
            // 200 for sluttpunkt (kun dersom slutt ikke er spesial). Vi
            // kompenserer derfor her ved å justere med ±100 for spesial og
            // vanlige celler slik at summene blir symmetriske.
            double extraDistance = 0;

            if (startPoint is SpecialPoint spStart)
            {
                // Spesialpunkt som start: motor = 400, dør = 900. Når dette legges
                // sammen med 100 som PathFinder.CalculateDistance gir for
                // startkoblingen, får vi 500/1000 mm som ønsket.
                if (allCells.TryGetValue((spStart.GlobalRow, spStart.GlobalCol), out startCell))
                {
                    extraDistance += spStart.Type == SpecialPointType.Door ? 900 : 400;
                }
            }
            else if (startPoint is Button btnStart)
            {
                // Vanlig celle som start: legg til 100 mm ekstra slik at
                // totale startkostnaden blir 200 mm (100 fra CalculateDistance + 100 her)
                startCell = allCells.Values.FirstOrDefault(c => c.ButtonRef == btnStart);
                if (startCell != null)
                    extraDistance += 100;
            }

            if (endPoint is SpecialPoint spEnd)
            {
                // Sluttpunkt som spesial: motor = 500, dør = 1000. Ingen justering
                // nødvendig her da CalculateDistance ikke legger til 200 når
                // sluttpunktet er spesial.
                if (allCells.TryGetValue((spEnd.GlobalRow, spEnd.GlobalCol), out endCell))
                {
                    extraDistance += spEnd.Type == SpecialPointType.Door ? 1000 : 500;
                }
            }
            else if (endPoint is Button btnEnd)
            {
                // Vanlig celle som slutt: ingen ekstra distanse her, da
                // CalculateDistance allerede legger til 200 mm for sluttkoblingen.
                endCell = allCells.Values.FirstOrDefault(c => c.ButtonRef == btnEnd);
            }

            if (startCell == null || endCell == null)
            {
                ResultText.Text = "Invalid start or end point.";
                ResetSelection();
                return;
            }

            // Special handling when both points resolve to the same physical cell.
            // In this case we always want the distance to be 300 mm regardless of
            // whether the points are regular cells, motors or doors.  PathFinder
            // would otherwise return a zero-length path which, combined with the
            // existing compensation logic, yields much larger values for motors
            // and doors (e.g. 1050 mm).  By overriding here we provide a
            // consistent result for identical start and end points.
            if (startCell.Row == endCell.Row && startCell.Col == endCell.Col)
            {
                // Highlight just this cell
                HighlightPath(new List<Cell> { startCell });
                // Fast distanse for identiske start- og sluttpunkter
                double totalDistanceSame = 300;
                // Display and log the measurement
                ResultText.Text = $"Shortest path: {totalDistanceSame:F2} mm";
                FindNextAvailableRow();
                LogMeasurementToExcel(totalDistanceSame);
                _currentExcelRow++;
                UpdateExcelDisplayText();
                return;
            }

            var path = PathFinder.FindShortestPath(startCell, endCell, allCells, HasHorizontalNeighbor);
            if (path == null)
            {
                ResultText.Text = "No valid path found.";
                ResetSelection();
                return;
            }

            bool endsInSpecial = endPoint is SpecialPoint;
            // Calculate the base distance using the same logic as PathFinder.  This
            // includes a fixed 100 mm start connection and, if the end point is not
            // a special point, a 200 mm end connection.  For measurements
            // between two regular cells we want the total to be exactly 100 mm
            // (start) + path distance + 200 mm (end) and NOT include any extra
            // compensation or constant.  For other combinations (involving motors
            // or doors) we preserve the existing logic with extra adjustments and
            // the additional 50 mm.
            double baseDistance = PathFinder.CalculateDistance(path, endsInSpecial, HasHorizontalNeighbor);
            double totalDistance;
            bool startIsButton = startPoint is Button;
            bool endIsButton = endPoint is Button;
            if (startIsButton && endIsButton)
            {
                // Both points are regular cells: use only the base distance.
                totalDistance = baseDistance;
            }
            else
            {
                // At least one point is a motor or door; use existing compensation and
                // add the constant 50 mm as originally implemented.
                totalDistance = baseDistance + extraDistance + 50;
            }
            HighlightPath(path);

            if (startPoint is SpecialPoint sp1 && endPoint is SpecialPoint sp2)
            {
                if ((sp1.Type == SpecialPointType.Motor && sp2.Type == SpecialPointType.Door) ||
                    (sp1.Type == SpecialPointType.Door && sp2.Type == SpecialPointType.Motor))
                {
                    sp1.Button.Background = new SolidColorBrush(Color.FromRgb(0, 178, 148));
                    sp2.Button.Background = new SolidColorBrush(Color.FromRgb(0, 178, 148));
                }
            }

            ResultText.Text = $"Shortest path: {totalDistance:F2} mm";

            FindNextAvailableRow();
            LogMeasurementToExcel(totalDistance);
            _currentExcelRow++;
            UpdateExcelDisplayText();
        }

        private void FindNextAvailableRow()
        {
            try
            {
                if (string.IsNullOrEmpty(SelectedExcelFile) || worksheet == null)
                    return;

                while (_currentExcelRow <= 1000)
                {
                    var cellA = (worksheet.Cells[_currentExcelRow, 1] as Excel.Range)?.Value;
                    if (cellA == null || string.IsNullOrEmpty(cellA.ToString()))
                    {
                        break;
                    }
                    _currentExcelRow++;
                }
            }
            catch (Exception)
            {
            }
        }

        private void ResetSelection()
        {
            foreach (var cell in allCells.Values)
            {
                cell.ButtonRef.Background = new SolidColorBrush(Color.FromRgb(74, 90, 91));
            }
            foreach (var cell in mappingCells.Values)
            {
                if (cell.ButtonRef.Content?.ToString() == "T")
                    cell.ButtonRef.Background = new SolidColorBrush(Color.FromRgb(80, 90, 100));
                else if (cell.ButtonRef.Content?.ToString() == "B")
                    cell.ButtonRef.Background = new SolidColorBrush(Color.FromRgb(60, 70, 80));
            }
            foreach (var dp in doorPoints)
                dp.Button.Background = (Brush)FindResource("OrangeButtonBrush");
            foreach (var mp in motorPoints)
                mp.Button.Background = (Brush)FindResource("OrangeButtonBrush");
            startPoint = _lockedPointA;
            endPoint = null;
            if (_lockedPointA != null)
            {
                startPoint = _lockedPointA;
                _lockedPointA.Button.Background = new SolidColorBrush(Color.FromRgb(0, 120, 212));
                ResultText.Text = $"Locked point A: {_lockedPointA.Type} {_lockedPointA.SectionIndex + 1}";
            }
            else
            {
                ResultText.Text = "";
            }
        }

        private void HighlightPath(List<Cell> path)
        {
            foreach (var cell in allCells.Values)
            {
                cell.ButtonRef.Background = new SolidColorBrush(Color.FromRgb(74, 90, 91));
            }
            foreach (var cell in mappingCells.Values)
            {
                if (cell.ButtonRef.Content?.ToString() == "T")
                    cell.ButtonRef.Background = new SolidColorBrush(Color.FromRgb(80, 90, 100));
                else if (cell.ButtonRef.Content?.ToString() == "B")
                    cell.ButtonRef.Background = new SolidColorBrush(Color.FromRgb(60, 70, 80));
            }
            foreach (var dp in doorPoints)
                dp.Button.Background = (Brush)FindResource("OrangeButtonBrush");
            foreach (var mp in motorPoints)
                mp.Button.Background = (Brush)FindResource("OrangeButtonBrush");
            foreach (var cell in path)
                cell.ButtonRef.Background = new SolidColorBrush(Color.FromRgb(0, 178, 148));
            if (startPoint is Button b1)
                b1.Background = new SolidColorBrush(Color.FromRgb(0, 120, 212));
            if (endPoint is Button b2)
                b2.Background = new SolidColorBrush(Color.FromRgb(232, 17, 35));
            if (startPoint is SpecialPoint sp1)
                sp1.Button.Background = new SolidColorBrush(Color.FromRgb(0, 120, 212));
            if (endPoint is SpecialPoint sp2)
                sp2.Button.Background = new SolidColorBrush(Color.FromRgb(232, 17, 35));
        }

        private void LogMeasurementToExcel(double distance)
        {
            try
            {
                if (string.IsNullOrEmpty(SelectedExcelFile) || worksheet == null)
                {
                    MessageBox.Show("Please select an Excel file first.", "Error", MessageBoxButton.OK, MessageBoxImage.Warning);
                    return;
                }
                worksheet.Cells[_currentExcelRow, 1] = distance;
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Error writing to Excel: {ex.Message}", "Error", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        private void OpenComponentMapping_Click(object sender, RoutedEventArgs e)
        {
            if (string.IsNullOrEmpty(SelectedExcelFile))
            {
                MessageBox.Show("Velg først en Excel-fil", "Feil",
                               MessageBoxButton.OK, MessageBoxImage.Warning);
                return;
            }

            if (_componentMappingManager == null)
            {
                _componentMappingManager = new ComponentMappingManager(this, SelectedExcelFile);
            }

            // Store reference to the window
            _componentMappingWindow = new ComponentMappingWindow(this, _componentMappingManager);
            _componentMappingWindow.Show();
        }

        // NY METODE - Refresh Excel knapp
        private void RefreshExcel_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                if (string.IsNullOrEmpty(SelectedExcelFile) || worksheet == null)
                {
                    MessageBox.Show("Velg først en Excel-fil", "Feil",
                                   MessageBoxButton.OK, MessageBoxImage.Warning);
                    return;
                }

                // Find next available row by scanning from row 2
                _currentExcelRow = 2;
                FindNextAvailableRow();
                UpdateExcelDisplayText();

                ResultText.Text = $"Excel oppdatert - neste rad: {_currentExcelRow}";
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Feil under oppdatering: {ex.Message}", "Feil",
                               MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        public void StartInteractiveMapping(string excelReference, string description, Action<string, string> onCompleted)
        {
            _currentMappingReference = excelReference;
            _isInMappingMode = true;
            _mappingCompletedCallback = onCompleted;

            // Show mapping indicator with reference text
            MappingIndicatorText.Text = $"Mapper: {excelReference}";
            MappingIndicator.Visibility = Visibility.Visible;

            // Show cancel button only if not in sequential mapping mode
            if (!_isInSequentialMappingMode)
            {
                CancelMappingButton.Visibility = Visibility.Visible;
            }

            if (_isInSequentialMappingMode)
            {
                ResultText.Text = $"Klikk på grid-posisjonen for {excelReference}";
            }
            else
            {
                ResultText.Text = $"Klikk på grid-posisjonen for {excelReference} (mellomrader, Door eller Motor knapper)";
            }

            // Highlight mapping cells
            foreach (var cell in mappingCells.Values)
            {
                cell.ButtonRef.Background = new SolidColorBrush(Color.FromRgb(100, 120, 140));
            }

            // Disable regular cells
            foreach (var cell in allCells.Values)
            {
                cell.ButtonRef.IsEnabled = false;
                cell.ButtonRef.Background = new SolidColorBrush(Color.FromRgb(40, 40, 40));
            }

            // Highlight Door and Motor buttons as mappable
            foreach (var dp in doorPoints)
            {
                dp.Button.Background = new SolidColorBrush(Color.FromRgb(255, 140, 0)); // Orange highlight
            }
            foreach (var mp in motorPoints)
            {
                mp.Button.Background = new SolidColorBrush(Color.FromRgb(255, 140, 0)); // Orange highlight
            }

            this.Activate();
            this.Focus();
        }

        /// <summary>
        /// Start bulk mapping selection.  The user is prompted to select one or more
        /// grid cells that represent a physical rail carrying a range of terminal
        /// numbers (prefix:start-end).  Once the selection is completed via
        /// FinishBulkMappingSelection, the callback is invoked with the prefix,
        /// start and end values and the list of selected cell coordinates.
        /// </summary>
        public void StartBulkMappingSelection(string prefix, int startNumber, int endNumber,
                                              Action<string, int, int, List<(int Row, int Col)>> onCompleted)
        {
            // Ensure any existing mapping or removal modes are disabled
            EndMappingMode();
            _isRemovingCells = false;

            _isBulkMappingSelectionMode = true;
            _bulkMappingPrefix = prefix;
            _bulkMappingStart = startNumber;
            _bulkMappingEnd = endNumber;
            _bulkSelectedCells = new List<(int Row, int Col)>();
            _bulkMappingCompletedCallback = onCompleted;

            // Reset selected side state for this bulk selection
            _bulkMappingSelectedIsTop = null;
            BulkMappingSelectedIsTop = false;

            // Display instructions to the user
            MappingIndicatorText.Text = $"Bulk mapping: {prefix}:{startNumber}-{endNumber}. Velg flere celler i gridet.";
            MappingIndicator.Visibility = Visibility.Visible;
            CancelMappingButton.Visibility = Visibility.Visible;

            // Show the Complete Bulk Mapping button in the overlay so the user can
            // finish the bulk mapping directly from the main window without
            // switching back to the ComponentMappingWindow.  It will be
            // hidden again when bulk mapping ends.
            if (CompleteBulkMappingButton != null)
            {
                CompleteBulkMappingButton.Visibility = Visibility.Visible;
            }

            // Enable all grid cells for selection and visually highlight them slightly
            foreach (var cell in allCells.Values)
            {
                if (cell.ButtonRef != null)
                {
                    cell.ButtonRef.IsEnabled = true;
                    // Light grey-blue highlight to indicate selectable state
                    cell.ButtonRef.Background = new SolidColorBrush(Color.FromRgb(80, 120, 140));
                }
            }
            // Enable mapping cells and highlight them similarly
            foreach (var kvp in mappingCells)
            {
                var mappingCell = kvp.Value;
                if (mappingCell.ButtonRef != null)
                {
                    mappingCell.ButtonRef.IsEnabled = true;
                    mappingCell.ButtonRef.Background = new SolidColorBrush(Color.FromRgb(80, 120, 140));
                }
            }

            // Disable door and motor buttons to avoid accidental selection during bulk mapping
            foreach (var dp in doorPoints)
                dp.Button.IsEnabled = false;
            foreach (var mp in motorPoints)
                mp.Button.IsEnabled = false;

            // Inform the user in the ResultText as well
            ResultText.Text = $"Bulk mapping: Velg grid-celler for {prefix}:{startNumber}-{endNumber}. Klikk igjen på et valgt celle for å fjerne den.";
        }

        /// <summary>
        /// Completes the bulk mapping selection and calls the registered callback.  This
        /// should be invoked by the ComponentMappingWindow when the user clicks the
        /// 'Ferdig bulk mapping' button.  The selection will be cleared and the UI
        /// restored to its normal state.
        /// </summary>
        public void FinishBulkMappingSelection()
        {
            if (!_isBulkMappingSelectionMode)
                return;

            // Capture state locally before resetting
            var prefix = _bulkMappingPrefix;
            var startNumber = _bulkMappingStart;
            var endNumber = _bulkMappingEnd;
            var selectedCells = _bulkSelectedCells != null ? new List<(int Row, int Col)>(_bulkSelectedCells) : new List<(int, int)>();

            // End bulk mapping mode and restore UI
            EndBulkMappingMode();

            // Invoke callback if provided
            _bulkMappingCompletedCallback?.Invoke(prefix, startNumber, endNumber, selectedCells);
        }

        /// <summary>
        /// Cancels bulk mapping selection without invoking the callback.  This
        /// method is called when the user presses the 'Avbryt Mapping' button
        /// or presses Escape while in bulk mapping mode.
        /// </summary>
        private void CancelBulkMappingSelection()
        {
            if (!_isBulkMappingSelectionMode)
                return;

            EndBulkMappingMode();
        }

        /// <summary>
        /// Internal helper to exit bulk mapping mode and restore the grid
        /// appearance and enabled state.  Does not call any callbacks.
        /// </summary>
        private void EndBulkMappingMode()
        {
            // Reset selection mode flag
            _isBulkMappingSelectionMode = false;
            _bulkMappingPrefix = null;
            _bulkMappingStart = 0;
            _bulkMappingEnd = 0;
            _bulkSelectedCells = null;
            _bulkMappingCompletedCallback = null;
            _bulkMappingSelectedIsTop = null;
            BulkMappingSelectedIsTop = false;

            // Restore grid cell colours and enablement
            foreach (var cell in allCells.Values)
            {
                if (cell.ButtonRef != null)
                {
                    // Only reset cells that are still enabled (not removed)
                    if (cell.ButtonRef.IsEnabled)
                    {
                        cell.ButtonRef.Background = new SolidColorBrush(Color.FromRgb(74, 90, 91));
                    }
                }
            }
            foreach (var kvp in mappingCells)
            {
                var mappingCell = kvp.Value;
                if (mappingCell.ButtonRef != null)
                {
                    var content = mappingCell.ButtonRef.Content?.ToString();
                    if (content == "T")
                        mappingCell.ButtonRef.Background = new SolidColorBrush(Color.FromRgb(80, 90, 100));
                    else if (content == "B")
                        mappingCell.ButtonRef.Background = new SolidColorBrush(Color.FromRgb(60, 70, 80));
                }
            }

            // Re-enable door and motor buttons
            foreach (var dp in doorPoints)
                dp.Button.IsEnabled = true;
            foreach (var mp in motorPoints)
                mp.Button.IsEnabled = true;

            // Hide mapping indicator
            MappingIndicator.Visibility = Visibility.Collapsed;
            CancelMappingButton.Visibility = Visibility.Collapsed;
            // Hide Complete Bulk Mapping button when ending bulk selection
            if (CompleteBulkMappingButton != null)
            {
                CompleteBulkMappingButton.Visibility = Visibility.Collapsed;
            }
            ResultText.Text = string.Empty;
        }

        private void CancelMapping_Click(object sender, RoutedEventArgs e)
        {
            // Cancel whichever mapping mode is active.  If a single interactive mapping
            // is in progress, end it.  If a bulk mapping selection is in progress,
            // cancel it without invoking any callbacks.  Otherwise do nothing.
            if (_isInMappingMode)
            {
                EndMappingMode();
            }
            else if (_isBulkMappingSelectionMode)
            {
                CancelBulkMappingSelection();
            }
        }

        /// <summary>
        /// Handles the click on the "Ferdig bulk mapping" button that appears in the
        /// main window during bulk mapping.  This simply finishes the bulk
        /// mapping selection, invoking any registered callback and restoring
        /// the UI to its normal state.  The button is hidden again once
        /// the selection is complete.
        /// </summary>
        private void CompleteBulkMappingButton_Click(object sender, RoutedEventArgs e)
        {
            FinishBulkMappingSelection();
        }

        /// <summary>
        /// Toggle cell removal mode on or off.  When removal mode is active
        /// clicking on a regular grid cell will remove it from the UI and
        /// underlying data structure so it is no longer considered in
        /// subsequent measurements.  Toggling the mode off restores normal
        /// measurement behaviour.  The button's label is updated to reflect
        /// the current state and a message is shown in the result text to
        /// guide the user.
        /// </summary>
        private void RemoveCellsButton_Click(object sender, RoutedEventArgs e)
        {
            _isRemovingCells = !_isRemovingCells;
            if (_isRemovingCells)
            {
                // Enter removal mode: update button text and inform user
                if (RemoveCellsButton != null)
                {
                    RemoveCellsButton.Content = "Ferdig fjerning";
                }
                // Clear any current selection and locked points
                startPoint = null;
                endPoint = null;
                ResultText.Text = "Klikk på cellene du vil fjerne.";
            }
            else
            {
                // Exit removal mode: restore button text and clear message
                if (RemoveCellsButton != null)
                {
                    RemoveCellsButton.Content = "Fjern celler";
                }
                ResultText.Text = _lockedPointA != null ? $"Locked point A: {_lockedPointA.Type} {_lockedPointA.SectionIndex + 1}" : "";
                // Reset any temporary highlights
                ResetSelection();
            }
        }

        public void UpdateExcelDisplayText()
        {
            if (string.IsNullOrEmpty(SelectedExcelFile) || worksheet == null)
            {
                ExcelDisplayText = "";
                return;
            }

            FindNextAvailableRowForDisplay();

            try
            {
                string cellB = (worksheet.Cells[_currentExcelRow, 2] as Excel.Range)?.Value?.ToString() ?? "";
                string cellC = (worksheet.Cells[_currentExcelRow, 3] as Excel.Range)?.Value?.ToString() ?? "";
                ExcelDisplayText = string.IsNullOrEmpty(cellB) && string.IsNullOrEmpty(cellC) ? "" : $"{cellB} - {cellC}".Trim();
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Error reading Excel data for row {_currentExcelRow}: {ex.Message}", "Error", MessageBoxButton.OK, MessageBoxImage.Error);
                ExcelDisplayText = "";
            }
        }

        public Dictionary<(int, int), Cell> GetAllCells()
        {
            return allCells;
        }

        public Dictionary<(int, int), Cell> GetMappingCells()
        {
            return mappingCells;
        }

        public SpecialPoint GetMotorPoint(int index)
        {
            if (index < motorPoints.Count)
                return motorPoints[index];
            return null;
        }

        public SpecialPoint GetDoorPoint(int index)
        {
            if (index < doorPoints.Count)
                return doorPoints[index];
            return null;
        }

        private void FindNextAvailableRowForDisplay()
        {
            try
            {
                if (string.IsNullOrEmpty(SelectedExcelFile) || worksheet == null)
                    return;

                while (_currentExcelRow <= 1000)
                {
                    var cellA = (worksheet.Cells[_currentExcelRow, 1] as Excel.Range)?.Value;
                    if (cellA == null || string.IsNullOrEmpty(cellA.ToString()))
                    {
                        var cellB = (worksheet.Cells[_currentExcelRow, 2] as Excel.Range)?.Value?.ToString() ?? "";
                        var cellC = (worksheet.Cells[_currentExcelRow, 3] as Excel.Range)?.Value?.ToString() ?? "";

                        if (!string.IsNullOrEmpty(cellB) || !string.IsNullOrEmpty(cellC))
                        {
                            break;
                        }
                    }
                    _currentExcelRow++;
                }
            }
            catch (Exception)
            {
            }
        }

        private bool InitializeExcel(string filePath)
        {
            try
            {
                CleanupExcel();
                excelApp = new Excel.Application { Visible = true };
                if (File.Exists(filePath))
                {
                    workbook = excelApp.Workbooks.Open(filePath);
                }
                else
                {
                    workbook = excelApp.Workbooks.Add();
                }
                worksheet = (Excel.Worksheet)workbook.Worksheets[1];
                return true;
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Failed to initialize Excel: {ex.Message}", "Error", MessageBoxButton.OK, MessageBoxImage.Error);
                return false;
            }
        }

        protected override void OnClosed(EventArgs e)
        {
            base.OnClosed(e);
            CleanupExcel();
        }

        private void CleanupExcel()
        {
            try
            {
                if (worksheet != null)
                {
                    Marshal.ReleaseComObject(worksheet);
                }
                if (workbook != null)
                {
                    Marshal.ReleaseComObject(workbook);
                }
                if (excelApp != null)
                {
                    Marshal.ReleaseComObject(excelApp);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Error cleaning up Excel: {ex.Message}", "Error", MessageBoxButton.OK, MessageBoxImage.Error);
            }
            finally
            {
                worksheet = null;
                workbook = null;
                excelApp = null;
            }
        }

        public event PropertyChangedEventHandler PropertyChanged;
        private void OnPropertyChanged(string propertyName) =>
            PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(propertyName));

        private void Rebuild_Click(object sender, RoutedEventArgs e)
        {
            if (!int.TryParse(SectionBox.Text, out int s) || s <= 0 || s > 20)
            {
                MessageBox.Show("Please enter a valid number of sections (1-20).", "Invalid Input", MessageBoxButton.OK, MessageBoxImage.Warning);
                return;
            }
            if (!int.TryParse(RowBox.Text, out int r) || r <= 0 || r > 20)
            {
                MessageBox.Show("Please enter a valid number of rows (1-20).", "Invalid Input", MessageBoxButton.OK, MessageBoxImage.Warning);
                return;
            }
            if (!int.TryParse(ColBox.Text, out int c) || c <= 0 || c > 10)
            {
                MessageBox.Show("Please enter a valid number of columns (1-10).", "Invalid Input", MessageBoxButton.OK, MessageBoxImage.Warning);
                return;
            }
            Sections = s;
            Rows = r;
            Cols = c;
            _lockedPointA = null;
            _lockedButton = null;
            _currentExcelRow = 2;
            BuildAllSections();
            // Reset removal mode when rebuilding the grid to avoid stale state
            _isRemovingCells = false;
            if (RemoveCellsButton != null)
            {
                RemoveCellsButton.Content = "Fjern celler";
            }
            // Reset sequential mapping mode when rebuilding
            EndSequentialMappingMode();
            UpdateExcelDisplayText();
        }

        private void SelectExcelFile_Click(object sender, RoutedEventArgs e)
        {
            var openFileDialog = new OpenFileDialog
            {
                Filter = "Excel Files (*.xlsx;*.xls)|*.xlsx;*.xls|All Files (*.*)|*.*",
                Title = "Select Excel File for Measurements"
            };
            if (openFileDialog.ShowDialog() == true)
            {
                SelectedExcelFile = openFileDialog.FileName;
                if (InitializeExcel(SelectedExcelFile))
                {
                    // Fjernet success melding som ønsket
                    _currentExcelRow = 2;
                    _componentMappingManager = new ComponentMappingManager(this, SelectedExcelFile);
                    UpdateExcelDisplayText();
                }
                else
                {
                    SelectedExcelFile = null;
                    ExcelDisplayText = "";
                    _componentMappingManager = null;
                }
            }
        }

        private void DeleteLastMeasurement_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                if (string.IsNullOrEmpty(SelectedExcelFile) || worksheet == null)
                {
                    MessageBox.Show("Please select an Excel file first.", "Error", MessageBoxButton.OK, MessageBoxImage.Warning);
                    return;
                }
                int lastRow = _currentExcelRow - 1;
                if (lastRow >= 2 && (worksheet.Cells[lastRow, 1] as Excel.Range)?.Value != null)
                {
                    ((Excel.Range)worksheet.Cells[lastRow, 1]).Clear();
                    _currentExcelRow = lastRow;
                    object cellValue = (worksheet.Cells[lastRow, 1] as Excel.Range)?.Value;
                    double? lastValue = cellValue != null && double.TryParse(cellValue.ToString(), out double parsedValue) ? parsedValue : null;
                    ResultText.Text = lastValue.HasValue ? $"Shortest path: {lastValue:F2} mm" : "";
                    UpdateExcelDisplayText();
                }
                else
                {
                    MessageBox.Show("No measurements to delete.", "Info", MessageBoxButton.OK, MessageBoxImage.Information);
                    _currentExcelRow = 2;
                    ResultText.Text = _lockedPointA != null ? $"Locked point A: {_lockedPointA.Type} {_lockedPointA.SectionIndex + 1}" : "";
                    UpdateExcelDisplayText();
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Error deleting last measurement: {ex.Message}", "Error", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        // === Eksport og reset-knapper ===
        private void ExportDataToNord_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                if (string.IsNullOrEmpty(SelectedExcelFile) || worksheet == null)
                {
                    MessageBox.Show("Velg først en Excel-fil (kildefil A).", "Mangler kildefil",
                                    MessageBoxButton.OK, MessageBoxImage.Warning);
                    return;
                }
                var dlg = new Microsoft.Win32.OpenFileDialog
                {
                    Filter = "Excel Files (*.xlsx;*.xls)|*.xlsx;*.xls|All Files (*.*)|*.*",
                    Title = "Velg eksportfil for Nord (B)"
                };
                if (dlg.ShowDialog() != true) return;
                ExportToNord(worksheet, dlg.FileName);
                MessageBox.Show("Eksport til Nord fullført.", "Ferdig",
                                MessageBoxButton.OK, MessageBoxImage.Information);
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Feil under eksport til Nord: {ex.Message}", "Feil",
                                MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        private void ExportDataToDurapart_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                if (string.IsNullOrEmpty(SelectedExcelFile) || worksheet == null)
                {
                    MessageBox.Show("Velg først en Excel-fil (kildefil A).", "Mangler kildefil",
                                    MessageBoxButton.OK, MessageBoxImage.Warning);
                    return;
                }
                var dlg = new Microsoft.Win32.OpenFileDialog
                {
                    Filter = "Excel Files (*.xlsx;*.xls)|*.xlsx;*.xls|All Files (*.*)|*.*",
                    Title = "Velg eksportfil for Durapart (B)"
                };
                if (dlg.ShowDialog() != true) return;
                ExportToDurapart(worksheet, dlg.FileName);
                MessageBox.Show("Eksport til Durapart fullført.", "Ferdig",
                                MessageBoxButton.OK, MessageBoxImage.Information);
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Feil under eksport til Durapart: {ex.Message}", "Feil",
                                MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        private static string SafeCellToString(Excel.Worksheet ws, int row, int col)
        {
            try
            {
                var v = (ws.Cells[row, col] as Excel.Range)?.Value2;
                return v?.ToString() ?? "";
            }
            catch
            {
                return "";
            }
        }

        /// <summary>
        /// Helper to write a value to a worksheet cell and optionally apply a background colour.
        /// Some Excel.Range objects may be null (e.g. when worksheet is protected); exceptions
        /// are ignored for colour assignment.  The value parameter can be null or empty.
        /// </summary>
        private static void WriteCellWithColour(Excel.Worksheet sheet, int row, int col, string value, uint colour)
        {
            try
            {
                var rng = (Excel.Range)sheet.Cells[row, col];
                rng.Value2 = value;
                // Apply colour only if non-zero (zero indicates no colour or error)
                if (colour != 0)
                {
                    rng.Interior.Color = colour;
                }
            }
            catch
            {
                // Ignore errors applying colour (e.g. locked cells)
                try
                {
                    var rng = (Excel.Range)sheet.Cells[row, col];
                    rng.Value2 = value;
                }
                catch
                {
                    // swallow any error
                }
            }
        }

        private void ExportToNord(Excel.Worksheet sourceSheet, string destPath)
        {
            Excel.Application app = null;
            Excel.Workbook wb = null;
            Excel.Worksheet ws = null;
            try
            {
                // Reuse existing Excel application if available to keep files open after export
                app = excelApp ?? new Excel.Application();
                app.Visible = true;
                wb = app.Workbooks.Open(destPath);

                // Finn riktig ark å skrive til
                ws = FindWorksheetByName(wb, "mal ferdig ledning");
                if (ws == null)
                    throw new InvalidOperationException("Finner ikke arket 'mal ferdig ledning' i eksportfilen.");

                // Finn siste rad i kildefilens data ved å bruke UsedRange
                var used = sourceSheet.UsedRange;
                int lastRow = used != null ? used.Rows.Count : 2;
                if (lastRow < 2) lastRow = 2;

                // Iterer gjennom alle rader fra rad 2 og kopier de som har data i minst én av de 7 første kolonnene
                for (int srcRow = 2; srcRow <= lastRow; srcRow++)
                {
                    bool hasData = false;
                    // Kontroller om raden har tekst eller tall i de første 7 kolonnene
                    for (int c = 1; c <= 7; c++)
                    {
                        var val = SafeCellToString(sourceSheet, srcRow, c);
                        if (!string.IsNullOrWhiteSpace(val))
                        {
                            hasData = true;
                            break;
                        }
                    }
                    if (!hasData) continue;

                    // Les verdier og farger fra kildefilen for de 7 kolonnene
                    string[] src = new string[7];
                    uint[] colours = new uint[7];
                    for (int i = 0; i < 7; i++)
                    {
                        src[i] = SafeCellToString(sourceSheet, srcRow, i + 1);
                        try
                        {
                            var range = (Excel.Range)sourceSheet.Cells[srcRow, i + 1];
                            colours[i] = (uint)range.Interior.Color;
                        }
                        catch
                        {
                            colours[i] = 0;
                        }
                    }

                    // Destinasjonsrad: flytt alle kilderader én rad ned (rad 2 -> dest 3, rad 3 -> dest 4, osv.)
                    int destRow = srcRow + 1;

                    // Skriv verdier og farger til destinasjonsarket
                    // Mapping: kolonne 1->C, 2->F, 3->G, 4->B, 5->D, 6->E, 7->J
                    WriteCellWithColour(ws, destRow, 3, src[0], colours[0]); // C
                    WriteCellWithColour(ws, destRow, 6, src[1], colours[1]); // F
                    WriteCellWithColour(ws, destRow, 7, src[2], colours[2]); // G
                    WriteCellWithColour(ws, destRow, 2, src[3], colours[3]); // B
                    WriteCellWithColour(ws, destRow, 4, src[4], colours[4]); // D
                    WriteCellWithColour(ws, destRow, 5, src[5], colours[5]); // E
                    WriteCellWithColour(ws, destRow, 10, src[6], colours[6]); // J
                }

                // Lagre endringer i destinasjonsfilen
                wb.Save();
            }
            finally
            {
                // Slipp referansen til arket; la workbook og app forbli åpne slik at filen ikke lukkes
                if (ws != null) Marshal.ReleaseComObject(ws);
                ws = null;
                wb = null;
                app = null;
            }
        }

        private void ExportToDurapart(Excel.Worksheet sourceSheet, string destPath)
        {
            Excel.Application app = null;
            Excel.Workbook wb = null;
            Excel.Worksheet ws = null;
            try
            {
                // Reuse existing Excel application if available to keep files open after export
                app = excelApp ?? new Excel.Application();
                app.Visible = true;
                wb = app.Workbooks.Open(destPath);

                // Velg ark: prøv "Durapart", ellers første ark
                ws = FindWorksheetByName(wb, "Durapart");
                if (ws == null)
                    ws = (Excel.Worksheet)wb.Worksheets[1];

                // Finn siste rad i kildefilen
                var used = sourceSheet.UsedRange;
                int lastRow = used != null ? used.Rows.Count : 2;
                if (lastRow < 2) lastRow = 2;

                // Iterer gjennom kilderader og kopier data og farge til destinasjonen
                for (int srcRow = 2; srcRow <= lastRow; srcRow++)
                {
                    bool hasData = false;
                    for (int c = 1; c <= 7; c++)
                    {
                        var val = SafeCellToString(sourceSheet, srcRow, c);
                        if (!string.IsNullOrWhiteSpace(val))
                        {
                            hasData = true;
                            break;
                        }
                    }
                    if (!hasData) continue;

                    string[] src = new string[7];
                    uint[] colours = new uint[7];
                    for (int i = 0; i < 7; i++)
                    {
                        src[i] = SafeCellToString(sourceSheet, srcRow, i + 1);
                        try
                        {
                            var range = (Excel.Range)sourceSheet.Cells[srcRow, i + 1];
                            colours[i] = (uint)range.Interior.Color;
                        }
                        catch
                        {
                            colours[i] = 0;
                        }
                    }

                    int destRow = srcRow + 1;

                    // Mapping for Durapart: kolonne 1->G, 2->D, 3->I, 4->F, 5->E, 6->J
                    WriteCellWithColour(ws, destRow, 7, src[0], colours[0]); // G
                    WriteCellWithColour(ws, destRow, 4, src[1], colours[1]); // D
                    WriteCellWithColour(ws, destRow, 9, src[2], colours[2]); // I
                    WriteCellWithColour(ws, destRow, 6, src[3], colours[3]); // F
                    WriteCellWithColour(ws, destRow, 5, src[4], colours[4]); // E
                    WriteCellWithColour(ws, destRow, 10, src[5], colours[5]); // J
                    // src[6] (kolonne 7 i A) flyttes ikke
                }

                wb.Save();
            }
            finally
            {
                if (ws != null) Marshal.ReleaseComObject(ws);
                ws = null;
                wb = null;
                app = null;
            }
        }

        private Excel.Worksheet FindWorksheetByName(Excel.Workbook wb, string name)
        {
            foreach (Excel.Worksheet sheet in wb.Worksheets)
            {
                if (string.Equals(sheet.Name, name, StringComparison.OrdinalIgnoreCase))
                    return sheet;
            }
            return null;
        }

        private void ResetRemovedCells_Click(object sender, RoutedEventArgs e)
        {
            int restored = 0;
            _isRemovingCells = false;
            if (RemoveCellsButton != null)
                RemoveCellsButton.Content = "Fjern celler";
            for (int s = 0; s < MainPanel.Children.Count; s++)
            {
                if (MainPanel.Children[s] is StackPanel sectionPanel)
                {
                    var grids = sectionPanel.Children.OfType<Grid>().ToList();
                    foreach (var grid in grids)
                    {
                        foreach (var child in grid.Children)
                        {
                            if (child is Button btn && btn.Height >= 40)
                            {
                                if (btn.Visibility == Visibility.Hidden || btn.IsEnabled == false)
                                {
                                    btn.Visibility = Visibility.Visible;
                                    btn.IsEnabled = true;
                                    btn.Background = new SolidColorBrush(Color.FromRgb(74, 90, 91));
                                    int localRow = Grid.GetRow(btn);
                                    int localCol = Grid.GetColumn(btn);
                                    int globalRow = localRow;
                                    int globalCol = s * Cols + localCol;
                                    if (!allCells.ContainsKey((globalRow, globalCol)))
                                    {
                                        allCells[(globalRow, globalCol)] = new Cell(globalRow, globalCol, btn);
                                    }
                                    restored++;
                                }
                            }
                        }
                    }
                }
            }
            ResetSelection();
            ResultText.Text = $"Tilbakestilt {restored} celler.";
        }
    }

    // Resten av koden forblir uendret (PathFinder, Cell, SpecialPoint, PriorityQueue, ExcelConnectionProcessor)
    // Dette er i en egen fil som følger...
    // Dette er fortsettelsen av MainWindow.xaml.cs - alle hjelpeklasser

    public static class PathFinder
    {
        public static List<Cell> FindShortestPath(Cell start, Cell end, Dictionary<(int, int), Cell> allCells, Func<int, int, bool> hasHorizontalNeighbor)
        {
            var dist = new Dictionary<Cell, double> { { start, 0 } };
            var prev = new Dictionary<Cell, Cell>();
            var queue = new PriorityQueue<Cell, double>();
            queue.Enqueue(start, 0);

            foreach (var cell in allCells.Values)
            {
                if (cell != start)
                    dist[cell] = double.PositiveInfinity;
                prev[cell] = null;
            }

            while (queue.Count > 0)
            {
                var u = queue.Dequeue();
                if (u == end) break;

                foreach (var neighbor in GetNeighbors(u, allCells))
                {
                    double weight = hasHorizontalNeighbor(neighbor.Row, neighbor.Col) ? 100 : 50;
                    double alt = dist[u] + weight;
                    if (alt < dist[neighbor])
                    {
                        dist[neighbor] = alt;
                        prev[neighbor] = u;
                        queue.Enqueue(neighbor, alt);
                    }
                }
            }

            if (double.IsInfinity(dist[end])) return null;

            var path = new List<Cell>();
            for (var curr = end; curr != null; curr = prev[curr])
                path.Add(curr);
            path.Reverse();
            return path;
        }

        private static List<Cell> GetNeighbors(Cell cell, Dictionary<(int, int), Cell> allCells)
        {
            var directions = new[] { (-1, 0), (1, 0), (0, -1), (0, 1) };
            var neighbors = new List<Cell>();

            foreach (var (dr, dc) in directions)
            {
                int nr = cell.Row + dr;
                int nc = cell.Col + dc;
                if (allCells.TryGetValue((nr, nc), out var neighbor))
                    neighbors.Add(neighbor);
            }
            return neighbors;
        }

        public static double CalculateDistance(List<Cell> path, bool endsInSpecial, Func<int, int, bool> hasHorizontalNeighbor)
        {
            if (path == null || path.Count == 0) return 0;

            double distance = 100; // Start connection

            for (int i = 1; i < path.Count - 1; i++)
                distance += hasHorizontalNeighbor(path[i].Row, path[i].Col) ? 100 : 50;

            if (!endsInSpecial)
                distance += 200; // End connection

            return distance;
        }
    }

    public class Cell
    {
        public int Row { get; }
        public int Col { get; }
        public Button ButtonRef { get; }

        public Cell(int row, int col, Button button)
        {
            Row = row;
            Col = col;
            ButtonRef = button;
        }
    }

    public enum SpecialPointType { Door, Motor }

    public class SpecialPoint
    {
        public int SectionIndex { get; set; }
        public SpecialPointType Type { get; set; }
        public Button Button { get; set; }
        public int GlobalRow { get; set; }
        public int GlobalCol { get; set; }
    }

    public class PriorityQueue<TItem, TPriority> where TPriority : IComparable<TPriority>
    {
        private readonly List<(TItem Item, TPriority Priority)> elements = new();
        public int Count => elements.Count;

        public void Enqueue(TItem item, TPriority priority)
        {
            elements.Add((item, priority));
        }

        public TItem Dequeue()
        {
            int bestIndex = 0;
            for (int i = 1; i < elements.Count; i++)
            {
                if (elements[i].Priority.CompareTo(elements[bestIndex].Priority) < 0)
                    bestIndex = i;
            }
            var result = elements[bestIndex].Item;
            elements.RemoveAt(bestIndex);
            return result;
        }
    }

    // ExcelConnectionProcessor klassen med støtte for bulk mapping
    public class ExcelConnectionProcessor
    {
        private readonly MainWindow _mainWindow;
        private readonly ComponentMappingManager _mappingManager;

        public ExcelConnectionProcessor(MainWindow mainWindow, ComponentMappingManager mappingManager)
        {
            _mainWindow = mainWindow;
            _mappingManager = mappingManager;
        }

        public int ProcessAllConnections()
        {
            if (_mainWindow.worksheet == null)
                throw new InvalidOperationException("Ingen Excel‑fil er åpen");

            int processedCount = 0;
            var allCells = _mainWindow.GetAllCells();

            try
            {
                var usedRange = _mainWindow.worksheet.UsedRange;
                if (usedRange == null) return 0;

                var lastRow = usedRange.Rows.Count;

                for (int row = 2; row <= lastRow; row++)
                {
                    try
                    {
                        // Les kolonne B og C (referanser) og A (måleverdi)
                        string cellB = "";
                        string cellC = "";
                        object cellA = null;

                        try { cellB = _mainWindow.worksheet.Cells[row, 2].Value?.ToString() ?? ""; } catch { }
                        try { cellC = _mainWindow.worksheet.Cells[row, 3].Value?.ToString() ?? ""; } catch { }
                        try { cellA = _mainWindow.worksheet.Cells[row, 1].Value; } catch { }

                        // Skipp dersom cellen allerede har en beregnet verdi
                        if (cellA != null && !string.IsNullOrEmpty(cellA.ToString()) &&
                            double.TryParse(cellA.ToString(), out _))
                            continue;

                        // Skipp hvis ingen referanser
                        if (string.IsNullOrWhiteSpace(cellB) && string.IsNullOrWhiteSpace(cellC))
                            continue;

                        // Hent mappinger
                        var mappingA = _mappingManager.GetMapping(cellB);
                        var mappingB = _mappingManager.GetMapping(cellC);
                        var bulkA = _mappingManager.GetBulkRangeMappingForReference(cellB);
                        var bulkB = _mappingManager.GetBulkRangeMappingForReference(cellC);

                        // Hvis begge referanser tilhører bulk‑range
                        if (bulkA != null && bulkB != null)
                        {
                            // Samle kandidat-koordinater for hver referanse
                            var candidateA = GetCandidateCellsForBulkReference(bulkA, cellB, allCells);
                            var candidateB = GetCandidateCellsForBulkReference(bulkB, cellC, allCells);

                            double connA = GetConnectionDistance(cellB, mappingA);
                            double connB = GetConnectionDistance(cellC, mappingB);
                            double maxDistance = 0;

                            // Finn LENGSTE avstand mellom alle kombinasjoner
                            foreach (var posA in candidateA)
                            {
                                foreach (var posB in candidateB)
                                {
                                    // Samme posisjon = 300 mm
                                    if (posA.Row == posB.Row && posA.Col == posB.Col)
                                    {
                                        maxDistance = Math.Max(maxDistance, 300.0);
                                        continue;
                                    }

                                    if (!allCells.TryGetValue((posA.Row, posA.Col), out var startCell) ||
                                        !allCells.TryGetValue((posB.Row, posB.Col), out var endCell))
                                        continue;

                                    var path = PathFinder.FindShortestPath(startCell, endCell, allCells, _mainWindow.HasHorizontalNeighbor);
                                    if (path == null || path.Count == 0)
                                        continue;

                                    double pathDist = 0;
                                    for (int i = 1; i < path.Count - 1; i++)
                                        pathDist += _mainWindow.HasHorizontalNeighbor(path[i].Row, path[i].Col) ? 100 : 50;

                                    double total = pathDist + connA + connB + 50;
                                    maxDistance = Math.Max(maxDistance, total);
                                }
                            }

                            if (maxDistance > 0)
                            {
                                _mainWindow.worksheet.Cells[row, 1].Value = maxDistance;
                                processedCount++;
                                continue;
                            }
                        }
                        // Hvis en er bulk og den andre er vanlig mapping
                        else if ((bulkA != null && mappingB != null) || (mappingA != null && bulkB != null))
                        {
                            double maxDistance = 0;

                            if (bulkA != null && mappingB != null)
                            {
                                // A er bulk, B er vanlig
                                var candidateA = GetCandidateCellsForBulkReference(bulkA, cellB, allCells);
                                var posB = GetGridPositionFromMapping(mappingB);

                                if (posB.HasValue)
                                {
                                    double connA = GetConnectionDistance(cellB, null);
                                    double connB = GetConnectionDistance(cellC, mappingB);

                                    foreach (var posA in candidateA)
                                    {
                                        if (posA.Row == posB.Value.Row && posA.Col == posB.Value.Col)
                                        {
                                            maxDistance = Math.Max(maxDistance, 300.0);
                                            continue;
                                        }

                                        if (!allCells.TryGetValue((posA.Row, posA.Col), out var startCell) ||
                                            !allCells.TryGetValue(posB.Value, out var endCell))
                                            continue;

                                        var path = PathFinder.FindShortestPath(startCell, endCell, allCells, _mainWindow.HasHorizontalNeighbor);
                                        if (path != null && path.Count > 0)
                                        {
                                            double pathDist = 0;
                                            for (int i = 1; i < path.Count - 1; i++)
                                                pathDist += _mainWindow.HasHorizontalNeighbor(path[i].Row, path[i].Col) ? 100 : 50;

                                            double total = pathDist + connA + connB + 50;
                                            maxDistance = Math.Max(maxDistance, total);
                                        }
                                    }
                                }
                            }
                            else if (mappingA != null && bulkB != null)
                            {
                                // A er vanlig, B er bulk
                                var posA = GetGridPositionFromMapping(mappingA);
                                var candidateB = GetCandidateCellsForBulkReference(bulkB, cellC, allCells);

                                if (posA.HasValue)
                                {
                                    double connA = GetConnectionDistance(cellB, mappingA);
                                    double connB = GetConnectionDistance(cellC, null);

                                    foreach (var posB in candidateB)
                                    {
                                        if (posA.Value.Row == posB.Row && posA.Value.Col == posB.Col)
                                        {
                                            maxDistance = Math.Max(maxDistance, 300.0);
                                            continue;
                                        }

                                        if (!allCells.TryGetValue(posA.Value, out var startCell) ||
                                            !allCells.TryGetValue((posB.Row, posB.Col), out var endCell))
                                            continue;

                                        var path = PathFinder.FindShortestPath(startCell, endCell, allCells, _mainWindow.HasHorizontalNeighbor);
                                        if (path != null && path.Count > 0)
                                        {
                                            double pathDist = 0;
                                            for (int i = 1; i < path.Count - 1; i++)
                                                pathDist += _mainWindow.HasHorizontalNeighbor(path[i].Row, path[i].Col) ? 100 : 50;

                                            double total = pathDist + connA + connB + 50;
                                            maxDistance = Math.Max(maxDistance, total);
                                        }
                                    }
                                }
                            }

                            if (maxDistance > 0)
                            {
                                _mainWindow.worksheet.Cells[row, 1].Value = maxDistance;
                                processedCount++;
                                continue;
                            }
                        }

                        // Hvis kun vanlig mapping (ikke bulk)
                        if (mappingA != null && mappingB != null)
                        {
                            var posA = GetGridPositionFromMapping(mappingA);
                            var posB = GetGridPositionFromMapping(mappingB);

                            if (posA.HasValue && posB.HasValue)
                            {
                                if (posA.Value.Equals(posB.Value))
                                {
                                    _mainWindow.worksheet.Cells[row, 1].Value = 300.0;
                                    processedCount++;
                                    continue;
                                }

                                if (allCells.TryGetValue(posA.Value, out var startCell) &&
                                    allCells.TryGetValue(posB.Value, out var endCell))
                                {
                                    var path = PathFinder.FindShortestPath(startCell, endCell, allCells, _mainWindow.HasHorizontalNeighbor);
                                    if (path != null && path.Count > 0)
                                    {
                                        // Kun grid-avstand (uten start- og slutt-tilkobling)
                                        double pathDistance = 0;
                                        for (int i = 1; i < path.Count - 1; i++)
                                            pathDistance += _mainWindow.HasHorizontalNeighbor(path[i].Row, path[i].Col) ? 100 : 50;

                                        bool aIsNormal = mappingA.GridRow != -1 && mappingA.GridRow != -2;
                                        bool bIsNormal = mappingB.GridRow != -1 && mappingB.GridRow != -2;
                                        double totalDistance;

                                        if (aIsNormal && bIsNormal)
                                        {
                                            totalDistance = pathDistance + 100 + 200;
                                        }
                                        else
                                        {
                                            double startDistance = mappingA.GridRow switch
                                            {
                                                -1 => 500,
                                                -2 => 1000,
                                                _ => 200
                                            };
                                            double endDistance = mappingB.GridRow switch
                                            {
                                                -1 => 500,
                                                -2 => 1000,
                                                _ => 200
                                            };
                                            totalDistance = pathDistance + startDistance + endDistance + 50;
                                        }

                                        _mainWindow.worksheet.Cells[row, 1].Value = totalDistance;
                                        processedCount++;
                                    }
                                }
                            }
                        }
                    }
                    catch (Exception ex)
                    {
                        System.Diagnostics.Debug.WriteLine($"Error processing row {row}: {ex.Message}");
                    }
                }
            }
            catch (Exception ex)
            {
                throw new InvalidOperationException($"Excel processing failed: {ex.Message}");
            }

            return processedCount;
        }

        private (int Row, int Col)? GetGridPositionFromMapping(ComponentMapping mapping)
        {
            // Handle special points (Motors and Doors)
            if (mapping.GridRow == -1 && mapping.GridColumn >= 1000) // Motor
            {
                var motorIndex = mapping.GridColumn - 1000;
                var motorPoint = _mainWindow.GetMotorPoint(motorIndex);
                return motorPoint != null ? (motorPoint.GlobalRow, motorPoint.GlobalCol) : null;
            }
            else if (mapping.GridRow == -2 && mapping.GridColumn >= 1000) // Door
            {
                var doorIndex = mapping.GridColumn - 1000;
                var doorPoint = _mainWindow.GetDoorPoint(doorIndex);
                return doorPoint != null ? (doorPoint.GlobalRow, doorPoint.GlobalCol) : null;
            }
            else
            {
                // Handle bottom side mappings (negative rows)
                if (mapping.GridRow < 0)
                {
                    // Convert negative row back to positive for allCells lookup
                    int actualRow = -(mapping.GridRow + 1);
                    return (actualRow, mapping.GridColumn);
                }
                else
                {
                    // Handle top side (T) mappings
                    int adjustedRow = mapping.GridRow;
                    var allCells = _mainWindow.GetAllCells();

                    // Try the exact position first
                    if (allCells.ContainsKey((mapping.GridRow, mapping.GridColumn)))
                    {
                        return (mapping.GridRow, mapping.GridColumn);
                    }

                    // If DefaultToBottom, try the row below
                    if (mapping.DefaultToBottom && allCells.ContainsKey((mapping.GridRow + 1, mapping.GridColumn)))
                    {
                        return (mapping.GridRow + 1, mapping.GridColumn);
                    }

                    // Try the row above
                    if (allCells.ContainsKey((mapping.GridRow - 1, mapping.GridColumn)))
                    {
                        return (mapping.GridRow - 1, mapping.GridColumn);
                    }

                    // Return original if nothing else works
                    return (mapping.GridRow, mapping.GridColumn);
                }
            }
        }

        private double GetConnectionDistance(string text, ComponentMapping mapping = null)
        {
            if (string.IsNullOrEmpty(text)) return 0;

            bool hasAsterisk = text.Contains("*");

            // Special points get their fixed distances
            if (mapping != null)
            {
                if (mapping.GridRow == -1) // Motor
                    return 500;
                else if (mapping.GridRow == -2) // Door
                    return 1000;
            }

            // Component-based distances
            if (text.Contains("F", StringComparison.OrdinalIgnoreCase))
                return hasAsterisk ? 50 : 30;
            else if (text.Contains("X", StringComparison.OrdinalIgnoreCase))
                return hasAsterisk ? 40 : 20;
            else if (text.Contains("K", StringComparison.OrdinalIgnoreCase))
                return 60;
            else if (text.Contains("A", StringComparison.OrdinalIgnoreCase))
                return 45;
            else if (text.Contains("Motor", StringComparison.OrdinalIgnoreCase))
                return 500;
            else if (text.Contains("Door", StringComparison.OrdinalIgnoreCase))
                return 1000;
            else
                return 25;
        }

        /// <summary>
        /// Returns a list of candidate grid cell coordinates for a reference within a bulk
        /// range.  The returned coordinates depend on whether the reference ends
        /// with an asterisk (*), which indicates that the reference should be
        /// mapped to the opposite side of the selected bulk mapping cells.
        /// </summary>
        private List<(int Row, int Col)> GetCandidateCellsForBulkReference(ComponentMappingManager.BulkRangeMapping bulkRange, string excelReference, Dictionary<(int, int), Cell> allCells)
        {
            var result = new List<(int Row, int Col)>();
            if (bulkRange == null) return result;

            bool hasStar = !string.IsNullOrWhiteSpace(excelReference) && excelReference.Trim().EndsWith("*");
            bool selectedIsTop = bulkRange.SelectedIsTop;

            foreach (var cell in bulkRange.Cells)
            {
                int row = cell.Row;
                int col = cell.Col;

                if (hasStar)
                {
                    // Star references map to the opposite side
                    int oppositeRow = selectedIsTop ? row + 1 : row - 1;
                    if (allCells.ContainsKey((oppositeRow, col)))
                    {
                        result.Add((oppositeRow, col));
                    }
                }
                else
                {
                    // Use the selected side as-is if the cell exists
                    if (allCells.ContainsKey((row, col)))
                    {
                        result.Add((row, col));
                    }
                }
            }

            // If no candidates were added, fall back to the original selected cells
            if (result.Count == 0)
            {
                foreach (var cell in bulkRange.Cells)
                {
                    if (allCells.ContainsKey((cell.Row, cell.Col)))
                    {
                        result.Add((cell.Row, cell.Col));
                    }
                }
            }

            // Remove any duplicates
            return result.Distinct().ToList();
        }

        public string TestProcessing()
        {
            if (_mainWindow.worksheet == null)
                return "Ingen Excel-fil er åpen";

            var results = new List<string>();
            var allMappings = _mappingManager.GetAllMappings();
            var bulkMappings = _mappingManager.GetAllBulkRanges();
            var usedRange = _mainWindow.worksheet.UsedRange;

            if (usedRange == null)
                return "Ingen data i Excel-arket";

            var lastRow = Math.Min(usedRange.Rows.Count, 10);

            results.Add("=== TEST PROSESSERING ===\n");
            results.Add($"Antall vanlige mappings: {allMappings.Count}");
            results.Add($"Antall bulk mappings: {bulkMappings.Count}");

            foreach (var mapping in allMappings)
            {
                if (mapping.GridRow == -1)
                    results.Add($"  {mapping.ExcelReference} -> Motor {mapping.GridColumn - 1000}");
                else if (mapping.GridRow == -2)
                    results.Add($"  {mapping.ExcelReference} -> Door {mapping.GridColumn - 1000}");
                else
                    results.Add($"  {mapping.ExcelReference} -> ({mapping.GridRow}, {mapping.GridColumn}) {(mapping.DefaultToBottom ? "(B)" : "(T)")}");
            }

            foreach (var bulk in bulkMappings)
            {
                results.Add($"  BULK: {bulk.Prefix}:{bulk.StartIndex}-{bulk.EndIndex} -> {bulk.Cells.Count} celler ({(bulk.SelectedIsTop ? "T" : "B")})");
            }

            results.Add("");

            for (int row = 2; row <= lastRow; row++)
            {
                try
                {
                    string cellB = "";
                    string cellC = "";
                    object cellA = null;

                    try { cellB = (_mainWindow.worksheet.Cells[row, 2] as Excel.Range)?.Value?.ToString() ?? ""; } catch { }
                    try { cellC = (_mainWindow.worksheet.Cells[row, 3] as Excel.Range)?.Value?.ToString() ?? ""; } catch { }
                    try { cellA = (_mainWindow.worksheet.Cells[row, 1] as Excel.Range)?.Value; } catch { }

                    results.Add($"RAD {row}:");
                    results.Add($"  Kolonne B: '{cellB}'");
                    results.Add($"  Kolonne C: '{cellC}'");

                    if (cellA != null && !string.IsNullOrEmpty(cellA.ToString()))
                    {
                        results.Add($"  -> HAR ALLEREDE MÅLEVERDI: {cellA}\n");
                        continue;
                    }

                    if (string.IsNullOrWhiteSpace(cellB) && string.IsNullOrWhiteSpace(cellC))
                    {
                        results.Add("  -> TOM RAD\n");
                        continue;
                    }

                    var mappingA = _mappingManager.GetMapping(cellB);
                    var mappingB = _mappingManager.GetMapping(cellC);
                    var bulkA = _mappingManager.GetBulkRangeMappingForReference(cellB);
                    var bulkB = _mappingManager.GetBulkRangeMappingForReference(cellC);

                    results.Add($"  Match A: {mappingA?.ExcelReference ?? (bulkA != null ? "BULK MAPPING" : "IKKE FUNNET")}");
                    results.Add($"  Match B: {mappingB?.ExcelReference ?? (bulkB != null ? "BULK MAPPING" : "IKKE FUNNET")}");

                    if ((mappingA != null || bulkA != null) && (mappingB != null || bulkB != null))
                    {
                        results.Add("  -> KAN PROSESSERES");
                    }
                    else
                    {
                        results.Add("  -> IKKE BEGGE MAPPINGS FUNNET - HOPPER OVER");
                    }
                    results.Add("");
                }
                catch (Exception ex)
                {
                    results.Add($"  FEIL: {ex.Message}\n");
                }
            }

            return string.Join("\n", results);
        }
    }
}