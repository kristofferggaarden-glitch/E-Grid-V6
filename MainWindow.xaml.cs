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

        private bool _isInSequentialMappingMode = false;
        private Stack<string> _sequentialMappingUndoStack = new Stack<string>();
        private ComponentMappingWindow _componentMappingWindow;

        private bool _isRemovingCells = false;

        private bool _isBulkMappingSelectionMode = false;
        private string _bulkMappingPrefix;
        private int _bulkMappingStart;
        private int _bulkMappingEnd;
        private List<(int Row, int Col)> _bulkSelectedCells;
        private Action<string, int, int, List<(int Row, int Col)>> _bulkMappingCompletedCallback;
        private bool? _bulkMappingSelectedIsTop;

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

                var grid = new Grid
                {
                    Background = Brushes.Transparent,
                    Margin = new Thickness(0, 10, 0, 10)
                };
                sectionPanel.Children.Add(grid);

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

                grid.RowDefinitions.Clear();
                grid.ColumnDefinitions.Clear();
                for (int r = 0; r < Rows; r++)
                    grid.RowDefinitions.Add(new RowDefinition { Height = GridLength.Auto });
                for (int c = 0; c < Cols; c++)
                    grid.ColumnDefinitions.Add(new ColumnDefinition { Width = GridLength.Auto });

                for (int row = 0; row < Rows; row++)
                {
                    if (row % 2 == 0)
                    {
                        for (int col = 0; col < Cols; col++)
                            AddCell(grid, row, col, s);
                    }
                    else
                    {
                        AddCell(grid, row, 0, s);
                        for (int col = 1; col < Cols; col++)
                            AddMappingCell(grid, row, col, s);
                    }
                }

                doorPoints.Add(new SpecialPoint
                {
                    SectionIndex = s,
                    Type = SpecialPointType.Door,
                    Button = doorBtn,
                    GlobalRow = 0,
                    GlobalCol = s * Cols
                });

                motorPoints.Add(new SpecialPoint
                {
                    SectionIndex = s,
                    Type = SpecialPointType.Motor,
                    Button = motorBtn,
                    GlobalRow = Rows - 1,
                    GlobalCol = s * Cols + (Cols - 1)
                });

                MainPanel.Children.Add(sectionPanel);
            }

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

            int globalRow = localRow;
            int globalCol = sectionIndex * Cols + localCol;

            mappingCells[(globalRow, globalCol)] = new Cell(globalRow, globalCol, btnTop);
            mappingCells[(-globalRow - 1, globalCol)] = new Cell(-globalRow - 1, globalCol, btnBottom);
        }

        public void StartSequentialMappingMode()
        {
            _isInSequentialMappingMode = true;
            _sequentialMappingUndoStack.Clear();
            SequentialMappingControls.Visibility = Visibility.Visible;
            UpdateUndoButtonState();
        }

        public void EndSequentialMappingMode()
        {
            _isInSequentialMappingMode = false;
            _sequentialMappingUndoStack.Clear();
            SequentialMappingControls.Visibility = Visibility.Collapsed;
        }

        private void UpdateUndoButtonState()
        {
            if (UndoMappingButton != null)
            {
                UndoMappingButton.IsEnabled = _sequentialMappingUndoStack.Count > 0;
            }
        }

        private void RecordMappingForUndo(string excelReference)
        {
            if (_isInSequentialMappingMode)
            {
                _sequentialMappingUndoStack.Push(excelReference);
                UpdateUndoButtonState();
            }
        }

        private void UndoMapping_Click(object sender, RoutedEventArgs e)
        {
            if (!_isInSequentialMappingMode || _sequentialMappingUndoStack.Count == 0)
                return;

            var lastMapping = _sequentialMappingUndoStack.Pop();

            if (_componentMappingManager != null)
            {
                _componentMappingManager.RemoveMapping(lastMapping);

                if (_componentMappingWindow != null)
                {
                    try
                    {
                        _componentMappingWindow.LoadExistingMappings();
                    }
                    catch { }
                }
            }

            if (_componentMappingWindow != null)
            {
                _componentMappingWindow.PutReferenceBackInQueue(lastMapping);
            }

            UpdateUndoButtonState();

            _currentMappingReference = lastMapping;
            MappingIndicatorText.Text = $"Mapper: {lastMapping}";
            ResultText.Text = $"Angret - mapper nå '{lastMapping}' på nytt. Klikk på grid-posisjon.";
        }

        private void FinishMapping_Click(object sender, RoutedEventArgs e)
        {
            EndSequentialMappingMode();
            EndMappingMode();

            if (_componentMappingWindow != null && !_componentMappingWindow.IsLoaded)
            {
                _componentMappingWindow = new ComponentMappingWindow(this, _componentMappingManager);
                _componentMappingWindow.Show();
            }
            else if (_componentMappingWindow != null)
            {
                _componentMappingWindow.WindowState = WindowState.Normal;
                _componentMappingWindow.Activate();
                _componentMappingWindow.OnSequentialMappingFinished();
            }
            else
            {
                OpenComponentMapping_Click(null, null);
            }
        }

        private void MappingCell_Click(object sender, RoutedEventArgs e)
        {
            var btn = sender as Button;
            if (btn == null) return;

            if (_isBulkMappingSelectionMode)
            {
                HandleBulkMappingCellClick(btn);
                return;
            }

            if (_isInMappingMode)
            {
                Cell cell = null;
                bool isBottomSide = false;

                foreach (var kvp in mappingCells)
                {
                    if (kvp.Value.ButtonRef == btn)
                    {
                        cell = kvp.Value;
                        isBottomSide = kvp.Key.Item1 < 0;
                        break;
                    }
                }

                if (cell != null && !string.IsNullOrEmpty(_currentMappingReference))
                {
                    int actualRow = isBottomSide ? -(cell.Row + 1) : cell.Row;

                    _componentMappingManager?.AddMapping(_currentMappingReference, actualRow, cell.Col, isBottomSide);

                    RecordMappingForUndo(_currentMappingReference);

                    _mappingCompletedCallback?.Invoke(_currentMappingReference, "");
                    EndMappingMode();
                }
            }
        }

        private void HandleBulkMappingCellClick(Button btn)
        {
            if (btn == null || !_isBulkMappingSelectionMode) return;

            var mappingEntry = mappingCells.FirstOrDefault(kvp => kvp.Value.ButtonRef == btn);
            if (mappingEntry.Equals(default(KeyValuePair<(int, int), Cell>)))
            {
                return;
            }

            bool clickedIsTop = mappingEntry.Key.Item1 >= 0;

            if (_bulkMappingSelectedIsTop == null)
            {
                _bulkMappingSelectedIsTop = clickedIsTop;
                BulkMappingSelectedIsTop = clickedIsTop;
            }
            else if (_bulkMappingSelectedIsTop != clickedIsTop)
            {
                MessageBox.Show("Du kan kun velge enten T eller B celler i samme bulk mapping.",
                               "Ugyldig valg", MessageBoxButton.OK, MessageBoxImage.Warning);
                return;
            }

            int actualRow;
            if (clickedIsTop)
            {
                actualRow = mappingEntry.Key.Item1;
            }
            else
            {
                actualRow = -(mappingEntry.Key.Item1 + 1);
            }

            int column = mappingEntry.Key.Item2;

            bool alreadySelected = _bulkSelectedCells != null && _bulkSelectedCells.Any(c => c.Row == actualRow && c.Col == column);

            if (alreadySelected)
            {
                _bulkSelectedCells.RemoveAll(c => c.Row == actualRow && c.Col == column);
                btn.Background = new SolidColorBrush(Color.FromRgb(80, 120, 140));
            }
            else
            {
                if (_bulkSelectedCells == null)
                    _bulkSelectedCells = new List<(int Row, int Col)>();

                _bulkSelectedCells.Add((actualRow, column));
                btn.Background = new SolidColorBrush(Color.FromRgb(0, 178, 148));
            }

            int count = _bulkSelectedCells?.Count ?? 0;
            ResultText.Text = $"Bulk mapping: {count} celler valgt ({(clickedIsTop ? "T" : "B")} side). Klikk 'Ferdig bulk mapping' når ferdig.";
        }

        public void StartBulkMappingSelection(string prefix, int startNumber, int endNumber,
                                              Action<string, int, int, List<(int Row, int Col)>> onCompleted)
        {
            EndMappingMode();
            _isRemovingCells = false;

            _isBulkMappingSelectionMode = true;
            _bulkMappingPrefix = prefix;
            _bulkMappingStart = startNumber;
            _bulkMappingEnd = endNumber;
            _bulkSelectedCells = new List<(int Row, int Col)>();
            _bulkMappingCompletedCallback = onCompleted;

            _bulkMappingSelectedIsTop = null;
            BulkMappingSelectedIsTop = false;

            MappingIndicatorText.Text = $"Bulk mapping: {prefix}:{startNumber}-{endNumber}. Velg T eller B celler.";
            MappingIndicator.Visibility = Visibility.Visible;
            CancelMappingButton.Visibility = Visibility.Visible;

            if (CompleteBulkMappingButton != null)
            {
                CompleteBulkMappingButton.Visibility = Visibility.Visible;
            }

            foreach (var kvp in mappingCells)
            {
                var mappingCell = kvp.Value;
                if (mappingCell.ButtonRef != null)
                {
                    mappingCell.ButtonRef.IsEnabled = true;
                    mappingCell.ButtonRef.Background = new SolidColorBrush(Color.FromRgb(80, 120, 140));
                }
            }

            foreach (var dp in doorPoints)
                dp.Button.IsEnabled = false;
            foreach (var mp in motorPoints)
                mp.Button.IsEnabled = false;

            ResultText.Text = $"Bulk mapping: Velg T eller B celler for {prefix}:{startNumber}-{endNumber}. Klikk på celler for å velge.";

            this.Activate();
            this.Focus();
        }

        public void FinishBulkMappingSelection()
        {
            if (!_isBulkMappingSelectionMode)
            {
                MessageBox.Show("Ikke i bulk mapping modus.", "Feil", MessageBoxButton.OK, MessageBoxImage.Warning);
                return;
            }

            int cellCount = _bulkSelectedCells?.Count ?? 0;

            if (_bulkSelectedCells == null || _bulkSelectedCells.Count == 0)
            {
                MessageBox.Show("Ingen celler ble valgt.\n\nKlikk på T eller B celler før du fullfører bulk mapping.",
                               "Ingen celler valgt", MessageBoxButton.OK, MessageBoxImage.Warning);
                return;
            }

            var prefix = _bulkMappingPrefix;
            var startNumber = _bulkMappingStart;
            var endNumber = _bulkMappingEnd;
            var selectedCells = new List<(int Row, int Col)>(_bulkSelectedCells);
            var callback = _bulkMappingCompletedCallback;

            EndBulkMappingMode();

            if (callback != null)
            {
                try
                {
                    callback.Invoke(prefix, startNumber, endNumber, selectedCells);
                }
                catch (Exception ex)
                {
                    MessageBox.Show($"Feil ved lagring av bulk mapping:\n\n{ex.Message}",
                                   "Feil", MessageBoxButton.OK, MessageBoxImage.Error);
                }
            }
        }

        private void CancelBulkMappingSelection()
        {
            if (!_isBulkMappingSelectionMode)
                return;

            EndBulkMappingMode();
            MessageBox.Show("Bulk mapping avbrutt.", "Avbrutt", MessageBoxButton.OK, MessageBoxImage.Information);
        }

        private void EndBulkMappingMode()
        {
            _isBulkMappingSelectionMode = false;
            _bulkMappingPrefix = null;
            _bulkMappingStart = 0;
            _bulkMappingEnd = 0;
            _bulkSelectedCells = null;
            _bulkMappingCompletedCallback = null;
            _bulkMappingSelectedIsTop = null;
            BulkMappingSelectedIsTop = false;

            foreach (var cell in allCells.Values)
            {
                if (cell.ButtonRef != null && cell.ButtonRef.IsEnabled)
                {
                    cell.ButtonRef.Background = new SolidColorBrush(Color.FromRgb(74, 90, 91));
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

            foreach (var dp in doorPoints)
                dp.Button.IsEnabled = true;
            foreach (var mp in motorPoints)
                mp.Button.IsEnabled = true;

            MappingIndicator.Visibility = Visibility.Collapsed;
            CancelMappingButton.Visibility = Visibility.Collapsed;
            if (CompleteBulkMappingButton != null)
            {
                CompleteBulkMappingButton.Visibility = Visibility.Collapsed;
            }
            ResultText.Text = string.Empty;
        }

        private void CancelMapping_Click(object sender, RoutedEventArgs e)
        {
            if (_isInMappingMode)
            {
                EndMappingMode();
            }
            else if (_isBulkMappingSelectionMode)
            {
                CancelBulkMappingSelection();
            }
        }

        private void CompleteBulkMappingButton_Click(object sender, RoutedEventArgs e)
        {
            FinishBulkMappingSelection();
        }

        private void RemoveCellsButton_Click(object sender, RoutedEventArgs e)
        {
            _isRemovingCells = !_isRemovingCells;
            if (_isRemovingCells)
            {
                if (RemoveCellsButton != null)
                {
                    RemoveCellsButton.Content = "Ferdig fjerning";
                }
                startPoint = null;
                endPoint = null;
                ResultText.Text = "Klikk på cellene du vil fjerne.";
            }
            else
            {
                if (RemoveCellsButton != null)
                {
                    RemoveCellsButton.Content = "Fjern celler";
                }
                ResultText.Text = _lockedPointA != null ? $"Locked point A: {_lockedPointA.Type} {_lockedPointA.SectionIndex + 1}" : "";
                ResetSelection();
            }
        }

        public void EndMappingMode()
        {
            _isInMappingMode = false;
            _currentMappingReference = "";
            _mappingCompletedCallback = null;

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

            if (_isBulkMappingSelectionMode)
            {
                return;
            }

            if (_isRemovingCells)
            {
                if (btn != null)
                {
                    var cellEntry = allCells.FirstOrDefault(kvp => kvp.Value.ButtonRef == btn);
                    if (!cellEntry.Equals(default(KeyValuePair<(int globalRow, int globalCol), Cell>)))
                    {
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

            if (_isInMappingMode && !string.IsNullOrEmpty(_currentMappingReference))
            {
                int specialRow = special.Type == SpecialPointType.Door ? -2 : -1;
                int specialCol = 1000 + sectionIndex;

                _componentMappingManager?.AddMapping(_currentMappingReference, specialRow, specialCol, false);

                RecordMappingForUndo(_currentMappingReference);

                _mappingCompletedCallback?.Invoke(_currentMappingReference, "");
                EndMappingMode();
                return;
            }

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

            double extraDistance = 0;

            if (startPoint is SpecialPoint spStart)
            {
                if (allCells.TryGetValue((spStart.GlobalRow, spStart.GlobalCol), out startCell))
                {
                    extraDistance += spStart.Type == SpecialPointType.Door ? 900 : 400;
                }
            }
            else if (startPoint is Button btnStart)
            {
                startCell = allCells.Values.FirstOrDefault(c => c.ButtonRef == btnStart);
                if (startCell != null)
                    extraDistance += 100;
            }

            if (endPoint is SpecialPoint spEnd)
            {
                if (allCells.TryGetValue((spEnd.GlobalRow, spEnd.GlobalCol), out endCell))
                {
                    extraDistance += spEnd.Type == SpecialPointType.Door ? 1000 : 500;
                }
            }
            else if (endPoint is Button btnEnd)
            {
                endCell = allCells.Values.FirstOrDefault(c => c.ButtonRef == btnEnd);
            }

            if (startCell == null || endCell == null)
            {
                ResultText.Text = "Invalid start or end point.";
                ResetSelection();
                return;
            }

            if (startCell.Row == endCell.Row && startCell.Col == endCell.Col)
            {
                HighlightPath(new List<Cell> { startCell });
                double totalDistanceSame = 300;
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
            double baseDistance = PathFinder.CalculateDistance(path, endsInSpecial, HasHorizontalNeighbor);
            double totalDistance;
            bool startIsButton = startPoint is Button;
            bool endIsButton = endPoint is Button;
            if (startIsButton && endIsButton)
            {
                totalDistance = baseDistance;
            }
            else
            {
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

            _componentMappingWindow = new ComponentMappingWindow(this, _componentMappingManager);
            _componentMappingWindow.Show();
        }

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

            MappingIndicatorText.Text = $"Mapper: {excelReference}";
            MappingIndicator.Visibility = Visibility.Visible;

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

            foreach (var cell in mappingCells.Values)
            {
                cell.ButtonRef.Background = new SolidColorBrush(Color.FromRgb(100, 120, 140));
            }

            foreach (var cell in allCells.Values)
            {
                cell.ButtonRef.IsEnabled = false;
                cell.ButtonRef.Background = new SolidColorBrush(Color.FromRgb(40, 40, 40));
            }

            foreach (var dp in doorPoints)
            {
                dp.Button.Background = new SolidColorBrush(Color.FromRgb(255, 140, 0));
            }
            foreach (var mp in motorPoints)
            {
                mp.Button.Background = new SolidColorBrush(Color.FromRgb(255, 140, 0));
            }

            this.Activate();
            this.Focus();
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
            _isRemovingCells = false;
            if (RemoveCellsButton != null)
            {
                RemoveCellsButton.Content = "Fjern celler";
            }
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

        private static void WriteCellWithColour(Excel.Worksheet sheet, int row, int col, string value, uint colour)
        {
            try
            {
                var rng = (Excel.Range)sheet.Cells[row, col];
                rng.Value2 = value;
                if (colour != 0)
                {
                    rng.Interior.Color = colour;
                }
            }
            catch
            {
                try
                {
                    var rng = (Excel.Range)sheet.Cells[row, col];
                    rng.Value2 = value;
                }
                catch
                {
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
                app = excelApp ?? new Excel.Application();
                app.Visible = true;
                wb = app.Workbooks.Open(destPath);

                ws = FindWorksheetByName(wb, "mal ferdig ledning");
                if (ws == null)
                    throw new InvalidOperationException("Finner ikke arket 'mal ferdig ledning' i eksportfilen.");

                var used = sourceSheet.UsedRange;
                int lastRow = used != null ? used.Rows.Count : 2;
                if (lastRow < 2) lastRow = 2;

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

                    WriteCellWithColour(ws, destRow, 3, src[0], colours[0]);
                    WriteCellWithColour(ws, destRow, 6, src[1], colours[1]);
                    WriteCellWithColour(ws, destRow, 7, src[2], colours[2]);
                    WriteCellWithColour(ws, destRow, 2, src[3], colours[3]);
                    WriteCellWithColour(ws, destRow, 4, src[4], colours[4]);
                    WriteCellWithColour(ws, destRow, 5, src[5], colours[5]);
                    WriteCellWithColour(ws, destRow, 10, src[6], colours[6]);
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

        private void ExportToDurapart(Excel.Worksheet sourceSheet, string destPath)
        {
            Excel.Application app = null;
            Excel.Workbook wb = null;
            Excel.Worksheet ws = null;
            try
            {
                app = excelApp ?? new Excel.Application();
                app.Visible = true;
                wb = app.Workbooks.Open(destPath);

                ws = FindWorksheetByName(wb, "Durapart");
                if (ws == null)
                    ws = (Excel.Worksheet)wb.Worksheets[1];

                var used = sourceSheet.UsedRange;
                int lastRow = used != null ? used.Rows.Count : 2;
                if (lastRow < 2) lastRow = 2;

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

                    WriteCellWithColour(ws, destRow, 7, src[0], colours[0]);
                    WriteCellWithColour(ws, destRow, 4, src[1], colours[1]);
                    WriteCellWithColour(ws, destRow, 9, src[2], colours[2]);
                    WriteCellWithColour(ws, destRow, 6, src[3], colours[3]);
                    WriteCellWithColour(ws, destRow, 5, src[4], colours[4]);
                    WriteCellWithColour(ws, destRow, 10, src[5], colours[5]);
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

            double distance = 100;

            for (int i = 1; i < path.Count - 1; i++)
                distance += hasHorizontalNeighbor(path[i].Row, path[i].Col) ? 100 : 50;

            if (!endsInSpecial)
                distance += 200;

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
                        string cellB = "";
                        string cellC = "";
                        object cellA = null;

                        try { cellB = _mainWindow.worksheet.Cells[row, 2].Value?.ToString() ?? ""; } catch { }
                        try { cellC = _mainWindow.worksheet.Cells[row, 3].Value?.ToString() ?? ""; } catch { }
                        try { cellA = _mainWindow.worksheet.Cells[row, 1].Value; } catch { }

                        if (cellA != null && !string.IsNullOrEmpty(cellA.ToString()) &&
                            double.TryParse(cellA.ToString(), out _))
                            continue;

                        if (string.IsNullOrWhiteSpace(cellB) && string.IsNullOrWhiteSpace(cellC))
                            continue;

                        var mappingA = _mappingManager.GetMapping(cellB);
                        var mappingB = _mappingManager.GetMapping(cellC);
                        var bulkA = _mappingManager.GetBulkRangeMappingForReference(cellB);
                        var bulkB = _mappingManager.GetBulkRangeMappingForReference(cellC);

                        // BULK TIL BULK: FINN LENGSTE VEI
                        if (bulkA != null && bulkB != null)
                        {
                            var candidateA = GetCandidateCellsForBulkReference(bulkA, cellB, allCells);
                            var candidateB = GetCandidateCellsForBulkReference(bulkB, cellC, allCells);

                            double maxDistance = 0;

                            foreach (var posA in candidateA)
                            {
                                foreach (var posB in candidateB)
                                {
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

                                    // Standard pathfinding som vanlige celler
                                    double pathDist = 100; // Start cell
                                    for (int i = 1; i < path.Count - 1; i++)
                                        pathDist += _mainWindow.HasHorizontalNeighbor(path[i].Row, path[i].Col) ? 100 : 50;
                                    pathDist += 200; // End cell

                                    maxDistance = Math.Max(maxDistance, pathDist);
                                }
                            }

                            if (maxDistance > 0)
                            {
                                _mainWindow.worksheet.Cells[row, 1].Value = maxDistance;
                                processedCount++;
                                continue;
                            }
                        }
                        // BULK TIL VANLIG ELLER VANLIG TIL BULK: FINN LENGSTE VEI
                        else if ((bulkA != null && mappingB != null) || (mappingA != null && bulkB != null))
                        {
                            double maxDistance = 0;

                            if (bulkA != null && mappingB != null)
                            {
                                // Bulk START → Vanlig SLUTT
                                var candidateA = GetCandidateCellsForBulkReference(bulkA, cellB, allCells);
                                var posB = GetGridPositionFromMapping(mappingB);

                                if (posB.HasValue)
                                {
                                    // Sjekk om mappingB er Motor eller Door (SLUTT-punkt)
                                    bool bIsMotor = mappingB.GridRow == -1;
                                    bool bIsDoor = mappingB.GridRow == -2;

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
                                            // Standard målelogikk
                                            double pathDist = 100;
                                            for (int i = 1; i < path.Count - 1; i++)
                                                pathDist += _mainWindow.HasHorizontalNeighbor(path[i].Row, path[i].Col) ? 100 : 50;

                                            // SLUTT-punkt: Motor/Door har forskjellige verdier
                                            if (bIsMotor)
                                            {
                                                pathDist += 500; // Motor som SLUTT-punkt
                                                pathDist += 50;  // Special point buffer
                                            }
                                            else if (bIsDoor)
                                            {
                                                pathDist += 1000; // Door som SLUTT-punkt
                                                pathDist += 50;   // Special point buffer
                                            }
                                            else
                                            {
                                                pathDist += 200; // Normal slutt-celle
                                            }

                                            maxDistance = Math.Max(maxDistance, pathDist);
                                        }
                                    }
                                }
                            }
                            else if (mappingA != null && bulkB != null)
                            {
                                // Vanlig START → Bulk SLUTT
                                var posA = GetGridPositionFromMapping(mappingA);
                                var candidateB = GetCandidateCellsForBulkReference(bulkB, cellC, allCells);

                                if (posA.HasValue)
                                {
                                    // Sjekk om mappingA er Motor eller Door (START-punkt)
                                    bool aIsMotor = mappingA.GridRow == -1;
                                    bool aIsDoor = mappingA.GridRow == -2;

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
                                            double pathDist = 100;
                                            for (int i = 1; i < path.Count - 1; i++)
                                                pathDist += _mainWindow.HasHorizontalNeighbor(path[i].Row, path[i].Col) ? 100 : 50;
                                            pathDist += 200; // Slutt-celle (bulk-siden)

                                            // START-punkt: Motor/Door har forskjellige verdier
                                            if (aIsMotor)
                                            {
                                                pathDist += 400; // Motor som START-punkt
                                                pathDist += 50;  // Special point buffer
                                            }
                                            else if (aIsDoor)
                                            {
                                                pathDist += 900; // Door som START-punkt
                                                pathDist += 50;  // Special point buffer
                                            }

                                            maxDistance = Math.Max(maxDistance, pathDist);
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

                        // VANLIG MAPPING TIL VANLIG MAPPING
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
            if (mapping.GridRow == -1 && mapping.GridColumn >= 1000)
            {
                var motorIndex = mapping.GridColumn - 1000;
                var motorPoint = _mainWindow.GetMotorPoint(motorIndex);
                return motorPoint != null ? (motorPoint.GlobalRow, motorPoint.GlobalCol) : null;
            }
            else if (mapping.GridRow == -2 && mapping.GridColumn >= 1000)
            {
                var doorIndex = mapping.GridColumn - 1000;
                var doorPoint = _mainWindow.GetDoorPoint(doorIndex);
                return doorPoint != null ? (doorPoint.GlobalRow, doorPoint.GlobalCol) : null;
            }
            else
            {
                if (mapping.GridRow < 0)
                {
                    int actualRow = -(mapping.GridRow + 1);
                    return (actualRow, mapping.GridColumn);
                }
                else
                {
                    int adjustedRow = mapping.GridRow;
                    var allCells = _mainWindow.GetAllCells();

                    if (allCells.ContainsKey((mapping.GridRow, mapping.GridColumn)))
                    {
                        return (mapping.GridRow, mapping.GridColumn);
                    }

                    if (mapping.DefaultToBottom && allCells.ContainsKey((mapping.GridRow + 1, mapping.GridColumn)))
                    {
                        return (mapping.GridRow + 1, mapping.GridColumn);
                    }

                    if (allCells.ContainsKey((mapping.GridRow - 1, mapping.GridColumn)))
                    {
                        return (mapping.GridRow - 1, mapping.GridColumn);
                    }

                    return (mapping.GridRow, mapping.GridColumn);
                }
            }
        }

        private double GetConnectionDistance(string text, ComponentMapping mapping = null)
        {
            if (string.IsNullOrEmpty(text)) return 0;

            bool hasAsterisk = text.Contains("*");

            if (mapping != null)
            {
                if (mapping.GridRow == -1)
                    return 500;
                else if (mapping.GridRow == -2)
                    return 1000;
            }

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

        private List<(int Row, int Col)> GetCandidateCellsForBulkReference(ComponentMappingManager.BulkRangeMapping bulkRange, string excelReference, Dictionary<(int, int), Cell> allCells)
        {
            var result = new List<(int Row, int Col)>();
            if (bulkRange == null) return result;

            bool hasStar = !string.IsNullOrWhiteSpace(excelReference) && excelReference.Trim().EndsWith("*");
            bool selectedIsTop = bulkRange.SelectedIsTop;

            // VIKTIG LOGIKK:
            // Hvis bulk mappet til T-celler (selectedIsTop=true):
            //   - UTEN stjerne: Bruk B-siden (undersiden) 
            //   - MED stjerne: Bruk T-siden (oversiden)
            //
            // Hvis bulk mappet til B-celler (selectedIsTop=false):
            //   - UTEN stjerne: Bruk T-siden (oversiden)
            //   - MED stjerne: Bruk B-siden (undersiden)

            bool useSideIsTop;
            if (selectedIsTop)
            {
                // Mappet til T-celler
                useSideIsTop = hasStar; // Med stjerne: T-siden, uten stjerne: B-siden
            }
            else
            {
                // Mappet til B-celler
                useSideIsTop = !hasStar; // Med stjerne: B-siden, uten stjerne: T-siden
            }

            foreach (var cell in bulkRange.Cells)
            {
                int mappingRow = cell.Row;
                int col = cell.Col;

                int actualCellRow;

                // Finn vanlig celle basert på mappingRow og hvilken side vi skal bruke
                if (mappingRow >= 0)
                {
                    // Positiv mappingRow = T-mapping-celle
                    if (useSideIsTop)
                    {
                        // Bruk T-siden → vanlig celle OVER mapping-cellen
                        actualCellRow = mappingRow - 1;
                    }
                    else
                    {
                        // Bruk B-siden → vanlig celle UNDER mapping-cellen
                        actualCellRow = mappingRow + 1;
                    }
                }
                else
                {
                    // Negativ mappingRow = B-mapping-celle
                    // Konverter til positiv rad først
                    int posRow = -(mappingRow + 1);

                    if (useSideIsTop)
                    {
                        // Bruk T-siden → vanlig celle OVER mapping-cellen
                        actualCellRow = posRow - 1;
                    }
                    else
                    {
                        // Bruk B-siden → vanlig celle UNDER mapping-cellen
                        actualCellRow = posRow + 1;
                    }
                }

                // Debug: sjekk om cellen finnes
                if (allCells.ContainsKey((actualCellRow, col)))
                {
                    result.Add((actualCellRow, col));
                }
                else
                {
                    // Prøv alternativ hvis ikke funnet
                    if (mappingRow >= 0)
                    {
                        // Prøv både over og under
                        if (allCells.ContainsKey((mappingRow - 1, col)))
                            result.Add((mappingRow - 1, col));
                        else if (allCells.ContainsKey((mappingRow + 1, col)))
                            result.Add((mappingRow + 1, col));
                    }
                    else
                    {
                        int posRow = -(mappingRow + 1);
                        if (allCells.ContainsKey((posRow - 1, col)))
                            result.Add((posRow - 1, col));
                        else if (allCells.ContainsKey((posRow + 1, col)))
                            result.Add((posRow + 1, col));
                    }
                }
            }

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