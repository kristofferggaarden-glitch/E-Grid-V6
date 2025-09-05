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

        // Når true vil klikk på vanlige celler «fjerne» cellen (skjules og deaktiveres)
        // i stedet for å starte måling.
        private bool _isRemovingCells = false;

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

        private void MappingCell_Click(object sender, RoutedEventArgs e)
        {
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

                    _mappingCompletedCallback?.Invoke(_currentMappingReference, "");
                    EndMappingMode();
                }
            }
        }

        public void EndMappingMode()
        {
            _isInMappingMode = false;
            _currentMappingReference = "";
            _mappingCompletedCallback = null;

            // Hide mapping indicator and cancel button
            MappingIndicator.Visibility = Visibility.Collapsed;
            CancelMappingButton.Visibility = Visibility.Collapsed;

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

            ResultText.Text = "";
        }

        private void ResetCellColors()
        {
            foreach (var cell in allCells.Values)
            {
                // Ikke reaktiver celler som er «fjernet»
                if (cell.ButtonRef.IsEnabled)
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

            // Fjerning av celler uten å kollapse griden:
            // - Ikke fjern fra Grid.Children
            // - Ikke sett Visibility=Collapsed (det kollapser layout)
            // - Bruk Visibility.Hidden så plassen beholdes
            // - Fjern cellen fra allCells slik at pathfinding ignorerer den
            if (_isRemovingCells)
            {
                if (btn != null)
                {
                    var cellEntry = allCells.FirstOrDefault(kvp => kvp.Value.ButtonRef == btn);
                    if (!cellEntry.Equals(default(KeyValuePair<(int globalRow, int globalCol), Cell>)))
                    {
                        btn.IsEnabled = false;
                        btn.Visibility = Visibility.Hidden;  // beholder plass i layout
                        btn.Opacity = 0.0;
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

            // Sjekk mapping-modus
            if (_isInMappingMode && !string.IsNullOrEmpty(_currentMappingReference))
            {
                int specialRow = special.Type == SpecialPointType.Door ? -2 : -1;
                int specialCol = 1000 + sectionIndex; // Spesiell koding for Door/Motor

                _componentMappingManager?.AddMapping(_currentMappingReference, specialRow, specialCol, false);

                _mappingCompletedCallback?.Invoke(_currentMappingReference, "");
                EndMappingMode();
                return;
            }

            // Vanlig klikk for manuell måling
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

            // Ekstra distanse for å justere start- og sluttdistanser (symmetri)
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

            // 💡 Ny regel: samme celle = 300 mm (f.eks. Motor1–Motor1 eller celle–samme-celle)
            if (startCell == endCell)
            {
                HighlightPath(new List<Cell> { startCell });
                const double sameRefDistance = 300.0;
                ResultText.Text = $"Shortest path: {sameRefDistance:F2} mm";
                FindNextAvailableRow();
                LogMeasurementToExcel(sameRefDistance);
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

            double totalDistance; // << kun én deklarasjon (fix for CS0136)
            bool startIsButton = startPoint is Button;
            bool endIsButton = endPoint is Button;

            if (startIsButton && endIsButton)
            {
                // Begge er vanlige celler: bruk kun baseDistance
                totalDistance = baseDistance;
            }
            else
            {
                // Minst én er motor/dør: bruk kompensasjon + 50 mm
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
                if (cell.ButtonRef.IsEnabled)
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
                if (cell.ButtonRef.IsEnabled)
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
                if (cell.ButtonRef.IsEnabled)
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

            var mappingWindow = new ComponentMappingWindow(this, _componentMappingManager);
            mappingWindow.Show();
        }

        public void StartInteractiveMapping(string excelReference, string description, Action<string, string> onCompleted)
        {
            _currentMappingReference = excelReference;
            _isInMappingMode = true;
            _mappingCompletedCallback = onCompleted;

            // Show mapping indicator with reference text
            MappingIndicatorText.Text = $"Mapper: {excelReference}";
            MappingIndicator.Visibility = Visibility.Visible;
            CancelMappingButton.Visibility = Visibility.Visible;

            ResultText.Text = $"Klikk på grid-posisjonen for {excelReference} (mellomrader, Door eller Motor knapper)";

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

        private void CancelMapping_Click(object sender, RoutedEventArgs e)
        {
            EndMappingMode();
        }

        /// <summary>
        /// Toggle cell removal mode on or off.
        /// </summary>
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

        private void AutomaticMeasureAll_Click(object sender, RoutedEventArgs e)
        {
            if (string.IsNullOrEmpty(SelectedExcelFile) || worksheet == null)
            {
                MessageBox.Show("Velg først en Excel-fil", "Feil",
                               MessageBoxButton.OK, MessageBoxImage.Warning);
                return;
            }

            if (_componentMappingManager == null)
            {
                MessageBox.Show("Du må først sette opp component mappings. Bruk 'Component Mapping' knappen.",
                               "Mappings mangler", MessageBoxButton.OK, MessageBoxImage.Warning);
                return;
            }

            var result = MessageBox.Show(
                "Dette vil automatisk måle alle ledninger som har component mappings. Fortsette?",
                "Automatisk måling",
                MessageBoxButton.YesNo,
                MessageBoxImage.Question);

            if (result == MessageBoxResult.Yes)
            {
                try
                {
                    var processor = new ExcelConnectionProcessor(this, _componentMappingManager);
                    var processedCount = processor.ProcessAllConnections();

                    MessageBox.Show($"Automatisk måling fullført!\nProsesserte {processedCount} ledninger.",
                                   "Ferdig", MessageBoxButton.OK, MessageBoxImage.Information);

                    UpdateExcelDisplayText();
                }
                catch (Exception ex)
                {
                    MessageBox.Show($"Feil under automatisk måling: {ex.Message}", "Feil",
                                   MessageBoxButton.OK, MessageBoxImage.Error);
                }
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
            // Reset removal mode when rebuilding the grid
            _isRemovingCells = false;
            if (RemoveCellsButton != null)
            {
                RemoveCellsButton.Content = "Fjern celler";
            }
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

    // BEHOLDT ORIGINAL ExcelConnectionProcessor med tillegg for 300 mm ved samme referanse
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
                throw new InvalidOperationException("Ingen Excel-fil er åpen");

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

                        // Hopp over hvis allerede målt
                        if (cellA != null && !string.IsNullOrEmpty(cellA.ToString()) &&
                            double.TryParse(cellA.ToString(), out _))
                            continue;

                        // Hopp over hvis tom
                        if (string.IsNullOrWhiteSpace(cellB) && string.IsNullOrWhiteSpace(cellC))
                            continue;

                        // Begge sider må ha mapping
                        var mappingA = _mappingManager.GetMapping(cellB);
                        var mappingB = _mappingManager.GetMapping(cellC);

                        if (mappingA != null && mappingB != null)
                        {
                            var posA = GetGridPositionFromMapping(mappingA);
                            var posB = GetGridPositionFromMapping(mappingB);

                            if (posA.HasValue && posB.HasValue)
                            {
                                // 💡 Ny regel: Samme posisjon = 300 mm
                                if (posA.Value == posB.Value)
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
                                        // Beregn path-avstand (kun mellom grid-punkter)
                                        double pathDistance = 0;
                                        for (int i = 1; i < path.Count - 1; i++)
                                            pathDistance += _mainWindow.HasHorizontalNeighbor(path[i].Row, path[i].Col) ? 100 : 50;

                                        bool aIsNormal = mappingA.GridRow != -1 && mappingA.GridRow != -2;
                                        bool bIsNormal = mappingB.GridRow != -1 && mappingB.GridRow != -2;
                                        double totalDistance;

                                        if (aIsNormal && bIsNormal)
                                        {
                                            // Vanlige celler begge ender
                                            totalDistance = pathDistance + 100 + 200;
                                        }
                                        else
                                        {
                                            // Motor/Dør-involvering: behold opprinnelig logikk
                                            double startDistance;
                                            if (mappingA.GridRow == -1) // Motor
                                                startDistance = 500;
                                            else if (mappingA.GridRow == -2) // Door
                                                startDistance = 1000;
                                            else
                                                startDistance = 200; // Normal celle

                                            double endDistance;
                                            if (mappingB.GridRow == -1) // Motor
                                                endDistance = 500;
                                            else if (mappingB.GridRow == -2) // Door
                                                endDistance = 1000;
                                            else
                                                endDistance = 200; // Normal celle

                                            totalDistance = pathDistance + startDistance + endDistance + 50;
                                        }

                                        if (totalDistance > 0)
                                        {
                                            _mainWindow.worksheet.Cells[row, 1].Value = totalDistance;
                                            processedCount++;
                                        }
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

        // ORIGINAL GetGridPositionFromMapping (med forbedringer for T/B)
        private (int Row, int Col)? GetGridPositionFromMapping(ComponentMapping mapping)
        {
            // Special points (Motors and Doors)
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
                // Bottom side mappings (negative rows)
                if (mapping.GridRow < 0)
                {
                    int actualRow = -(mapping.GridRow + 1);
                    return (actualRow, mapping.GridColumn);
                }
                else
                {
                    // Top-side mapping, prøv smart fallback
                    var allCells = _mainWindow.GetAllCells();

                    if (allCells.ContainsKey((mapping.GridRow, mapping.GridColumn)))
                        return (mapping.GridRow, mapping.GridColumn);

                    if (mapping.DefaultToBottom && allCells.ContainsKey((mapping.GridRow + 1, mapping.GridColumn)))
                        return (mapping.GridRow + 1, mapping.GridColumn);

                    if (allCells.ContainsKey((mapping.GridRow - 1, mapping.GridColumn)))
                        return (mapping.GridRow - 1, mapping.GridColumn);

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

        public string TestProcessing()
        {
            if (_mainWindow.worksheet == null)
                return "Ingen Excel-fil er åpen";

            var results = new List<string>();
            var allMappings = _mappingManager.GetAllMappings();
            var usedRange = _mainWindow.worksheet.UsedRange;

            if (usedRange == null)
                return "Ingen data i Excel-arket";

            var lastRow = Math.Min(usedRange.Rows.Count, 10);

            results.Add("=== TEST PROSESSERING ===\n");
            results.Add($"Antall mappings: {allMappings.Count}");

            foreach (var mapping in allMappings)
            {
                if (mapping.GridRow == -1)
                    results.Add($"  {mapping.ExcelReference} -> Motor {mapping.GridColumn - 1000}");
                else if (mapping.GridRow == -2)
                    results.Add($"  {mapping.ExcelReference} -> Door {mapping.GridColumn - 1000}");
                else
                    results.Add($"  {mapping.ExcelReference} -> ({mapping.GridRow}, {mapping.GridColumn}) {(mapping.DefaultToBottom ? "(B)" : "(T)")}");
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

                    results.Add($"  Match A: {mappingA?.ExcelReference ?? "IKKE FUNNET"}");
                    results.Add($"  Match B: {mappingB?.ExcelReference ?? "IKKE FUNNET"}");

                    if (mappingA != null && mappingB != null)
                    {
                        var posA = GetGridPositionFromMapping(mappingA);
                        var posB = GetGridPositionFromMapping(mappingB);

                        results.Add($"  Pos A: {posA?.ToString() ?? "IKKE FUNNET"}");
                        results.Add($"  Pos B: {posB?.ToString() ?? "IKKE FUNNET"}");

                        if (posA.HasValue && posB.HasValue)
                        {
                            if (posA.Value == posB.Value)
                            {
                                results.Add($"  -> SAMME REF: 300.00 mm\n");
                                continue;
                            }

                            var allCells = _mainWindow.GetAllCells();
                            if (allCells.TryGetValue(posA.Value, out var startCell) &&
                                allCells.TryGetValue(posB.Value, out var endCell))
                            {
                                var path = PathFinder.FindShortestPath(startCell, endCell, allCells, _mainWindow.HasHorizontalNeighbor);
                                if (path != null && path.Count > 0)
                                {
                                    double baseDistance = PathFinder.CalculateDistance(path, false, _mainWindow.HasHorizontalNeighbor);
                                    double connectionDistanceA = GetConnectionDistance(cellB, mappingA);
                                    double connectionDistanceB = GetConnectionDistance(cellC, mappingB);

                                    // Speil logikken i manuell/auto (ekstra 50 mm)
                                    double totalDistance = baseDistance + connectionDistanceA + connectionDistanceB + 50;
                                    results.Add($"  BASE AVSTAND: {baseDistance:F2} mm");
                                    results.Add($"  TILKOBLINGS A: {connectionDistanceA:F2} mm");
                                    results.Add($"  TILKOBLINGS B: {connectionDistanceB:F2} mm");
                                    results.Add($"  TOTAL AVSTAND: {totalDistance:F2} mm");
                                }
                                else
                                {
                                    results.Add("  -> INGEN STI FUNNET");
                                }
                            }
                            else
                            {
                                results.Add("  -> CELLER IKKE FUNNET I GRID");
                            }
                        }
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
