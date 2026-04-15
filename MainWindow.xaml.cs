using Microsoft.VisualBasic;
using Microsoft.Win32;
using System.ComponentModel;
using System.Data;
using System.Globalization;
using System.IO;
using System.Runtime.CompilerServices;
using System.Text.Json;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Media;
using System.Windows.Media.Imaging;

namespace WpfDemo
{
    public partial class MainWindow : Window, INotifyPropertyChanged
    {
        private const string NumberColumn = "번호";
        private const string ViewerColumn = "시청자";
        private const string GreetingColumn = "방종 인사";
        private const string ReactionColumn = "리액션";
        private const string ActionColumn = "행동";
        private const string AttendanceColumn = "출석";
        private const string ExtraColumn = "기타";
        private const string TotalColumn = "총점";
        private const string MemoColumn = "메모";

        private readonly JsonSerializerOptions jsonOptions = new() { WriteIndented = true };
        private readonly string autoSaveFilePath = Path.Combine(
            Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),
            "WpfDemo",
            "upi-viewer-autosave.json");

        private readonly DataTable table = new("ViewerScores");

        private BitmapImage? upiImageSource;
        private BitmapImage? ppuyoImageSource;
        private string? upiImagePath;
        private string? ppuyoImagePath;

        public MainWindow()
        {
            InitializeComponent();
            DataContext = this;
            Closing += MainWindow_Closing;

            InitializeTable();
            TableGrid.ItemsSource = table.DefaultView;

            if (!TryLoadAutoSave())
            {
                LoadSampleData();
                UpdateStatus("샘플 점수표를 불러왔습니다.");
            }
        }

        public BitmapImage? UpiImageSource
        {
            get => upiImageSource;
            set
            {
                if (SetField(ref upiImageSource, value))
                {
                    OnPropertyChanged(nameof(UpiImageHintVisibility));
                }
            }
        }

        public BitmapImage? PpuyoImageSource
        {
            get => ppuyoImageSource;
            set
            {
                if (SetField(ref ppuyoImageSource, value))
                {
                    OnPropertyChanged(nameof(PpuyoImageHintVisibility));
                }
            }
        }

        public Visibility UpiImageHintVisibility => UpiImageSource is null ? Visibility.Visible : Visibility.Collapsed;

        public Visibility PpuyoImageHintVisibility => PpuyoImageSource is null ? Visibility.Visible : Visibility.Collapsed;

        private void InitializeTable()
        {
            table.Columns.Clear();

            table.Columns.Add(NumberColumn, typeof(int));
            table.Columns[NumberColumn]!.Caption = NumberColumn;

            table.Columns.Add(ViewerColumn, typeof(string));
            table.Columns[ViewerColumn]!.Caption = ViewerColumn;

            table.Columns.Add(GreetingColumn, typeof(int));
            table.Columns[GreetingColumn]!.Caption = GreetingColumn;

            table.Columns.Add(ReactionColumn, typeof(int));
            table.Columns[ReactionColumn]!.Caption = ReactionColumn;

            table.Columns.Add(ActionColumn, typeof(int));
            table.Columns[ActionColumn]!.Caption = ActionColumn;

            table.Columns.Add(AttendanceColumn, typeof(int));
            table.Columns[AttendanceColumn]!.Caption = AttendanceColumn;

            table.Columns.Add(ExtraColumn, typeof(int));
            table.Columns[ExtraColumn]!.Caption = ExtraColumn;

            DataColumn totalColumn = table.Columns.Add(TotalColumn, typeof(int));
            totalColumn.Expression = $"[{GreetingColumn}] + [{ReactionColumn}] + [{ActionColumn}] + [{AttendanceColumn}] + [{ExtraColumn}]";
            totalColumn.Caption = TotalColumn;

            table.Columns.Add(MemoColumn, typeof(string));
            table.Columns[MemoColumn]!.Caption = MemoColumn;
        }

        private void AddRow_Click(object sender, RoutedEventArgs e)
        {
            DataRow row = table.NewRow();
            FillDefaultRow(row);
            table.Rows.Add(row);
            RenumberRows();
            RefreshOverallScore();
            TableGrid.SelectedIndex = table.Rows.Count - 1;
            UpdateStatus("새 시청자 행을 추가했습니다.");
        }

        private void InsertAbove_Click(object sender, RoutedEventArgs e)
        {
            InsertRelativeRow(0);
        }

        private void InsertBelow_Click(object sender, RoutedEventArgs e)
        {
            InsertRelativeRow(1);
        }

        private void DeleteRow_Click(object sender, RoutedEventArgs e)
        {
            DataRowView? rowView = GetCurrentRowView();
            if (rowView is null)
            {
                UpdateStatus("먼저 행을 선택해주세요.");
                return;
            }

            table.Rows.Remove(rowView.Row);
            RenumberRows();
            RefreshOverallScore();
            UpdateStatus("선택한 행을 삭제했습니다.");
        }

        private void AddColumn_Click(object sender, RoutedEventArgs e)
        {
            string name = Interaction.InputBox("추가할 열 이름을 입력해주세요.", "열 추가", $"추가정보{Math.Max(1, table.Columns.Count - 8)}").Trim();

            if (string.IsNullOrWhiteSpace(name))
            {
                UpdateStatus("열 추가를 취소했습니다.");
                return;
            }

            if (table.Columns.Contains(name))
            {
                UpdateStatus("같은 이름의 열이 이미 있습니다.");
                return;
            }

            DataColumn column = table.Columns.Add(name, typeof(string));
            column.DefaultValue = string.Empty;
            column.Caption = name;
            RefreshGridBinding();
            UpdateStatus($"'{name}' 열을 추가했습니다.");
        }

        private void RenameColumn_Click(object sender, RoutedEventArgs e)
        {
            string? currentColumnName = GetSelectedColumnName();
            if (string.IsNullOrWhiteSpace(currentColumnName))
            {
                UpdateStatus("이름을 바꿀 열의 셀을 먼저 선택해주세요.");
                return;
            }

            DataColumn currentColumn = table.Columns[currentColumnName]!;
            string currentHeader = string.IsNullOrWhiteSpace(currentColumn.Caption) ? currentColumn.ColumnName : currentColumn.Caption;

            string newName = Interaction.InputBox("새 열 이름을 입력해주세요.", "열 이름 변경", currentHeader).Trim();
            if (string.IsNullOrWhiteSpace(newName))
            {
                UpdateStatus("열 이름 변경을 취소했습니다.");
                return;
            }

            if (string.Equals(currentHeader, newName, StringComparison.Ordinal))
            {
                UpdateStatus("같은 이름으로는 변경되지 않습니다.");
                return;
            }

            bool duplicateCaption = table.Columns
                .Cast<DataColumn>()
                .Any(column => !ReferenceEquals(column, currentColumn)
                    && string.Equals(string.IsNullOrWhiteSpace(column.Caption) ? column.ColumnName : column.Caption, newName, StringComparison.Ordinal));

            if (duplicateCaption)
            {
                UpdateStatus("화면에 같은 이름의 열이 이미 있습니다.");
                return;
            }

            currentColumn.Caption = newName;
            RefreshGridBinding();
            UpdateStatus($"열 이름을 '{currentHeader}'에서 '{newName}'으로 변경했습니다.");
        }

        private void DeleteColumn_Click(object sender, RoutedEventArgs e)
        {
            string? columnName = GetSelectedColumnName();
            if (string.IsNullOrWhiteSpace(columnName))
            {
                UpdateStatus("삭제할 열의 셀을 먼저 선택해주세요.");
                return;
            }

            if (IsProtectedColumn(columnName))
            {
                UpdateStatus("기본 점수 열과 총점 열은 삭제할 수 없습니다.");
                return;
            }

            table.Columns.Remove(columnName);
            RefreshGridBinding();
            UpdateStatus($"'{columnName}' 열을 삭제했습니다.");
        }

        private void LoadSample_Click(object sender, RoutedEventArgs e)
        {
            LoadSampleData();
            UpdateStatus("샘플 점수표를 불러왔습니다.");
        }

        private void Save_Click(object sender, RoutedEventArgs e)
        {
            var dialog = new SaveFileDialog
            {
                Filter = "JSON 파일 (*.json)|*.json",
                DefaultExt = ".json",
                FileName = "upi-viewer-scores.json"
            };

            if (dialog.ShowDialog() != true)
            {
                UpdateStatus("파일 저장을 취소했습니다.");
                return;
            }

            File.WriteAllText(dialog.FileName, JsonSerializer.Serialize(CreateAppState(), jsonOptions));
            UpdateStatus($"파일을 저장했습니다: {dialog.FileName}");
        }

        private void Load_Click(object sender, RoutedEventArgs e)
        {
            var dialog = new OpenFileDialog
            {
                Filter = "JSON 파일 (*.json)|*.json",
                DefaultExt = ".json"
            };

            if (dialog.ShowDialog() != true)
            {
                UpdateStatus("파일 불러오기를 취소했습니다.");
                return;
            }

            string json = File.ReadAllText(dialog.FileName);
            if (TryLoadStateFromJson(json, out ViewerScoreAppState? state) && state is not null)
            {
                ApplyAppState(state);
                UpdateStatus($"파일을 불러왔습니다: {dialog.FileName}");
                return;
            }

            UpdateStatus("파일 형식을 읽지 못했습니다.");
        }

        private void LoadUpiImage_Click(object sender, RoutedEventArgs e)
        {
            BitmapImage? image = SelectImageFile();
            if (image is null)
            {
                UpdateStatus("UPI 이미지 불러오기를 취소했습니다.");
                return;
            }

            UpiImageSource = image;
            upiImagePath = image.UriSource?.LocalPath;
            UpdateStatus("UPI 이미지를 바꿨습니다.");
        }

        private void LoadPpuyoImage_Click(object sender, RoutedEventArgs e)
        {
            BitmapImage? image = SelectImageFile();
            if (image is null)
            {
                UpdateStatus("뿌요 이미지 불러오기를 취소했습니다.");
                return;
            }

            PpuyoImageSource = image;
            ppuyoImagePath = image.UriSource?.LocalPath;
            UpdateStatus("뿌요 이미지를 바꿨습니다.");
        }

        private void InsertRelativeRow(int offset)
        {
            int insertIndex;

            DataRowView? rowView = GetCurrentRowView();
            if (rowView is not null)
            {
                int selectedIndex = table.Rows.IndexOf(rowView.Row);
                insertIndex = Math.Clamp(selectedIndex + offset, 0, table.Rows.Count);
            }
            else
            {
                insertIndex = offset == 0 ? 0 : table.Rows.Count;
            }

            DataRow row = table.NewRow();
            FillDefaultRow(row);
            table.Rows.InsertAt(row, insertIndex);
            RenumberRows();
            RefreshOverallScore();
            TableGrid.SelectedIndex = insertIndex;
            UpdateStatus(offset == 0 ? "행을 위쪽에 삽입했습니다." : "행을 아래쪽에 삽입했습니다.");
        }

        private void FillDefaultRow(DataRow row)
        {
            row[NumberColumn] = table.Rows.Count + 1;
            row[ViewerColumn] = "새 시청자";
            row[GreetingColumn] = 0;
            row[ReactionColumn] = 0;
            row[ActionColumn] = 0;
            row[AttendanceColumn] = 0;
            row[ExtraColumn] = 0;
            row[MemoColumn] = string.Empty;

            foreach (DataColumn column in table.Columns)
            {
                if (column.ColumnName == TotalColumn)
                {
                    continue;
                }

                if (column.DataType == typeof(string) && row.IsNull(column))
                {
                    row[column] = string.Empty;
                }
            }
        }

        private void LoadSampleData()
        {
            table.Rows.Clear();

            AddSampleRow(1, "시청자A", 3, 5, 2, 4, 1, "방종 인사가 좋았음");
            AddSampleRow(2, "시청자B", 1, 2, 5, 3, 2, "행동 점수가 높음");
            AddSampleRow(3, "시청자C", 4, 4, 4, 5, 0, "전반적으로 활발함");

            RenumberRows();
            RefreshOverallScore();
        }

        private void AddSampleRow(int number, string viewer, int greeting, int reaction, int action, int attendance, int extra, string memo)
        {
            DataRow row = table.NewRow();
            row[NumberColumn] = number;
            row[ViewerColumn] = viewer;
            row[GreetingColumn] = greeting;
            row[ReactionColumn] = reaction;
            row[ActionColumn] = action;
            row[AttendanceColumn] = attendance;
            row[ExtraColumn] = extra;
            row[MemoColumn] = memo;

            foreach (DataColumn column in table.Columns)
            {
                if (column.ColumnName == TotalColumn)
                {
                    continue;
                }

                if (column.DataType == typeof(string) && row.IsNull(column))
                {
                    row[column] = string.Empty;
                }
            }

            table.Rows.Add(row);
        }

        private void TableGrid_AutoGeneratingColumn(object sender, DataGridAutoGeneratingColumnEventArgs e)
        {
            e.Column.IsReadOnly = e.PropertyName == NumberColumn || e.PropertyName == TotalColumn;
            DataColumn sourceColumn = table.Columns[e.PropertyName]!;
            e.Column.Header = string.IsNullOrWhiteSpace(sourceColumn.Caption) ? sourceColumn.ColumnName : sourceColumn.Caption;

            if (e.Column is DataGridTextColumn textColumn)
            {
                textColumn.Width = e.PropertyName == MemoColumn ? new DataGridLength(220) : new DataGridLength(120);
            }

            if (e.PropertyName == ViewerColumn)
            {
                e.Column.Width = new DataGridLength(160);
            }

            if (e.PropertyName == NumberColumn)
            {
                e.Column.Width = new DataGridLength(70);
            }

            if (e.PropertyName == TotalColumn && e.Column is DataGridTextColumn totalTextColumn)
            {
                totalTextColumn.ElementStyle = BuildTotalCellStyle();
            }
        }

        private void TableGrid_AutoGeneratedColumns(object? sender, EventArgs e)
        {
            EnsurePpuyoImageColumn();
        }

        private void TableGrid_PreparingCellForEdit(object sender, DataGridPreparingCellForEditEventArgs e)
        {
            if (e.Column.SortMemberPath is not (GreetingColumn or ReactionColumn or ActionColumn or AttendanceColumn or ExtraColumn))
            {
                return;
            }

            if (e.EditingElement is TextBox textBox)
            {
                textBox.SelectAll();
            }
        }

        private void TableGrid_CellEditEnding(object sender, DataGridCellEditEndingEventArgs e)
        {
            Dispatcher.BeginInvoke(new Action(() =>
            {
                NormalizeNumericColumns();
                RefreshOverallScore();
            }));
        }

        private void NormalizeNumericColumns()
        {
            foreach (DataRow row in table.Rows)
            {
                row[GreetingColumn] = ParseInt(row[GreetingColumn]);
                row[ReactionColumn] = ParseInt(row[ReactionColumn]);
                row[ActionColumn] = ParseInt(row[ActionColumn]);
                row[AttendanceColumn] = ParseInt(row[AttendanceColumn]);
                row[ExtraColumn] = ParseInt(row[ExtraColumn]);
            }
        }

        private static int ParseInt(object value)
        {
            if (value is int intValue)
            {
                return intValue;
            }

            return int.TryParse(value?.ToString(), out int parsed) ? parsed : 0;
        }

        private void RenumberRows()
        {
            for (int i = 0; i < table.Rows.Count; i++)
            {
                table.Rows[i][NumberColumn] = i + 1;
            }
        }

        private void RefreshOverallScore()
        {
            int total = 0;
            foreach (DataRow row in table.Rows)
            {
                total += ParseInt(row[TotalColumn]);
            }

            OverallScoreTextBlock.Text = total.ToString();
        }

        private void MainWindow_Closing(object? sender, CancelEventArgs e)
        {
            TryAutoSave();
        }

        private bool TryLoadAutoSave()
        {
            try
            {
                if (!File.Exists(autoSaveFilePath))
                {
                    return false;
                }

                string json = File.ReadAllText(autoSaveFilePath);
                if (!TryLoadStateFromJson(json, out ViewerScoreAppState? state) || state is null)
                {
                    return false;
                }

                ApplyAppState(state);
                UpdateStatus("마지막 작업 상태를 복원했습니다.");
                return true;
            }
            catch
            {
                return false;
            }
        }

        private void TryAutoSave()
        {
            try
            {
                string? directory = Path.GetDirectoryName(autoSaveFilePath);
                if (!string.IsNullOrWhiteSpace(directory))
                {
                    Directory.CreateDirectory(directory);
                }

                File.WriteAllText(autoSaveFilePath, JsonSerializer.Serialize(CreateAppState(), jsonOptions));
            }
            catch
            {
            }
        }

        private ViewerScoreAppState CreateAppState()
        {
            var state = new ViewerScoreAppState
            {
                UpiImagePath = upiImagePath,
                PpuyoImagePath = ppuyoImagePath
            };

            foreach (DataColumn column in table.Columns)
            {
                state.Columns.Add(new ColumnState
                {
                    Name = column.ColumnName,
                    Header = string.IsNullOrWhiteSpace(column.Caption) ? column.ColumnName : column.Caption,
                    DataType = column.DataType == typeof(int) ? "int" : "string"
                });
            }

            foreach (DataRow row in table.Rows)
            {
                var rowState = new RowState();
                foreach (DataColumn column in table.Columns)
                {
                    if (column.ColumnName == TotalColumn)
                    {
                        continue;
                    }

                    rowState.Values[column.ColumnName] = row[column]?.ToString() ?? string.Empty;
                }

                state.Rows.Add(rowState);
            }

            return state;
        }

        private bool TryLoadStateFromJson(string json, out ViewerScoreAppState? state)
        {
            state = null;

            try
            {
                state = JsonSerializer.Deserialize<ViewerScoreAppState>(json);
                return state is not null;
            }
            catch
            {
                return false;
            }
        }

        private void ApplyAppState(ViewerScoreAppState state)
        {
            BuildTableFromState(state);
            upiImagePath = state.UpiImagePath;
            ppuyoImagePath = state.PpuyoImagePath;
            UpiImageSource = LoadBitmapFromPath(upiImagePath);
            PpuyoImageSource = LoadBitmapFromPath(ppuyoImagePath);
            RefreshOverallScore();
        }

        private void BuildTableFromState(ViewerScoreAppState state)
        {
            table.Clear();
            table.Columns.Clear();

            foreach (ColumnState column in state.Columns)
            {
                if (column.Name == TotalColumn)
                {
                    continue;
                }

                Type type = column.DataType == "int" ? typeof(int) : typeof(string);
                DataColumn createdColumn = table.Columns.Add(column.Name, type);
                createdColumn.Caption = string.IsNullOrWhiteSpace(column.Header) ? column.Name : column.Header;
            }

            EnsureRequiredColumns();

            DataColumn totalColumn = table.Columns.Contains(TotalColumn)
                ? table.Columns[TotalColumn]!
                : table.Columns.Add(TotalColumn, typeof(int));
            totalColumn.Expression = $"[{GreetingColumn}] + [{ReactionColumn}] + [{ActionColumn}] + [{AttendanceColumn}] + [{ExtraColumn}]";

            foreach (RowState rowState in state.Rows)
            {
                DataRow row = table.NewRow();
                foreach (DataColumn column in table.Columns)
                {
                    if (column.ColumnName == TotalColumn)
                    {
                        continue;
                    }

                    if (!rowState.Values.TryGetValue(column.ColumnName, out string? value))
                    {
                        row[column] = column.DataType == typeof(int) ? 0 : string.Empty;
                        continue;
                    }

                    row[column] = column.DataType == typeof(int) ? ParseInt(value ?? "0") : value ?? string.Empty;
                }

                table.Rows.Add(row);
            }

            RenumberRows();
            RefreshGridBinding();
        }

        private void EnsureRequiredColumns()
        {
            AddMissingColumn(NumberColumn, typeof(int), 0);
            AddMissingColumn(ViewerColumn, typeof(string), string.Empty);
            AddMissingColumn(GreetingColumn, typeof(int), 0);
            AddMissingColumn(ReactionColumn, typeof(int), 0);
            AddMissingColumn(ActionColumn, typeof(int), 0);
            AddMissingColumn(AttendanceColumn, typeof(int), 0);
            AddMissingColumn(ExtraColumn, typeof(int), 0);
            AddMissingColumn(MemoColumn, typeof(string), string.Empty);
        }

        private void AddMissingColumn(string columnName, Type type, object defaultValue)
        {
            if (table.Columns.Contains(columnName))
            {
                return;
            }

            DataColumn column = table.Columns.Add(columnName, type);
            column.DefaultValue = defaultValue;
            column.Caption = columnName;
        }

        private BitmapImage? SelectImageFile()
        {
            var dialog = new OpenFileDialog
            {
                Filter = "이미지 파일|*.png;*.jpg;*.jpeg;*.bmp;*.gif;*.webp",
                DefaultExt = ".png"
            };

            if (dialog.ShowDialog() != true)
            {
                return null;
            }

            return LoadBitmapFromPath(dialog.FileName);
        }

        private static BitmapImage? LoadBitmapFromPath(string? path)
        {
            if (string.IsNullOrWhiteSpace(path) || !File.Exists(path))
            {
                return null;
            }

            var image = new BitmapImage();
            image.BeginInit();
            image.CacheOption = BitmapCacheOption.OnLoad;
            image.UriSource = new Uri(path);
            image.EndInit();
            image.Freeze();
            return image;
        }

        private string? GetSelectedColumnName()
        {
            if (TableGrid.CurrentCell.Column?.SortMemberPath is string name && !string.IsNullOrWhiteSpace(name))
            {
                return name;
            }

            if (TableGrid.SelectedCells.Count > 0)
            {
                return TableGrid.SelectedCells[0].Column.SortMemberPath;
            }

            return null;
        }

        private DataRowView? GetCurrentRowView()
        {
            if (TableGrid.SelectedItem is DataRowView selectedRow)
            {
                return selectedRow;
            }

            if (TableGrid.CurrentItem is DataRowView currentRow)
            {
                return currentRow;
            }

            if (TableGrid.CurrentCell.Item is DataRowView currentCellRow)
            {
                return currentCellRow;
            }

            if (TableGrid.SelectedCells.Count > 0 && TableGrid.SelectedCells[0].Item is DataRowView selectedCellRow)
            {
                return selectedCellRow;
            }

            return null;
        }

        private static bool IsProtectedColumn(string columnName)
        {
            return columnName is NumberColumn
                or ViewerColumn
                or GreetingColumn
                or ReactionColumn
                or ActionColumn
                or AttendanceColumn
                or ExtraColumn
                or TotalColumn
                or MemoColumn;
        }

        private static Style BuildTotalCellStyle()
        {
            var style = new Style(typeof(TextBlock));
            style.Setters.Add(new Setter(TextBlock.HorizontalAlignmentProperty, HorizontalAlignment.Center));
            style.Setters.Add(new Setter(TextBlock.FontWeightProperty, FontWeights.SemiBold));
            style.Setters.Add(new Setter(TextBlock.ForegroundProperty, new Binding(TotalColumn)
            {
                Converter = new ScoreToBrushConverter()
            }));
            return style;
        }

        private void EnsurePpuyoImageColumn()
        {
            const string headerText = "뿌요";

            for (int i = TableGrid.Columns.Count - 1; i >= 0; i--)
            {
                if (Equals(TableGrid.Columns[i].Header, headerText))
                {
                    TableGrid.Columns.RemoveAt(i);
                }
            }

            var imageFactory = new FrameworkElementFactory(typeof(Image));
            imageFactory.SetBinding(Image.SourceProperty, new Binding(nameof(PpuyoImageSource))
            {
                Source = this
            });
            imageFactory.SetValue(Image.WidthProperty, 26.0);
            imageFactory.SetValue(Image.HeightProperty, 26.0);
            imageFactory.SetValue(Image.StretchProperty, Stretch.Uniform);
            imageFactory.SetValue(Image.HorizontalAlignmentProperty, HorizontalAlignment.Center);
            imageFactory.SetValue(Image.VerticalAlignmentProperty, VerticalAlignment.Center);

            var template = new DataTemplate
            {
                VisualTree = imageFactory
            };

            var column = new DataGridTemplateColumn
            {
                Header = headerText,
                Width = new DataGridLength(54),
                IsReadOnly = true,
                CellTemplate = template
            };

            int targetIndex = TableGrid.Columns.Count > 0 ? 1 : 0;
            TableGrid.Columns.Insert(targetIndex, column);
        }

        private void RefreshGridBinding()
        {
            DataView currentView = table.DefaultView;
            TableGrid.ItemsSource = null;
            TableGrid.Columns.Clear();
            TableGrid.ItemsSource = currentView;
            EnsurePpuyoImageColumn();
        }

        private void UpdateStatus(string message)
        {
            StatusTextBlock.Text = message;
        }

        public event PropertyChangedEventHandler? PropertyChanged;

        private bool SetField<T>(ref T field, T value, [CallerMemberName] string? propertyName = null)
        {
            if (EqualityComparer<T>.Default.Equals(field, value))
            {
                return false;
            }

            field = value;
            OnPropertyChanged(propertyName);
            return true;
        }

        private void OnPropertyChanged([CallerMemberName] string? propertyName = null)
        {
            PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(propertyName));
        }
    }

    public class ViewerScoreAppState
    {
        public List<ColumnState> Columns { get; set; } = new();
        public List<RowState> Rows { get; set; } = new();
        public string? UpiImagePath { get; set; }
        public string? PpuyoImagePath { get; set; }
    }

    public class ColumnState
    {
        public string Name { get; set; } = string.Empty;
        public string? Header { get; set; }
        public string DataType { get; set; } = "string";
    }

    public class RowState
    {
        public Dictionary<string, string> Values { get; set; } = new();
    }

    public class ScoreToBrushConverter : IValueConverter
    {
        public object Convert(object value, Type targetType, object parameter, CultureInfo culture)
        {
            int score = 0;

            if (value is int intValue)
            {
                score = intValue;
            }
            else if (value is string text && int.TryParse(text, out int parsed))
            {
                score = parsed;
            }

            if (score >= 18)
            {
                return new SolidColorBrush(Color.FromRgb(13, 148, 74));
            }

            if (score >= 10)
            {
                return new SolidColorBrush(Color.FromRgb(11, 92, 173));
            }

            if (score >= 5)
            {
                return new SolidColorBrush(Color.FromRgb(202, 138, 4));
            }

            return new SolidColorBrush(Color.FromRgb(185, 28, 28));
        }

        public object ConvertBack(object value, Type targetType, object parameter, CultureInfo culture)
        {
            throw new NotSupportedException();
        }
    }
}
