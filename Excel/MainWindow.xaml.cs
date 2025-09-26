using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.ComponentModel;
using System.Globalization;
using System.Linq;
using System.Runtime.CompilerServices;
using System.Windows;
using System.Windows.Data;
using ClosedXML.Excel;
using Microsoft.Win32;
using System.Text.RegularExpressions;

namespace Excel
{
    public partial class MainWindow : Window, INotifyPropertyChanged
    {
        private static readonly CultureInfo[] SupportedCultures =
        {
            CultureInfo.CurrentCulture,
            CultureInfo.GetCultureInfo("ru-RU"),
            CultureInfo.InvariantCulture
        };

        private const string DefaultProductName = "Без указания продукции";
        private const decimal AmountChangeThresholdPercent = 5m;

        private static readonly Regex RomanNumeralPrefixRegex = new("^[IVXLCDM]+\\.", RegexOptions.IgnoreCase | RegexOptions.Compiled);
        private static readonly string[] ExcludedLabelPrefixes =
        {
            "итого",
            "в т.ч."
        };

        private string? _firstFilePath;
        private string? _secondFilePath;
        private string _statusMessage = "Выберите файлы и нажмите \"Сравнить\".";
        private IReadOnlyDictionary<string, ProductSummary> _productSummaries = new Dictionary<string, ProductSummary>(StringComparer.OrdinalIgnoreCase);

        public MainWindow()
        {
            InitializeComponent();
            DataContext = this;
            DiffRowsView = CollectionViewSource.GetDefaultView(DiffRows);
            DiffRowsView.GroupDescriptions.Add(new PropertyGroupDescription(nameof(DiffRow.Product)));
            DiffRows.CollectionChanged += (_, __) => OnPropertyChanged(nameof(HasResults));
            ProductSummaries = new Dictionary<string, ProductSummary>(StringComparer.OrdinalIgnoreCase);
        }

        public event PropertyChangedEventHandler? PropertyChanged;

        public ObservableCollection<DiffRow> DiffRows { get; } = new();

        public ICollectionView DiffRowsView { get; }
        public IReadOnlyDictionary<string, ProductSummary> ProductSummaries
        {
            get => _productSummaries;
            private set
            {
                if (!ReferenceEquals(_productSummaries, value))
                {
                    _productSummaries = value;
                    OnPropertyChanged();
                }
            }
        }

        public string? FirstFilePath
        {
            get => _firstFilePath;
            set
            {
                if (_firstFilePath != value)
                {
                    _firstFilePath = value;
                    OnPropertyChanged();
                }
            }
        }

        public string? SecondFilePath
        {
            get => _secondFilePath;
            set
            {
                if (_secondFilePath != value)
                {
                    _secondFilePath = value;
                    OnPropertyChanged();
                }
            }
        }

        public bool HasResults => DiffRows.Count > 0;

        public string StatusMessage
        {
            get => _statusMessage;
            set
            {
                if (_statusMessage != value)
                {
                    _statusMessage = value;
                    OnPropertyChanged();
                }
            }
        }

        private void BrowseFirstFile_Click(object sender, RoutedEventArgs e)
        {
            var dialog = CreateOpenFileDialog();
            if (dialog.ShowDialog() == true)
            {
                FirstFilePath = dialog.FileName;
            }
        }

        private void BrowseSecondFile_Click(object sender, RoutedEventArgs e)
        {
            var dialog = CreateOpenFileDialog();
            if (dialog.ShowDialog() == true)
            {
                SecondFilePath = dialog.FileName;
            }
        }

        private void CompareButton_Click(object sender, RoutedEventArgs e)
        {
            if (string.IsNullOrWhiteSpace(FirstFilePath) || string.IsNullOrWhiteSpace(SecondFilePath))
            {
                MessageBox.Show("Выберите оба файла перед сравнением.", "Внимание", MessageBoxButton.OK, MessageBoxImage.Warning);
                return;
            }

            try
            {
                StatusMessage = "Выполняется сравнение...";
                var firstRows = LoadExcelRows(FirstFilePath);
                var secondRows = LoadExcelRows(SecondFilePath);
                var summaries = CalculateProductSummaries(firstRows, secondRows);
                var visibleProducts = new HashSet<string>(
                    summaries
                        .Where(kvp => kvp.Value.ShouldDisplay(AmountChangeThresholdPercent))
                        .Select(kvp => kvp.Key),
                    StringComparer.OrdinalIgnoreCase);

                var differences = BuildDifferences(firstRows, secondRows)
                    .Where(row => visibleProducts.Contains(row.Product))
                    .ToList();

                ProductSummaries = visibleProducts.Count == 0
                    ? new Dictionary<string, ProductSummary>(StringComparer.OrdinalIgnoreCase)
                    : visibleProducts.ToDictionary(
                        product => product,
                        product => summaries[product],
                        StringComparer.OrdinalIgnoreCase);

                DiffRows.Clear();
                foreach (var row in differences)
                {
                    DiffRows.Add(row);
                }

                DiffRowsView.Refresh();

                if (DiffRows.Count == 0)
                {
                    StatusMessage = "Различий не обнаружено.";
                }
                else
                {
                    var changed = DiffRows.Count(r => r.HasDifference);
                    StatusMessage = changed == 0
                        ? "Данные идентичны."
                        : $"Найдено {changed} строк с отличиями.";
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Не удалось сравнить файлы. {ex.Message}", "Ошибка", MessageBoxButton.OK, MessageBoxImage.Error);
                StatusMessage = "Ошибка при сравнении.";
            }
        }

        private void ExportButton_Click(object sender, RoutedEventArgs e)
        {
            if (!HasResults)
            {
                MessageBox.Show("Сначала выполните сравнение.", "Внимание", MessageBoxButton.OK, MessageBoxImage.Information);
                return;
            }

            var dialog = new SaveFileDialog
            {
                Filter = "Excel (*.xlsx)|*.xlsx",
                FileName = "Различия.xlsx"
            };

            if (dialog.ShowDialog() != true)
            {
                return;
            }

            try
            {
                using var workbook = new XLWorkbook();
                var worksheet = workbook.Worksheets.Add("Различия");

                var headers = new[]
                {
                    "Позиция",
                    "Кол-во (старое)",
                    "Кол-во (новое)",
                    "Δ Кол-во",
                    "Цена (старая)",
                    "Цена (новая)",
                    "Δ Цена",
                    "Сумма (старая)",
                    "Сумма (новая)",
                    "Δ Сумма",
                    "Δ %",
                    "Статус"
                };

                for (var i = 0; i < headers.Length; i++)
                {
                    worksheet.Cell(1, i + 1).Value = headers[i];
                    worksheet.Cell(1, i + 1).Style.Font.Bold = true;
                }
                for (var column = 2; column <= 10; column++)
                {
                    worksheet.Column(column).Style.NumberFormat.Format = "0.00";
                }

                worksheet.Column(11).Style.NumberFormat.Format = "0.00\\%";
                var rowIndex = 2;
                string? currentProduct = null;

                foreach (var diff in DiffRows)
                {
                    if (!string.Equals(currentProduct, diff.Product, StringComparison.OrdinalIgnoreCase))
                    {
                        if (currentProduct != null)
                        {
                            rowIndex++;
                        }

                        currentProduct = diff.Product;
                        var summaryRange = worksheet.Range(rowIndex, 1, rowIndex, headers.Length);
                        summaryRange.Style.Font.Bold = true;
                        summaryRange.Style.Fill.BackgroundColor = XLColor.FromHtml("#EFEFEF");

                        var nameRange = worksheet.Range(rowIndex, 1, rowIndex, 7);
                        nameRange.Merge();
                        nameRange.Value = currentProduct;

                        if (ProductSummaries.TryGetValue(currentProduct, out var summary))
                        {
                            worksheet.Cell(rowIndex, 8).Value = summary.OldAmount;
                            worksheet.Cell(rowIndex, 9).Value = summary.NewAmount;
                            worksheet.Cell(rowIndex, 10).Value = summary.AmountDelta;
                            worksheet.Cell(rowIndex, 11).Value = summary.AmountDeltaPercent;
                        }
                        rowIndex++;
                    }

                    worksheet.Cell(rowIndex, 1).Value = diff.Item;
                    worksheet.Cell(rowIndex, 2).Value = diff.OldQuantity;
                    worksheet.Cell(rowIndex, 3).Value = diff.NewQuantity;
                    worksheet.Cell(rowIndex, 4).Value = diff.QuantityDelta;
                    worksheet.Cell(rowIndex, 5).Value = diff.OldPrice;
                    worksheet.Cell(rowIndex, 6).Value = diff.NewPrice;
                    worksheet.Cell(rowIndex, 7).Value = diff.PriceDelta;
                    worksheet.Cell(rowIndex, 8).Value = diff.OldAmount;
                    worksheet.Cell(rowIndex, 9).Value = diff.NewAmount;
                    worksheet.Cell(rowIndex, 10).Value = diff.AmountDelta;
                    worksheet.Cell(rowIndex, 11).Value = diff.AmountDeltaPercent;
                    worksheet.Cell(rowIndex, 12).Value = diff.StatusText;
                    rowIndex++;
                }

                worksheet.Columns().AdjustToContents();
                workbook.SaveAs(dialog.FileName);

                StatusMessage = $"Файл с результатами сохранен: {dialog.FileName}";
                MessageBox.Show("Файл успешно сохранен.", "Готово", MessageBoxButton.OK, MessageBoxImage.Information);
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Не удалось сохранить файл. {ex.Message}", "Ошибка", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        private static OpenFileDialog CreateOpenFileDialog() => new()
        {
            Filter = "Excel файлы (*.xlsx;*.xlsm;*.xltx;*.xltm)|*.xlsx;*.xlsm;*.xltx;*.xltm|Все файлы (*.*)|*.*"
        };
        private static IReadOnlyDictionary<string, ProductSummary> CalculateProductSummaries(List<ExcelRow> firstRows, List<ExcelRow> secondRows)
        {
            var products = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
            foreach (var row in firstRows)
            {
                products.Add(row.Product);
            }

            foreach (var row in secondRows)
            {
                products.Add(row.Product);
            }

            var result = new Dictionary<string, ProductSummary>(StringComparer.OrdinalIgnoreCase);
            foreach (var product in products)
            {
                var oldAmount = SumAmounts(firstRows, product);
                var newAmount = SumAmounts(secondRows, product);
                result[product] = new ProductSummary(product, oldAmount, newAmount);
            }

            return result;
        }

        private static decimal SumAmounts(IEnumerable<ExcelRow> rows, string product)
        {
            return rows
                .Where(r => string.Equals(r.Product, product, StringComparison.OrdinalIgnoreCase))
                .Where(r => !ShouldExcludeRow(r.BaseLabel))
                .Sum(r => r.Amount ?? 0m);
        }
        private static List<DiffRow> BuildDifferences(List<ExcelRow> firstRows, List<ExcelRow> secondRows)
        {
            var result = new List<DiffRow>();

            var firstProducts = firstRows
                .GroupBy(r => r.Product, StringComparer.OrdinalIgnoreCase)
                .ToDictionary(g => g.Key, g => g.OrderBy(r => r.RowNumber).ToList(), StringComparer.OrdinalIgnoreCase);

            var secondProducts = secondRows
                .GroupBy(r => r.Product, StringComparer.OrdinalIgnoreCase)
                .ToDictionary(g => g.Key, g => g.OrderBy(r => r.RowNumber).ToList(), StringComparer.OrdinalIgnoreCase);

            var allProducts = new HashSet<string>(firstProducts.Keys, StringComparer.OrdinalIgnoreCase);
            allProducts.UnionWith(secondProducts.Keys);

            var orderedProducts = allProducts
                .Select(product => new
                {
                    Product = product,
                    Order = Math.Min(
                        firstProducts.TryGetValue(product, out var firstList) && firstList.Count > 0 ? firstList[0].RowNumber : int.MaxValue,
                        secondProducts.TryGetValue(product, out var secondList) && secondList.Count > 0 ? secondList[0].RowNumber : int.MaxValue)
                })
                .OrderBy(item => item.Order)
                .ThenBy(item => item.Product, StringComparer.OrdinalIgnoreCase)
                .Select(item => item.Product)
                .ToList();

            foreach (var product in orderedProducts)
            {
                firstProducts.TryGetValue(product, out var firstList);
                secondProducts.TryGetValue(product, out var secondList);

                var firstByLabel = (firstList ?? new List<ExcelRow>())
                    .GroupBy(r => r.BaseLabel, StringComparer.OrdinalIgnoreCase)
                    .ToDictionary(g => g.Key, g => g.OrderBy(r => r.RowNumber).ToList(), StringComparer.OrdinalIgnoreCase);

                var secondByLabel = (secondList ?? new List<ExcelRow>())
                    .GroupBy(r => r.BaseLabel, StringComparer.OrdinalIgnoreCase)
                    .ToDictionary(g => g.Key, g => g.OrderBy(r => r.RowNumber).ToList(), StringComparer.OrdinalIgnoreCase);

                var allKeys = new HashSet<string>(firstByLabel.Keys, StringComparer.OrdinalIgnoreCase);
                allKeys.UnionWith(secondByLabel.Keys);

                var keyOrder = allKeys
                    .Select(key => new
                    {
                        Key = key,
                        Order = Math.Min(
                            firstByLabel.TryGetValue(key, out var firstRowsForKey) && firstRowsForKey.Count > 0 ? firstRowsForKey[0].RowNumber : int.MaxValue,
                            secondByLabel.TryGetValue(key, out var secondRowsForKey) && secondRowsForKey.Count > 0 ? secondRowsForKey[0].RowNumber : int.MaxValue)
                    })
                    .OrderBy(item => item.Order)
                    .ThenBy(item => item.Key, StringComparer.OrdinalIgnoreCase)
                    .Select(item => item.Key)
                    .ToList();

                foreach (var key in keyOrder)
                {
                    firstByLabel.TryGetValue(key, out var firstRowsForKey);
                    secondByLabel.TryGetValue(key, out var secondRowsForKey);

                    var maxCount = Math.Max(firstRowsForKey?.Count ?? 0, secondRowsForKey?.Count ?? 0);

                    for (var i = 0; i < maxCount; i++)
                    {
                        var firstRow = firstRowsForKey != null && i < firstRowsForKey.Count ? firstRowsForKey[i] : null;
                        var secondRow = secondRowsForKey != null && i < secondRowsForKey.Count ? secondRowsForKey[i] : null;
                        if (ShouldExcludeRow(firstRow?.BaseLabel ?? secondRow?.BaseLabel ?? key))
                        {
                            continue;
                        }
                        var display = firstRow?.DisplayLabel ?? secondRow?.DisplayLabel ?? (i == 0 ? key : $"{key} ({i + 1})");
                        result.Add(new DiffRow(product, display, firstRow, secondRow));
                    }
                }
            }

            return result;
        }
        private static bool ShouldExcludeRow(string? label)
        {
            if (string.IsNullOrWhiteSpace(label))
            {
                return false;
            }

            var trimmed = label.Trim();
            var lower = trimmed.ToLowerInvariant();

            if (lower.Contains("итого"))
            {
                return true;
            }

            foreach (var prefix in ExcludedLabelPrefixes)
            {
                if (lower.StartsWith(prefix))
                {
                    return true;
                }
            }

            return RomanNumeralPrefixRegex.IsMatch(trimmed) && trimmed.Contains(':');
        }
        private static List<ExcelRow> LoadExcelRows(string path)
        {
            var rows = new List<ExcelRow>();

            using var workbook = new XLWorkbook(path);
            var worksheet = workbook.Worksheets.First();
            var usedRange = worksheet.RangeUsed();
            if (usedRange == null)
            {
                return rows;
            }

            var keyCounters = new Dictionary<string, int>(StringComparer.OrdinalIgnoreCase);
            string? currentProduct = null;

            foreach (var row in usedRange.Rows())
            {
                var worksheetRow = row.WorksheetRow();
                var baseLabel = GetRowLabel(worksheetRow);

                var numericCells = new List<(IXLCell Cell, decimal Value)>();
                foreach (var cell in row.Cells())
                {
                    if (TryGetDecimal(cell, out var value))
                    {
                        numericCells.Add((cell, value));
                    }
                }

                if (TryExtractProductName(worksheetRow, out var productName))
                {
                    currentProduct = productName;
                    if (numericCells.Count == 0)
                    {
                        continue;
                    }
                }

                if (numericCells.Count == 0)
                {
                    continue;
                }

                if (string.IsNullOrWhiteSpace(baseLabel))
                {
                    baseLabel = $"Строка {row.RowNumber()}";
                }

                var effectiveProductName = currentProduct ?? DefaultProductName;
                var counterKey = $"{effectiveProductName}\u001F{baseLabel}";
                keyCounters.TryGetValue(counterKey, out var index);
                keyCounters[counterKey] = index + 1;

                var displayLabel = index == 0 ? baseLabel : $"{baseLabel} ({index + 1})";

                var orderedValues = numericCells
                    .OrderBy(item => item.Cell.Address.ColumnNumber)
                    .Select(item => (decimal?)item.Value)
                    .ToList();

                rows.Add(new ExcelRow(
                    effectiveProductName,
                    baseLabel,
                    displayLabel,
                    row.RowNumber(),
                    orderedValues.ElementAtOrDefault(0),
                    orderedValues.ElementAtOrDefault(1),
                    orderedValues.ElementAtOrDefault(2)));
            }

            return rows;
        }

        private static string GetRowLabel(IXLRow row)
        {
            var lastColumn = row.LastCellUsed()?.Address.ColumnNumber ?? row.CellCount();
            for (var column = 1; column <= lastColumn; column++)
            {
                var cell = row.Cell(column);
                var text = cell.GetString().Trim();
                if (string.IsNullOrEmpty(text))
                {
                    continue;
                }

                if (TryGetDecimal(cell, out _))
                {
                    continue;
                }

                return text;
            }

            return string.Empty;
        }

        private static bool TryExtractProductName(IXLRow row, out string productName)
        {
            productName = string.Empty;
            var lastColumn = row.LastCellUsed()?.Address.ColumnNumber ?? row.CellCount();

            for (var column = 1; column <= lastColumn; column++)
            {
                var cell = row.Cell(column);
                var text = cell.GetString().Trim();
                if (string.IsNullOrEmpty(text))
                {
                    continue;
                }

                if (text.IndexOf("продукц", StringComparison.OrdinalIgnoreCase) < 0)
                {
                    continue;
                }

                var valuePart = ExtractProductNameFromCellText(text);
                if (!string.IsNullOrWhiteSpace(valuePart))
                {
                    productName = valuePart.Trim();
                    return true;
                }

                for (var nextColumn = column + 1; nextColumn <= lastColumn; nextColumn++)
                {
                    var nextText = row.Cell(nextColumn).GetString().Trim();
                    if (string.IsNullOrEmpty(nextText))
                    {
                        continue;
                    }

                    productName = nextText;
                    return true;
                }
            }

            productName = string.Empty;
            return false;
        }

        private static string? ExtractProductNameFromCellText(string text)
        {
            var separators = new[] { ':', '-', '—', '–' };

            foreach (var separator in separators)
            {
                var index = text.IndexOf(separator);
                if (index < 0)
                {
                    continue;
                }

                var candidate = text[(index + 1)..].Trim();
                if (!string.IsNullOrWhiteSpace(candidate))
                {
                    return candidate;
                }
            }

            var lines = text.Split(new[] { '\n', '\r' }, StringSplitOptions.RemoveEmptyEntries);
            if (lines.Length > 1)
            {
                var lastLine = lines[^1].Trim();
                if (!string.IsNullOrWhiteSpace(lastLine) && lastLine.IndexOf("продукц", StringComparison.OrdinalIgnoreCase) < 0)
                {
                    return lastLine;
                }
            }

            return null;
        }

        private static bool TryGetDecimal(IXLCell cell, out decimal value)
        {
            if (cell.DataType == XLDataType.Number)
            {
                value = cell.GetValue<decimal>();
                return true;
            }

            var text = cell.GetString().Trim();
            if (string.IsNullOrWhiteSpace(text))
            {
                value = 0m;
                return false;
            }

            var normalized = text.Replace("\u00A0", string.Empty).Replace(" ", string.Empty);

            foreach (var culture in SupportedCultures)
            {
                if (decimal.TryParse(normalized, NumberStyles.Any, culture, out value))
                {
                    return true;
                }
            }

            value = 0m;
            return false;
        }

        private void OnPropertyChanged([CallerMemberName] string? propertyName = null)
        {
            PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(propertyName));
        }
    }

    public enum DiffStatus
    {
        Unchanged,
        Added,
        Removed,
        Changed
    }

    public sealed class DiffRow
    {
        public DiffRow(string product, string item, ExcelRow? oldRow, ExcelRow? newRow)
        {
            Product = product;
            Item = item;
            OldQuantity = DecimalHelper.Round2(oldRow?.Quantity);
            NewQuantity = DecimalHelper.Round2(newRow?.Quantity);
            OldPrice = DecimalHelper.Round2(oldRow?.Price);
            NewPrice = DecimalHelper.Round2(newRow?.Price);
            OldAmount = DecimalHelper.Round2(oldRow?.Amount);
            NewAmount = DecimalHelper.Round2(newRow?.Amount);

            QuantityDelta = CalculateDelta(NewQuantity, OldQuantity);
            PriceDelta = CalculateDelta(NewPrice, OldPrice);
            AmountDelta = CalculateDelta(NewAmount, OldAmount);
            AmountDeltaPercent = DecimalHelper.CalculatePercentChange(NewAmount, OldAmount);

            Status = oldRow == null && newRow != null
                ? DiffStatus.Added
                : oldRow != null && newRow == null
                    ? DiffStatus.Removed
                    : HasValueDifference()
                        ? DiffStatus.Changed
                        : DiffStatus.Unchanged;
        }

        public string Product { get; }
        public string Item { get; }
        public decimal? OldQuantity { get; }
        public decimal? NewQuantity { get; }
        public decimal? QuantityDelta { get; }
        public decimal? OldPrice { get; }
        public decimal? NewPrice { get; }
        public decimal? PriceDelta { get; }
        public decimal? OldAmount { get; }
        public decimal? NewAmount { get; }
        public decimal? AmountDelta { get; }
        public decimal? AmountDeltaPercent { get; }
        public DiffStatus Status { get; }
        public string StatusText => Status switch
        {
            DiffStatus.Added => "Добавлено",
            DiffStatus.Removed => "Удалено",
            DiffStatus.Changed => "Изменено",
            _ => "Без изменений"
        };

        public bool HasDifference => Status != DiffStatus.Unchanged;

        private bool HasValueDifference()
        {
            return !AreEqual(OldQuantity, NewQuantity)
                   || !AreEqual(OldPrice, NewPrice)
                   || !AreEqual(OldAmount, NewAmount);
        }

        private static decimal? CalculateDelta(decimal? newValue, decimal? oldValue)
        {
            if (newValue.HasValue && oldValue.HasValue)
            {
                return DecimalHelper.Round2(newValue.Value - oldValue.Value);
            }

            return null;
        }

        private static bool AreEqual(decimal? left, decimal? right)
        {
            if (!left.HasValue && !right.HasValue)
            {
                return true;
            }

            if (!left.HasValue || !right.HasValue)
            {
                return false;
            }

            return DecimalHelper.IsEffectivelyZero(left.Value - right.Value);
        }
    }

    public sealed class ExcelRow
    {
        public ExcelRow(string product, string baseLabel, string displayLabel, int rowNumber, decimal? quantity, decimal? price, decimal? amount)
        {
            Product = product;
            BaseLabel = baseLabel;
            DisplayLabel = displayLabel;
            RowNumber = rowNumber;
            Quantity = quantity;
            Price = price;
            Amount = amount;
        }

        public string Product { get; }
        public string BaseLabel { get; }
        public string DisplayLabel { get; }
        public int RowNumber { get; }
        public decimal? Quantity { get; }
        public decimal? Price { get; }
        public decimal? Amount { get; }
    }
}
