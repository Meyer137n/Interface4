using Microsoft.Win32;
using OP10FormApp;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.ComponentModel;
using System.Linq;
using System.Windows;
using ClosedXML.Excel;
using System.IO;
using System.Windows.Media;
using Excel = Microsoft.Office.Interop.Excel;
using System.Windows.Controls;
using System;
using System.Windows.Data;

namespace KitchenReportForm
{
    public partial class MainWindow : Window
    {
        public ObservableCollection<KitchenItem> KitchenItems { get; set; }

        // Словарь: Название → Код
        public static readonly Dictionary<string, string> NameToCode = new Dictionary<string, string>
        {
            { "Котлета по-киевски", "101" },
            { "Суп борщ", "102" },
            { "Пюре картофельное", "103" },
            { "Омлет с сыром", "104" }
        };

        public static readonly Dictionary<string, string> CodeToName = new Dictionary<string, string>
        {
            { "101", "Котлета по-киевски" },
            { "102", "Суп борщ" },
            { "103", "Пюре картофельное" },
            { "104", "Омлет с сыром" }
        };

        public MainWindow()
        {
            InitializeComponent();

            KitchenItems = new ObservableCollection<KitchenItem>
            {
                new KitchenItem { Number = 1 }
            };

            DataContext = this;
        }

        private void AddRowButton_Click(object sender, RoutedEventArgs e)
        {
            KitchenItems.Add(new KitchenItem { Number = KitchenItems.Count + 1 });
        }

        private void ImportFromExcel_Click(object sender, RoutedEventArgs e)
        {
            var openFileDialog = new OpenFileDialog
            {
                Title = "Выберите файл Excel",
                Filter = "Excel файл (*.xlsx)|*.xlsx"
            };

            if (openFileDialog.ShowDialog() == true)
            {
                var workbook = new XLWorkbook(openFileDialog.FileName);
                var worksheet = workbook.Worksheet(1);

                // Шаг 1: загрузка данных из A90+ в TextBox'ы
                var textBoxes = new List<TextBox>();

                void CollectTextBoxes(DependencyObject parent)
                {
                    int count = VisualTreeHelper.GetChildrenCount(parent);
                    for (int i = 0; i < count; i++)
                    {
                        var child = VisualTreeHelper.GetChild(parent, i);
                        if (child is TextBox tb)
                            textBoxes.Add(tb);
                        else
                            CollectTextBoxes(child);
                    }
                }

                CollectTextBoxes(this);

                int row = 90;
                foreach (var tb in textBoxes)
                {
                    var cellValue = worksheet.Cell($"A{row}").GetString();
                    tb.Text = cellValue;
                    row++;
                }

                // Шаг 2: загрузка строк таблицы из A110+
                int tableRow = 110;
                var items = new List<KitchenItem>();

                while (!worksheet.Cell($"A{tableRow}").IsEmpty())
                {
                    try
                    {
                        var item = new KitchenItem
                        {
                            Number = int.TryParse(worksheet.Cell(tableRow, 1).GetString(), out var num) ? num : 0,
                            Name = worksheet.Cell(tableRow, 2).GetString(),
                            Code = worksheet.Cell(tableRow, 3).GetString(),
                            Price = double.TryParse(worksheet.Cell(tableRow, 4).GetString(), out var p) ? p : 0,
                            QuantityNal = double.TryParse(worksheet.Cell(tableRow, 5).GetString(), out var qn) ? qn : 0,
                            SumNal = double.TryParse(worksheet.Cell(tableRow, 6).GetString(), out var sn) ? sn : 0,
                            QuantityBufet = double.TryParse(worksheet.Cell(tableRow, 7).GetString(), out var qb) ? qb : 0,
                            SumBufet = double.TryParse(worksheet.Cell(tableRow, 8).GetString(), out var sb) ? sb : 0,
                            QuantityOrg = double.TryParse(worksheet.Cell(tableRow, 9).GetString(), out var qo) ? qo : 0,
                            SumOrg = double.TryParse(worksheet.Cell(tableRow, 10).GetString(), out var so) ? so : 0,
                            QuantityTotal = double.TryParse(worksheet.Cell(tableRow, 11).GetString(), out var qt) ? qt : 0,
                            SumTotal = double.TryParse(worksheet.Cell(tableRow, 12).GetString(), out var st) ? st : 0,
                            AccountingPrice = double.TryParse(worksheet.Cell(tableRow, 13).GetString(), out var ap) ? ap : 0,
                            AccountingSum = double.TryParse(worksheet.Cell(tableRow, 14).GetString(), out var asum) ? asum : 0
                        };

                        items.Add(item);
                        tableRow++;
                    }
                    catch
                    {
                        break; // прерываем при ошибке или пустой строке
                    }
                }

                KitchenDataGrid.ItemsSource = items;
            }
        }

        private void ExportTextBoxesToExcel()
        {
            string templatePath = "TemplateFormOP10.xlsx";

            if (!File.Exists(templatePath))
            {
                MessageBox.Show($"Не найден шаблон Excel по пути: {templatePath}", "Ошибка", MessageBoxButton.OK, MessageBoxImage.Error);
                return;
            }

            var textBoxValues = new List<string>();

            void Traverse(DependencyObject parent)
            {
                int count = VisualTreeHelper.GetChildrenCount(parent);
                for (int i = 0; i < count; i++)
                {
                    var child = VisualTreeHelper.GetChild(parent, i);
                    if (child is TextBox tb)
                        textBoxValues.Add(tb.Text);
                    else
                        Traverse(child);
                }
            }

            Traverse(this);

            var saveFileDialog = new SaveFileDialog
            {
                Title = "Сохранить заполненный файл",
                Filter = "Excel файл (*.xlsx)|*.xlsx",
                FileName = "ФормаОП10_заполненная.xlsx"
            };

            if (saveFileDialog.ShowDialog() == true)
            {
                using (var workbook = new XLWorkbook(templatePath))
                {
                    var worksheet = workbook.Worksheet(1);

                    // Запись TextBox'ов в A90+
                    int row = 90;
                    foreach (var value in textBoxValues)
                    {
                        worksheet.Cell($"A{row}").Value = value;
                        row++;
                    }

                    // Прямой доступ к ItemsSource
                    var itemsSource = KitchenDataGrid.ItemsSource as IEnumerable<KitchenItem>;
                    if (itemsSource != null)
                    {
                        var items = itemsSource.Take(18).ToList(); // максимум 18 строк

                        for (int i = 0; i < items.Count; i++)
                        {
                            int targetRow = (i < 11) ? 27 + i : 47 + (i - 11); // строки A27–A37 и A47–A53

                            var item = items[i];

                            worksheet.Cell($"A{targetRow}").Value = item.Number;
                            worksheet.Cell($"E{targetRow}").Value = item.Name;
                            worksheet.Cell($"P{targetRow}").Value = item.Code;
                            worksheet.Cell($"S{targetRow}").Value = item.Price;
                            worksheet.Cell($"X{targetRow}").Value = item.QuantityNal;
                            worksheet.Cell($"AB{targetRow}").Value = item.SumNal;
                            worksheet.Cell($"AG{targetRow}").Value = item.QuantityBufet;
                            worksheet.Cell($"AK{targetRow}").Value = item.SumBufet;
                            worksheet.Cell($"AP{targetRow}").Value = item.QuantityOrg;
                            worksheet.Cell($"AT{targetRow}").Value = item.SumOrg;
                            worksheet.Cell($"AY{targetRow}").Value = item.QuantityTotal;
                            worksheet.Cell($"BC{targetRow}").Value = item.SumTotal;
                            worksheet.Cell($"BG{targetRow}").Value = item.AccountingPrice;
                            worksheet.Cell($"BK{targetRow}").Value = item.AccountingSum;
                            worksheet.Cell($"BO{targetRow}").Value = item.PricePrice;
                            worksheet.Cell($"BT{targetRow}").Value = item.PriceSumm;
                        }
                    }

                    workbook.SaveAs(saveFileDialog.FileName);
                }

                MessageBox.Show("Файл успешно сохранён:\n" + saveFileDialog.FileName, "Успех", MessageBoxButton.OK, MessageBoxImage.Information);
            }
        }

        private void ExportToExcel_Click(object sender, RoutedEventArgs e)
        {
            ExportTextBoxesToExcel();
        }


    }

    public class KitchenItem : INotifyPropertyChanged
    {
        private string _name;
        private string _code;
        private double _price;

        public double QuantityNal { get; set; }
        public double SumNal { get; set; }
        public double QuantityBufet { get; set; }
        public double SumBufet { get; set; }
        public double QuantityOrg { get; set; }
        public double SumOrg { get; set; }
        public double QuantityTotal { get; set; }
        public double SumTotal { get; set; }
        public double AccountingPrice { get; set; }
        public double AccountingSum { get; set; }
        public double PricePrice { get; set; }
        public double PriceSumm { get; set; }

        public int Number { get; set; }

        public string Name
        {
            get => _name;
            set
            {
                if (_name != value)
                {
                    _name = value;
                    OnPropertyChanged(nameof(Name));

                    // Автозаполнение кода
                    if (MainWindow.NameToCode.TryGetValue(_name, out var matchedCode))
                    {
                        _code = matchedCode;
                        OnPropertyChanged(nameof(Code));
                    }
                }
            }
        }

        public string Code
        {
            get => _code;
            set
            {
                if (_code != value)
                {
                    _code = value;
                    OnPropertyChanged(nameof(Code));

                    // Автозаполнение названия
                    if (MainWindow.CodeToName.TryGetValue(_code, out var matchedName))
                    {
                        _name = matchedName;
                        OnPropertyChanged(nameof(Name));
                    }
                }
            }
        }

        public double Price
        {
            get => _price;
            set { _price = value; OnPropertyChanged(nameof(Price)); }
        }

        public event PropertyChangedEventHandler PropertyChanged;
        private void OnPropertyChanged(string propertyName) =>
            PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(propertyName));
    }
}
