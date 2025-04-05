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
using System.Globalization;
using System.Text.RegularExpressions;
using System.Windows.Input;

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
            if (KitchenItems.Count >= 18)
            {
                MessageBox.Show("Нельзя добавить больше 18 строк.");
                return;
            }

            KitchenItems.Add(new KitchenItem { Number = KitchenItems.Count + 1 });
        }

        private void DeleteRowButton_Click(object sender, RoutedEventArgs e)
        {
            var selectedItem = KitchenDataGrid.SelectedItem as KitchenItem;
            if (selectedItem != null)
            {
                KitchenItems.Remove(selectedItem);
            }
            else
            {
                MessageBox.Show("Выберите строку для удаления.");
            }
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

                    //Экспорт значений из элементов формы
                    worksheet.Cell("A6").Value = OrganizationTextBox.Text;        // (A90) Название организации
                    worksheet.Cell("A8").Value = DepartmentTextBox.Text;          // (B90) Структурное подразделение
                    worksheet.Cell("AQ14").Value = DocumentNumberTextBox.Text;      // (C90) Номер документа
                    if (ApprovalDatePicker.SelectedDate is DateTime date)
                    {
                        var culture = new CultureInfo("ru-RU");
                        worksheet.Cell("BK17").Value = date.Day.ToString("00");               // (D90) День
                        worksheet.Cell("BM17").Value = date.ToString("MMMM", culture);        // (E90) Месяц текстом (например, "апрель")
                        worksheet.Cell("BU17").Value = date.Year;                             // (F90) Год
                    }
                    worksheet.Cell("AY14").Value = ApprovalDatePicker2.SelectedDate?.ToString("dd.MM.yyyy");  // (E90) Дата составления акта

                    worksheet.Cell("BQ6").Value = OkpoTextBox.Text;                // (F90) Код по ОКПО
                    worksheet.Cell("BQ9").Value = OkdpTextBox.Text;                // (G90) Вид деятельности по ОКДП
                    worksheet.Cell("BQ10").Value = OperationTextBox.Text;           // (H90) Вид операции

                    worksheet.Cell("BJ13").Value = PositionTextBox.Text;            // (I90) Должность руководителя

                    worksheet.Cell("A58").Value = ReceivedRubTextBox.Text;         // (J90) Получено за приготовление (руб.)
                    worksheet.Cell("BP58").Value = ReceivedKopTextBox.Text;         // (K90) Получено за приготовление (коп.)
                    worksheet.Cell("AE59").Value = TotalRubTextBox.Text;            // (L90) Итого за день (руб.)
                    worksheet.Cell("BP59").Value = TotalKopTextBox.Text;            // (M90) Итого за день (коп.)

                    worksheet.Cell("V62").Value = SpicesPercentTextBox.Text;
                    worksheet.Cell("AK62").Value = SpicesRubTextBox.Text;           // (N90) Специи (руб.)
                    worksheet.Cell("BC62").Value = SpicesKopTextBox.Text;           // (O90) Специи (коп.)

                    worksheet.Cell("U64").Value = SaltPercentTextBox.Text;
                    worksheet.Cell("AK64").Value = SaltRubTextBox.Text;             // (P90) Соль (руб.)
                    worksheet.Cell("BC64").Value = SaltKopTextBox.Text;             // (Q90) Соль (коп.)

                    // Суммируем
                    int spicesRub = 0;
                    int spicesKop = 0;
                    int saltRub = 0;
                    int saltKop = 0;

                    int.TryParse(SpicesRubTextBox.Text, out spicesRub);
                    int.TryParse(SpicesKopTextBox.Text, out spicesKop);
                    int.TryParse(SaltRubTextBox.Text, out saltRub);
                    int.TryParse(SaltKopTextBox.Text, out saltKop);

                    int totalKop = spicesKop + saltKop;
                    int totalRub = spicesRub + saltRub + (totalKop / 100);
                    totalKop = totalKop % 100;

                    // Записываем
                    worksheet.Cell("AK66").Value = totalRub; // Итого руб.
                    worksheet.Cell("BC66").Value = totalKop; // Итого коп.


                    // Комиссия
                    worksheet.Cell("A73").Value = JobTitleComisionTextBox.Text;        // (A70) Должность члена комиссии

                    // Касса
                    worksheet.Cell("I76").Value = CashRubTextBox.Text;                 // (B72) руб.
                    worksheet.Cell("BR76").Value = CashKopTextBox.Text;                 // (C72) коп.

                    // Приложения
                    worksheet.Cell("I80").Value = InvoiceNumberTextBox.Text;           // (B74) Накладная №
                    worksheet.Cell("BH80").Value = DailySumTextBoxRub.Text;             // (C74) сумма руб.
                    worksheet.Cell("BR80").Value = DailySumTextBoxCop.Text;             // (D74) сумма коп.
                    worksheet.Cell("L82").Value = SheetNumberTextBox.Text;             // (B75) Заборный лист №


                    // Прямой доступ к ItemsSource
                    var itemsSource = KitchenDataGrid.ItemsSource as IEnumerable<KitchenItem>;
                    if (itemsSource != null)
                    {
                        var items = itemsSource.Take(18).ToList(); // максимум 18 строк

                        for (int i = 0; i < items.Count; i++)
                        {
                            int targetRow = (i < 11) ? 27 + i : 47 + (i - 11);
                            var item = items[i];

                            worksheet.Cell($"A{targetRow}").Value = item.Number;
                            worksheet.Cell($"E{targetRow}").Value = item.Name;
                            worksheet.Cell($"P{targetRow}").Value = item.Code;

                            worksheet.Cell($"S{targetRow}").Value = item.Price;
                            worksheet.Cell($"S{targetRow}").Style.NumberFormat.Format = "#,##0.00 ₽";

                            worksheet.Cell($"X{targetRow}").Value = item.QuantityNal;
                            worksheet.Cell($"AB{targetRow}").Value = item.SumNal;
                            worksheet.Cell($"AB{targetRow}").Style.NumberFormat.Format = "#,##0.00 ₽";

                            worksheet.Cell($"AG{targetRow}").Value = item.QuantityBufet;
                            worksheet.Cell($"AK{targetRow}").Value = item.SumBufet;
                            worksheet.Cell($"AK{targetRow}").Style.NumberFormat.Format = "#,##0.00 ₽";

                            worksheet.Cell($"AP{targetRow}").Value = item.QuantityOrg;
                            worksheet.Cell($"AT{targetRow}").Value = item.SumOrg;
                            worksheet.Cell($"AT{targetRow}").Style.NumberFormat.Format = "#,##0.00 ₽";

                            worksheet.Cell($"AY{targetRow}").Value = item.QuantityTotal;
                            worksheet.Cell($"BC{targetRow}").Value = item.SumTotal;
                            worksheet.Cell($"BC{targetRow}").Style.NumberFormat.Format = "#,##0.00 ₽";

                            worksheet.Cell($"BG{targetRow}").Value = item.AccountingPrice;
                            worksheet.Cell($"BG{targetRow}").Style.NumberFormat.Format = "#,##0.00 ₽";

                            worksheet.Cell($"BK{targetRow}").Value = item.AccountingSum;
                            worksheet.Cell($"BK{targetRow}").Style.NumberFormat.Format = "#,##0.00 ₽";

                            worksheet.Cell($"BO{targetRow}").Value = item.PricePrice;
                            worksheet.Cell($"BO{targetRow}").Style.NumberFormat.Format = "#,##0.00 ₽";

                            worksheet.Cell($"BT{targetRow}").Value = item.PriceSum;
                            worksheet.Cell($"BT{targetRow}").Style.NumberFormat.Format = "#,##0.00 ₽";
                        }

                    }

                    try
                    {
                        workbook.SaveAs(saveFileDialog.FileName);
                        MessageBox.Show("Файл успешно сохранён:\n" + saveFileDialog.FileName, "Успех", MessageBoxButton.OK, MessageBoxImage.Information);
                    }
                    catch (IOException ex)
                    {
                        MessageBox.Show($"Не удалось сохранить файл. Он может быть открыт в другой программе или заблокирован.\n{ex.Message}", "Ошибка", MessageBoxButton.OK, MessageBoxImage.Error);
                    }

                }

                //MessageBox.Show("Файл успешно сохранён:\n" + saveFileDialog.FileName, "Успех", MessageBoxButton.OK, MessageBoxImage.Information);
            }
        }

        private void ExportToExcel_Click(object sender, RoutedEventArgs e)
        {
            ExportTextBoxesToExcel();
        }

        private void DecimalTextBox_PreviewTextInput(object sender, TextCompositionEventArgs e)
        {
            // Проверка на ввод допустимого десятичного числа (с точкой)
            var textBox = sender as TextBox;
            string fullText = textBox.Text.Insert(textBox.SelectionStart, e.Text);

            e.Handled = !IsValidDecimalInput(fullText);
        }

        private bool IsValidDecimalInput(string input)
        {
            // Разрешаем до 2 знаков после запятой, точка — разделитель
            return Regex.IsMatch(input, @"^\d*\.?\d{0,2}$");
        }

        private void OnlyNumbers_PreviewTextInput(object sender, TextCompositionEventArgs e)
        {
            // Разрешаем только цифры и запятую (для десятичных)
            e.Handled = !Regex.IsMatch(e.Text, @"^[0-9,]+$");
        }

        private void Percentage_PreviewTextInput(object sender, TextCompositionEventArgs e)
        {
            // Ввод только цифр
            e.Handled = !Regex.IsMatch(e.Text, "^[0-9]+$");
        }

        private void Percentage_LostFocus(object sender, RoutedEventArgs e)
        {
            if (sender is TextBox tb)
            {
                if (int.TryParse(tb.Text, out int value))
                {
                    if (value < 0 || value > 100)
                        tb.Text = "100"; // Автоматически ограничиваем
                }
                else
                {
                    tb.Text = "0"; // Если не число — сбрасываем
                }
            }
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
        public double PriceSum { get; set; }

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
