using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Navigation;
using System.Windows.Shapes;

namespace DTE.View
{
    /// <summary>
    /// Interaktionslogik für MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {
        private string _templatePath = AppDomain.CurrentDomain.BaseDirectory + "template.xls";

        public MainWindow()
        {
            InitializeComponent();

            yearText.Text = "" + (DateTime.Today.Month < 12 ? DateTime.Today.Year : DateTime.Today.Year + 1);
            monthCombo.SelectedIndex = DateTime.Today.Month % 12;
        }

        private void ComboBox_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            if (yearText == null) return;
            TemplateResetter.SpecifyDate(monthCombo.SelectedIndex + 1, int.Parse(yearText.Text));
        }

        private void NumberValidationTextBox(object sender, TextCompositionEventArgs e)
        {
            if (e.Text == null) e.Handled = false;
            e.Handled = IsNumber(e.Text);
        }

        private bool IsNumber(string text)
        {
            Regex regex = new Regex("[^0-9]+");
            return regex.IsMatch(text);
        }

        private void YearText_TextChanged(object sender, TextChangedEventArgs e)
        {
            if (yearText == null || yearText.Text.Equals("")) return;
            TemplateResetter.SpecifyDate(monthCombo.SelectedIndex + 1, int.Parse(yearText.Text));
        }

        private void Button_Click(object sender, RoutedEventArgs e)
        {
            ExcelInterface.SetTemplatePath(_templatePath);
            ExcelInterface.OpenExcelApplication();
            TemplateResetter.ResetEmployeeTableToWhite();
            TemplateResetter.FillInWeekdays();
            ExcelInterface.MakeVisible();
            ExcelInterface.CollectGarbage();
        }

        private void EmployeeNumberText_TextChanged(object sender, TextChangedEventArgs e)
        {
            if (employeeNumberText == null || employeeNumberText.Text.Equals("")) return;
            TemplateResetter.numberEmployees = int.Parse(employeeNumberText.Text);
        }

        private void Button_Click_1(object sender, RoutedEventArgs e)
        {
            if (employeeNumberText == null) return;
            int employeeNumber = int.Parse(employeeNumberText.Text);
            employeeNumber--;
            employeeNumberText.Text = "" + employeeNumber;
        }

        private void Button_Click_2(object sender, RoutedEventArgs e)
        {
            if (employeeNumberText == null) return;
            int employeeNumber = int.Parse(employeeNumberText.Text);
            employeeNumber++;
            employeeNumberText.Text = "" + employeeNumber;
        }
    }
}
