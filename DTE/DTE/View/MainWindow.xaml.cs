using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
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

            //ExcelInterface.SetTemplatePath(_templatePath);
            //ExcelInterface.OpenExcelApplication();

            //TemplateResetter.SpecifyDate(2, 2020);

            //TemplateResetter.numberEmployees = 11;
            //TemplateResetter.ResetEmployeeTableToWhite();

            //TemplateResetter.FillInWeekdays();
            
            //ExcelInterface.MakeVisible();
        }

        private void Button_Click(object sender, RoutedEventArgs e)
        {

        }
    }
}
