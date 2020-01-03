using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Excel = Microsoft.Office.Interop.Excel;
using System.Reflection;

namespace DTE
{
    public static class ExcelInterface
    {
        private static Excel.Application _exApp;
        private static Excel.Workbooks _exBooks;
        private static Excel.Workbook _exBook;

        private static string _templatePath { get; set; }

        public static void SetTemplatePath(string templatePath)
        {
            _templatePath = templatePath;
        }

        public static void OpenExcelApplication()
        {
            _exApp = new Excel.Application();
            
            _exBooks = _exApp.Workbooks;

            _exBook = _exBooks.Open(_templatePath);
        }

        public static void MakeVisible()
        {
            _exApp.Visible = true;
        }

        public static void EditCellValue(int col, int row, string value)
        {
            _exApp.Cells[row, col] = value;
        }

        public static void EditCellValueInRange(int colIdxStart, int rowIdxStart, int colIdxEnd, int rowIdxEnd, string value)
        {
            _exApp.Range[_exApp.Cells[rowIdxStart, colIdxStart], _exApp.Cells[rowIdxEnd, colIdxEnd]] = value;
        }

        public static void EditCellColorInRange(int colIdxStart, int rowIdxStart, int colIdxEnd, int rowIdxEnd, Excel.XlRgbColor value)
        {
            var cellRange = _exApp.Range[_exApp.Cells[rowIdxStart, colIdxStart], _exApp.Cells[rowIdxEnd, colIdxEnd]];

            cellRange.Interior.Color = value;
        }


        public static void EditCellColor(int col, int row, Excel.XlRgbColor value)
        {

        }
    }
}
