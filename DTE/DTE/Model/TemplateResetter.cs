using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Excel = Microsoft.Office.Interop.Excel;

namespace DTE
{
    public static class TemplateResetter
    {
        public static int days { get; set; }
        public static int numberEmployees { get; set; }

        //changing the position of the first employee requires this to be modified
        private const int STARTING_ROW = 26;
        private const int STARTING_COL = 5;

        private const int CELLS_IN_DAY = 5;

        private static DateTime date;
        private static Month currentMonth;
        private static int currentYear;

        public static void ResetEmployeeTableToWhite()
        {
            //clear every cell in editable area
            ExcelInterface.EditCellValueInRange(
                STARTING_COL, STARTING_ROW-2, 
                STARTING_COL + 30, STARTING_ROW + (numberEmployees * CELLS_IN_DAY) - 1, 
                string.Empty);

            //color every cell in editable area white
            ExcelInterface.EditCellColorInRange(
                 STARTING_COL, STARTING_ROW-2, 
                 STARTING_COL + 30, STARTING_ROW + (numberEmployees * CELLS_IN_DAY) - 1,
                 Excel.XlRgbColor.rgbWhite);

            //color every other employee slighty orange for contrast
            for (int i = STARTING_ROW + CELLS_IN_DAY; i < STARTING_ROW + (numberEmployees * CELLS_IN_DAY); i += 2 * CELLS_IN_DAY)
            {
                ExcelInterface.EditCellColorInRange(
                    2, i,
                    STARTING_COL + 30, i + CELLS_IN_DAY -1,
                    Excel.XlRgbColor.rgbFloralWhite);
            }

            //grey color for unused days
            //skip if month has 31 days
            if (days >= 31) return;
            ExcelInterface.EditCellColorInRange(
                 STARTING_COL + days, STARTING_ROW - 2, 
                 STARTING_COL + 30, STARTING_ROW + (numberEmployees * CELLS_IN_DAY) - 1,
                 Excel.XlRgbColor.rgbDarkGrey);
        }
       

        public static void SpecifyDate(int month, int year)
        {
            currentMonth = (Month)month;
            currentYear = year;

            date = new DateTime(currentYear, (int)currentMonth, 1);

            days = DateTime.DaysInMonth(date.Year, date.Month);
        }

        public static void FillInWeekdays()
        {
            ExcelInterface.EditCellValue(20, 1, currentMonth.ToString() + " " + currentYear);


            for(int i = STARTING_COL; i <= days + STARTING_COL - 1; i++)
            {
                //day of the month
                string dayOfMonth = i - STARTING_COL + 1 + "";

                ExcelInterface.EditCellValue(i, STARTING_ROW - 1, dayOfMonth + ".");

                for(int k = STARTING_ROW + 4; k < STARTING_ROW + (numberEmployees * CELLS_IN_DAY); k += 5)
                {
                    ExcelInterface.EditCellValue(i, k, dayOfMonth + '.');
                }

                Weekday day = Weekday.EMPTY;
                date = new DateTime(currentYear, (int)currentMonth, i - STARTING_COL + 1);

                switch (date.DayOfWeek)
                {
                    case DayOfWeek.Monday:
                        day = Weekday.Mo;
                        break;
                    case DayOfWeek.Tuesday:
                        day = Weekday.Di;
                        break;
                    case DayOfWeek.Wednesday:
                        day = Weekday.Mi;
                        break;
                    case DayOfWeek.Thursday:
                        day = Weekday.Do;
                        break;
                    case DayOfWeek.Friday:
                        day = Weekday.Fr;
                        break;
                    case DayOfWeek.Saturday:
                        day = Weekday.Sa;
                        ExcelInterface.EditCellColorInRange(
                            i, STARTING_ROW - 2,
                            i, STARTING_ROW + (numberEmployees * CELLS_IN_DAY) - 1,
                            Excel.XlRgbColor.rgbPaleTurquoise);
                        break;
                    case DayOfWeek.Sunday:
                        day = Weekday.So;
                        ExcelInterface.EditCellColorInRange(
                            i, STARTING_ROW - 2,
                            i, STARTING_ROW + (numberEmployees * CELLS_IN_DAY) - 1,
                            Excel.XlRgbColor.rgbPaleTurquoise);
                        break;
                    default: break;
                }

                ExcelInterface.EditCellValue(i, STARTING_ROW - 2, day.ToString() + ".");
            }
        }

        public enum Month
        {
            EMPTY,
            Januar,
            Februar,
            März,
            April,
            Mai,
            Juni,
            Juli,
            August,
            September,
            Oktober,
            November,
            Dezember
        }

        public enum Weekday
        {
            Mo,
            Di,
            Mi,
            Do,
            Fr,
            Sa,
            So,
            EMPTY
        }
    }
}
