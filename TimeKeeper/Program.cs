
using System;
using System.IO;
using Microsoft.Office.Interop.Excel;

namespace TimeKeeper
{
    class Program
    {
        static void Main(string[] args)
        {
            string workbookPath = Directory.GetCurrentDirectory() + @"/" + Properties.Settings.Default.TemplateFileName;
            Workbook excelWorkbook = null;
            try
            {
                Console.WriteLine("Loading Template...");
                excelWorkbook = new Application().Workbooks.Open(workbookPath, 0, false, 5, "", "", false, XlPlatform.xlWindows, "", true, false, 0, true, false, false);

                Sheets excelSheets = excelWorkbook.Worksheets;
                Worksheet excelWorksheet = excelSheets[Properties.Settings.Default.SheetNumber];

                //Write dates
                Console.WriteLine("Writing file...");
                int year = DateTime.Today.Year;
                int month = DateTime.Today.Month;
                int startday = 0;
                int endDay = 0;
                if (DateTime.Today.Day < 6)
                {
                    month--;
                    startday = 21;
                    endDay = 5;
                }
                else if (DateTime.Today.Day > 20)
                {
                    month++;
                    startday = 21;
                    endDay = 5;
                }
                else
                {
                    startday = 6;
                    endDay = 20;
                }

                Console.WriteLine("Writing Employee Details...");
                (excelWorksheet.Cells[1, 1] as Range).Value = Properties.Settings.Default.EmployeeNumber;
                (excelWorksheet.Cells[1, 2] as Range).Value = Properties.Settings.Default.FamilyName + ", " + Properties.Settings.Default.FirstName;
                DateTime currentDay = new DateTime(year, month, startday);
                string labelStartDate = currentDay.ToString("yyyy-MM-dd");
                for (int currentRow = 3; currentDay.Day != endDay + 1; currentRow++)
                {
                    Console.WriteLine("Writing " + currentDay.ToShortDateString());
                    (excelWorksheet.Cells[currentRow, 1] as Range).Value = currentDay.ToShortDateString();
                    if (!currentDay.DayOfWeek.Equals(DayOfWeek.Saturday) && !currentDay.DayOfWeek.Equals(DayOfWeek.Sunday))
                    {
                        (excelWorksheet.Cells[currentRow, 3] as Range).Value = "REGULAR WORK";
                        (excelWorksheet.Cells[currentRow, 4] as Range).Value = Properties.Settings.Default.TimeIn;
                    }
                    else
                    {
                        (excelWorksheet.Cells[currentRow, 3] as Range).Value = "REST DAY"; 
                    }
                    currentDay = currentDay.AddDays(1);
                }
                string labelEndDate = currentDay.AddDays(-1).ToString("yyyy-MM-dd");

                Console.WriteLine("Saving new file...");
                excelWorkbook.SaveAs(GenerateFileName(labelStartDate, labelEndDate));
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.Message);
            }

            if (excelWorkbook != null)
            {
                excelWorkbook.Close();
            }
            Console.Write("Completed! Press any key to exit...");
            Console.ReadKey();
            
        }

        static string GenerateFileName(string labelStartDate, string labelEndDate)
        {
            string directoryPath = string.IsNullOrWhiteSpace(Properties.Settings.Default.CustomSavePath) ?
                System.Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments) :
                Properties.Settings.Default.CustomSavePath;

            return directoryPath + 
                @"/" + Properties.Settings.Default.FamilyName + 
                "_VPI Timesheet - " + 
                labelStartDate + " to " + 
                labelEndDate + "_" + 
                DateTime.Now.ToFileTime() + ".xlsm";
        }
    }
}
