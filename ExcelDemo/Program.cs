using Microsoft.Office.Interop.Excel;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading.Tasks;

namespace ExcelDemo
{
    /// <summary>
    /// http://csharp.net-informations.com/excel/csharp-format-excel.htm
    /// </summary>
    class Program
    {
        static void Main(string[] args)
        {
            Example2();
        }

        private static void Example1()
        { 
            Microsoft.Office.Interop.Excel.Application xlApp = new Microsoft.Office.Interop.Excel.Application();

            if (xlApp == null)
            {
                Console.WriteLine("Excel is not properly installed!!");
                return;
            }


            Microsoft.Office.Interop.Excel.Workbook xlWorkBook;
            Microsoft.Office.Interop.Excel.Worksheet xlWorkSheet1;
            object misValue = System.Reflection.Missing.Value;

            xlWorkBook = xlApp.Workbooks.Add(misValue);
            xlWorkSheet1 = (Microsoft.Office.Interop.Excel.Worksheet)xlWorkBook.Worksheets.get_Item(1);
            xlWorkSheet1.Name = "Sheet1";

            xlWorkSheet1.Cells[1, 1] = "ID";
            xlWorkSheet1.Cells[1, 2] = "Name";
            xlWorkSheet1.Cells[2, 1] = "1";
            xlWorkSheet1.Cells[2, 2] = "One";
            xlWorkSheet1.Cells[3, 1] = "2";
            xlWorkSheet1.Cells[3, 2] = "Two";

            xlWorkBook.SaveAs("c:\\csharp-Excel.xls", Microsoft.Office.Interop.Excel.XlFileFormat.xlWorkbookNormal, misValue, misValue, misValue, misValue, Microsoft.Office.Interop.Excel.XlSaveAsAccessMode.xlExclusive, misValue, misValue, misValue, misValue, misValue);
            xlWorkBook.Close(true, misValue, misValue);
            xlApp.Quit();

            Marshal.ReleaseComObject(xlWorkSheet1);
            Marshal.ReleaseComObject(xlWorkBook);
            Marshal.ReleaseComObject(xlApp);
        }

        private static void Example2()
        {
            try
            {
                Microsoft.Office.Interop.Excel.Application xlApp = new Microsoft.Office.Interop.Excel.Application();

                if (xlApp == null)
                {
                    Console.WriteLine("Excel is not properly installed!!");
                    return;
                }

                Application ExcelApp = new Application();
                Workbook ExcelWorkBook = null;
                Worksheet ExcelWorkSheet = null;
                //ExcelApp.Visible = true;

                //SheetNames
                List<string> SheetNames = new List<string> {"Autoprofi","Autofit","Meisterhaft","Castrol","AutoCheck" };
                //Columns headers
                List<string> headers = new List<string>
                {
                    "Nr.",
                    "Werkstattname (hat KKZ)",
                    "Im AppManagement registriert",
                    "In der App sichtbar",
                    "Anzahl Kunden (gesamt)",
                    "Anzahl neue Kunden (aktueller Monat)",
                    "Veränderung Kunden (zum Vormonat)",
                    "Anzahl Termine (gesamt)",
                    "Anzahl Termine (aktueller Monat)",
                    "Veränderung Termine (zum Vormonat)"
                };

                //GarageList
                List<string> garageList = new List<string> { "Autohaus Gerlach", "AUTOHAUS WELBERS", "Autoservice Penderak", "BRUNKE", "D & D CAR CENTER" };


                ExcelWorkBook = ExcelApp.Workbooks.Add(XlWBATemplate.xlWBATWorksheet);
                for (int i = 1; i < SheetNames.Count; i++)
                    ExcelWorkBook.Worksheets.Add(); //Adding New sheet in Excel Workbook

                

                for (int i = 0; i < SheetNames.Count; i++)
                {
                    ExcelWorkSheet = ExcelWorkBook.Worksheets[i + 1 ];
                    ExcelWorkSheet.Name = SheetNames[i];

                    //Adding Custom header to the excel file
                    //ExcelWorkSheet.Cells[1, 1] = "September 2017";
                    ExcelWorkSheet.get_Range("a1", "i1").Merge(false);
                    Range chartRange = ExcelWorkSheet.get_Range("a1", "i1");
                    chartRange.FormulaR1C1 = "September 2017";
                    chartRange.HorizontalAlignment = 3;
                    chartRange.VerticalAlignment = 3;
                    chartRange.Font.Size = 14;
                    chartRange.Font.Bold = true;

                    //initialize Excel row start position = 1;
                    int r = 4;


                    //writing columns name un excel sheet
                    for (int col = 1; col < headers.Count; col++)
                    {
                        ExcelWorkSheet.Cells[r, col] = headers[col-1];
                    }
                    r++;

                    //Set header background color
                    ExcelWorkSheet.get_Range("a4", "i4").Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.LightGray);

                    //Bold entire row
                    ExcelWorkSheet.get_Range("a4").EntireRow.Font.Bold = true;

                    //Writing rows into Excel sheet
                    for (int row = 0; row < garageList.Count; row++)
                    {
                        ExcelWorkSheet.Cells[r, 1] = row + 1;
                        ExcelWorkSheet.Cells[r, 2] = garageList[row];
                        r++;
                    }

                    ExcelWorkSheet.Cells[r, 1] = "Summe";
                    ExcelWorkSheet.get_Range($"a{r}").EntireRow.Font.Bold = true;
                    //Set footer background color
                    ExcelWorkSheet.get_Range($"a{r}", $"i{r}").Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.LightGray);



                    //Border around multiple cells in excel
                    Range borderRange = ExcelWorkSheet.get_Range("a4", $"i{r}");
                    borderRange.BorderAround(XlLineStyle.xlContinuous, XlBorderWeight.xlMedium, XlColorIndex.xlColorIndexAutomatic, XlColorIndex.xlColorIndexAutomatic);
                    //adjust column width
                    ExcelWorkSheet.Columns.AutoFit();
                }

                ExcelWorkBook.SaveAs($"c:\\csharp-Excel.xlsx");
                ExcelWorkBook.Close();
                ExcelApp.Quit();

                releaseObject(ExcelWorkSheet);
                releaseObject(ExcelWorkBook);
                releaseObject(ExcelApp);
            }
            catch (Exception)
            {

            }
            finally
            {

                foreach (Process process in Process.GetProcessesByName("Excel"))
                    process.Kill();
            }
        }

        private static void releaseObject(object obj)
        {
            try
            {
                System.Runtime.InteropServices.Marshal.ReleaseComObject(obj);
                obj = null;
            }
            catch (Exception ex)
            {
                obj = null;
                Console.WriteLine("Exception Occured while releasing object " + ex.ToString());
            }
            finally
            {
                GC.Collect();
            }
        }
    }
}
