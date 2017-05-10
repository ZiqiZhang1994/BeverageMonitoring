using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Excel = Microsoft.Office.Interop.Excel;
using System.IO;
using System.Reflection;

namespace MyProject
{
    class ExcelExport
    {
        public static void AutomateExcel(List<Data> dataList)
        {
            AutomateExcelImpl(dataList);


            // Clean up the unmanaged Excel COM resources by forcing a garbage 
            // collection as soon as the calling function is off the stack (at 
            // which point these objects are no longer rooted).

            GC.Collect();
            GC.WaitForPendingFinalizers();
            // GC needs to be called twice in order to get the Finalizers called 
            // - the first time in, it simply makes a list of what is to be 
            // finalized, the second time in, it actually is finalizing. Only 
            // then will the object do its automatic ReleaseComObject.
            GC.Collect();
            GC.WaitForPendingFinalizers();
        }

        private static void AutomateExcelImpl(List<Data> dataList)
        {
            object missing = Type.Missing;

            try
            {
                // Create an instance of Microsoft Excel and make it invisible.

                Excel.Application oXL = new Excel.Application();
                oXL.Visible = false;
                Console.WriteLine("Excel.Application is started");

                // Create a new Workbook.

                Excel.Workbook oWB = oXL.Workbooks.Add(missing);
                Console.WriteLine("A new workbook is created");

                // Get the active Worksheet and set its name.

                Excel.Worksheet oSheet = oWB.ActiveSheet as Excel.Worksheet;
                oSheet.Name = "Report";
                Console.WriteLine("The active worksheet is renamed as Report");

                // Fill data into the worksheet's cells.

                Console.WriteLine("Filling data into the worksheet ...");

                // Set the column header
                oSheet.Cells[1, 1] = "Brand";
                oSheet.Cells[1, 2] = "ProductName";
                oSheet.Cells[1, 3] = "PackageSize";
                oSheet.Cells[1, 4] = "Quantity";
                oSheet.Cells[1, 5] = "Litre Price";
                oSheet.Cells[1, 6] = "IfSale";
                oSheet.Cells[1, 7] = "Saved Price";
                oSheet.Cells[1, 8] = "Current Price";

                //List insert
                int rowIndex = 2;
                foreach(Data a in dataList)
                {
                    int colIndex = 1;
                    oSheet.Cells[rowIndex, colIndex] = a.Brand;
                    colIndex++;
                    oSheet.Cells[rowIndex, colIndex] = a.Title;
                    colIndex++;
                    oSheet.Cells[rowIndex, colIndex] = a.Size;
                    colIndex++;
                    oSheet.Cells[rowIndex, colIndex] = a.Quantity;
                    colIndex++;
                    oSheet.Cells[rowIndex, colIndex] = a.PerL;
                    colIndex++;
                    oSheet.Cells[rowIndex, colIndex] = a.IfSale;
                    colIndex++;
                    oSheet.Cells[rowIndex, colIndex] = a.SavedPrice;
                    colIndex++;
                    oSheet.Cells[rowIndex, colIndex] = a.Price;
                    rowIndex++;
                }


                //// Construct an array of user names
                //string[,] saNames = new string[,] {
                //{"John", "Smith"},
                //{"Tom", "Brown"},
                //{"Sue", "Thomas"},
                //{"Jane", "Jones"},
                //{"Adam", "Johnson"}};

                //// Fill A2:B6 with an array of values (First and Last Names).
                //oSheet.get_Range("A2", "B6").Value2 = saNames;

                //// Fill C2:C6 with a relative formula (=A2 & " " & B2).
                //oSheet.get_Range("C2", "C6").Formula = "=A2 & \" \" & B2";

                // Save the workbook as a xlsx file and close it.

                Console.WriteLine("Save and close the workbook");

                string fileName = Path.GetDirectoryName(
                    Assembly.GetExecutingAssembly().Location) + "\\Test.xlsx";
                oWB.SaveAs(fileName, Excel.XlFileFormat.xlOpenXMLWorkbook,
                    missing, missing, missing, missing,
                    Excel.XlSaveAsAccessMode.xlNoChange,
                    missing, missing, missing, missing, missing);
                oWB.Close(missing, missing, missing);

                // Quit the Excel application.

                Console.WriteLine("Quit the Excel application");

                // Excel will stick around after Quit if it is not under user 
                // control and there are outstanding references. When Excel is 
                // started or attached programmatically and 
                // Application.Visible = false, Application.UserControl is false. 
                // The UserControl property can be explicitly set to True which 
                // should force the application to terminate when Quit is called, 
                // regardless of outstanding references.
                oXL.UserControl = true;

                oXL.Quit();
            }
            catch (Exception ex)
            {
                Console.WriteLine("ExcelExport.AutomateExcel throws the error: {0}",
                    ex.Message);
            }
        }
    }
}
