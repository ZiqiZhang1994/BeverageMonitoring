using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Excel = Microsoft.Office.Interop.Excel;
using System.IO;
using System.Reflection;
using System.Diagnostics;

namespace MyProject
{
    class ExcelExport
    {

        //Coles extraction data save controller
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

        //Coles demo extraction data save controller
        public static void AutomateExcel_Test(List<Data> dataList)
        {
            AutomateExcelImpl_Test(dataList);


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

        //Woolworths test extraction data save controller
        public static void AutomateExcel_WWsTest(List<Data_WWs> dataList)
        {
            AutomateExcelImpl_WWsTest(dataList);


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

        //Woolworths extraction data save controller
        public static void AutomateExcel_WWs(List<Data_WWs> dataList)
        {
            AutomateExcelImpl_WWs(dataList);


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

        //Coles extraction data save
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
                oSheet.Cells[1, 9] = "Original Price";
                oSheet.Cells[1, 10] = "Barcode";
                oSheet.Cells[1, 11] = "Page Link";

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
                    colIndex++;
                    oSheet.Cells[rowIndex, colIndex] = a.OriginalPrice;
                    colIndex++;
                    oSheet.Cells[rowIndex, colIndex] = a.Barcode;
                    colIndex++;
                    oSheet.Cells[rowIndex, colIndex] = a.Link;
                    rowIndex++;
                }

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

        //Coles demo extraction data save
        private static void AutomateExcelImpl_Test(List<Data> dataList)
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
                oSheet.Cells[1, 9] = "Original Price";
                oSheet.Cells[1, 10] = "Barcode";
                oSheet.Cells[1, 11] = "Page Link";

                //List insert
                int rowIndex = 2;
                foreach (Data a in dataList)
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
                    colIndex++;
                    oSheet.Cells[rowIndex, colIndex] = a.OriginalPrice;
                    colIndex++;
                    oSheet.Cells[rowIndex, colIndex] = a.Barcode;
                    colIndex++;
                    oSheet.Cells[rowIndex, colIndex] = a.Link;
                    rowIndex++;
                }

                // Save the workbook as a xlsx file and close it.

                Console.WriteLine("Save and close the workbook");

                string fileName = Path.GetDirectoryName(
                    Assembly.GetExecutingAssembly().Location) + "\\Demo_Coles.xlsx";
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

        //Woolworths extraction data save
        private static void AutomateExcelImpl_WWs(List<Data_WWs> dataList)
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
                oSheet.Cells[1, 1] = "ProductName";
                oSheet.Cells[1, 2] = "PackageSize";
                oSheet.Cells[1, 3] = "Quantity";
                oSheet.Cells[1, 4] = "Litre Price";
                oSheet.Cells[1, 5] = "IfSale";
                oSheet.Cells[1, 6] = "Saved Price";
                oSheet.Cells[1, 7] = "Current Price";
                oSheet.Cells[1, 8] = "Original Price";
                oSheet.Cells[1, 9] = "Quantity Special Sale";
                oSheet.Cells[1, 10] = "Product Image Link";


                //List insert
                int rowIndex = 2;
                foreach (Data_WWs a in dataList)
                {
                    int colIndex = 1;
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
                    colIndex++;
                    oSheet.Cells[rowIndex, colIndex] = a.OriginalPrice;
                    colIndex++;
                    oSheet.Cells[rowIndex, colIndex] = a.QuantitySpecial;
                    colIndex++;
                    oSheet.Cells[rowIndex, colIndex] = a.ImageUrl;
                    rowIndex++;
                }

                // Save the workbook as a xlsx file and close it.

                Console.WriteLine("Save and close the workbook");

                string fileName = Path.GetDirectoryName(
                    Assembly.GetExecutingAssembly().Location) + "\\Test_WWs.xlsx";
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

        //Woolworths test extraction data save
        private static void AutomateExcelImpl_WWsTest(List<Data_WWs> dataList)
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
                oSheet.Cells[1, 1] = "ProductName";
                oSheet.Cells[1, 2] = "PackageSize";
                oSheet.Cells[1, 3] = "Quantity";
                oSheet.Cells[1, 4] = "Litre Price";
                oSheet.Cells[1, 5] = "IfSale";
                oSheet.Cells[1, 6] = "Saved Price";
                oSheet.Cells[1, 7] = "Current Price";
                oSheet.Cells[1, 8] = "Original Price";
                oSheet.Cells[1, 9] = "Quantity Special Sale";
                oSheet.Cells[1, 10] = "Product Image Link";


                //List insert
                int rowIndex = 2;
                foreach (Data_WWs a in dataList)
                {
                    int colIndex = 1;
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
                    colIndex++;
                    oSheet.Cells[rowIndex, colIndex] = a.OriginalPrice;
                    colIndex++;
                    oSheet.Cells[rowIndex, colIndex] = a.QuantitySpecial;
                    colIndex++;
                    oSheet.Cells[rowIndex, colIndex] = a.ImageUrl;
                    rowIndex++;
                }

                // Save the workbook as a xlsx file and close it.

                Console.WriteLine("Save and close the workbook");

                string fileName = Path.GetDirectoryName(
                    Assembly.GetExecutingAssembly().Location) + "\\Demo_WWs.xlsx";
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

        public static List<Data_WWs> ReadExcel()
        {
            List<Data_WWs> dataList = new List<Data_WWs>();
            string fileName = Path.GetDirectoryName(
                    Assembly.GetExecutingAssembly().Location) + "\\Test_WWs.xlsx";
            object missing = System.Reflection.Missing.Value;
            Excel.Application excel = new Excel.Application();//lauch excel application
            if (excel == null)
            {
                Console.WriteLine("alert('Can't access excel')");
            }
            else
            {
                excel.Visible = false; excel.UserControl = true;
                // open file as only read
                Excel.Workbook wb = excel.Application.Workbooks.Open(fileName, missing, true, missing, missing, missing,
                 missing, missing, missing, true, missing, missing, missing, missing, missing);
                //take worksheet 
                Excel.Worksheet ws = (Excel.Worksheet)wb.Worksheets.get_Item(1);


                //take total row (include title)
                int rowsint = ws.UsedRange.Cells.Rows.Count; //take row
                                                             
                //int columnsint = mySheet.UsedRange.Cells.Columns.Count;//take column


                //take range data (not include title) 
                Excel.Range rng1 = ws.Cells.get_Range("A2", "A" + rowsint); 
                Excel.Range rng2 = ws.Cells.get_Range("B2", "B" + rowsint);
                Excel.Range rng3 = ws.Cells.get_Range("C2", "C" + rowsint);
                Excel.Range rng4 = ws.Cells.get_Range("D2", "D" + rowsint);
                Excel.Range rng5 = ws.Cells.get_Range("E2", "E" + rowsint);
                Excel.Range rng6 = ws.Cells.get_Range("F2", "F" + rowsint);
                Excel.Range rng7 = ws.Cells.get_Range("G2", "G" + rowsint);
                Excel.Range rng8 = ws.Cells.get_Range("H2", "H" + rowsint);
                Excel.Range rng9 = ws.Cells.get_Range("I2", "I" + rowsint);
                Excel.Range rng10 = ws.Cells.get_Range("J2", "J" + rowsint);
                object[,] arryName = (object[,])rng1.Value2;   //get range's value
                object[,] arrySize = (object[,])rng2.Value2;   //get range's value
                object[,] arryQuantity = (object[,])rng3.Value2;   //get range's value
                object[,] arryPerL = (object[,])rng4.Value2;   //get range's value
                object[,] arryIfSale = (object[,])rng5.Value2;   //get range's value
                object[,] arrySavedPrice = (object[,])rng6.Value2;   //get range's value
                object[,] arryNowPrice = (object[,])rng7.Value2;   //get range's value
                object[,] arryOriPrice = (object[,])rng8.Value2;   //get range's value
                object[,] arryQS = (object[,])rng9.Value2;   //get range's value
                object[,] arryLink = (object[,])rng10.Value2;   //get range's value
                                                               
                //insert array in list
                
                for (int i = 1; i <= rowsint - 1; i++)
                {
                    Data_WWs data = new Data_WWs();

                    if(arryName[i, 1] != null)
                    {
                        data.Title = arryName[i, 1].ToString();
                    }
                    else
                    {
                        data.Title = "0";
                    }

                    if (arrySize[i, 1] != null)
                    {
                        data.Size = arrySize[i, 1].ToString();
                    }
                    else
                    {
                        data.Size = "0";
                    }

                    if (arryQuantity[i, 1] != null)
                    {
                        data.Quantity = arryQuantity[i, 1].ToString();
                    }
                    else
                    {
                        data.Quantity = "0";
                    }

                    if (arryPerL[i, 1] != null)
                    {
                        data.PerL = arryPerL[i, 1].ToString();
                    }
                    else
                    {
                        data.PerL = "0";
                    }

                    if (arryIfSale[i, 1] != null)
                    {
                        data.IfSale = arryIfSale[i, 1].ToString();
                    }
                    else
                    {
                        data.IfSale = "N";
                    }

                    if (arrySavedPrice[i, 1] != null)
                    {
                        data.SavedPrice = arrySavedPrice[i, 1].ToString();
                    }
                    else
                    {
                        data.SavedPrice = "0";
                    }

                    if (arryNowPrice[i, 1] != null)
                    {
                        data.Price = arryNowPrice[i, 1].ToString();
                    }
                    else
                    {
                        data.Price = "0";
                    }

                    if (arryOriPrice[i, 1] != null)
                    {
                        data.OriginalPrice = arryOriPrice[i, 1].ToString();
                    }
                    else
                    {
                        data.OriginalPrice = "0";
                    }

                    if (arryQS[i, 1] != null)
                    {
                        data.QuantitySpecial = arryQS[i, 1].ToString();
                    }
                    else
                    {
                        data.QuantitySpecial = "0";
                    }

                    if (arryLink[i, 1] != null)
                    {
                        data.ImageUrl = arryLink[i, 1].ToString();
                    }
                    else
                    {
                        data.ImageUrl = "0";
                    }



                    dataList.Add(data);
                    data.ToString();
                }
            }
            excel.Quit(); excel = null;
            Process[] procs = Process.GetProcessesByName("excel");


            foreach (Process pro in procs)
            {
                pro.Kill();//kill process
            }
            GC.Collect();

            return dataList;
        }

        public static List<Data> ReadExcel_Coles()
        {
            List<Data> dataList = new List<Data>();
            string fileName = Path.GetDirectoryName(
                    Assembly.GetExecutingAssembly().Location) + "\\Test.xlsx";
            object missing = System.Reflection.Missing.Value;
            Excel.Application excel = new Excel.Application();//lauch excel application
            if (excel == null)
            {
                Console.WriteLine("alert('Can't access excel')");
            }
            else
            {
                excel.Visible = false; excel.UserControl = true;
                // open file as only read
                Excel.Workbook wb = excel.Application.Workbooks.Open(fileName, missing, true, missing, missing, missing,
                 missing, missing, missing, true, missing, missing, missing, missing, missing);
                //take worksheet 
                Excel.Worksheet ws = (Excel.Worksheet)wb.Worksheets.get_Item(1);


                //take total row (include title)
                int rowsint = ws.UsedRange.Cells.Rows.Count; //take row
                //int columnsint = mySheet.UsedRange.Cells.Columns.Count;//take column


               //take range data (not include title)  
                Excel.Range rng1 = ws.Cells.get_Range("A2", "A" + rowsint);
                Excel.Range rng2 = ws.Cells.get_Range("B2", "B" + rowsint);
                Excel.Range rng3 = ws.Cells.get_Range("C2", "C" + rowsint);
                Excel.Range rng4 = ws.Cells.get_Range("D2", "D" + rowsint);
                Excel.Range rng5 = ws.Cells.get_Range("E2", "E" + rowsint);
                Excel.Range rng6 = ws.Cells.get_Range("F2", "F" + rowsint);
                Excel.Range rng7 = ws.Cells.get_Range("G2", "G" + rowsint);
                Excel.Range rng8 = ws.Cells.get_Range("H2", "H" + rowsint);
                Excel.Range rng9 = ws.Cells.get_Range("I2", "I" + rowsint);
                Excel.Range rng10 = ws.Cells.get_Range("J2", "J" + rowsint);
                Excel.Range rng11 = ws.Cells.get_Range("K2", "K" + rowsint);
                object[,] arryBrand = (object[,])rng1.Value2;   //get range's value
                object[,] arryProductName = (object[,])rng2.Value2;   //get range's value
                object[,] arrySize = (object[,])rng3.Value2;   //get range's value
                object[,] arryQuantity = (object[,])rng4.Value2;   //get range's value
                object[,] arryPerL = (object[,])rng5.Value2;   //get range's value
                object[,] arryIfSale = (object[,])rng6.Value2;   //get range's value
                object[,] arrySavedPrice = (object[,])rng7.Value2;   //get range's value
                object[,] arryNowPrice = (object[,])rng8.Value2;   //get range's value
                object[,] arryOriPrice = (object[,])rng9.Value2;   //get range's value
                object[,] arryCode = (object[,])rng10.Value2;   //get range's value
                object[,] arryLink = (object[,])rng11.Value2;   //get range's value

                //insert array in list

                for (int i = 1; i <= rowsint - 1; i++)
                {
                    Data data = new Data();

                    if (arryBrand[i, 1] != null)
                    {
                        data.Brand = arryBrand[i, 1].ToString();
                    }
                    else
                    {
                        data.Brand = "0";
                    }

                    if (arryProductName[i, 1] != null)
                    {
                        data.Title = arryProductName[i, 1].ToString();
                    }
                    else
                    {
                        data.Title = "0";
                    }

                    if (arrySize[i, 1] != null)
                    {
                        data.Size = arrySize[i, 1].ToString();
                    }
                    else
                    {
                        data.Size = "0";
                    }

                    if (arryQuantity[i, 1] != null)
                    {
                        data.Quantity = arryQuantity[i, 1].ToString();
                    }
                    else
                    {
                        data.Quantity = "0";
                    }

                    if (arryPerL[i, 1] != null)
                    {
                        data.PerL = arryPerL[i, 1].ToString();
                    }
                    else
                    {
                        data.PerL = "0";
                    }

                    if (arryIfSale[i, 1] != null)
                    {
                        data.IfSale = arryIfSale[i, 1].ToString();
                    }
                    else
                    {
                        data.IfSale = "N";
                    }

                    if (arrySavedPrice[i, 1] != null)
                    {
                        data.SavedPrice = arrySavedPrice[i, 1].ToString();
                    }
                    else
                    {
                        data.SavedPrice = "0";
                    }

                    if (arryNowPrice[i, 1] != null)
                    {
                        data.Price = arryNowPrice[i, 1].ToString();
                    }
                    else
                    {
                        data.Price = "0";
                    }

                    if (arryOriPrice[i, 1] != null)
                    {
                        data.OriginalPrice = arryOriPrice[i, 1].ToString();
                    }
                    else
                    {
                        data.OriginalPrice = "0";
                    }

                    if (arryCode[i, 1] != null)
                    {
                        data.Barcode = arryCode[i, 1].ToString();
                    }
                    else
                    {
                        data.Barcode = "0";
                    }

                    if (arryLink[i, 1] != null)
                    {
                        data.Link = arryLink[i, 1].ToString();
                    }
                    else
                    {
                        data.Link = "0";
                    }

                    dataList.Add(data);
                    data.ToString();
                }
            }
            excel.Quit(); excel = null;
            Process[] procs = Process.GetProcessesByName("excel");


            foreach (Process pro in procs)
            {
                pro.Kill();//kill process
            }
            GC.Collect();

            return dataList;
        }

    }
}
