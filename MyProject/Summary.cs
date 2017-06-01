using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace MyProject
{
    class Summary
    {
        //sumary writer
        public static void SpecialProductSummary()
        {
            //get data from woolworths excel file
            List<Data_WWs> dataList_WWs = new List<Data_WWs>();
            dataList_WWs = ExcelExport.ReadExcel();

            //get data from coles excel file
            List<Data> dataList = new List<Data>();
            dataList = ExcelExport.ReadExcel_Coles();

            int specialProductNum = 0;
            int specialProductNum_Coles = 0;

            //special product list
            foreach (Data_WWs a in dataList_WWs)
            {
                if (a.IfSale == "Y")
                {
                    specialProductNum++;
                    Console.WriteLine("Product: " + a.Title + " Sale Price: $" + a.Price + " You can save: $" + a.SavedPrice);
                }
            }

            foreach (Data b in dataList)
            {
                if (b.IfSale == "Y")
                {
                    specialProductNum_Coles++;
                    Console.WriteLine("Brand: " + b.Brand +" Product: " + b.Title + " Sale Price: $" + b.Price + " You can save: $" + b.SavedPrice);
                }
            }

            //summary
            Console.WriteLine("Current Total Special Product (Coles) Number is " + specialProductNum_Coles.ToString());
            Console.WriteLine("Current Total Special Product (WoolWorths) Number is " + specialProductNum.ToString());
            Console.WriteLine("Current Total Special Product (Coles and WoolWorths) Number is " + (specialProductNum + specialProductNum_Coles).ToString());
            Console.WriteLine("Current Total Product (Coles) Number is " + dataList.Count.ToString());
            Console.WriteLine("Current Total Product (WoolWorths) Number is " + dataList_WWs.Count.ToString());
            Console.WriteLine("Current Total Product (Coles and WoolWorths) Number is " + (dataList.Count + dataList_WWs.Count).ToString());

        }
    }
}
