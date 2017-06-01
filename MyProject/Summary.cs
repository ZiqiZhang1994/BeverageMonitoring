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
            double totalSalePrice = 0;
            double totalSalePrice_WWs = 0;
            //special product list
            foreach (Data_WWs a in dataList_WWs)
            {
                totalSalePrice_WWs = totalSalePrice_WWs + Convert.ToDouble(a.Price);
                if (a.IfSale == "Y")
                {
                    specialProductNum++;
                    Console.WriteLine("***********************************************************");
                    Console.WriteLine("Product: " + a.Title + " Sale Price: $" + a.Price + " You can save: $" + a.SavedPrice);
                }
            }

            foreach (Data b in dataList)
            {
                totalSalePrice = totalSalePrice + Convert.ToDouble(b.Price);
                if (b.IfSale == "Y")
                {
                    specialProductNum_Coles++;
                    Console.WriteLine("***********************************************************");
                    Console.WriteLine("Brand: " + b.Brand +" Product: " + b.Title + " Sale Price: $" + b.Price + " You can save: $" + b.SavedPrice);
                }
            }


            //calculation
            double totalAVG = (totalSalePrice_WWs + totalSalePrice) / Convert.ToDouble((dataList_WWs.Count + dataList.Count));
            double colesAVG = totalSalePrice / Convert.ToDouble(dataList.Count);
            double wWsAVG = totalSalePrice_WWs / Convert.ToDouble(dataList_WWs.Count);
            totalAVG = Math.Round(totalAVG, 2);
            colesAVG = Math.Round(colesAVG, 2);
            wWsAVG = Math.Round(wWsAVG, 2);
            //summary
            Console.WriteLine("Current Total Average Product (Coles) Price is $" + colesAVG.ToString());
            Console.WriteLine("Current Total Average Product (WoolWorths) Price is $" + wWsAVG.ToString());
            Console.WriteLine("Current Total Average Product (Coles and WoolWorths) Price is $" + totalAVG.ToString());
            Console.WriteLine("Current Total Special Product (Coles) Number is " + specialProductNum_Coles.ToString());
            Console.WriteLine("Current Total Special Product (WoolWorths) Number is " + specialProductNum.ToString());
            Console.WriteLine("Current Total Special Product (Coles and WoolWorths) Number is " + (specialProductNum + specialProductNum_Coles).ToString());
            Console.WriteLine("Current Total Product (Coles) Number is " + dataList.Count.ToString());
            Console.WriteLine("Current Total Product (WoolWorths) Number is " + dataList_WWs.Count.ToString());
            Console.WriteLine("Current Total Product (Coles and WoolWorths) Number is " + (dataList.Count + dataList_WWs.Count).ToString());

        }
    }
}
