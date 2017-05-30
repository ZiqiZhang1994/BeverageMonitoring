using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace MyProject
{
    class WoolworthsExtraction
    {
        public static List<string> FirstCategoryLinkExtraction(string source)
        {
            //parameter
            List<string> linkList = new List<string>();

            //cut off the important part
            string startKeyWord = "<!----><div class=\"categoryList-aisles\" ng-if=\"isAisleOpen(aisle)\">";
            string endKeyWord = "</div><!---->\n</div></wow-categories-spinner-aisle-mf><!----><wow-categories-spinner-aisle-mf ng-repeat=\"aisle in categories track by aisle.UrlFriendlyName\">";
            int startIndex = source.IndexOf(startKeyWord) + startKeyWord.Length;
            int lastIndex = source.IndexOf(endKeyWord);

            //important part
            string linkDiv = source.Substring(startIndex, lastIndex - startIndex);

            //extract link from important part
            string startLinkKeyWord = "ng-href=\"";
            string endLindKeyWord = "\" href=\"";
            int startLinkIndex = 0;
            int lastLinkIndex = 0;
            while (lastLinkIndex != -1)
            {
                startLinkIndex = linkDiv.IndexOf(startLinkKeyWord) + startLinkKeyWord.Length;
                lastLinkIndex = linkDiv.IndexOf(endLindKeyWord);
                if (lastLinkIndex != -1)
                {
                    string link = linkDiv.Substring(startLinkIndex, lastLinkIndex - startLinkIndex);
                    linkList.Add(link);
                    linkDiv = linkDiv.Remove(startLinkIndex - startLinkKeyWord.Length, startLinkKeyWord.Length + link.Length + endLindKeyWord.Length);
                }
            }

            return linkList;

        }

        public static string PageExtraction(string source)
        {
            //cut off the important part
            string startKeyWord = "<div class=\"cardList-loadMore\">";
            string endKeyWord = "</wow-card-list>";
            int startIndex = source.IndexOf(startKeyWord) + startKeyWord.Length;
            int lastIndex = source.IndexOf(endKeyWord);

            //important part
            string linkDiv = source.Substring(startIndex, lastIndex - startIndex);

            //extract link from important part
            string startLinkKeyWord = "ng-href=\"";
            string endLindKeyWord = "\" ng-show=\"!isLoading &amp;&amp; !isInfiniteScrollingEnabled\"";
            string link = null;
            int startLinkIndex = 0;
            int lastLinkIndex = 0;
            startLinkIndex = linkDiv.IndexOf(startLinkKeyWord) + startLinkKeyWord.Length;
            lastLinkIndex = linkDiv.IndexOf(endLindKeyWord);
            if (lastLinkIndex != -1)
            {
                link = linkDiv.Substring(startLinkIndex, lastLinkIndex - startLinkIndex);
            }

            return link;
        }

        public static List<Data_WWs> DetailExtraction(string source)
        {
            //parameter
            List<Data_WWs> productLinkList = new List<Data_WWs>();


            //cut off the important part
            string startKeyWord = "<!----><div ng-if=\"!breakpointsService.isInSmallFormat\"";
            string endKeyWord = "<div class=\"cardList-loadMore\">";
            int startIndex = source.IndexOf(startKeyWord) + startKeyWord.Length;
            int lastIndex = source.IndexOf(endKeyWord);

            //important part
            string linkDiv = source.Substring(startIndex, lastIndex - startIndex);

            //extract link from important part
            string startGridKeyWord = "card in cards track by card.isotopeId";
            string endGridKeyWord = "</div></wow-card>";
            int startGridIndex = 0;
            int lastGridIndex = 0;

            

            while (lastGridIndex != -1)
            {
                startGridIndex = linkDiv.IndexOf(startGridKeyWord) + startGridKeyWord.Length;
                lastGridIndex = linkDiv.IndexOf(endGridKeyWord);
                string gridDiv = linkDiv.Substring(startGridIndex, lastGridIndex - startGridIndex);
                if (lastGridIndex != -1)
                {
                    string detailGrid = linkDiv.Substring(startGridIndex, lastGridIndex - startGridIndex);

                    //extrat detail here
                    if (detailGrid.IndexOf("Shop now") == -1 )
                    {
                        if (detailGrid.IndexOf(" View All Sizes") != -1)
                        {
                            //cut off the important part
                            string startChildGridKeyWord = "<wow-shelf-bundle bundle=\"bundle\" initial-product=\"bundle.Products[0].Stockcode\" flipped=\"flipped\">";
                            string endChildGridKeyWord = "</wow-shelf-bundle>";
                            int startChildGridIndex = detailGrid.IndexOf(startChildGridKeyWord) + startChildGridKeyWord.Length;
                            int lastChildGridIndex = detailGrid.IndexOf(endChildGridKeyWord);

                            //important part
                            string childDiv = detailGrid.Substring(startChildGridIndex, lastChildGridIndex - startChildGridIndex);

                            //detail extraction
                            string startSeperationKeyWord = "<li class=\"productVariants - item clearfix\"";
                            string endSeperationKeyWord = "</wow-shelf-variant>";
                            int startSeperationIndex = 0;
                            int lastSeperationIndex = 0;

                            //catch title
                            string startTitleKeyWord = "<h2 class=\"productVariants-title\" ng-bind-html=\"bundle.Name | sanitize\"> ";
                            string endTitleKeyWord = " </h2>\n\n    <ul class=\"productVariants-list\">";
                            int startTitleIndex = childDiv.IndexOf(startTitleKeyWord) + startTitleKeyWord.Length;
                            int lastTitleIndex = childDiv.IndexOf(endTitleKeyWord);
                            string title = childDiv.Substring(startTitleIndex, lastTitleIndex - startTitleIndex);

                            //catch detail from different quantity
                            while (lastSeperationIndex != -1)
                            {
                                startSeperationIndex = childDiv.IndexOf(startSeperationKeyWord) + startSeperationKeyWord.Length;
                                lastSeperationIndex = childDiv.IndexOf(endSeperationKeyWord);
                                string smallGrid = childDiv.Substring(startSeperationIndex, lastSeperationIndex - startSeperationIndex);
                                string startSpecialPriceKeyWord = "<div class=\"productVariant-price u-special\" ng-class=\"specialClass\">\n                $";

                                if (smallGrid.IndexOf(startSpecialPriceKeyWord) == -1)
                                { 
                                    //catch size
                                    string startSizeKeyWord = "<div class=\"productVariant-variant\">\n            ";
                                    string endSizeKeyWord = "\n        </div>\n\n        <div class=\"productVariant-prices\">";
                                    int startSizeIndex = smallGrid.IndexOf(startSizeKeyWord) + startSizeKeyWord.Length;
                                    int lastSizeIndex = smallGrid.IndexOf(endSizeKeyWord);
                                    string size = smallGrid.Substring(startSizeIndex, lastSizeIndex - startSizeIndex);

                                    //if sale?
                                    string ifSale = "N";


                                    //catch price
                                    string price = "0";
                                    string startPriceKeyWord = "<div class=\"productVariant-price\" ng-class=\"specialClass\">\n                $";
                                    string endPriceKeyWord = "\n        </div>\n\n        <div class=\"productVariant-prices\">";



                                    int startPriceIndex = smallGrid.IndexOf(startPriceKeyWord) + startPriceKeyWord.Length;
                                    int lastPriceIndex = smallGrid.IndexOf(endPriceKeyWord);
                                    price = smallGrid.Substring(startPriceIndex, lastPriceIndex - startPriceIndex);


                                    //catch perL
                                    string startPerLKeyWord = "<div class=\"productVariant-cup\" ng-show=\"product.HasCupPrice\">\n                $";
                                    string endPerLKeyWord = "\n            </div>\n        </div>\n    </div>\n\n    <div class=\"productVariant-thumbnail\">\n";
                                    int startPerLIndex = smallGrid.IndexOf(startPerLKeyWord) + startPerLKeyWord.Length;
                                    int lastPerLIndex = smallGrid.IndexOf(endPerLKeyWord);
                                    string perL = smallGrid.Substring(startPerLIndex, lastPerLIndex - startPerLIndex);

                                    //catch ImageUrl
                                    string startImageUrlKeyWord = "<!----><img ng-if=\"size != 'xsmall'\" ng-src=\"";
                                    string endImageUrlKeyWord = "\" alt=\"" + title + "\"";
                                    int startImageUrlIndex = smallGrid.IndexOf(startImageUrlKeyWord) + startImageUrlKeyWord.Length;
                                    int lastImageUrlIndex = smallGrid.IndexOf(endImageUrlKeyWord);
                                    string imageUrl = smallGrid.Substring(startImageUrlIndex, lastImageUrlIndex - startImageUrlIndex);

                                    Data_WWs tempWWsData = new Data_WWs(title, price, size, "1", perL, ifSale, "0", imageUrl);

                                    childDiv = childDiv.Remove(startSeperationIndex - startSeperationKeyWord.Length, startSeperationKeyWord.Length + smallGrid.Length + endSeperationKeyWord.Length);

                                }
                            } 
                        }
                    }






                    linkDiv = linkDiv.Remove(startGridIndex - startGridKeyWord.Length, startGridKeyWord.Length + detailGrid.Length + endGridKeyWord.Length);
                }
            }

            return null;
        }

        private static Data ViewMoreDetailPageExtraction(string source)
        {

            //catch quantity
            string quantity;
            string startQuantityKeyWord = "product-qty\">";
            string endQuantityKeyWord = "</span>\n\t\t\t<span class=\"product-text\">for</span>\n\t\t\t<span>";
            string endQuantityKeyWord2 = "</span>\n\t\t\t<span class=\"accessibility\" data-ng-bind=\"::fatControllerVM.product.prodA11yHeading\">";
            int startQuantityIndex = detailDiv.IndexOf(startQuantityKeyWord) + startQuantityKeyWord.Length;
            int endQuantityIndex = detailDiv.IndexOf(endQuantityKeyWord);
            int endQuantityIndex2 = detailDiv.IndexOf(endQuantityKeyWord2);
            if (endQuantityIndex == -1 && endQuantityIndex2 != -1)
            {
                quantity = detailDiv.Substring(startQuantityIndex, endQuantityIndex2 - startQuantityIndex);
            }
            else if (endQuantityIndex2 == -1 && endQuantityIndex != -1)
            {
                quantity = detailDiv.Substring(startQuantityIndex, endQuantityIndex - startQuantityIndex);
            }
            else
            {
                quantity = "0";
            }

            //catch price
            string price;
            if (Convert.ToInt32(quantity) > 1)
            {
                string startPriceKeyWord = "product-price\">";
                string endPriceKeyWord = "</strong></span>\n\t\t\t<span aria-hidden=\"true\">";
                int startPriceIndex = detailDiv.IndexOf(startPriceKeyWord) + startPriceKeyWord.Length;
                int endPriceIndex = detailDiv.IndexOf(endPriceKeyWord);
                price = detailDiv.Substring(startPriceIndex, endPriceIndex - startPriceIndex);
            }
            else if (Convert.ToInt32(quantity) == 1)
            {
                string startPriceKeyWord = "product-price\">";
                string endPriceKeyWord = "</strong>\n\t\t\t<span class=\"accessibility\">to the trolley.</span>\n\t\t\t<!---->";
                int startPriceIndex = detailDiv.IndexOf(startPriceKeyWord) + startPriceKeyWord.Length;
                int endPriceIndex = detailDiv.IndexOf(endPriceKeyWord);
                price = detailDiv.Substring(startPriceIndex, endPriceIndex - startPriceIndex);
            }
            else
            {
                price = "0";
            }

            //catch size
            string size;
            string startSizeKeyWord = "<h2 class=\"product-specific-heading\" data-ng-bind=\"::productSpecific.name\">Size:</h2>\n                                <span data-ng-bind-html=\"::productSpecific.value.sanitizeHtml()\">";
            string endSizeKeyWord = "</span>";
            string startSizeKeyWord2 = "productDisplayVM.product.showOnlineSizeDesc\" data-ng-bind=\"::productDisplayVM.product.sizeDescription\">";
            string endSizeKeyWord2 = "</span><!----> \n";
            int startSizeIndex = detailDiv.IndexOf(startSizeKeyWord);
            int startSizeIndex2 = detailDiv.IndexOf(startSizeKeyWord2);
            if (startSizeIndex > -1 && startSizeIndex2 == -1)
            {
                startSizeIndex = startSizeIndex + startSizeKeyWord.Length;
                string cutStep1 = detailDiv.Substring(startSizeIndex);
                int endSizeIndex = cutStep1.IndexOf(endSizeKeyWord);
                size = cutStep1.Remove(endSizeIndex);
            }
            else if (startSizeIndex2 > -1 && startSizeIndex == -1)
            {

                startSizeIndex2 = detailDiv.IndexOf(startSizeKeyWord2) + startSizeKeyWord2.Length;
                string cutStep2 = detailDiv.Substring(startSizeIndex2);
                int endSizeIndex2 = cutStep2.IndexOf(endSizeKeyWord2);
                size = cutStep2.Remove(endSizeIndex2);
            }
            else if (startSizeIndex > -1 && startSizeIndex2 > -1)
            {
                startSizeIndex2 = startSizeIndex2 + startSizeKeyWord2.Length;
                string cutStep2 = detailDiv.Substring(startSizeIndex2);
                int endSizeIndex2 = cutStep2.IndexOf(endSizeKeyWord2);
                size = cutStep2.Remove(endSizeIndex2);
            }
            else
            {
                size = null;
            }


            //catch per L
            string startPerLKeyWord = "productDisplayVM.product.unitPrice\">";
            string endPerLKeyWord = "</span>\n\t                                <span class=\"accessibility-inline\"";
            int startPerLIndex = detailDiv.IndexOf(startPerLKeyWord) + startPerLKeyWord.Length;
            int endPerLIndex = detailDiv.IndexOf(endPerLKeyWord);
            string perL = detailDiv.Substring(startPerLIndex, endPerLIndex - startPerLIndex);

            //if sale
            string ifSale = "N";
            string salePrice = "";
            if (detailDiv.IndexOf("data-colrs-fat-controller-responsive-component=\"Save\">") != -1)
            {
                string startSaveKeyWord = "\n\t\t\t\tsave \n\t\t\t\t<strong>";
                string endSaveKeyWord = "</strong>\n\t\t\t</span>";
                int startSaveIndex = detailDiv.IndexOf(startSaveKeyWord) + startSaveKeyWord.Length;
                int endSaveIndex = detailDiv.IndexOf(endSaveKeyWord);
                salePrice = detailDiv.Substring(startSaveIndex, endSaveIndex - startSaveIndex);
                ifSale = "Y";
            }
            else if (Convert.ToInt32(quantity) > 1)
            {
                salePrice = "0";
                ifSale = "Y";
            }


            //catch barcode
            string startCodeKeyWord = "Code</h2>\n\t\t\t\t\t\t\t\t<p data-ng-bind=\"::productDisplayVM.product.partNumber\">";
            string endCodeKeyWord = "</p>\n\t\t\t\t\t\t\t</div>\n\t\t\t\t\t\t\t\n";
            int startCodeIndex = detailDiv.IndexOf(startCodeKeyWord) + startCodeKeyWord.Length;
            int endCodeIndex = detailDiv.IndexOf(endCodeKeyWord);
            string barcode = detailDiv.Substring(startCodeIndex, endCodeIndex - startCodeIndex);

            Data data = new Data(brand, title, price, size, quantity.ToString(), perL, ifSale, salePrice, barcode);

            return data;
        }
    }
}
