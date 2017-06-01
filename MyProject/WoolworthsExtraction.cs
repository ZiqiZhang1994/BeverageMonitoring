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
            string startKeyWord = "<!----><div ng-if=\"!breakpointsService.isInSmallFormat\" class=\"cardList-cards text-left cardList-isotopeContainer\" ng-class=\"{'isCompact' : stampSize == cardSizes.Compact, 'cardList-isotopeContainer': enableIsotope }\" style=\"height: auto;\">";
            string endKeyWord = "<div class=\"_cardDetail-content\"></div>\n\n</div></wow-card><!---->\n    </div><!---->\n\n    <div class=\"cardList-loadMore\">\n";
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
                if (lastGridIndex != -1)
                {
                    string detailGrid = linkDiv.Substring(startGridIndex, lastGridIndex - startGridIndex);

                    //extrat detail here
                    if (detailGrid.IndexOf("inspirationCard-link") == -1 )
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
                            string startSeperationKeyWord = "<!----><li class=\"productVariants-item clearfix\"";
                            string endSeperationKeyWord = "</wow-shelf-variant>";
                            int startSeperationIndex = 0;
                            int lastSeperationIndex = 0;

                            //catch title
                            string startTitleKeyWord = "<h2 class=\"productVariants-title\" ng-bind-html=\"bundle.Name | sanitize\"> ";
                            string endTitleKeyWord = " </h2>\n\n    <ul class=\"productVariants-list\">";
                            int startTitleIndex = childDiv.IndexOf(startTitleKeyWord) + startTitleKeyWord.Length;
                            int lastTitleIndex = childDiv.IndexOf(endTitleKeyWord);
                            string title = childDiv.Substring(startTitleIndex, lastTitleIndex - startTitleIndex);
                            if(title.IndexOf("<br>") != -1)
                            {
                                title = title.Replace("<br>", " ");
                            }
                            if (title == "" || title == " " || title == null)
                            {
                                title = "0";
                            }

                            //catch detail from different quantity
                            while (lastSeperationIndex != -1)
                            {
                                startSeperationIndex = childDiv.IndexOf(startSeperationKeyWord) + startSeperationKeyWord.Length;
                                lastSeperationIndex = childDiv.IndexOf(endSeperationKeyWord);
                              
                                if (lastSeperationIndex != -1)
                                {
                                    string smallGrid = childDiv.Substring(startSeperationIndex, lastSeperationIndex - startSeperationIndex);
                                    string startSpecialPriceKeyWord = "<div class=\"productVariant-price u-special\" ng-class=\"specialClass\">\n                ";
                                    if (smallGrid.IndexOf(startSpecialPriceKeyWord) == -1)
                                    {
                                        //catch size
                                        string startSizeKeyWord = "<div class=\"productVariant-variant\">\n            ";
                                        string endSizeKeyWord = "\n        </div>\n\n        <div class=\"productVariant-prices\">";
                                        int startSizeIndex = smallGrid.IndexOf(startSizeKeyWord) + startSizeKeyWord.Length;
                                        int lastSizeIndex = smallGrid.IndexOf(endSizeKeyWord);
                                        string size = smallGrid.Substring(startSizeIndex, lastSizeIndex - startSizeIndex);
                                        if (size == "" || size == " " || size == null)
                                        {
                                            size = "0";
                                        }
                                        //if sale?
                                        string ifSale = "N";


                                        //catch price
                                        string startPriceKeyWord = "<div class=\"productVariant-price\" ng-class=\"specialClass\">\n                ";
                                        string endPriceKeyWord = "\n            </div>\n\n            <div class=\"productVariant-cup\" ng-show=\"product.HasCupPrice\">\n";
                                        int startPriceIndex = smallGrid.IndexOf(startPriceKeyWord) + startPriceKeyWord.Length;
                                        int lastPriceIndex = smallGrid.IndexOf(endPriceKeyWord);
                                        string price = smallGrid.Substring(startPriceIndex, lastPriceIndex - startPriceIndex);
                                        if (price == "" || price == " " || price == null)
                                        {
                                            price = "0";
                                        }

                                        //catch perL
                                        string startPerLKeyWord = "<div class=\"productVariant-cup\" ng-show=\"product.HasCupPrice\">\n                ";
                                        string endPerLKeyWord = "\n            </div>\n        </div>\n    </div>\n\n    <div class=\"productVariant-thumbnail\">\n";
                                        int startPerLIndex = smallGrid.IndexOf(startPerLKeyWord) + startPerLKeyWord.Length;
                                        int lastPerLIndex = smallGrid.IndexOf(endPerLKeyWord);
                                        string perL = smallGrid.Substring(startPerLIndex, lastPerLIndex - startPerLIndex);
                                        if (perL == "" || perL == " " || perL == null)
                                        {
                                            perL = "0";
                                        }
                                        //catch ImageUrl
                                        string imageUrl = "";
                                        string startImageUrlKeyWord = "<!----><img ng-if=\"size != 'xsmall'\" ng-src=\"";
                                        string endImageUrlKeyWord = "\" alt=\"";
                                        int startImageUrlIndex = smallGrid.IndexOf(startImageUrlKeyWord) + startImageUrlKeyWord.Length;
                                        int lastImageUrlIndex = smallGrid.IndexOf(endImageUrlKeyWord);
                                        if (lastImageUrlIndex < startImageUrlIndex)
                                        {
                                            lastImageUrlIndex = smallGrid.IndexOf(endImageUrlKeyWord, lastImageUrlIndex + 1);
                                            imageUrl = smallGrid.Substring(startImageUrlIndex, lastImageUrlIndex - startImageUrlIndex);
                                        }
                                        else
                                        {
                                            imageUrl = smallGrid.Substring(startImageUrlIndex, lastImageUrlIndex - startImageUrlIndex);
                                        }
                                        if (imageUrl == "" || imageUrl == " " || imageUrl == null)
                                        {
                                            imageUrl = "0";
                                        }

                                        //data log
                                        Data_WWs tempWWsData = new Data_WWs(title, price, size, "1", perL, ifSale, "0", "0", imageUrl);
                                        productLinkList.Add(tempWWsData);
                                        childDiv = childDiv.Remove(startSeperationIndex - startSeperationKeyWord.Length, startSeperationKeyWord.Length + smallGrid.Length + endSeperationKeyWord.Length);
                                        Console.WriteLine(tempWWsData.ToString());
                                        
                                    }
                                    else
                                    {
                                        childDiv = childDiv.Remove(startSeperationIndex - startSeperationKeyWord.Length, startSeperationKeyWord.Length + smallGrid.Length + endSeperationKeyWord.Length);
                                    }
                                }
                            } 
                        }
                        else
                        {
                            //catch title
                            string startTitleKeyWord = "<!----><span class=\"shelfProductStamp-productDetailsLink\" ng-if=\"displaySmallFormatVersion() || isdetail || !hasVariants\" ng-bind-html=\"::(product.SmallFormatDescription | sanitize)\">";
                            string endTitleKeyWord = "</span><!---->\n                        <br>\n                        <!---->";
                            int startTitleIndex = detailGrid.IndexOf(startTitleKeyWord) + startTitleKeyWord.Length;
                            int lastTitleIndex = detailGrid.IndexOf(endTitleKeyWord);
                            string title = detailGrid.Substring(startTitleIndex, lastTitleIndex - startTitleIndex);
                            if (title == "" || title == " " || title == null)
                            {
                                title = "0";
                            }
                            //catch size
                            string startSizeKeyWord = "<!----><span class=\"shelfProductStamp-productDetailsPackageSize\" ng-if=\"displaySmallFormatVersion() || isdetail || !hasVariants\"> ";
                            string endSizeKeyWord = "</span><!---->\n                    </a><!---->\n                    <!---->\n                </h3>\n\n                <!---->\n            </div>\n";
                            int startSizeIndex = detailGrid.IndexOf(startSizeKeyWord) + startSizeKeyWord.Length;
                            int lastSizeIndex = detailGrid.IndexOf(endSizeKeyWord);
                            string size = detailGrid.Substring(startSizeIndex, lastSizeIndex - startSizeIndex);
                            if (size == "" || size == " " || size == null)
                            {
                                size = "0";
                            }



                            //if sale? & price
                            string ifSale = "N";
                            string price = "0";
                            string savedPrice = "0";
                            if (detailGrid.IndexOf("<span class=\"pricingContainer-saveHeading\">") != -1)
                            {
                                ifSale = "Y";
                                //catch price
                                string startPriceKeyWord = "<span class=\"pricingContainer-priceAmount u-special\" ng-class=\"::specialClass\">";
                                string endPriceKeyWord = "</span>\n        <!----><span ng-if=\"::product.CupPrice\" class=\"pricingContainer-priceCup\">\n";
                                int startPriceIndex = detailGrid.IndexOf(startPriceKeyWord) + startPriceKeyWord.Length;
                                int lastPriceIndex = detailGrid.IndexOf(endPriceKeyWord);
                                price = detailGrid.Substring(startPriceIndex, lastPriceIndex - startPriceIndex);
                                if (price == "" || price == " " || price == null)
                                {
                                    price = "0";
                                }
                                //catch saved price
                                string startSavedPriceKeyWord = "</span>\n            <span class=\"pricingContainer-savePrice\">";
                                string endSavedPriceKeyWord = "</span>\n        </div>\n    </div><!---->\n    <div class=\"pricingContainer-priceContainer\">\n";
                                int startSavedPriceIndex = detailGrid.IndexOf(startSavedPriceKeyWord) + startSavedPriceKeyWord.Length;
                                int lastSavedPriceIndex = detailGrid.IndexOf(endSavedPriceKeyWord);
                                savedPrice = detailGrid.Substring(startSavedPriceIndex, lastSavedPriceIndex - startSavedPriceIndex);
                                if(savedPrice.IndexOf("¢") != -1)
                                {
                                    savedPrice = savedPrice.Remove(savedPrice.IndexOf("¢"), 1);
                                    savedPrice = "$" + (Convert.ToDouble(savedPrice) / 100).ToString();
                                }
                                if (savedPrice == "" || savedPrice == " " || savedPrice == null)
                                {
                                    savedPrice = "0";
                                }

                            }
                            else
                            {
                                //catch price
                                string startPriceKeyWord = "<span class=\"pricingContainer-priceAmount\" ng-class=\"::specialClass\">";
                                string endPriceKeyWord = "</span>\n        <!----><span ng-if=\"::product.CupPrice\" class=\"pricingContainer-priceCup\">\n";
                                string secondEndPriceKeyWord = "</span>\n        <!---->\n    </div>\n</div><!----></wow-shelf-product-pricing-container>";
                                int startPriceIndex = detailGrid.IndexOf(startPriceKeyWord) + startPriceKeyWord.Length;
                                int lastPriceIndex = detailGrid.IndexOf(endPriceKeyWord);
                                if(lastPriceIndex == -1)
                                {
                                    lastPriceIndex = detailGrid.IndexOf(secondEndPriceKeyWord);
                                }
                                price = detailGrid.Substring(startPriceIndex, lastPriceIndex - startPriceIndex);
                                if (price == "" || price == " " || price == null)
                                {
                                    price = "0";
                                }
                            }

                            //catch perL
                            string perL = "0";
                            string startPerLKeyWord = "<!----><span ng-if=\"::product.CupPrice\" class=\"pricingContainer-priceCup\">\n            ";
                            string endPerLKeyWord = "\n        </span><!---->\n    </div>\n</div><!----></wow-shelf-product-pricing-container>\n\n";
                            int startPerLIndex = detailGrid.IndexOf(startPerLKeyWord) + startPerLKeyWord.Length;
                            int lastPerLIndex = detailGrid.IndexOf(endPerLKeyWord);
                            if(lastPerLIndex != -1)
                            {
                                 perL = detailGrid.Substring(startPerLIndex, lastPerLIndex - startPerLIndex);
                            }
                            if (perL == "" || perL == " " || perL == null)
                            {
                                perL = "0";
                            }
                            //catch quantity special price
                            string quantitySpecialPrice = "0";
                            if (detailGrid.IndexOf("<!----><span ng-if=\"!enabledCenterProductTag || !tag.TagLink\" ng-bind-html=\"::tag.displayContent\" ng-click=\"htmlTagClick($event)\" class=\"shelfProductStamp-centerProductTagHtmlInner\"><span style=\"COLOR: #e2001a\">") == -1 && detailGrid.IndexOf("<span style=\"COLOR: #e2001a\">") != -1)
                            {
                                string startQuantitySpecialKeyWord = "<span style=\"COLOR: #e2001a\">";
                                string endQuantitySpecialKeyWord = "</a></span><!---->\n    </div><!---->\n</div><!----></wow-shelf-product-producttag><!---->\n";
                                int startQuantitySpecialIndex = detailGrid.IndexOf(startQuantitySpecialKeyWord) + startQuantitySpecialKeyWord.Length;
                                int lastQuantitySpecialIndex = detailGrid.IndexOf(endQuantitySpecialKeyWord);

                                quantitySpecialPrice = detailGrid.Substring(startQuantitySpecialIndex, lastQuantitySpecialIndex - startQuantitySpecialIndex);
                                quantitySpecialPrice = quantitySpecialPrice.Remove(quantitySpecialPrice.IndexOf("</span>", "</span>".Length));
                                ifSale = "Y";
                            }
                            if (quantitySpecialPrice == "" || quantitySpecialPrice == " " || quantitySpecialPrice == null)
                            {
                                quantitySpecialPrice = "0";
                            }

                            //catch ImageUrl
                            string startImageUrlKeyWord = "<div class=\"shelfProductStamp-imageTagsContainer\" ng-switch=\"::(!!product.ImageTag.TagLink)\">\n                <img ng-src=\"";
                            string endImageUrlKeyWord = "\" alt=\"\" class=\"shelfProductStamp-productImage";
                            int startImageUrlIndex = detailGrid.IndexOf(startImageUrlKeyWord) + startImageUrlKeyWord.Length;
                            int lastImageUrlIndex = detailGrid.IndexOf(endImageUrlKeyWord);
                            string imageUrl = detailGrid.Substring(startImageUrlIndex, lastImageUrlIndex - startImageUrlIndex);
                            if (imageUrl == "" || imageUrl == " " || imageUrl == null)
                            {
                                imageUrl = "0";
                            }
                            //data log
                            Data_WWs tempWWsData = new Data_WWs(title, price, size, "1", perL, ifSale, savedPrice, quantitySpecialPrice, imageUrl);
                            productLinkList.Add(tempWWsData);
                            Console.WriteLine(tempWWsData.ToString());
                        }

                    }

                    //remove used part
                    linkDiv = linkDiv.Remove(startGridIndex - startGridKeyWord.Length, startGridKeyWord.Length + detailGrid.Length + endGridKeyWord.Length);
                }
            }

            return productLinkList;
        }


    }
}
