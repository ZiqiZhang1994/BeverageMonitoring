using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace MyProject
{
    class ColesExtraction
    {
        public static List<string> FirstCategoryLinkExtraction(string source)
        {
            //parameter
            List<string> linkList = new List<string>();

            //cut off the important part
            string startKeyWord = "</span><!----></span></a><!---->\n   \t\t\n\t\t<!---->\n\t\t<!----><div id=\"cat-nav-list-2\" class=\"colrs-animate nav-animate-item item-l2\" colrs-category-list=\"\" data-catgroup=\"catNavVM.model.tabView.currentCatGrps[1]\" ng-switch-when=\"1\"><ul><!----><li class=\"cat-nav-item\" data-ng-class=\"{'is-disabled': (catNavLiVM.getCategoryCount() === 0)}\" aria-hidden=\"false\" data-level=\"0\" data-ng-repeat=\"category in catgroup.children\">\n    <";
            string endKeyWord = "<div class=\"utils-sorted is-disabled\" colrs-sort-view=\"\" ng-class=\"{'is-disabled' : !sortViewVM.model.active.isVisible}\" aria-hidden=\"true\">\n\t\n<p id=\"sort-options\">";
            int startIndex = source.IndexOf(startKeyWord) + startKeyWord.Length;
            int lastIndex = source.IndexOf(endKeyWord);

            //important part
            string linkDiv = source.Substring(startIndex, lastIndex - startIndex);

            //extract link from important part
            string startLinkKeyWord = "ng-href=\"";
            string endLindKeyWord = "\" tabindex=\"";
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
            string startKeyWord = "\n\t\t\t\t\t<!----><span class=\"button is-active\" data-ng-if=\"::paginationVM.isActive($index)\" tabindex=\"0\">\n\t\t\t\t\t\t<span class=\"accessibility-inline\">\n\t\t\t\t\t\t\tCurrent page page";
            string endKeyWord = "<!----><!----><div class=\"results-view ng-hide\" id=\"results-bought-before\" ng-repeat=\"tab in mainViewVM.model.tabs track by tab.id\" data-ui-view=\"results-bought-before\"";
            int startIndex = source.IndexOf(startKeyWord) + startKeyWord.Length;
            int lastIndex = source.IndexOf(endKeyWord);

            //important part
            string linkDiv = source.Substring(startIndex, lastIndex - startIndex);

            //extract link from important part
            string startLinkKeyWord = "data-ng-click=\"productListVM.pagination(paginationObject.number,$event)\" href=\"";
            string endLindKeyWord = "\">\n\t\t\t\t\t\t<span class=\"accessibility-inline\">Page</span>";
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

        public static List<string> ProductLinkExtraction(string source)
        {
            //parameter
            List<string> productLinkList = new List<string>();


            //cut off the important part
            string startKeyWord = "< !--Begin Page-- >";
            string endKeyWord = "<div class=\"pagination-container\" data-colrs-pagination=\"\">";
            int startIndex = source.IndexOf(startKeyWord) + startKeyWord.Length;
            int lastIndex = source.IndexOf(endKeyWord);

            //important part
            string linkDiv = source.Substring(startIndex, lastIndex - startIndex);

            //extract link from important part
            string startLinkKeyWord = "<a role=\"presentation\" aria-hidden=\"true\" tabindex=\"-1\" data-ng-attr-href=\"{{::productTileVM.pdpHref}}\" data-ng-click=\"productTileVM.openProduct($event)\" href=\"";
            string endLindKeyWord = "\">\n\t\t\t\t\t<img role=\"presentation\" alt=\"";
            int startLinkIndex = 0;
            int lastLinkIndex = 0;
            while (lastLinkIndex != -1)
            {
                startLinkIndex = linkDiv.IndexOf(startLinkKeyWord) + startLinkKeyWord.Length;
                lastLinkIndex = linkDiv.IndexOf(endLindKeyWord);
                if (lastLinkIndex != -1)
                {
                    string link = linkDiv.Substring(startLinkIndex, lastLinkIndex - startLinkIndex);
                    productLinkList.Add(link);
                    linkDiv = linkDiv.Remove(startLinkIndex - startLinkKeyWord.Length, startLinkKeyWord.Length + link.Length + endLindKeyWord.Length);
                }
            }

            return productLinkList;
        }

        public static Data DetailPageExtraction(string source)
        {
            //catch title
            string startTitleKeyWord = "data-ng-bind=\"mainVM.metadata.title\">";
            string endTitleKeyWord = " | Coles Online</div>";
            int startTitleIndex = source.IndexOf(startTitleKeyWord) + startTitleKeyWord.Length;
            int lastTitleIndex = source.IndexOf(endTitleKeyWord);
            string title = source.Substring(startTitleIndex, lastTitleIndex - startTitleIndex);

            //catch brand
            string startDetailKeyWord = "<span class=\"product-name\" aria-hidden=\"true\" data-ng-bind=\"::productDisplayVM.product.displayNameText\">" + title + "</span>";
            string endDetailKeyWord = "Product display page disclaimer.";
            string startBarndKeyWord = "<span class=\"product-brand\" aria-hidden=\"true\" data-ng-bind=\"::productDisplayVM.product.displayBrandText\">";
            string endBrandKeyWord = "</span>\n\t                                " + startDetailKeyWord;
            int startBrandIndex = source.IndexOf(startBarndKeyWord) + startBarndKeyWord.Length;
            int endBrandIndex = source.IndexOf(endBrandKeyWord);

            string brand = source.Substring(startBrandIndex, endBrandIndex - startBrandIndex);

            //cut off necessary part
            int startDetailIndex = source.IndexOf(startDetailKeyWord) + startDetailKeyWord.Length;
            int endDetailIndex = source.IndexOf(endDetailKeyWord);
            string detailDiv = source.Substring(startDetailIndex, endDetailIndex - startDetailIndex);

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
