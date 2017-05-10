using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using CefSharp;
using CefSharp.OffScreen;
using System.IO;
using System.Diagnostics;
using System.Threading;

namespace MyProject
{
    public class Program
    {
        private const string TestUrl = "https://www.google.com/";
        private const string shopUrl = "https://shop.coles.com.au";
        public static int Main(string[] args)
        {
            bool isRunning = true; // For exit function
            string MenuSelection = ""; // For MenuSelection Case Function
            string Response = "";

            Console.WriteLine("Test network for {0}", TestUrl);
            Console.WriteLine("You may see a lot of Chromium debugging output, please wait...");
            Console.WriteLine();

            var settings = new CefSettings()
            {
                //By default CefSharp will use an in-memory cache, you need to specify a Cache Folder to persist data
                CachePath = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData), "CefSharp\\Cache")
            };

            //Perform dependency check to make sure all relevant resources are in our output directory.
            Cef.Initialize(settings, performDependencyCheck: true, browserProcessHandler: null);

            //Menu loop
            while (isRunning == true)
            {
                DisplayMenu();
                MenuSelection = Console.ReadLine();
                MenuSelection = MenuSelection.ToLower();

                switch (MenuSelection)
                {
                    case "1":
                        DemoMainAsync("cachePath2");
                        break;
                    case "2":
                        FullColesMainAsync("cachePatch1");
                        break;
                    case "3":
                        WoolworthsTestMainAsync("cachePatch3");
                        break;
                    // Exit Function
                    case "x":
                        //confirm exit
                        while (isRunning)
                        {
                            Console.Write("You have chosen to exit. Do you wish to exit? (Y or N) : ");
                            Response = Console.ReadLine();
                            Response = Response.ToLower();
                            if (Response == "y")
                            {
                                isRunning = false;
                            }
                            else if (Response == "n")
                            {
                                break;
                            }
                        }
                        break;
                    // Index Selection Verification
                    default:
                        Console.WriteLine("Not a valid selection. Press ENTER to continue");
                        break;

                }
            }

            
            //Demo showing Zoom Level of 3.0
            //Using seperate request contexts allows the urls from the same domain to have independent zoom levels
            //otherwise they would be the same - default behaviour of Chromium
            //MainAsync("cachePath2", 3.0);

            // We have to wait for something, otherwise the process will exit too soon.
            //Console.ReadKey();

            // Clean up Chromium objects.  You need to call this in your application otherwise
            // you will get a crash when closing.
            Cef.Shutdown();

            //Success
            return 0;
        }

        private static async void DemoMainAsync(string cachePath)
        {
            int delayTimer = 2500;
            List<string> cateLinkList = new List<string>();
            List<string> productPageLinkList = new List<string>();
            List<Data> dataList = new List<Data>();
            var browserSettings = new BrowserSettings();
            //Reduce rendering speed to one frame per second so it's easier to take screen shots
            browserSettings.WindowlessFrameRate = 1;
            var requestContextSettings = new RequestContextSettings { CachePath = cachePath };

            // RequestContext can be shared between browser instances and allows for custom settings
            // e.g. CachePath
            using (var requestContext = new RequestContext(requestContextSettings))
            using (var browser = new ChromiumWebBrowser(TestUrl, browserSettings, requestContext))
            {
                await LoadPageAsync(browser);

                //Check preferences on the CEF UI Thread
                await Cef.UIThreadTaskFactory.StartNew(delegate
                {
                    var preferences = requestContext.GetAllPreferences(true);

                    //Check do not track status
                    var doNotTrack = (bool)preferences["enable_do_not_track"];

                    Debug.WriteLine("DoNotTrack:" + doNotTrack);
                });

                var onUi = Cef.CurrentlyOnThread(CefThreadIds.TID_UI);

                // For Google.com pre-pupulate the search text box
                await browser.EvaluateScriptAsync("document.getElementById('lst-ib').value = 'CefSharp Was Here!'");

                // Wait for the screenshot to be taken,
                // if one exists ignore it, wait for a new one to make sure we have the most up to date
                //await browser.ScreenshotAsync(true).ContinueWith(DisplayBitmap);

                //await LoadPageAsync(browser, "https://shop.coles.com.au/a/a-national/everything/browse/drinks?pageNumber=1");
                //Thread.Sleep(delayTimer);

                //Gets a wrapper around the underlying CefBrowser instance
                var cefBrowser = browser.GetBrowser();
                // Gets a warpper around the CefBrowserHost instance
                // You can perform a lot of low level browser operations using this interface
                var cefHost = cefBrowser.GetHost();

                //You can call Invalidate to redraw/refresh the image
                cefHost.Invalidate(PaintElementType.View);

                //demo presentation
                string a = "/a/a-national/everything/browse/drinks/non-alcoholic-3314656?pageNumber=1";
                string pageLink = "";
                List<string> pageLinkList = new List<string>();
                pageLinkList.Add(a);
                await LoadPageAsync(browser, shopUrl + a);
                Thread.Sleep(delayTimer);
                pageLink = ColesExtraction.PageExtraction(await browser.GetSourceAsync());

                if (pageLink != null && pageLink != "")
                {
                    pageLinkList.Add(pageLink);
                }

                Console.WriteLine("************************category link list*************************");
                Console.WriteLine(shopUrl + a);
                Console.WriteLine("********************************************************************");
                Console.WriteLine("************************First page link*************************");
                Console.WriteLine(shopUrl + pageLink);
                Console.WriteLine("********************************************************************");

                while (pageLink != null)
                {

                    await LoadPageAsync(browser, shopUrl + pageLink);
                    Thread.Sleep(delayTimer);
                    pageLink = ColesExtraction.PageExtraction(await browser.GetSourceAsync());
                    if (pageLink != null)
                    {
                        Console.WriteLine("****************Page Link******************************");
                        Console.WriteLine(shopUrl + pageLink);
                        Console.WriteLine("********************************************************");

                        pageLinkList.Add(pageLink);
                    }
                }

                foreach (string b in pageLinkList)
                {
                    await LoadPageAsync(browser, shopUrl + b);
                    Thread.Sleep(delayTimer);
                    Console.WriteLine("*****************Page link list****************************");
                    Console.WriteLine(shopUrl + b);
                    Console.WriteLine("***********************************************************");

                    productPageLinkList.AddRange(ColesExtraction.ProductLinkExtraction(await browser.GetSourceAsync()));
                }

                foreach (string c in productPageLinkList)
                {
                    await LoadPageAsync(browser, shopUrl + c);
                    Thread.Sleep(delayTimer);
                    Console.WriteLine("*****************Product detail page link list****************************");
                    Console.WriteLine(shopUrl + c);
                    Console.WriteLine("***********************************************************");
                    dataList.Add(ColesExtraction.DetailPageExtraction(await browser.GetSourceAsync()));
                }
                //demo end
                ExcelExport.AutomateExcel(dataList);
            }
        }

        private static async void FullColesMainAsync(string cachePath)
        {
            int delayTimer = 2500;
            List<string> cateLinkList = new List<string>();
            List<string> productPageLinkList = new List<string>();
            List<Data> dataList = new List<Data>();
            var browserSettings = new BrowserSettings();
            //Reduce rendering speed to one frame per second so it's easier to take screen shots
            browserSettings.WindowlessFrameRate = 1;
            var requestContextSettings = new RequestContextSettings { CachePath = cachePath };

            // RequestContext can be shared between browser instances and allows for custom settings
            // e.g. CachePath
            using (var requestContext = new RequestContext(requestContextSettings))
            using (var browser = new ChromiumWebBrowser(TestUrl, browserSettings, requestContext))
            {
                //if (zoomLevel > 1)
                //{
                //    browser.FrameLoadStart += (s, argsi) =>
                //    {
                //        var b = (ChromiumWebBrowser)s;
                //        if (argsi.Frame.IsMain)
                //        {
                //            b.SetZoomLevel(zoomLevel);
                //        }
                //    };
                //}
                await LoadPageAsync(browser);

                //Check preferences on the CEF UI Thread
                await Cef.UIThreadTaskFactory.StartNew(delegate
                {
                    var preferences = requestContext.GetAllPreferences(true);

                    //Check do not track status
                    var doNotTrack = (bool)preferences["enable_do_not_track"];

                    Debug.WriteLine("DoNotTrack:" + doNotTrack);
                });

                var onUi = Cef.CurrentlyOnThread(CefThreadIds.TID_UI);

                // For Google.com pre-pupulate the search text box
                await browser.EvaluateScriptAsync("document.getElementById('lst-ib').value = 'CefSharp Was Here!'");

                // Wait for the screenshot to be taken,
                // if one exists ignore it, wait for a new one to make sure we have the most up to date
                //await browser.ScreenshotAsync(true).ContinueWith(DisplayBitmap);

                await LoadPageAsync(browser, "https://shop.coles.com.au/a/a-national/everything/browse/drinks?pageNumber=1");
                Thread.Sleep(delayTimer);

                //Gets a wrapper around the underlying CefBrowser instance
                var cefBrowser = browser.GetBrowser();
                // Gets a warpper around the CefBrowserHost instance
                // You can perform a lot of low level browser operations using this interface
                var cefHost = cefBrowser.GetHost();

                //You can call Invalidate to redraw/refresh the image
                cefHost.Invalidate(PaintElementType.View);


                
                // Wait for the screenshot to be taken.
                cateLinkList = ColesExtraction.FirstCategoryLinkExtraction(await browser.GetSourceAsync());
                
                foreach (string a in cateLinkList)
                {
                    string pageLink = "";
                    List<string> pageLinkList = new List<string>();
                    pageLinkList.Add(a);
                    await LoadPageAsync(browser, shopUrl + a);
                    Thread.Sleep(delayTimer);
                    pageLink = ColesExtraction.PageExtraction(await browser.GetSourceAsync());

                    if (pageLink != null && pageLink != "")
                    {
                        pageLinkList.Add(pageLink);
                    }

                    Console.WriteLine("************************category link list*************************");
                    Console.WriteLine(shopUrl + a);
                    Console.WriteLine("********************************************************************");
                    Console.WriteLine("************************First page link*************************");
                    Console.WriteLine(shopUrl + pageLink);
                    Console.WriteLine("********************************************************************");

                    while (pageLink != null)
                    {

                        await LoadPageAsync(browser, shopUrl + pageLink);
                        Thread.Sleep(delayTimer);
                        pageLink = ColesExtraction.PageExtraction(await browser.GetSourceAsync());
                        if (pageLink != null)
                        {
                            Console.WriteLine("****************Page Link******************************");
                            Console.WriteLine(shopUrl + pageLink);
                            Console.WriteLine("********************************************************");

                            pageLinkList.Add(pageLink);
                        }
                    }

                    foreach (string b in pageLinkList)
                    {
                        await LoadPageAsync(browser, shopUrl + b);
                        Thread.Sleep(delayTimer);
                        Console.WriteLine("*****************Page link list****************************");
                        Console.WriteLine(shopUrl + b);
                        Console.WriteLine("***********************************************************");

                        productPageLinkList.AddRange(ColesExtraction.ProductLinkExtraction(await browser.GetSourceAsync()));
                    }
                }

                foreach (string c in productPageLinkList)
                {
                    await LoadPageAsync(browser, shopUrl + c);
                    Thread.Sleep(delayTimer);
                    Console.WriteLine("*****************Product detail page link list****************************");
                    Console.WriteLine(shopUrl + c);
                    Console.WriteLine("***********************************************************");
                    dataList.Add(ColesExtraction.DetailPageExtraction(await browser.GetSourceAsync()));
                }
            
            
                ExcelExport.AutomateExcel(dataList);
            }
        }

        private static async void WoolworthsTestMainAsync(string cachePath)
        {
            int delayTimer = 2500;
            List<string> cateLinkList = new List<string>();
            List<string> productPageLinkList = new List<string>();
            List<Data> dataList = new List<Data>();
            var browserSettings = new BrowserSettings();
            //Reduce rendering speed to one frame per second so it's easier to take screen shots
            browserSettings.WindowlessFrameRate = 1;
            var requestContextSettings = new RequestContextSettings { CachePath = cachePath };

            // RequestContext can be shared between browser instances and allows for custom settings
            // e.g. CachePath
            using (var requestContext = new RequestContext(requestContextSettings))
            using (var browser = new ChromiumWebBrowser(TestUrl, browserSettings, requestContext))
            {
                //if (zoomLevel > 1)
                //{
                //    browser.FrameLoadStart += (s, argsi) =>
                //    {
                //        var b = (ChromiumWebBrowser)s;
                //        if (argsi.Frame.IsMain)
                //        {
                //            b.SetZoomLevel(zoomLevel);
                //        }
                //    };
                //}
                await LoadPageAsync(browser);

                //Check preferences on the CEF UI Thread
                await Cef.UIThreadTaskFactory.StartNew(delegate
                {
                    var preferences = requestContext.GetAllPreferences(true);

                    //Check do not track status
                    var doNotTrack = (bool)preferences["enable_do_not_track"];

                    Debug.WriteLine("DoNotTrack:" + doNotTrack);
                });

                var onUi = Cef.CurrentlyOnThread(CefThreadIds.TID_UI);

                // For Google.com pre-pupulate the search text box
                await browser.EvaluateScriptAsync("document.getElementById('lst-ib').value = 'CefSharp Was Here!'");

                // Wait for the screenshot to be taken,
                // if one exists ignore it, wait for a new one to make sure we have the most up to date
                //await browser.ScreenshotAsync(true).ContinueWith(DisplayBitmap);

                await LoadPageAsync(browser, "https://www.woolworths.com.au/Shop/Browse/drinks/black-tea");
                Thread.Sleep(delayTimer);

                //Gets a wrapper around the underlying CefBrowser instance
                var cefBrowser = browser.GetBrowser();
                // Gets a warpper around the CefBrowserHost instance
                // You can perform a lot of low level browser operations using this interface
                var cefHost = cefBrowser.GetHost();

                string tempSource = await browser.GetSourceAsync();

                cateLinkList = WoolworthsExtraction.FirstCategoryLinkExtraction(tempSource);

            }
        }


        private static void dataExportForTXT(List<Data> dataList)
        {
            if (!File.Exists("DemoDataTxt.txt"))
            {
                FileStream fs1 = new FileStream("DemoDataTxt.txt", FileMode.Create, FileAccess.Write);
                StreamWriter sw = new StreamWriter(fs1);
                foreach (Data a in dataList)
                {
                    sw.WriteLine(a.ToString());
                }
                sw.Close();
                fs1.Close();
            }
            else
            {
                FileStream fs = new FileStream("DemoDataTxt.txt", FileMode.Open, FileAccess.Write);
                StreamWriter sr = new StreamWriter(fs);
                foreach (Data b in dataList)
                {
                    sr.WriteLine(b.ToString());
                }
                sr.Close();
                fs.Close();
            }
        }



        public static Task LoadPageAsync(IWebBrowser browser, string address = null)
        {
            //If using .Net 4.6 then use TaskCreationOptions.RunContinuationsAsynchronously
            //and switch to tcs.TrySetResult below - no need for the custom extension method
            var tcs = new TaskCompletionSource<bool>();

            EventHandler<LoadingStateChangedEventArgs> handler = null;
            handler = (sender, args) =>
            {
                //Wait for while page to finish loading not just the first frame
                if (!args.IsLoading)
                {
                    browser.LoadingStateChanged -= handler;
                    //This is required when using a standard TaskCompletionSource
                    //Extension method found in the CefSharp.Internals namespace
                    tcs.TrySetResultAsync(true);
                }
            };

            browser.LoadingStateChanged += handler;

            if (!string.IsNullOrEmpty(address))
            {
                browser.Load(address);
            }
            return tcs.Task;
        }


        static void DisplayMenu()
        {
            // Inedex
            Console.Clear();
            Console.WriteLine("---------> menu <---------");
            Console.WriteLine("(1)Demo Extraction");
            Console.WriteLine("(2)Coles Beverage Full Extraction");
            Console.WriteLine("(3)Woolworth Beverage Test Extraction");
            Console.WriteLine("(X)Exit");
            Console.WriteLine("");
            Console.Write("Choose a selection: ");
            // Case Function For Index Choice
        }

    }
}