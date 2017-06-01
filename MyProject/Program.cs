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
        //preset url
        private const string TestUrl = "https://www.google.com/";
        private const string shopUrl = "https://shop.coles.com.au";
        private const string shopUrl_WWS = "https://www.woolworths.com.au";


        public static int Main(string[] args)
        {
            bool isRunning = true; // For exit function
            string MenuSelection = ""; // For MenuSelection Case Function
            string Response = "";//read choice of menu

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
                    case "4":
                        WoolworthsRealTestMainAsync("cachePatch3");
                        break;
                    case "5":
                        Summary.SpecialProductSummary();
                        Console.ReadKey();
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
            //parameters
            int delayTimer = 50;
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


                // Wait for the screenshot to be taken,


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


                //extract first page
                bool loadfail = false;
                while (loadfail == false)
                {
                    try
                    {
                        await LoadPageAsync(browser, shopUrl + a);
                        Thread.Sleep(delayTimer);
                        pageLink = ColesExtraction.PageExtraction(await browser.GetSourceAsync());
                        loadfail = true;
                    }
                    catch (Exception e)
                    {
                        delayTimer += 50;
                        Console.WriteLine("***********************************************************");
                        Console.WriteLine("Error: " + e);
                        Console.WriteLine("***********************************************************");
                        Console.WriteLine("Page Loading Fail, Web is Reloading, or Set up delay time");
                        Console.WriteLine("***********************************************************");
                        loadfail = false;
                    }
                }



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
                    int delayTimer1 = 50;
                    bool loadfail1 = false;
                    while (loadfail1 == false)
                    {
                        try
                        {
                            await LoadPageAsync(browser, shopUrl + pageLink);
                            Thread.Sleep(delayTimer1);
                            pageLink = ColesExtraction.PageExtraction(await browser.GetSourceAsync());
                            loadfail1 = true;
                        }
                        catch (Exception e)
                        {
                            delayTimer1 += 50;
                            Console.WriteLine("***********************************************************");
                            Console.WriteLine("Error: " + e);
                            Console.WriteLine("***********************************************************");
                            Console.WriteLine("Page Loading Fail, Web is Reloading, or Set up delay time");
                            Console.WriteLine("***********************************************************");
                            loadfail1 = false;
                        }
                    }


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
                    int delayTimer1 = 50;
                    bool loadfail1 = false;
                    while (loadfail1 == false)
                    {
                        try
                        {
                            await LoadPageAsync(browser, shopUrl + b);
                            Thread.Sleep(delayTimer1);
                            Console.WriteLine("*****************Page link list****************************");
                            Console.WriteLine(shopUrl + b);
                            Console.WriteLine("***********************************************************");

                            productPageLinkList.AddRange(ColesExtraction.ProductLinkExtraction(await browser.GetSourceAsync()));
                            loadfail1 = true;
                        }
                        catch (Exception e)
                        {
                            delayTimer1 += 50;
                            Console.WriteLine("***********************************************************");
                            Console.WriteLine("Error: " + e);
                            Console.WriteLine("***********************************************************");
                            Console.WriteLine("Page Loading Fail, Web is Reloading, or Set up delay time");
                            Console.WriteLine("***********************************************************");
                            loadfail1 = false;
                        }
                    }


                }

                foreach (string c in productPageLinkList)
                {
                    int delayTimer1 = 50;
                    bool loadfail1 = false;
                    while (loadfail1 == false)
                    {
                        try
                        {
                            await LoadPageAsync(browser, shopUrl + c);
                            Thread.Sleep(delayTimer1);
                            Console.WriteLine("*****************Product detail page link list****************************");
                            Console.WriteLine(shopUrl + c);
                            Console.WriteLine("***********************************************************");
                            dataList.Add(ColesExtraction.DetailPageExtraction(await browser.GetSourceAsync(), shopUrl + c));
                            loadfail1 = true;
                        }
                        catch(Exception e)
                        {
                            delayTimer1 += 50;
                            Console.WriteLine("***********************************************************");
                            Console.WriteLine("Error: " + e);
                            Console.WriteLine("***********************************************************");
                            Console.WriteLine("Page Loading Fail, Web is Reloading, or Set up delay time");
                            Console.WriteLine("***********************************************************");
                            loadfail1 = false;
                        }
                    }

                }
                //demo end
                ExcelExport.AutomateExcel_Test(dataList);
            }
        }

        private static async void FullColesMainAsync(string cachePath)
        {
            //parameters
            int delayTimer = 0;
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

                //Gets a wrapper around the underlying CefBrowser instance
                var cefBrowser = browser.GetBrowser();
                // Gets a warpper around the CefBrowserHost instance
                // You can perform a lot of low level browser operations using this interface
                var cefHost = cefBrowser.GetHost();

                //You can call Invalidate to redraw/refresh the image
                cefHost.Invalidate(PaintElementType.View);

                //all drink category links extraction
                bool loadFail = false;
                int addedTimer = 100;
                while(loadFail == false)
                {
                    try
                    {
                        //web page load and extract
                        await LoadPageAsync(browser, "https://shop.coles.com.au/a/a-national/everything/browse/drinks?pageNumber=1");
                        Thread.Sleep(delayTimer);

                        cateLinkList = ColesExtraction.FirstCategoryLinkExtraction(await browser.GetSourceAsync());

                        //loading flag
                        loadFail = true;
                    }
                    catch(Exception e)
                    {
                        //error action
                        delayTimer += addedTimer;
                        Console.WriteLine("***********************************************************");
                        Console.WriteLine("Error: " + e);
                        Console.WriteLine("***********************************************************");
                        Console.WriteLine("Page Loading Fail, Web is Reloading, or Set up delay time");
                        Console.WriteLine("***********************************************************");
                        loadFail = false;
                    }
                }

                //start to extract each page link of each category
                foreach (string a in cateLinkList)
                {
                    string pageLink = "";
                    List<string> pageLinkList = new List<string>();
                    pageLinkList.Add(a);
                    //loop for reload web
                    int delayTimerpagelink = 10;
                    bool loadFailpagelink = false;
                    int addedTimerpagelink = 100;
                    //start from first page of category
                    while (loadFailpagelink == false)
                    {
                        try
                        {
                            //web page load and extract
                            await LoadPageAsync(browser, shopUrl + a);
                            Thread.Sleep(delayTimerpagelink);
                            pageLink = ColesExtraction.PageExtraction(await browser.GetSourceAsync());

                            //loading flag
                            loadFailpagelink = true;
                        }
                        catch (Exception e)
                        {
                            //error action
                            delayTimerpagelink += addedTimerpagelink;
                            Console.WriteLine("***********************************************************");
                            Console.WriteLine("Error: " + e);
                            Console.WriteLine("***********************************************************");
                            Console.WriteLine("Page Loading Fail, Web is Reloading, or Set up delay time");
                            Console.WriteLine("***********************************************************");
                            loadFailpagelink = false;
                        }
                    }



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

                    //extract follow next page of current category
                    while (pageLink != null)
                    {
                        //loop for reload web
                        int delayTimerpagelink1 = 10;
                        bool loadFailpagelink1 = false;
                        int addedTimerpagelink1 = 100;
                        while (loadFailpagelink1 == false)
                        {
                            try
                            {
                                //web page load and extract
                                await LoadPageAsync(browser, shopUrl + pageLink);
                                Thread.Sleep(delayTimerpagelink1);
                                pageLink = ColesExtraction.PageExtraction(await browser.GetSourceAsync());

                                //loading flag
                                loadFailpagelink1 = true;
                            }
                            catch (Exception e)
                            {
                                //error action
                                delayTimerpagelink1 += addedTimerpagelink1;
                                Console.WriteLine("***********************************************************");
                                Console.WriteLine("Error: " + e);
                                Console.WriteLine("***********************************************************");
                                Console.WriteLine("Page Loading Fail, Web is Reloading, or Set up delay time");
                                Console.WriteLine("***********************************************************");
                                loadFailpagelink1 = false;
                            }
                        }




                        if (pageLink != null)
                        {
                            Console.WriteLine("****************Page Link******************************");
                            Console.WriteLine(shopUrl + pageLink);
                            Console.WriteLine("********************************************************");

                            pageLinkList.Add(pageLink);
                        }
                    }

                    //extract each product link from all page
                    foreach (string b in pageLinkList)
                    {
                        //loop for reload web
                        int delayTimerpagelist = 10;
                        bool loadFailpagelist = false;
                        int addedTimerpagelist = 100;
                        while (loadFailpagelist == false)
                        {
                            try
                            {
                                //web page load and extract
                                await LoadPageAsync(browser, shopUrl + b);
                                Thread.Sleep(delayTimerpagelist);
                                Console.WriteLine("*****************Page link list****************************");
                                Console.WriteLine(shopUrl + b);
                                Console.WriteLine("***********************************************************");

                                productPageLinkList.AddRange(ColesExtraction.ProductLinkExtraction(await browser.GetSourceAsync()));

                                //loading flag
                                loadFailpagelist = true;
                            }
                            catch (Exception e)
                            {
                                //error action
                                delayTimerpagelist += addedTimerpagelist;
                                Console.WriteLine("***********************************************************");
                                Console.WriteLine("Error: " + e);
                                Console.WriteLine("***********************************************************");
                                Console.WriteLine("Page Loading Fail, Web is Reloading, or Set up delay time");
                                Console.WriteLine("***********************************************************");
                                loadFailpagelist = false;
                            }
                        }


                    }
                }

                //get in each product detail page by product link to extract product detail
                foreach (string c in productPageLinkList)
                {
                    //loop for reload web
                    int delayTimerpagelist = 10;
                    bool loadFailpagelist = false;
                    int addedTimerpagelist = 100;
                    while (loadFailpagelist == false)
                    {
                        try
                        {
                            //web page load and extract
                            await LoadPageAsync(browser, shopUrl + c);
                            Thread.Sleep(delayTimerpagelist);
                            Console.WriteLine("*****************Product detail page link list****************************");
                            Console.WriteLine(shopUrl + c);
                            Console.WriteLine("***********************************************************");
                            dataList.Add(ColesExtraction.DetailPageExtraction(await browser.GetSourceAsync(), shopUrl + c));

                            //loading flag
                            loadFailpagelist = true;
                        }
                        catch (Exception e)
                        {
                            //error action
                            delayTimerpagelist += addedTimerpagelist;
                            Console.WriteLine("***********************************************************");
                            Console.WriteLine("Error: " + e);
                            Console.WriteLine("***********************************************************");
                            Console.WriteLine("Page Loading Fail, Web is Reloading, or Set up delay time");
                            Console.WriteLine("***********************************************************");
                            loadFailpagelist = false;
                        }
                    }


                }

                //Export Extraction as Excel file
                ExcelExport.AutomateExcel(dataList);
            }
        }

        private static async void WoolworthsTestMainAsync(string cachePath)
        {
            //parameter
            int delayTimer = 2000;
            List<string> pageLinkList = new List<string>();
            List<string> cateLinkList = new List<string>();
            List<string> productPageLinkList = new List<string>();
            List<Data_WWs> dataList = new List<Data_WWs>();
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

                


                //Gets a wrapper around the underlying CefBrowser instance
                var cefBrowser = browser.GetBrowser();
                // Gets a warpper around the CefBrowserHost instance
                // You can perform a lot of low level browser operations using this interface
                var cefHost = cefBrowser.GetHost();

                //reload section
                bool loadFail = false;
                int addedTimer = 2000;
                //extract each subcategory link
                while (loadFail == false)
                {
                    try
                    {
                        //web page load and extract
                        await LoadPageAsync(browser, "https://www.woolworths.com.au/Shop/Browse/drinks/black-tea");
                        Thread.Sleep(delayTimer);
                        string tempSource = await browser.GetSourceAsync();
                        cateLinkList = WoolworthsExtraction.FirstCategoryLinkExtraction(tempSource);

                        //loading flag
                        loadFail = true;
                    }
                    catch (Exception e)
                    {
                        //error action
                        delayTimer += addedTimer;
                        Console.WriteLine("***********************************************************");
                        Console.WriteLine("Error: " + e);
                        Console.WriteLine("***********************************************************");
                        Console.WriteLine("Page Loading Fail, Web is Reloading, or Set up delay time");
                        Console.WriteLine("***********************************************************");
                        loadFail = false;
                    }
                }

                //extract each page of each subcategory
                foreach (string a in cateLinkList)
                {
                    string pageLink = "";
                    pageLinkList.Add(a);

                    //reload section
                    int delayTimercatelink = 2000;
                    bool loadFailcatelink = false;
                    int addedTimercatelink = 2000;
                    while (loadFailcatelink == false)
                    {
                        try
                        {
                            //web page load and extract
                            await LoadPageAsync(browser, shopUrl_WWS + a);
                            Thread.Sleep(delayTimercatelink);
                            pageLink = WoolworthsExtraction.PageExtraction(await browser.GetSourceAsync());

                            //loading flag
                            loadFailcatelink = true;
                        }
                        catch (Exception e)
                        {
                            //error action
                            delayTimercatelink += addedTimercatelink;
                            Console.WriteLine("***********************************************************");
                            Console.WriteLine("Error: " + e);
                            Console.WriteLine("***********************************************************");
                            Console.WriteLine("Page Loading Fail, Web is Reloading, or Set up delay time");
                            Console.WriteLine("***********************************************************");
                            loadFailcatelink = false;
                        }
                    }



                    if (pageLink != null && pageLink != "")
                    {
                        pageLinkList.Add(pageLink);
                    }

                    Console.WriteLine("************************First Page Link*************************");
                    Console.WriteLine(shopUrl_WWS + a);
                    Console.WriteLine("********************************************************************");
                    if (pageLink != null && pageLink != "")
                    { 

                    Console.WriteLine("************************Second page link*************************");
                    Console.WriteLine(shopUrl_WWS + pageLink);
                    Console.WriteLine("********************************************************************");

                    }

                    while (pageLink != null)
                    {
                        //reload section
                        int delayTimerpagelink = 2000;
                        bool loadFailpagelink = false;
                        int addedTimerpagelink = 2000;
                        while (loadFailpagelink == false)
                        {
                            try
                            {
                                //web page load and extract
                                await LoadPageAsync(browser, shopUrl_WWS + pageLink);
                                Thread.Sleep(delayTimerpagelink);
                                pageLink = WoolworthsExtraction.PageExtraction(await browser.GetSourceAsync());


                                //loading flag
                                loadFailpagelink = true;
                            }
                            catch (Exception e)
                            {
                                //error action
                                delayTimerpagelink += addedTimerpagelink;
                                Console.WriteLine("***********************************************************");
                                Console.WriteLine("Error: " + e);
                                Console.WriteLine("***********************************************************");
                                Console.WriteLine("Page Loading Fail, Web is Reloading, or Set up delay time");
                                Console.WriteLine("***********************************************************");
                                loadFailpagelink = false;
                            }
                        }

                        if (pageLink != null)
                        {
                            Console.WriteLine("****************Page Link******************************");
                            Console.WriteLine(shopUrl_WWS + pageLink);
                            Console.WriteLine("********************************************************");

                            pageLinkList.Add(pageLink);
                        }
                    }


                }

                //extract product detail from each subcategory page
                foreach (string b in pageLinkList)
                {
                    //reload section
                    int delayTimerpagelinklist = 2000;
                    bool loadFailpagelinklist = false;
                    int addedTimerpagelinklist = 2000;
                    while (loadFailpagelinklist == false)
                    {
                        try
                        {
                            //web page load and extract
                            await LoadPageAsync(browser, shopUrl_WWS + b);
                            Thread.Sleep(delayTimerpagelinklist);
                            string pageSource = await browser.GetSourceAsync();
                            Console.WriteLine("*****************Product detail****************************");
                            Console.WriteLine(shopUrl_WWS + b);
                            Console.WriteLine("***********************************************************");
                            dataList.AddRange(WoolworthsExtraction.DetailExtraction(pageSource));
                            //loading flag
                            loadFailpagelinklist = true;
                        }
                        catch (Exception e)
                        {
                            //error action
                            delayTimerpagelinklist += addedTimerpagelinklist;
                            Console.WriteLine("***********************************************************");
                            Console.WriteLine("Error: " + e);
                            Console.WriteLine("***********************************************************");
                            Console.WriteLine("Page Loading Fail, Web is Reloading, or Set up delay time");
                            Console.WriteLine("***********************************************************");
                            loadFailpagelinklist = false;
                        }
                    }
                }

            }
            //export data as Excel file
            ExcelExport.AutomateExcel_WWs(dataList);

        }

        private static async void WoolworthsRealTestMainAsync(string cachePath)
        {
            //parameter
            int delayTimer = 2000;
            List<string> pageLinkList = new List<string>();
            List<Data_WWs> cateLinkList = new List<Data_WWs>();
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




                //Gets a wrapper around the underlying CefBrowser instance
                var cefBrowser = browser.GetBrowser();
                // Gets a warpper around the CefBrowserHost instance
                // You can perform a lot of low level browser operations using this interface
                var cefHost = cefBrowser.GetHost();

                //reload section
                bool loadFail = false;
                int addedTimer = 2000;

                //extract all product detail from current subcategory link
                while (loadFail == false)
                {
                    try
                    {
                        //web page load and extract
                        await LoadPageAsync(browser, "https://www.woolworths.com.au/Shop/Browse/drinks/energy-drinks");
                        Thread.Sleep(delayTimer);
                        string tempSource = await browser.GetSourceAsync();
                        cateLinkList = WoolworthsExtraction.DetailExtraction(tempSource);

                        //loading flag
                        loadFail = true;
                    }
                    catch (Exception e)
                    {
                        //error action
                        delayTimer += addedTimer;
                        Console.WriteLine("***********************************************************");
                        Console.WriteLine("Error: " + e);
                        Console.WriteLine("***********************************************************");
                        Console.WriteLine("Page Loading Fail, Web is Reloading, or Set up delay time");
                        Console.WriteLine("***********************************************************");
                        loadFail = false;
                    }
                }
            }

            //export as Excel file
            ExcelExport.AutomateExcel_WWsTest(cateLinkList);
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
            Console.WriteLine("(3)Woolworth Beverage Full Extraction");
            Console.WriteLine("(4)Woolworth TEST");
            Console.WriteLine("(5)Woolworth Excel Extract TEST");
            Console.WriteLine("(X)Exit");
            Console.WriteLine("");
            Console.Write("Choose a selection: ");
            // Case Function For Index Choice
        }

    }
}