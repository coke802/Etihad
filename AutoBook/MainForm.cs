using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Runtime.InteropServices;
using Microsoft.Win32;
using System.Threading;
using System.Diagnostics;
using System.IO;
using System.Xml;

namespace AutoBook
{
    public partial class MainForm : Form
    {
        public struct BookSetting
        {
            public int title;
            public string firstName;
            public string lastName;
            public string telephone;
            public string email;
            public int flightType;
        };

        [DllImport("user32.dll")]
        public static extern int ShowWindow(IntPtr IntPtr, uint nCmdShow);

        [DllImport("user32.dll")]
        public static extern int SetActiveWindow(IntPtr IntPtr);

        [DllImport("USER32.DLL")]
        public static extern bool SetForegroundWindow(IntPtr IntPtr);
        [DllImport("user32.dll")]
        public static extern int SystemParametersInfo(int uAction, int uParam, IntPtr lpvParam, int fuWinIni);

        [DllImport("urlmon.dll", CharSet = CharSet.Ansi)]
        private static extern int UrlMkSetSessionOption(int dwOption, string pBuffer, int dwBufferLength, int dwReserved);

        [DllImport("Kernel32.dll")]
        public static extern bool Beep(int frequency, int duration);

        const int URLMON_OPTION_USERAGENT = 0x10000001;
        const string url = "https://www.etihad.com/zh-cn/";

        bool stopped = true;
        bool injected = false;
        int step = 0;
        int sleep = 0;
        int beep_count = 0;
        int elementPositionX;
        int elementPositionY;
        List<HtmlElement> elemFlights = null;


        public MainForm()
        {
            InitializeComponent();

            comboBox1.Text = "";
            elementPositionX = 0;
            elementPositionY = 0;
            elemFlights = new List<HtmlElement>();
        }

        public static void ChangeUserAgent(string userAgent)
        {
            UrlMkSetSessionOption(URLMON_OPTION_USERAGENT, userAgent, userAgent.Length, 0);
        }

        static void SetWebBrowserFeatures()
        {
            // don't change the registry if running in-proc inside Visual Studio
            if (LicenseManager.UsageMode != LicenseUsageMode.Runtime)
                return;

            var appName = System.IO.Path.GetFileName(System.Diagnostics.Process.GetCurrentProcess().MainModule.FileName);

            var featureControlRegKey = @"HKEY_CURRENT_USER\Software\Microsoft\Internet Explorer\Main\FeatureControl\";

            Registry.SetValue(featureControlRegKey + "FEATURE_BROWSER_EMULATION",
                appName, GetBrowserEmulationMode(), RegistryValueKind.DWord);

            // enable the features which are "On" for the full Internet Explorer browser

            Registry.SetValue(featureControlRegKey + "FEATURE_ENABLE_CLIPCHILDREN_OPTIMIZATION",
                appName, 1, RegistryValueKind.DWord);

            Registry.SetValue(featureControlRegKey + "FEATURE_AJAX_CONNECTIONEVENTS",
                appName, 1, RegistryValueKind.DWord);

            Registry.SetValue(featureControlRegKey + "FEATURE_GPU_RENDERING",
                appName, 1, RegistryValueKind.DWord);

            Registry.SetValue(featureControlRegKey + "FEATURE_WEBOC_DOCUMENT_ZOOM",
                appName, 1, RegistryValueKind.DWord);

            Registry.SetValue(featureControlRegKey + "FEATURE_NINPUT_LEGACYMODE",
                appName, 0, RegistryValueKind.DWord);
        }

        static UInt32 GetBrowserEmulationMode()
        {
            int browserVersion = 0;
            using (var ieKey = Registry.LocalMachine.OpenSubKey(@"SOFTWARE\Microsoft\Internet Explorer",
                RegistryKeyPermissionCheck.ReadSubTree,
                System.Security.AccessControl.RegistryRights.QueryValues))
            {
                var version = ieKey.GetValue("svcVersion");
                if (null == version)
                {
                    version = ieKey.GetValue("Version");
                    if (null == version)
                        throw new ApplicationException("Microsoft Internet Explorer is required!");
                }
                int.TryParse(version.ToString().Split('.')[0], out browserVersion);
            }

            if (browserVersion < 7)
            {
                throw new ApplicationException("Unsupported version of Microsoft Internet Explorer!");
            }

            UInt32 mode = 11000; // Internet Explorer 11. Webpages containing standards-based !DOCTYPE directives are displayed in IE11 Standards mode. 

            switch (browserVersion)
            {
                case 7:
                    mode = 7000; // Webpages containing standards-based !DOCTYPE directives are displayed in IE7 Standards mode. 
                    break;
                case 8:
                    mode = 8000; // Webpages containing standards-based !DOCTYPE directives are displayed in IE8 mode. 
                    break;
                case 9:
                    mode = 9000; // Internet Explorer 9. Webpages containing standards-based !DOCTYPE directives are displayed in IE9 mode.                    
                    break;
                case 10:
                    mode = 10000; // Internet Explorer 10.
                    break;
            }

            return mode;
        }

        private void OnLoad(object sender, EventArgs e)
        {
            BookSetting bookSetting = LoadFromXml();
            SetBookSetting(bookSetting);

            SetWebBrowserFeatures();
            //ChangeUserAgent("Dalvik/2.1.0 (Linux; U; Android 9; MI 8 Lite MIUI/V10.3.2.0.PDTCNXM)");
            //ChangeUserAgent("Mozilla/5.0 (iPhone; CPU iPhone OS 11_4_1 like Mac OS X) AppleWebKit/605.1.15 (KHTML, like Gecko) Version/11.0 Mobile/15E148 Safari/604.1");
            webBrowser1.ScriptErrorsSuppressed = true;

            webBrowser1.Navigate(url);
        }

        private BookSetting LoadFromXml()
        {
            BookSetting bookSetting;

            XmlDocument xmlDoc = new XmlDocument();
            xmlDoc.Load("AutoBook.xml");
            XmlNode dataNode1 = xmlDoc.SelectSingleNode("Data");
            bookSetting.title = Convert.ToInt32(dataNode1.SelectSingleNode("Title").InnerText);
            bookSetting.firstName = dataNode1.SelectSingleNode("FirstName").InnerText;
            bookSetting.lastName = dataNode1.SelectSingleNode("LastName").InnerText;
            bookSetting.telephone = dataNode1.SelectSingleNode("Telephone").InnerText;
            bookSetting.email = dataNode1.SelectSingleNode("Email").InnerText;
            bookSetting.flightType = Convert.ToInt32(dataNode1.SelectSingleNode("FlightType").InnerText);

            return bookSetting;
        }

        private void SaveToXml(BookSetting bookSetting)
        {
            XmlDocument xmlDoc = new XmlDocument();
            xmlDoc.Load("AutoBook.xml");

            XmlNode dataNode1 = xmlDoc.SelectSingleNode("Data");
            dataNode1.SelectSingleNode("Title").InnerText = bookSetting.title.ToString();
            dataNode1.SelectSingleNode("FirstName").InnerText = bookSetting.firstName;
            dataNode1.SelectSingleNode("LastName").InnerText = bookSetting.lastName;
            dataNode1.SelectSingleNode("Telephone").InnerText = bookSetting.telephone;
            dataNode1.SelectSingleNode("Email").InnerText = bookSetting.email;
            dataNode1.SelectSingleNode("FlightType").InnerText = bookSetting.flightType.ToString();
            xmlDoc.Save("AutoBook.xml");
        }

        private BookSetting GetBookSetting()
        {
            BookSetting bookSetting;
            bookSetting.title = comboBox1.SelectedIndex;
            bookSetting.firstName = textBox1.Text;
            bookSetting.lastName = textBox2.Text;
            bookSetting.telephone = textBox3.Text;
            bookSetting.email = textBox4.Text;
            bookSetting.flightType = comboBox2.SelectedIndex;
            return bookSetting;
        }

        private void SetBookSetting(BookSetting bookSetting)
        {
            comboBox1.SelectedIndex = bookSetting.title;
            textBox1.Text = bookSetting.firstName;
            textBox2.Text = bookSetting.lastName;
            textBox3.Text = bookSetting.telephone;
            textBox4.Text = bookSetting.email;
            comboBox2.SelectedIndex = bookSetting.flightType;
        }

        private void DocumentCompleted(object sender, WebBrowserDocumentCompletedEventArgs e)
        {
            string url = e.Url.AbsoluteUri.ToString();
            UtilsLog.Log("DocumentCompleted url={0}", url);
        }

        private void button1_Click(object sender, EventArgs e)
        {
            if (comboBox1.Text.Length == 0)
            {
                Form1 form1 = new Form1();
                form1.ShowDialog();
                return;
            }

            Run();
        }

        private void button2_Click(object sender, EventArgs e)
        {
            Stop();
        }

        private void button3_Click(object sender, EventArgs e)
        {
            webBrowser1.Navigate(url);
        }

        private void button4_Click(object sender, EventArgs e)
        {
            BookSetting bookSetting = GetBookSetting();
            SaveToXml(bookSetting);
        }

        public void Run()
        {
            if (stopped)
            {
                stopped = false;
                Thread t = new Thread(RunLoop);
                t.Start();
            }
        }

        public void Stop()
        {
            if (!stopped)
            {
                stopped = true;
            }
        }

        public void RunLoopSleep(int timeout)
        {
            UtilsLog.Log("RunLoopSleep ENTER {0}", timeout);
            sleep += timeout;
        }

        public void RunLoopBegin()
        {
            step = 0;
            sleep = 0;
        }

        public void RunLoopEnd()
        {
        }

        public delegate void InvokeWebBrowser();
        public void RunLoop()
        {
            RunLoopBegin();

            while (!stopped)
            {
                if (sleep > 0)
                {
                    if (sleep > 1000)
                    {
                        sleep -= 1000;
                    }
                    else
                    {
                        sleep = 0;
                    }
                }
                else
                {
                    webBrowser1.BeginInvoke(new InvokeWebBrowser(RunLoopMainEx));
                }

                Thread.Sleep(1000);
            }

            RunLoopEnd();
        }

        private void RunLoopMainEx()
        {
            try
            {
                RunLoopMain();
            }
            catch (Exception e)
            {
                UtilsLog.Log("RunLoopMainEx EXCEPTION" + e.Message);
            }
        }

        private void RunLoopMain()
        {
            switch (step)
            {
                case 0:
                    RunLoopMainStep0();
                    break;
                case 1:
                    RunLoopMainStep1();
                    break;
                case 2:
                    RunLoopMainStep2();
                    break;
                case 3:
                    RunLoopMainStep3();
                    break;
                case 4:
                    RunLoopMainStep4();
                    break;
                case 5:
                    RunLoopMainStep5();
                    break;
                case 6:
                    RunLoopMainStep6();
                    break;
                case 7:
                    RunLoopMainStep7();
                    break;
                case 8:
                    RunLoopMainStep8();
                    break;
                case 9:
                    RunLoopMainStep9();
                    break;
                case 10:
                    RunLoopMainStep10();
                    break;
                case 11:
                    RunLoopMainStep11();
                    break;
                case 12:
                    RunLoopMainStep12();
                    break;
                case 13:
                    RunLoopMainStep13();
                    break;
                case 14:
                    RunLoopMainStep14();
                    break;
                case 15:
                    RunLoopMainStep15();
                    break;
            }
        }


        private void InitJsScript()
        {
            if (injected)
            {
                return;
            }

            string funcName;
            string funcBody;

            funcName = "getWindowDevicePixelRatio";
            funcBody = "{return window.devicePixelRatio;}";
            CreateJsFunc(funcName, funcBody);

            funcName = "getDocumentScrollLeft";
            funcBody = "{return document.documentElement.scrollLeft;}";
            CreateJsFunc(funcName, funcBody);

            funcName = "getDocumentScrollTop";
            funcBody = "{return document.documentElement.scrollTop;}";
            CreateJsFunc(funcName, funcBody);

//             funcName = "getDocumentScreenLeft";
//             funcBody = "{return window.screenLeft;}";
//             CreateJsFunc(funcName, funcBody);
// 
//             funcName = "getDocumentScreenTop";
//             funcBody = "{return window.screenTop;}";
//             CreateJsFunc(funcName, funcBody);
// 
//             funcName = "getDocumentScrollLeft2";
//             funcBody = "{var e = document.getElementById(id); return e.getBoundingClientRect().left;}";
//             CreateJsFuncWithParam1(funcName, "id", funcBody);
// 
//             funcName = "getDocumentScrollTop2";
//             funcBody = "{var e = document.getElementById(id); return e.getBoundingClientRect().top;}";
//             CreateJsFuncWithParam1(funcName, "id", funcBody);
            
//             funcName = "testPrint";
//             funcBody = "{alert(x);}";
//             CreateJsFuncWithParam1(funcName, "x", funcBody);

            injected = true;
        }

        private double GetWindowDevicePixelRatio()
        {
            UtilsLog.Log("GetWindowDevicePixelRatio ENTER");

            object retngetWindowDevicePixelRatio = ExecJSFunc("getWindowDevicePixelRatio");
            UtilsLog.Log("GetWindowDevicePixelRatio retngetWindowDevicePixelRatio={0} {1}", retngetWindowDevicePixelRatio.GetType(), retngetWindowDevicePixelRatio.ToString());

            double devicePixelRatio = Convert.ToDouble(retngetWindowDevicePixelRatio);
            return devicePixelRatio;
        }

        private int GetDocumentScreenLeft()
        {
            object retnGetDocumentScreenLeft = ExecJSFunc("getDocumentScreenLeft");
            UtilsLog.Log("GetDocumentScreenLeft retnGetDocumentScreenLeft={0} {1}", retnGetDocumentScreenLeft.GetType(), retnGetDocumentScreenLeft.ToString());
            int screenLeft = Convert.ToInt32(retnGetDocumentScreenLeft);
            UtilsLog.Log("GetDocumentScreenLeft screenLeft={0}", screenLeft);
            return screenLeft;
        }

        private int GetDocumentScreenTop()
        {
            UtilsLog.Log("GetDocumentScreenTop ENTER");

            object retnGetDocumentScrollTop = ExecJSFunc("getDocumentScreenTop");
            UtilsLog.Log("GetDocumentScreenTop retnGetDocumentScrollTop={0} {1}", retnGetDocumentScrollTop.GetType(), retnGetDocumentScrollTop.ToString());
            int screenTop = Convert.ToInt32(retnGetDocumentScrollTop);
            UtilsLog.Log("GetDocumentScreenTop screenTop={0}", screenTop);
            return screenTop;
        }

        private int GetDocumentScrollLeft()
        {
            UtilsLog.Log("GetDocumentScrollLeft ENTER");

            object retnGetDocumentScrollLeft = ExecJSFunc("getDocumentScrollLeft");
            UtilsLog.Log("GetDocumentScrollLeft retnGetDocumentScrollLeft={0} {1}", retnGetDocumentScrollLeft.GetType(), retnGetDocumentScrollLeft.ToString());
            int scrollLeft = Convert.ToInt32(retnGetDocumentScrollLeft);
            UtilsLog.Log("GetDocumentScrollLeft scrollLeft={0}", scrollLeft);
            return scrollLeft;
        }

        private int GetDocumentScrollTop()
        {
            UtilsLog.Log("GetDocumentScrollTop ENTER");

            object retnGetDocumentScrollTop = ExecJSFunc("getDocumentScrollTop");
            UtilsLog.Log("GetDocumentScrollTop retnGetDocumentScrollTop={0} {1}", retnGetDocumentScrollTop.GetType(), retnGetDocumentScrollTop.ToString());
            int scrollTop = Convert.ToInt32(retnGetDocumentScrollTop);
            UtilsLog.Log("GetDocumentScrollTop scrollTop={0}", scrollTop);
            return scrollTop;
        }

        private int GetDocumentScrollLeft2(string elemId)
        {
            UtilsLog.Log("GetDocumentScrollLeft2 ENTER {0}", elemId);

            object retnGetDocumentScrollLeft = ExecJSFuncWithParam1("getDocumentScrollLeft2", elemId);
            UtilsLog.Log("GetDocumentScrollLeft2 retnGetDocumentScrollLeft={0} {1}", retnGetDocumentScrollLeft.GetType(), retnGetDocumentScrollLeft.ToString());

            double scrollLeft = Convert.ToDouble(retnGetDocumentScrollLeft);
            return (int)scrollLeft;
        }

        private int GetDocumentScrollTop2(string elemId)
        {
            UtilsLog.Log("GetDocumentScrollTop2 ENTER {0}", elemId);

            object retnGetDocumentScrollTop = ExecJSFuncWithParam1("getDocumentScrollTop2", elemId);
            UtilsLog.Log("GetDocumentScrollTop2 retnGetDocumentScrollTop={0} {1}", retnGetDocumentScrollTop.GetType(), retnGetDocumentScrollTop.ToString());

            double scrollTop = Convert.ToDouble(retnGetDocumentScrollTop);
            return (int)scrollTop;
        }

        private void ExecJsScript()
        {
            UtilsLog.Log("ExecJsScript ENTER");
            //GetWindowDevicePixelRatio();
            //GetDocumentScrollLeft();
            //GetDocumentScrollTop();
        }

        private void RunLoopMainStep0()
        {
            UtilsLog.Log("RunLoopMainStep0 ENTER");

            //InitJsScript();
            //ExecJsScript();

            HtmlElement elemoneWay = webBrowser1.Document.GetElementById("oneWay");
            if (elemoneWay == null)
            {
                return;
            }

            elemoneWay.Focus();
            elemoneWay.InvokeMember("click");
            step++;
        }

        private void RunLoopMainStep1()
        {
            UtilsLog.Log("RunLoopMainStep1 ENTER");

            List<HtmlElement> elems = new List<HtmlElement>();
            UtilsHtml.FindWindowHtmlElementsByTagAndText(webBrowser1.Document.Window, "button", "搜索航班", elems);
            if (elems.Count > 0)
            {
                elems[0].Focus();
                elems[0].InvokeMember("click");
                RunLoopSleep(2000);
                step++;
            }
        }

        private void RunLoopMainStep2()
        {
            UtilsLog.Log("RunLoopMainStep2 ENTER");

            HtmlElement eledxpdateselectionview = webBrowser1.Document.GetElementById("dxp-date-selection-view");
            if (eledxpdateselectionview != null)
            {
                List<HtmlElement> eledxpdateselectionviewItems = new List<HtmlElement>();
                UtilsHtml.FindElementHtmlElements(eledxpdateselectionview, ":div/:div/:div/:div/class:page-messages/:div/class:title-container/:div/class:title/:span", eledxpdateselectionviewItems);
                if (eledxpdateselectionviewItems.Count > 0)
                {
                    HtmlElement elemappItem = eledxpdateselectionviewItems[0];
                    if (elemappItem.InnerText.Contains("未找到航班"))
                    {
                        //失败
                        RunLoopSleep(5000);
                        step++;
                    }
                }
            }

            HtmlElement eledxpflighttablesection = webBrowser1.Document.GetElementById("dxp-flight-table-section");
            if (eledxpflighttablesection != null)
            {
                //成功
                RunLoopSleep(2000);
                step += 3;
            }
        }

        private void RunLoopMainStep3()
        {
            UtilsLog.Log("RunLoopMainStep3 ENTER");

            webBrowser1.Navigate(url);
            RunLoopSleep(3000);
            step++;
        }

        private void RunLoopMainStep4()
        {
            UtilsLog.Log("RunLoopMainStep4 ENTER");

            HtmlElement elemflightsearchcopy = webBrowser1.Document.GetElementById("flightsearch_copy");
            if (elemflightsearchcopy == null)
            {
                return;
            }

            step = 0;
        }


        private void RunLoopMainStep5()
        {
            UtilsLog.Log("RunLoopMainStep5 ENTER");
            HtmlElement eledxpflighttablesection = webBrowser1.Document.GetElementById("dxp-flight-table-section");
            if (eledxpflighttablesection == null)
            {
                return;
            }

            elemFlights.Clear();
            foreach (HtmlElement eledxpflighttablesectionChild in eledxpflighttablesection.Children)
            {
                string eledxpflighttablesectionChildClassName = eledxpflighttablesectionChild.GetAttribute("className");
                if (eledxpflighttablesectionChildClassName.Contains("ducp-component-panel")
                    && eledxpflighttablesectionChildClassName.Contains("spark-panel")
                    && eledxpflighttablesectionChildClassName.Contains("itinerary-part-offer"))
                {
                    elemFlights.Add(eledxpflighttablesectionChild);
                }
            }

            if (elemFlights.Count > 0)
            {
                if (comboBox2.Text == "商务舱")
                {
                    HtmlElement elemFlight = elemFlights[0];
                    List<HtmlElement> elemFlightItems = new List<HtmlElement>();
                    UtilsHtml.FindElementHtmlElements(elemFlight, ":div/:div/:div/:button", elemFlightItems);
                    foreach (HtmlElement elemFlightItem in elemFlightItems)
                    {
                        if (elemFlightItem.InnerText.Contains("下一個"))
                        {
                            elemFlightItem.Focus();
                            elemFlightItem.InvokeMember("click");
                        }
                    }
                }

                RunLoopSleep(2000);
                step++;
            }
        }

        private void RunLoopMainStep6()
        {
            UtilsLog.Log("RunLoopMainStep6 ENTER");

            HtmlElement elemFlight = elemFlights[0];

            List<HtmlElement> elemFlightItems = new List<HtmlElement>();
            UtilsHtml.FindElementHtmlElements(elemFlight, ":div/:div/:div/class:itinerary-part-offer-price/class:price-content-wrapper/data-test-id:cabin-offer-button", elemFlightItems);
            UtilsLog.Log("RunLoopMainStep6 elemFlightItems.Count={0}", elemFlightItems.Count);
            if (elemFlightItems.Count > 0)
            {
                elemFlightItems[0].Focus();
                elemFlightItems[0].InvokeMember("click");
                RunLoopSleep(2000);
                step++;
            }
        }

        private void RunLoopMainStep7()
        {
            UtilsLog.Log("RunLoopMainStep7 ENTER");

            HtmlElement elemFlight = elemFlights[0];

            List<HtmlElement> elemFlightItems = new List<HtmlElement>();
            UtilsHtml.FindElementHtmlElements(elemFlight, ":div/:div/class:shadow-box/:div/:div/class:brand-selection-button-container/:div/:button", elemFlightItems);
            UtilsLog.Log("RunLoopMainStep7 elemFlightItems.Count={0}", elemFlightItems.Count);
            if (elemFlightItems.Count > 0)
            {
                elemFlightItems[0].Focus();
                elemFlightItems[0].InvokeMember("click");
                RunLoopSleep(2000);
                step++;
            }
        }

        private void RunLoopMainStep8()
        {
            UtilsLog.Log("RunLoopMainStep8 ENTER");

            HtmlElement eledxpsharedflightselection = webBrowser1.Document.GetElementById("dxp-shared-flight-selection");
            if (eledxpsharedflightselection != null)
            {
                List<HtmlElement> eledxpsharedflightselectionItems = new List<HtmlElement>();
                UtilsHtml.FindElementHtmlElements(eledxpsharedflightselection, ":div/:div/:button", eledxpsharedflightselectionItems);
                UtilsLog.Log("RunLoopMainStep8 eledxpsharedflightselection.Count={0}", eledxpsharedflightselectionItems.Count);
                foreach (HtmlElement eledxpsharedflightselectionItem in eledxpsharedflightselectionItems)
                {
                    if (eledxpsharedflightselectionItem.InnerText == "乘客")
                    {
                        eledxpsharedflightselectionItem.Focus();
                        eledxpsharedflightselectionItem.InvokeMember("click");
                        RunLoopSleep(2000);
                        step++;
                    }
                }
            }
        }

        private void ForegroundWindow()
        {
            IntPtr mainWnd = this.Handle;
            SystemParametersInfo(0x2001, 0, IntPtr.Zero, 3);
            ShowWindow(mainWnd, 1);
            SetForegroundWindow(mainWnd);
            SetActiveWindow(mainWnd);
        }

        private void RunLoopMainStep9()
        {
            UtilsLog.Log("RunLoopMainStep9 ENTER");

            HtmlElement elemreactselect5value = webBrowser1.Document.GetElementById("react-select-5--value");
            if (elemreactselect5value == null)
            {
                return;
            }

            elemreactselect5value.Focus();
            RunLoopSleep(2000);

            string funcName;
            string funcBody;

            funcName = "getWindowDevicePixelRatio";
            funcBody = "{return window.devicePixelRatio;}";
            CreateJsFunc(funcName, funcBody);

            funcName = "getDocumentScrollLeft";
            funcBody = "{return document.documentElement.scrollLeft;}";
            CreateJsFunc(funcName, funcBody);

            funcName = "getDocumentScrollTop";
            funcBody = "{return document.documentElement.scrollTop;}";
            CreateJsFunc(funcName, funcBody);

            ForegroundWindow();

            step++;
        }

        private void RunLoopMainStep10()
        {
            UtilsLog.Log("RunLoopMainStep10 ENTER");

            HtmlElement elemreactselect5value = webBrowser1.Document.GetElementById("react-select-5--value");
            UtilsLog.Log("RunLoopMainStep10 elemreactselect5value={0}", elemreactselect5value);
            if (elemreactselect5value == null)
            {
                return;
            }

            int scrollLeft = GetDocumentScrollLeft();
            int scrollTop = GetDocumentScrollTop();
            UtilsLog.Log("RunLoopMainStep10 scrollLeft={0} scrollTop={1}", scrollLeft, scrollTop);
            double devicePixelRatio = GetWindowDevicePixelRatio();
            UtilsLog.Log("RunLoopMainStep10 devicePixelRatio={0}", devicePixelRatio);

            Point elementPosition = UtilsHtml.GetHtmlElementScreenPosition(elemreactselect5value, scrollLeft, scrollTop);
            elementPositionX = (int)(elementPosition.X * devicePixelRatio);
            elementPositionY = (int)(elementPosition.Y * devicePixelRatio);
            elementPositionX += 20;
            elementPositionY += 10;
            UtilsLog.Log("RunLoopMainStep10 elementPositionX={0} elementPositionY={1}", elementPositionX, elementPositionY);

            MouseEvent.EventMouseMove(elementPositionX, elementPositionY);
            Thread.Sleep(500);
            MouseEvent.EventMouseClick(elementPositionX, elementPositionY);

            RunLoopSleep(2000);
            step++;
        }

        private void RunLoopMainStep11()
        {
            HtmlElement elemreactselect5value = webBrowser1.Document.GetElementById("react-select-5--value");
            UtilsLog.Log("RunLoopMainStep11 elemreactselect5value={0}", elemreactselect5value);
            if (elemreactselect5value == null)
            {
                return;
            }

            int scrollLeft = GetDocumentScrollLeft();
            int scrollTop = GetDocumentScrollTop();
            UtilsLog.Log("RunLoopMainStep11 scrollLeft={0} scrollTop={1}", scrollLeft, scrollTop);
            double devicePixelRatio = GetWindowDevicePixelRatio();
            UtilsLog.Log("RunLoopMainStep11 devicePixelRatio={0}", devicePixelRatio);

            Point elementPosition = UtilsHtml.GetHtmlElementScreenPosition(elemreactselect5value, scrollLeft, scrollTop);
            elementPositionX = (int)(elementPosition.X * devicePixelRatio);
            elementPositionY = (int)(elementPosition.Y * devicePixelRatio);
            elementPositionX += 20;
            elementPositionY += 10;
            UtilsLog.Log("RunLoopMainStep11 elementPositionX={0} elementPositionY={1}", elementPositionX, elementPositionY);

            UtilsLog.Log("RunLoopMainStep11 comboBox1.SelectedIndex={0}", comboBox1.SelectedIndex);
            elementPositionY += (comboBox1.SelectedIndex + 1) * 68;

            MouseEvent.EventMouseMove(elementPositionX, elementPositionY);
            Thread.Sleep(500);
            MouseEvent.EventMouseClick(elementPositionX, elementPositionY);
            step++;
        }

        private void RunLoopMainStep12()
        {
            UtilsLog.Log("RunLoopMainStep12 ENTER");

            HtmlElement elemfirstNamePassengerItemAdt1BasicInfoEditFirstName = webBrowser1.Document.GetElementById("firstNamePassengerItemAdt1BasicInfoEditFirstName-passenger-item-ADT-1-basic-info-edit");
            UtilsLog.Log("RunLoopMainStep12 elemfirstNamePassengerItemAdt1BasicInfoEditFirstName={0}", elemfirstNamePassengerItemAdt1BasicInfoEditFirstName);
            HtmlElement elemlastNamePassengerItemAdt1BasicInfoEditLastName = webBrowser1.Document.GetElementById("lastNamePassengerItemAdt1BasicInfoEditLastName-passenger-item-ADT-1-basic-info-edit");
            UtilsLog.Log("RunLoopMainStep12 elemlastNamePassengerItemAdt1BasicInfoEditLastName={0}", elemlastNamePassengerItemAdt1BasicInfoEditLastName);
            HtmlElement elemphoneDefault = webBrowser1.Document.GetElementById("phoneDefaultInput-undefined-additional-contact-info-phone");
            UtilsLog.Log("RunLoopMainStep12 elemphoneDefault={0}", elemphoneDefault);
            HtmlElement elememailAdditionalContactInfoEmailEmail = webBrowser1.Document.GetElementById("emailAdditionalContactInfoEmailEmail-additional-contact-info-email");
            UtilsLog.Log("RunLoopMainStep12 elememailAdditionalContactInfoEmailEmail={0}", elememailAdditionalContactInfoEmailEmail);
            if (elemfirstNamePassengerItemAdt1BasicInfoEditFirstName == null
                || elemlastNamePassengerItemAdt1BasicInfoEditLastName == null
                || elemphoneDefault == null
                || elememailAdditionalContactInfoEmailEmail == null)
            {
                return;
            }

            HtmlElement elemreactselect5value = webBrowser1.Document.GetElementById("react-select-5--value");
            UtilsLog.Log("RunLoopMainStep12 elemreactselect5value={0}", elemreactselect5value);
            if (elemreactselect5value == null)
            {
                return;
            }

            HtmlElement elemreactselect5valueitem = webBrowser1.Document.GetElementById("react-select-5--value-item");
            UtilsLog.Log("RunLoopMainStep12 elemreactselect5valueitem={0}", elemreactselect5valueitem);
            if (elemreactselect5valueitem == null)
            {
                //elemreactselect5value.Children[0].InvokeMember("click");
                return;
            }

            BookSetting bookSetting = GetBookSetting();

            if (elemfirstNamePassengerItemAdt1BasicInfoEditFirstName.GetAttribute("value") == "")
            {
                elemfirstNamePassengerItemAdt1BasicInfoEditFirstName.Focus();
                elemfirstNamePassengerItemAdt1BasicInfoEditFirstName.SetAttribute("value", bookSetting.lastName);
            }
            else if (elemlastNamePassengerItemAdt1BasicInfoEditLastName.GetAttribute("value") == "")
            {
                elemlastNamePassengerItemAdt1BasicInfoEditLastName.Focus();
                elemlastNamePassengerItemAdt1BasicInfoEditLastName.SetAttribute("value", bookSetting.firstName);
            }
            else if (elemphoneDefault.GetAttribute("value").Length < 6)
            {
                elemphoneDefault.Focus();
                elemphoneDefault.SetAttribute("value", bookSetting.telephone);
            }
            else if (elememailAdditionalContactInfoEmailEmail.GetAttribute("value") == "")
            {
                elememailAdditionalContactInfoEmailEmail.Focus();
                elememailAdditionalContactInfoEmailEmail.SetAttribute("value", bookSetting.email);
            }
            else
            {
                HtmlElement elemdxppassengerviewskip = webBrowser1.Document.GetElementById("dxp-passenger-view-skip");
                if (elemdxppassengerviewskip != null)
                {
                    elemdxppassengerviewskip.Focus();
                    elemdxppassengerviewskip.InvokeMember("click");

                    RunLoopSleep(3000);
                    step++;
                    return;
                }
            }
        }

        private void RunLoopMainStep13()
        {
            UtilsLog.Log("RunLoopMainStep13 ENTER");

            HtmlElement elempaymentcreditcardform = webBrowser1.Document.GetElementById("payment-credit-card-form");
            if (elempaymentcreditcardform != null)
            {
                beep_count = 5;

                RunLoopSleep(2000);
                step++;
                return;
            }

            HtmlElement elemdxppagenavigationcontinuebutton = webBrowser1.Document.GetElementById("dxp-page-navigation-continue-button");
            UtilsLog.Log("RunLoopMainStep13 elemdxppagenavigationcontinuebutton={0}", elemdxppagenavigationcontinuebutton);
            if (elemdxppagenavigationcontinuebutton != null)
            {
                elemdxppagenavigationcontinuebutton.Focus();
                elemdxppagenavigationcontinuebutton.InvokeMember("click");
            }
        }

        private void RunLoopMainStep14()
        {
            UtilsLog.Log("RunLoopMainStep14 ENTER");

            if (beep_count == 0)
            {
                step++;
                return;
            }

            Beep(500, 700);
            beep_count--;
        }

        private void RunLoopMainStep15()
        {
            UtilsLog.Log("RunLoopMainStep15 ENTER");
        }


        private HtmlElement FindFlightContent()
        {
            HtmlElement elemapp = webBrowser1.Document.GetElementById("app");
            if (elemapp != null)
            {
                List<HtmlElement> elemappItems = new List<HtmlElement>();
                UtilsHtml.FindElementHtmlElements(elemapp, ":div/:div", elemappItems);
                foreach (HtmlElement elemappItem in elemappItems)
                {
                    string elemappItemClassName = elemappItem.GetAttribute("className");
                    if (elemappItemClassName != null && elemappItemClassName.Contains("flight-content"))
                    {
                        return elemappItem;
                    }
                }
            }

            return null;
        }

        public void CreateJsFunc(string funcName, string funcBody)
        {
            UtilsLog.Log("CreateJsFunc ENTER {0}", funcName);

            HtmlDocument document = webBrowser1.Document;
            string funcText = "function " + funcName + "(){" + funcBody + "}";
            HtmlElement elementScript = document.CreateElement("script");
            elementScript.SetAttribute("type", "text/javascript");
            elementScript.SetAttribute("text", funcText);
            HtmlElement elementHead = document.GetElementsByTagName("head")[0];
            elementHead.AppendChild(elementScript);
        }

        public void CreateJsFuncWithParam1(string funcName, string param1, string funcBody)
        {
            UtilsLog.Log("CreateJsFunc ENTER {0}", funcName);

            HtmlDocument document = webBrowser1.Document;
            string funcText = "function " + funcName + "(" + param1 + "){" + funcBody + "}";
            HtmlElement elementScript = document.CreateElement("script");
            elementScript.SetAttribute("type", "text/javascript");
            elementScript.SetAttribute("text", funcText);
            HtmlElement elementHead = document.GetElementsByTagName("head")[0];
            elementHead.AppendChild(elementScript);
        }

        public void CreateJsFuncWithParam2(string funcName, string param1, string param2, string funcBody)
        {
            UtilsLog.Log("CreateJsFunc ENTER {0}", funcName);

            HtmlDocument document = webBrowser1.Document;
            string funcText = "function " + funcName + "(" + param1 + ", " + param2 + "){" + funcBody + "}";
            HtmlElement elementScript = document.CreateElement("script");
            elementScript.SetAttribute("type", "text/javascript");
            elementScript.SetAttribute("text", funcText);
            HtmlElement elementHead = document.GetElementsByTagName("head")[0];
            elementHead.AppendChild(elementScript);
        }

        public object ExecJSFunc(string funcName)
        {
            HtmlDocument document = webBrowser1.Document;
            return document.InvokeScript(funcName);
        }

        public object ExecJSFuncWithParam1(string funcName, object param1)
        {
            object[] args = { param1};

            HtmlDocument document = webBrowser1.Document;
            return document.InvokeScript(funcName, args);
        }

        public object ExecJSFuncWithParam2(string funcName, object param1, object param2)
        {
            object[] args = { param1 , param2 };

            HtmlDocument document = webBrowser1.Document;
            return document.InvokeScript(funcName, args);
        }
    }
}
