using FluentAssertions.Common;
using FluentAssertions.Common;
using DocumentFormat.OpenXml.Office2013.PowerPoint.Roaming;
using Limilabs.Client.IMAP;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using Newrest.Winrest.FunctionalTests.PageObjects.Purchasing;
using Newrest.Winrest.FunctionalTests.PageObjects.Shared;
using Newrest.Winrest.FunctionalTests.Utils;
using OfficeOpenXml.FormulaParsing.Excel.Functions.Text;
using OpenQA.Selenium;
using OpenQA.Selenium.Interactions;
using OpenQA.Selenium.Support.PageObjects;
using OpenQA.Selenium.Support.UI;
using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using System.Threading;
using System.Runtime.CompilerServices;
using static System.Windows.Forms.VisualStyles.VisualStyleElement;

namespace Newrest.Winrest.FunctionalTests.PageObjects.Flights.Flight
{
    public class HedomadairePage : PageBase
    {
        private const string SORTBY = "cbSortBy";
        private const string FLIGHT_NUMBER = "//*[@id=\"list-item-with-action\"]/div/table/tbody/tr[*]/td[3]";
        private const string RESET_FILTER = "ResetFilter";
        private const string DATE_LABEL = "";
        private const string HIDE_CANCELLED_FLIGHT = "HideCancelledFlights";
        private const string FIRST_FLIGHT = "//*[@id=\"flightTable\"]/tbody/tr[{0}]";
        private const string UNFOLD_ALL = "//a[text()='Unfold All']";
        private const string EXTENDED_BTN_FOLD_UNFOLD = "//button[text()='...']";
        private const string EXTENDED_BTN = "//*[@id=\"tabContentItemContainer\"]/div[1]/div/div/div[2]/div[1]/button";
              private const string GUESTTYPE_QTY = "Items_Items_0__GuestTypes_0__Values_0__NbGuestTypes";//*[@id="Items_Items_0__GuestTypes_0__Values_0__NbGuestTypes"]
        private const string GUESTTYPE = "/html/body/div[2]/div/div[2]/div/div[3]/div/div/table/tbody/tr[2]/td/div/table/tbody/tr/td[2]/input[6]";

        private const string START_DATE_FILTER = "StartDate";
        private const string END_DATE_FILTER = "EndDate";
        private const string END_DATE_TABLE = "//*[@id=\"list-item-with-action\"]/div/table/thead/tr/th[20]";

        private const string START_DATE_TABLE = "//*[@id=\"list-item-with-action\"]/div/table/thead/tr/th[6]";
        private const string SEARCH_FILTER = "SearchPattern";
        private const string EXPORT = "btn-export-excel";
        private const string FIRST_FLIGHT_NUMBER = "//*[@id=\"list-item-with-action\"]/div/table/tbody/tr[1]/td[3]";
        private const string EXPORT_OP245 = "//*[@id=\"tabContentItemContainer\"]/div[1]/div/div/div/a[2]";

        private const string AIRLINES_FILTER = "SelectedCustomers_ms";
        private const string UNSELECT_ALL_AIRLINES = "/html/body/div[10]/div/ul/li[2]/a/span[2]";
        private const string SELECT_ALL_AIRLINES = "/html/body/div[10]/div/ul/li[1]/a/span[2]";
        private const string AIRLINES_SEARCH = "/html/body/div[10]/div/div/label/input";
        private const string AIRLINES_XPATH = "//*[@id=\"list-item-with-action\"]/div/table/tbody/tr[*]/td[5]";
        private const string FLIGHTS_FILTER_ETD_FROM = "ETDFrom";
        private const string FLIGHTS_FILTER_ETD_TO = "ETDTo";
        private const string DOWNLOADS_FIRST_ROW = "//*[starts-with(@id, 'printData')]";
        private const string DOWNLOAD_BUTTON = "header-print-button";
        private const string DOWNLOAD_PANEL = "//*[starts-with(@id,'popover')]";

        private const string SITES_FILTER = "cbSortSiteBy";

        private const string LEG = "//*[@id='list-item-with-action']/div/table/tbody/tr[*]/td[4]";


        private const string SEARCH_NAME = "SearchPattern";
        //Export Hobdomadaire
        private const string CLICKSHOWMENU = "//*[@id=\"tabContentItemContainer\"]/div[1]/div/div/button"; 
        private const string PRINTBYCUSTOMERCATEGORIES = "print-flight-period";

        public HedomadairePage(IWebDriver webDriver, TestContext testContext) : base(webDriver, testContext)
        {
        }

        [FindsBy(How = How.Id, Using = PRINTBYCUSTOMERCATEGORIES)]
        private IWebElement _printbycustomercategories;
        
        [FindsBy(How = How.XPath, Using = EXPORT_OP245)]
        private IWebElement _exportop245;

        [FindsBy(How = How.XPath, Using = CLICKSHOWMENU)]
        private IWebElement _clickshowmenu;

        [FindsBy(How = How.Id, Using = SORTBY)]
        private IWebElement _sortBy;

        [FindsBy(How = How.XPath, Using = START_DATE_TABLE)]
        private IWebElement _startDateTable;

        [FindsBy(How = How.Id, Using = RESET_FILTER)]
        private IWebElement _resetFilter;

        [FindsBy(How = How.XPath, Using = DATE_LABEL)]
        private IWebElement _dateLabel;

        [FindsBy(How = How.XPath, Using = FLIGHT_NUMBER)]
        private IWebElement _flightNumber;

        [FindsBy(How = How.Id, Using = HIDE_CANCELLED_FLIGHT)]
        private IWebElement _hideCancelledFlight;

        [FindsBy(How = How.Id, Using = START_DATE_FILTER)]
        private IWebElement _startDate;
        [FindsBy(How = How.Id, Using = SEARCH_FILTER)]
        private IWebElement _searchFilter;
        [FindsBy(How = How.Id, Using = SITES_FILTER)]
        private IWebElement _site;

        [FindsBy(How = How.XPath, Using = EXPORT)]
        private IWebElement _export;

        [FindsBy(How = How.XPath, Using = FIRST_FLIGHT_NUMBER)]
        private IWebElement _firstFlightNumber;

        [FindsBy(How = How.XPath, Using = UNFOLD_ALL)]
        private IWebElement _unFoldAll;
        [FindsBy(How = How.XPath, Using = EXTENDED_BTN_FOLD_UNFOLD)]
        private IWebElement _extendedButtonFold;
        [FindsBy(How = How.XPath, Using = EXTENDED_BTN)]
        private IWebElement _extendedButton;
        [FindsBy(How = How.Id, Using = GUESTTYPE_QTY)]
        private IWebElement _guestTypeQty;
        [FindsBy(How = How.XPath, Using = GUESTTYPE)]
        private IWebElement _guestType;
        [FindsBy(How = How.XPath, Using = END_DATE_FILTER)]
        private IWebElement _enddate;

        [FindsBy(How = How.Id, Using = AIRLINES_FILTER)]
        private IWebElement _airlinesFilter;

        [FindsBy(How = How.XPath, Using = UNSELECT_ALL_AIRLINES)]
        private IWebElement _unselectAllAirlines;

        [FindsBy(How = How.XPath, Using = SELECT_ALL_AIRLINES)]
        private IWebElement _selectAllAirlines;

        [FindsBy(How = How.XPath, Using = AIRLINES_SEARCH)]
        private IWebElement _searchAirlines;
        [FindsBy(How = How.Id, Using = FLIGHTS_FILTER_ETD_FROM)]
        private IWebElement _ETDFrom;

        [FindsBy(How = How.Id, Using = FLIGHTS_FILTER_ETD_TO)]
        private IWebElement _ETDTo;

        [FindsBy(How = How.Id, Using = DOWNLOAD_BUTTON)]
        private IWebElement _downloadButton;
        [FindsBy(How = How.XPath, Using = DOWNLOAD_PANEL)]
        private IWebElement _exportPanel;

        public enum FilterType
        {

            SortBy,
            HideCancelledFlight,
            Search,
            StartDate,
            EndDate,
            Airlines,
            AirlinesUncheck,
            AirlinesCheckAll,
            ETDFrom,
            ETDTo,
            Site



        }
        public void Filter(FilterType filterType, object value)
        {
            switch (filterType)
            {
                case FilterType.Search:
                    _searchFilter = WaitForElementIsVisible(By.Id(SEARCH_FILTER));
                    _searchFilter.SetValue(ControlType.TextBox, value);
                    break;
                case FilterType.Site:
                    _site = WaitForElementIsVisible(By.Id(SITES_FILTER));
                    _site.SetValue(ControlType.DropDownList, value);
                    break;

                case FilterType.SortBy:
                    _sortBy = WaitForElementIsVisible(By.Id(SORTBY));
                    _sortBy.Click();
                    var element = WaitForElementIsVisible(By.XPath("//option[contains(text(),'" + value + "')]"));
                    _sortBy.SetValue(ControlType.DropDownList, element.Text);
                    _sortBy.Click();
                    // pourquoi ?
                    WaitPageLoading();
                    break;
                case FilterType.HideCancelledFlight:
                    _hideCancelledFlight = WaitForElementExists(By.Id(HIDE_CANCELLED_FLIGHT));
                    _hideCancelledFlight.SetValue(ControlType.CheckBox, value);
                    // pourquoi ?
                    WaitPageLoading();
                    WaitForLoad();
                    break;

                case FilterType.StartDate:
                    _startDate = WaitForElementIsVisible(By.Id(START_DATE_FILTER));
                    _startDate.SetValue(ControlType.DateTime, value);
                    _startDate.SendKeys(Keys.Tab);
                    WaitPageLoading();
                    break;
                case FilterType.EndDate:
                    _startDate = WaitForElementIsVisible(By.Id(END_DATE_FILTER));
                    _startDate.SetValue(ControlType.DateTime, value);
                    _startDate.SendKeys(Keys.Tab);
                    WaitPageLoading();
                    break;

                case FilterType.AirlinesUncheck:
                    _airlinesFilter = WaitForElementIsVisible(By.Id(AIRLINES_FILTER));
                    _airlinesFilter.Click();

                    _unselectAllAirlines = WaitForElementIsVisible(By.XPath(UNSELECT_ALL_AIRLINES));
                    _unselectAllAirlines.Click();

                    _airlinesFilter.Click();
                    break;

                case FilterType.AirlinesCheckAll:
                    _airlinesFilter = WaitForElementIsVisible(By.Id(AIRLINES_FILTER));
                    _airlinesFilter.Click();

                    var selectAllAirlines = WaitForElementIsVisible(By.XPath(SELECT_ALL_AIRLINES));
                    selectAllAirlines.Click();

                    _airlinesFilter.Click();
                    break;

                case FilterType.Airlines:
                    ComboBoxSelectById(new ComboBoxOptions(AIRLINES_FILTER, (string)value));
                    break;
                case FilterType.ETDFrom:
                    _ETDFrom = WaitForElementIsVisible(By.Id(FLIGHTS_FILTER_ETD_FROM), nameof(FLIGHTS_FILTER_ETD_FROM));
                    _ETDFrom.Clear();
                    _ETDFrom.Click();
                    _ETDFrom.SendKeys((string)value);
                    // pourquoi ?
                    WaitPageLoading();
                    WaitForLoad();
                    break;
                case FilterType.ETDTo:
                    _ETDTo = WaitForElementIsVisible(By.Id(FLIGHTS_FILTER_ETD_TO), nameof(FLIGHTS_FILTER_ETD_TO));
                    _ETDTo.Clear();
                    _ETDTo.Click();
                    _ETDTo.SendKeys((string)value);
                    // pourquoi ?
                    WaitPageLoading();
                    WaitForLoad();
                    break;

                default:
                    throw new ArgumentOutOfRangeException(nameof(FilterType), filterType, null);
            }
            WaitPageLoading();
            WaitForLoad();

        }
        public void ResetFilter()
        {
            _resetFilter = WaitForElementIsVisible(By.Id(RESET_FILTER));
            _resetFilter.Click();
            WaitPageLoading();
            WaitForLoad();

        }
        public string SelectFirstFlightNumber()
        {
            var firstFlightNumber = WaitForElementIsVisible(By.XPath(FIRST_FLIGHT_NUMBER));

            return firstFlightNumber.Text.Trim();
        }
        public bool IsSortedByFlightNo()
        {
            int tot = CheckTotalNumber() > 12 ? 12 : CheckTotalNumber();

            if (tot == 0)
                return false;

            var listeFlightNumber = _webDriver.FindElements(By.XPath(FLIGHT_NUMBER));

            var ancienFlightNo = listeFlightNumber[0].GetAttribute("innerText");
            int compteur = 0;

            foreach (var elm in listeFlightNumber)
            {
                // Problème exploitation résultats à cause du type d'affichage des données de la catégorie (on ne prend que les 12 premiers)
                if (compteur >= tot)
                    break;

                if (String.Compare(ancienFlightNo, elm.GetAttribute("innerText")) > 0)
                    return false;


                ancienFlightNo = elm.GetAttribute("innerText");
                compteur++;
            }

            return true;
        }
        public bool IsSortedByDate()
        {
            // Take the date format from the datepicker element and use it to format the date column to avoid date errors
            var dateFormat = _startDate.GetAttribute("data-date-format");
            CultureInfo ci = dateFormat.Equals("dd/mm/yyyy") ? new CultureInfo("fr-FR") : new CultureInfo("en-US");

            bool valueBool = true;
            DateTime ancientDate = DateTime.MinValue;
            int tot;

            if (CheckTotalNumber() > 100)
            {
                tot = 100;
            }
            else
            {
                tot = CheckTotalNumber();
            }
            if (tot == 0)
                return false;

            WaitForLoad();

            var elements = _webDriver.FindElements(By.XPath(START_DATE_FILTER));


            foreach (var element in elements)
            {
                string dateText = element.Text;
                if (DateTime.TryParse(dateText, ci, DateTimeStyles.None, out DateTime date))
                {
                    if (DateTime.Compare(ancientDate, date) > 0)
                    {
                        valueBool = false;
                        break;
                    }

                    ancientDate = date;
                }
                else
                {
                    valueBool = false;
                    break;
                }
            }
            return valueBool;
        }
        public string ConvertDateFormat(DateTime dateInput)
        {
 
            // Define the output date format
            string outputFormat = "MM-dd (ddd)";

            // Convert the DateTime object to the desired output format
            string formattedDate = dateInput.ToString(outputFormat, CultureInfo.InvariantCulture);

            return formattedDate;
        }
        public string GetStartDateColon()
        {
            _startDateTable = WaitForElementIsVisible(By.XPath(START_DATE_TABLE));
           
            return _startDateTable.Text.Trim();
        }
       
        public bool IsFlightDisplayed()
        {
            WaitPageLoading();
            if (isElementVisible(By.XPath(FIRST_FLIGHT)))
            {
                return true;
            }
            else
            {
                return false;
            }
        }

        public override void ShowExtendedMenu()
        {
            _clickshowmenu = WaitForElementIsVisible(By.XPath(CLICKSHOWMENU));
            _clickshowmenu.Click();
            WaitForLoad();
        }
        public void ClickPrintServicesQuantitiesByCustomerAndCategories()
        {
            _printbycustomercategories = WaitForElementIsVisible(By.Id(PRINTBYCUSTOMERCATEGORIES));
            _printbycustomercategories.Click();
            WaitForLoad();
        }
        public string GetTitleOngletPrint()
        {
            WebDriverWait wait = new WebDriverWait(_webDriver, TimeSpan.FromSeconds(60));
            wait.Until(driver => driver.WindowHandles.Count > 1);
            string newTabHandle = _webDriver.WindowHandles[1];
            _webDriver.SwitchTo().Window(newTabHandle);
            string newTabUrl = _webDriver.Url;
            _webDriver.Navigate().GoToUrl(newTabUrl);
            string pdfFileName = new Uri(newTabUrl).Segments.Last();
            return pdfFileName;
        }
        public enum ExportType
        {
            Export,
            OP245
        }
        public void ExportHedomadaire(ExportType exportType, bool versionPrint)
        {
            ShowExtendedMenu();

            switch (exportType)
            {
                case ExportType.Export:
                    _export = WaitForElementIsVisible(By.Id(EXPORT));
                    _export.Click();
                    WaitForLoad();
                    break;

                case ExportType.OP245:
                    _exportop245 = WaitForElementIsVisible(By.XPath(EXPORT_OP245));
                    _exportop245.Click();
                    ComboBoxSelectById(new ComboBoxOptions("SelectedGuestTypes_ms", "ALL", false));
                    var export = WaitForElementIsVisible(By.Id("printButton"));
                    export.Click();                  
                    WaitForLoad();         
                    break;

                default:
                break;
            }

            if (versionPrint)
            {
                IsFileLoaded(By.CssSelector("[target='_blank'][class='far fa-file-excel']"));
                ClosePrintButton();
            }

            WaitForDownload();
            //Close();
        }
        public string GetFirstFlightNumber()
        {
            _firstFlightNumber = WaitForElementIsVisible(By.XPath(FIRST_FLIGHT_NUMBER));
            return _firstFlightNumber.Text;
        }
        public FileInfo GetExportExcelFile(FileInfo[] taskFiles)
        {
            List<FileInfo> correctDownloadFiles = new List<FileInfo>();

            StringBuilder sb = new StringBuilder();

            foreach (var file in taskFiles)
            {
                sb.Append(file.Name + " ");
                //  Test REGEX
                if (IsExcelFileCorrect(file.Name))
                {
                    correctDownloadFiles.Add(file);
                }
            }

            if (correctDownloadFiles.Count == 0)
            {
                return null;
            }

            var time = correctDownloadFiles[0].LastWriteTimeUtc;
            var correctFile = correctDownloadFiles[0];

            correctDownloadFiles.ForEach(file =>
            {
                if (time < file.LastWriteTimeUtc)
                {
                    time = file.LastWriteTimeUtc;
                    correctFile = file;
                }
            });

            return correctFile;
        }
        public bool IsExcelFileCorrect(string filePath)
        {
            Regex r = new Regex("[LPCarts\\s\\d._]", RegexOptions.IgnoreCase | RegexOptions.Singleline);
            Match m = r.Match(filePath);

            return m.Success;
        }
        public void ShowExtendedMenuFold()
        {
            var actions = new Actions(_webDriver);
            _extendedButton = WaitForElementIsVisible(By.XPath(EXTENDED_BTN_FOLD_UNFOLD), nameof(EXTENDED_BTN_FOLD_UNFOLD));
            actions.MoveToElement(_extendedButtonFold).MoveByOffset(-3, 0).Perform();
            WaitForLoad();
        }
        public void UnfoldAll()
        {
            ShowExtendedMenuFold();
            _unFoldAll = WaitForElementIsVisible(By.XPath(UNFOLD_ALL));
            _unFoldAll.Click();
            // animation
            WaitPageLoading();
            WaitForLoad();
        }
        public void SetGuestTypeQty(string guesttype)
        {
            _guestTypeQty = WaitForElementIsVisible(By.Id(GUESTTYPE_QTY));
            _guestTypeQty.Click();
            _guestTypeQty.ClearElement();
            _guestTypeQty.SendKeys(guesttype);
            WaitPageLoading();
            WaitPageLoading();
        }
        public string GetGuestTypeQty()
        {
            _guestTypeQty = WaitForElementIsVisible(By.Id(GUESTTYPE_QTY));
            _guestTypeQty.Click();
            WaitPageLoading();
            return _guestTypeQty.GetAttribute("value");

        }
        public string GetEndDateColumn()
        {
            var startDateColon = WaitForElementIsVisible(By.XPath("//*[@id='list-item-with-action']/div/table/thead/tr/th[6]"));
            return startDateColon.Text.Trim();
        }
        public void UncheckAllAirlines()
        {
            _airlinesFilter = WaitForElementIsVisible(By.Id(AIRLINES_FILTER));
            _airlinesFilter.Click();

            _unselectAllAirlines = WaitForElementIsVisible(By.XPath(UNSELECT_ALL_AIRLINES));
            _unselectAllAirlines.Click();

            _airlinesFilter.Click();

            WaitPageLoading();
            Thread.Sleep(2000);
        }
        public List<string> GetAirlinesFiltred()
        {
            var airlines = _webDriver.FindElements(By.XPath(AIRLINES_XPATH));

            if (airlines.Count == 0)
                return new List<string>();

            var filteredAirlines = new List<string>();

            foreach (var elm in airlines)
            {

                filteredAirlines.Add(elm.Text.Trim());

            }

            return filteredAirlines;
        }
        public bool CheckDownloadsExists()
        {
            if (!isElementExists(By.XPath(DOWNLOAD_PANEL)))
            {
                _downloadButton = WaitForElementIsVisible(By.Id(DOWNLOAD_BUTTON));
                _downloadButton.Click();
                Thread.Sleep(1000);
            }

            WaitForLoad();
            return isElementExists(By.XPath(DOWNLOADS_FIRST_ROW));
        }
        public List<string> GetDisplayedFlights()
        {
            var siteElements = _webDriver.FindElements(By.XPath(LEG));
            var displayedFlights = siteElements
                .Select(element => element.Text.Trim())
                .ToList();
            return displayedFlights;
        }

    }
}
