using Microsoft.VisualStudio.TestTools.UnitTesting;
using Newrest.Winrest.FunctionalTests.PageObjects.Shared;
using Newrest.Winrest.FunctionalTests.Utils;
using OpenQA.Selenium;
using OpenQA.Selenium.Support.PageObjects;
using OpenQA.Selenium.Support.UI;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading;

namespace Newrest.Winrest.FunctionalTests.PageObjects.Flights.Flight
{
    public class FlightMassiveDeleteModalPage : PageBase
    {
        private const string MASSIVE_DELETE_FLIGHT_NAME = "SearchPattern";
        private const string MASSIVE_DELETE_FLIGHT_NAME_XPATH = "/html/body/div[3]/div/div/div[2]/div/form/div/div[1]/div/input";
        private const string MASSIVE_DELETE_FLIGHT_SITE = "SelectedSiteIds_ms";
        private const string MASSIVE_DELETE_FLIGHT_CUSTOMER = "SelectedCustomersForDeletion_ms";
        private const string MASSIVE_DELETE_FLIGHT_STATUS = "SelectedStatus_ms";
        private const string MASSIVE_DELETE_FLIGHT_SEARCH_BUTTON = "SearchFlightsBtn"; 
        private const string MASSIVE_DELETE_FLIGHT_SHOW_INACTIVE_CUSTOMERS = "ShowInactiveCustomers";
        private const string MASSIVE_DELETE_FLIGHT_SHOW_INACTIVE_SITES = "ShowInactiveSites";
        private const string MASSIVE_DELETE_FLIGHT_STATUS_CHECK_INACTIVE_SITES = "/html/body/div[15]/ul/li[2]/label/span";
        private const string MASSIVE_DELETE_FLIGHT_STATUS_CHECK_INACTIVE_CUSTOMERS = "html/body/div[15]/ul/li[4]/label/span";
        ///html/body/div[15]/ul/li[4]/label/span
        private const string DATE_FROM = "dateFlightStart";
        private const string DATE_TO = "dateFlightEnd";
        private const string CUSTOMER_FILTER = "SelectedCustomersForDeletion_ms";  
        private const string SITEHEADER_BTN = "//*[@id=\"tableFlights\"]/thead/tr/th[5]/span/a";
        private const string CUSTOMERHEADER_BTN = "//*[@id=\"tableFlights\"]/thead/tr/th[6]/span/a";
        private const string MASSIVE_DELETE_PAGESIZE = "//*[@role=\"dialog\"]//*[@id=\"page-size-selector\"]";
        private const string MASSIVE_DELETE_FLIGHT_CUSTOMER_LIST = "//*[@id=\"tableFlights\"]/tbody/tr[*]/td[6]";
        private const string MASSIVE_DELETE_FLIGHT_SITE_LIST = "//*[@id=\"tableFlights\"]/tbody/tr[*]/td[5]";
        private const string SITE_FILTER = "SelectedSiteIds_ms";
        private const string FLIGHTHEADER_BTN = "//*[@id=\"tableFlights\"]/thead/tr/th[3]/span/a";
        private const string BARSETHEADER_BTN = "//*[@id=\"tableFlights\"]/thead/tr/th[4]/span/a";
        private const string MASSIVE_DELETE_PAGINATION = "//*[@id=\"list-flights-deletion\"]/nav/ul/li[position()=last()-2]/a";
        private const string DDL_PAGESIZE = "//*[@id=\"div-flightResultTable\"]/descendant::select[contains(@id, 'page-size-selector')]";
        private const string MASSIVE_DELETE_FLIGHT_NAME_LIST = "//*[@id=\"tableFlights\"]/tbody/tr[*]/td[3]/div";
        private const string MASSIVE_DELETE_FLIGHT_DATE_LIST = "//*[@id=\"tableFlights\"]/tbody/tr[*]/td[2]";
        private const string DATEHEADER_BTN = "//*[@id=\"tableFlights\"]/thead/tr/th[2]/span/a";
        private const string SELECT_ALL_BUTTON = "selectAll";
        private const string DELETE_FLIGHT_BTN = "deleteFlightsBtn";
        private const string CONFIRM_BTN = "dataConfirmOK";
        private const string OK_BTN = "/html/body/div[3]/div/div/div[3]/button";


        private const string FIRST_FLIGHT_NUMBER = "/html/body/div[3]/div/div/div[2]/div/form/div/div[8]/div/div/table/tbody/tr[1]/td[3]/div";
        private const string FIRST_FLIGHT_NUMBER_ID = "flights-flight-massive-delete-flight-number-1";
        private const string LAST_FLIGHT_NUMBER = "/html/body/div[3]/div/div/div[2]/div/form/div/div[8]/div/div/table/tbody/tr[15]/td[3]/div";
        private const string CLOSE = "//*[@id=\"modal-1\"]/div[1]/button/span";
        //dateFlightStart


        // Period
        [FindsBy(How = How.Id, Using = DATE_FROM)]
        private IWebElement _fromDate;

        [FindsBy(How = How.Id, Using = DATE_TO)]
        private IWebElement _toDate;

        [FindsBy(How = How.XPath, Using = CLOSE)]
        private IWebElement _close;

        public FlightMassiveDeleteModalPage(IWebDriver webDriver, TestContext testContext) : base(webDriver, testContext)
        {
        }
        public enum FilterType
        {
            flightName,
            From,
            To,
            Site,
            Customer,
            Status
        }

        public void SetDateFromAndTo(DateTime fromDate, DateTime toDate)
        {
            _fromDate = WaitForElementIsVisible(By.Id(DATE_FROM));
            _fromDate.SetValue(ControlType.DateTime, fromDate);
            _fromDate.SendKeys(Keys.Tab);
            WaitForLoad();

            _toDate = WaitForElementIsVisible(By.Id(DATE_TO));
            _toDate.SetValue(ControlType.DateTime, toDate);
            _toDate.SendKeys(Keys.Tab);
            WaitForLoad();
        }

        public void ClickSearchButton()
        {
            var searchBtn = WaitForElementIsVisible(By.Id(MASSIVE_DELETE_FLIGHT_SEARCH_BUTTON));
            searchBtn.Click();
            WaitPageLoading();
            WaitForLoad();
        }

        public void SelectInactiveSites()
        {
            var statusClick =  WaitForElementExists(By.Id(MASSIVE_DELETE_FLIGHT_STATUS));
            statusClick.Click();
            WaitForLoad();
            var inactiveSelect = _webDriver.FindElements(By.XPath(MASSIVE_DELETE_FLIGHT_STATUS_CHECK_INACTIVE_SITES));
            WaitForLoad();
            inactiveSelect.FirstOrDefault().Click();
        }

        public void SelectInactiveCustomers()
        {
            var statusClick = WaitForElementExists(By.Id(MASSIVE_DELETE_FLIGHT_STATUS));
            statusClick.Click();
            WaitForLoad();
            var inactiveSelect = _webDriver.FindElements(By.XPath(MASSIVE_DELETE_FLIGHT_STATUS_CHECK_INACTIVE_CUSTOMERS));
            WaitForLoad();
            inactiveSelect.FirstOrDefault().Click();
        }

        public void SelectSiteToFilter(string site = "ACE")
        {
            ComboBoxSelectById(new ComboBoxOptions(MASSIVE_DELETE_FLIGHT_SITE, site, false));
        }

        public bool CheckIfSitesAreInactive()
        {
            bool defaultval = true;
            var listsitesActif = WaitForElementExists(By.Id(MASSIVE_DELETE_FLIGHT_SITE));
            listsitesActif.Click();
            var xbout = "/html/body/div[16]/ul/li"; 
            WaitForLoad();
            var listsitactive = _webDriver.FindElements(By.XPath(xbout)).Select(e => e.Text).ToList();


            var listSiteResult = GetCurrentFlightSiteName().ToList();
           
            var noMatch = listSiteResult.Where(p => listsitactive.All(p2 => p2!= p)).Count();
            defaultval = noMatch==0;
            var nbpage = "//*[@id=\"list-flights-deletion\"]/nav/ul/li"; // //*[@id="list-flights-deletion"]/nav/ul/li[3]/a
            WaitForLoad();
            var nombrPages = _webDriver.FindElements(By.XPath(nbpage)).Select(e => e.Text).ToList().Count();

            for (int i = 3; i < nombrPages; i++)
            {
                var currentpage = "//*[@id=\"list-flights-deletion\"]/nav/ul/li[" + i + "]/a";
                WaitForLoad();
                var pagesuivante = _webDriver.FindElements(By.XPath(currentpage)).FirstOrDefault();
                pagesuivante.Click();

                var listSitePAgeCourante = GetCurrentFlightSiteName().ToList();
                var isMatchingItem = listSiteResult.Where(p => listsitactive.All(p2 => p2 != p)).Count();
                bool isMatching = isMatchingItem == 0;
                defaultval = defaultval && isMatching;
            }

            return defaultval;
        }

        public bool CheckIfCustomerAreInactive() 
        {
            bool defaultval = true;
            var listCustomersActif = WaitForElementExists(By.Id(MASSIVE_DELETE_FLIGHT_CUSTOMER));
            listCustomersActif.Click();
            var xbout = "/html/body/div[17]/ul/li";
            WaitForLoad();
            var listCustActive = _webDriver.FindElements(By.XPath(xbout)).Select(e => e.Text).ToList();

            var listCustomerResult = GetCurrentFlightCustomerName().ToList();

            var noMatch = listCustomerResult.Where(p => listCustActive.All(p2 => p2 != p)).Count();
            defaultval = noMatch == 0;
            var nbpage = "//*[@id=\"list-flights-deletion\"]/nav/ul/li";
            WaitForLoad();
            var nombrPages = _webDriver.FindElements(By.XPath(nbpage)).Select(e => e.Text).ToList().Count();

            for (int i = 3; i < nombrPages; i++)
            {
                var currentpage = "//*[@id=\"list-flights-deletion\"]/nav/ul/li[" + i + "]/a";
                WaitForLoad();
                var pagesuivante = _webDriver.FindElements(By.XPath(currentpage)).FirstOrDefault();
                pagesuivante.Click();

                var listSitePAgeCourante = GetCurrentFlightCustomerName().ToList();
                var isMatchingItem = listCustomerResult.Where(p => listCustActive.All(p2 => p2 != p)).Count();
                bool isMatching = isMatchingItem == 0;
                defaultval = defaultval && isMatching;
            }

            return defaultval;
        }

        public bool CheckIfAllLineContainTheSite(string siteName)
        {
            bool defaultval = true;
            var listSiteResult = GetCurrentFlightSiteName().ToList();

            
            var nbpage = "//*[@id=\"list-flights-deletion\"]/nav/ul/li"; 
            WaitForLoad();
            var nombrPages = _webDriver.FindElements(By.XPath(nbpage)).Select(e => e.Text).ToList().Count();
            defaultval = defaultval && listSiteResult.All(p => p == siteName);

            for (int i = 3; i < nombrPages; i++)
            {
                var currentpage = "//*[@id=\"list-flights-deletion\"]/nav/ul/li[" + i + "]/a";
                WaitForLoad();
                var pagesuivante = _webDriver.FindElements(By.XPath(currentpage)).FirstOrDefault();
                pagesuivante.Click();

                var listSitePAgeCourante = GetCurrentFlightSiteName().ToList();             
                defaultval = defaultval && listSitePAgeCourante.All(p => p == siteName);
            }

            return defaultval;
        }

        public IEnumerable<string> GetFirstLine()
        {

           var firstLine = "//*[@id=\"tableFlights\"]/tbody/tr[1]";
            var listCustActive = _webDriver.FindElements(By.XPath(firstLine)).Select(e => e.Text);

            return listCustActive;
        }

        public IEnumerable<string> GetDisplayedNumberFlight()
        {
            var numer = "//*[@id=\"flightCount\"]";
            var nbFlightSelectedDisplayed = _webDriver.FindElements(By.XPath(numer)).Select(e => e.Text);

            return nbFlightSelectedDisplayed;

        }


        public IEnumerable<String> GetCurrentFlightSiteName()
        {
            var cursite = "//*[@id=\"tableFlights\"]/tbody/tr/td[5]";
            WaitForLoad();
            return _webDriver.FindElements(By.XPath(cursite)).Select(e => e.Text);
        }

        public IEnumerable<string> GetCurrentFlightCustomerName()
        {

            var cursite = "//*[@id=\"tableFlights\"]/tbody/tr/td[6]";
            WaitForLoad();
            return _webDriver.FindElements(By.XPath(cursite)).Select(e => e.Text);
        }

        public IEnumerable<DateTime> GetCurrentDatesResult()
        {
            var result = new List<DateTime>();

            var curdate = "//*[@id=\"tableFlights\"]/tbody/tr/td[2]";
            WaitForLoad();
            string dateFormat = "dd/MM/yyyy";
            var listeDateString= _webDriver.FindElements(By.XPath(curdate)).Select(e => e.Text).ToList();
            foreach(var item in listeDateString)
            {
                DateTime dt = DateTime.ParseExact(item, dateFormat, null);
                result.Add(dt);
            }

            return result;
        }

        public bool CheckIfDateIsGood(DateTime deb, DateTime fin, List<DateTime> dates)
        {
            var result = true;
            foreach (var item in dates)
            {
                if (item < deb || item> fin)
                {
                    return false;
                }                 
            }

            return result;
        }

        public void ShowInactiveCustomers()
        {
            var checkbox = WaitForElementExists(By.XPath("//*/label[@for=\"ShowInactiveCustomers\"]/../div"));
            checkbox.Click();
        }

        public void SelectCustomer(string customer,bool selectAll=false)
        {
            // secousses
            Thread.Sleep(2000);
            ComboBoxSelectById(
               selectAll
                     ? new ComboBoxOptions(CUSTOMER_FILTER, customer, false) { ClickCheckAllAtStart = false, ClickCheckAllAfterSelection = true }
                     : new ComboBoxOptions(CUSTOMER_FILTER, customer, false)
                               );
            WaitForLoad();
        }

        public void ClickOnCustomerFilter()
        {
            var select = WaitForElementExists(By.Id(CUSTOMER_FILTER));
            select.Click();
        }

        public bool IsInactive()
        {
            var list = _webDriver.FindElements(By.XPath("/html/body/div[17]/ul/li[not(@class='ui-multiselect-excluded')]/label")).ToList();

            foreach (var element in list)
            {
                if(!element.Text.StartsWith("Inactive -"))
                {
                    return false;
                }

            }
            return true;
        }

        public bool VerifyCustomerSearch(string customer)
        {
            var listCustomers = _webDriver.FindElements(By.XPath(MASSIVE_DELETE_FLIGHT_CUSTOMER_LIST));

            return listCustomers.Where(s => s.Text.Equals(customer)).Count() == listCustomers.Count();
        }
        public new void PageSize(string size)
        {
            
            var PageSizeDdl = _webDriver.FindElement(By.XPath(MASSIVE_DELETE_PAGESIZE));
            PageSizeDdl.SetValue(ControlType.DropDownList, size);
            WaitForLoad();
        }
        public void ClickOnSiteHeader()
        {
            var siteHeader = WaitForElementExists(By.XPath(SITEHEADER_BTN));
            siteHeader.Click();
            WaitPageLoading();
        }
        private List<string> GetSitesCodes()
        {
            List<string> ids = new List<string>();
            string xPath = "//*[@id=\"tableFlights\"]/tbody/tr[*]/td[5]/div";
            var sitesCodes = _webDriver.FindElements(By.XPath(xPath));

            foreach (var code in sitesCodes)
            {
                ids.Add(code.Text);
            }
            return ids;
        }
        public bool VerifySiteCodeSort(SortType sortType)
        {
            PageSize("100");

            List<string> ids = this.GetSitesCodes();

            for (int i = 0; i < ids.Count - 1; i++)
            {
                int comparisonResult = string.Compare(ids[i], ids[i + 1], StringComparison.OrdinalIgnoreCase);
                if ((sortType == SortType.Ascending && comparisonResult > 0) || (sortType == SortType.Descending && comparisonResult < 0))
                {
                    return false;
                }
            }
            return true;
        }
        public void ClickOnCustomerHeader()
        {
            var customerHeader = WaitForElementExists(By.XPath(CUSTOMERHEADER_BTN));
            customerHeader.Click();
            WaitPageLoading();
        }
        private List<string> GetCustomersCodes()
        {
            List<string> ids = new List<string>();
            string xPath = "//*[@id=\"tableFlights\"]/tbody/tr[*]/td[6]/div";
            var customersCodes = _webDriver.FindElements(By.XPath(xPath));

            foreach (var code in customersCodes)
            {
                ids.Add(code.Text.Trim());
            }
            return ids;
        }
        public bool VerifyCustomerCodeSort(SortType sortType)
        {
            PageSize("100");

            List<string> ids = this.GetCustomersCodes();

            for (int i = 0; i < ids.Count - 1; i++)
            {
                int comparisonResult = string.Compare(ids[i], ids[i + 1], StringComparison.OrdinalIgnoreCase);
                if ((sortType == SortType.Ascending && comparisonResult > 0) || (sortType == SortType.Descending && comparisonResult < 0))
                {
                    return false;
                }
            }
            return true;
        }

        public void ShowInactiveSites()
        {
            var checkbox = WaitForElementExists(By.XPath("//*/label[@for=\"ShowInactiveSites\"]/../div"));
            checkbox.Click();
            WaitForLoad();
        }

        public void SelectSite(string site)
        {
            ComboBoxSelectById(new ComboBoxOptions(SITE_FILTER, site, false));
            WaitForLoad();
        }

        public void ClickOnSitesFilter()
        {
            var select = WaitForElementExists(By.Id(SITE_FILTER));
            select.Click();
        }

        public bool CheckData()
        {
            try
            {
                var text = WaitForElementExists(By.XPath("//*[@id='tableFlights']/tbody/tr/td[2]/b")).Text;

                if (!text.Equals("No data to display. Use the Search button to load data")) return true;
                return false;
            }
            catch
            {
                return true;
            }
        }

        public int CheckTotalPageNumber()
        {
            try
            {
                return int.Parse(WaitForElementExists(By.XPath(MASSIVE_DELETE_PAGINATION)).Text);
            }
            catch
            {
                return 0;
            }
        }

        public void GoToPage(int page)
        {
            string paginationxpath = MASSIVE_DELETE_PAGINATION;
            if (!isElementExists(By.XPath(paginationxpath)))
            {
                paginationxpath = MASSIVE_DELETE_PAGINATION.Replace("/div[8]/div/div/nav/ul/li", "/div[8]/div/div/div/nav/ul/li");
            }
            var pageLink = _webDriver.FindElements(By.XPath(paginationxpath)).Where(e => e.Text.Trim().Equals(page.ToString())).FirstOrDefault();
            if (pageLink != null)
            {
                pageLink.Click();
                WaitForLoad();
            }
        }

        public bool VerifySiteSearch(string site)
        {
            var listSites = _webDriver.FindElements(By.XPath(MASSIVE_DELETE_FLIGHT_SITE_LIST));

            return listSites.Where(s => s.Text.Equals(site)).Count() == listSites.Count();
        }

        public void ClickOnFlightHeader()
        {
            var customerHeader = WaitForElementExists(By.XPath(FLIGHTHEADER_BTN));
            customerHeader.Click();
            WaitPageLoading();
        }
        private List<string> GetFlightsCodes()
        {
            List<string> ids = new List<string>();
            string xPath = "//*[@id=\"tableFlights\"]/tbody/tr[*]/td[3]/div";
            var flightsCodes = _webDriver.FindElements(By.XPath(xPath));

            foreach (var code in flightsCodes)
            {
                    ids.Add(code.Text);
            }
            return ids;
        }
        public bool VerifyFlightCodeSort(SortType sortType)
        {
            PageSize("100");

            List<string> ids = this.GetFlightsCodes();

            // Sort the list based on the sortType
            List<string> sortedIds = sortType == SortType.Ascending
                ? ids.OrderBy(id => id, StringComparer.OrdinalIgnoreCase).ToList()
                : ids.OrderByDescending(id => id, StringComparer.OrdinalIgnoreCase).ToList();

            // Verify if the original list is sorted correctly
            for (int i = 0; i < ids.Count - 1; i++)
            {
                // donnée corompu
                if (ids[i].StartsWith("-") || ids[i + 1].StartsWith("-")) continue;
                int comparisonResult = string.Compare(ids[i], ids[i + 1], StringComparison.OrdinalIgnoreCase);
                if ((sortType == SortType.Ascending && comparisonResult > 0) || (sortType == SortType.Descending && comparisonResult < 0))
                {
                    return false;
                }
            }
            return true;
        }
        public void ClickOnBarsetHeader()
        {
            var barsetHeader = WaitForElementExists(By.XPath(BARSETHEADER_BTN));
            barsetHeader.Click();
            WaitPageLoading();
        }
        private List<string> GetBarsets()
        {
            List<string> ids = new List<string>();
            string xPath = "//*[@id=\"tableFlights\"]/tbody/tr[*]/td[4]"; 
            var barsets = _webDriver.FindElements(By.XPath(xPath));

            foreach (var code in barsets)
            {
                ids.Add(code.Text);
            }
            return ids;
        }
        public void GoToLastPage()
        {
            var lastPage = WaitForElementExists(By.XPath("//*[@id=\"list-flights-deletion\"]/nav/ul/li[8]/a")); 
            lastPage.Click();
            WaitPageLoading();
        }
        public bool VerifyOrganizeByBarset(string val)
        {
            PageSize("100");
            List<string> barsets = this.GetBarsets();
            foreach (var barset in barsets) 
            {
                if (barset.Equals(val) == false)
                {
                    GoToLastPage();
                    List<string> lastBarsets = this.GetBarsets();
                    if (lastBarsets[lastBarsets.Count - 1] == val)
                    {
                        return false;
                    }
                    break;
                }
            }
            return true;
        }
        public void ClickSelectAllButton()
        {
            var btn = WaitForElementIsVisible(By.Id(SELECT_ALL_BUTTON)); 
            btn.Click();
            WaitForLoad();
        }
        public int CheckTotalSelectCount()
        {
            var count = _webDriver.FindElement(By.Id("flightCount"));
            var totalnumber = count.Text.Trim();
            return Convert.ToInt32(totalnumber);
        }
        public bool IsPageSizeEqualsTo(string size)
        {
            var nbPages = WaitForElementExists(By.XPath(DDL_PAGESIZE));
            SelectElement select = new SelectElement(nbPages);
            IWebElement selectedOption = select.SelectedOption;
            string selectedValue = selectedOption.GetAttribute("value");

            return selectedValue == size;
        }
        public int GetTotalRowsForPagination()
        {
            var table = _webDriver.FindElement(By.Id("tableFlights"));

            var allRows = table.FindElements(By.TagName("tr"));

            var totalRows = (allRows.Count() - 1) / 2;
            return totalRows;
        }

        public void SetFlightName(string value)
        {
            WaitForLoad();
            var flightNameInput = WaitForElementIsVisible(By.XPath(MASSIVE_DELETE_FLIGHT_NAME_XPATH));
            flightNameInput.SendKeys(value.ToString());
            WaitForLoad();
        }

        public bool VerifyServiceNameSearch(string serviceName)
        {
            var listFlights = _webDriver.FindElements(By.XPath(MASSIVE_DELETE_FLIGHT_NAME_LIST));

            return listFlights.Where(s => s.Text.Equals(serviceName)).Count() == listFlights.Count();
        }

        public void ClickOnDateHeader()
        {
            var siteHeader = WaitForElementExists(By.XPath(DATEHEADER_BTN));
            siteHeader.Click();
            WaitPageLoading();
        }

        public string GetSortType()
        {
            var datesList = _webDriver.FindElements(By.XPath(MASSIVE_DELETE_FLIGHT_DATE_LIST));

            string dateFormat = "dd/MM/yyyy";
            List<DateTime> dateTimes = datesList.Select(element => DateTime.ParseExact(element.Text,dateFormat,null)).ToList();

            bool isSortedAscending = dateTimes.SequenceEqual(dateTimes.OrderBy(date => date));

            bool isSortedDescending = dateTimes.SequenceEqual(dateTimes.OrderByDescending(date => date));

            return isSortedAscending ? "asc" : "desc";
        }
        public string GetFirstRowDate()
        {
            var firstDate = _webDriver.FindElement(By.XPath("//*[@id=\"tableFlights\"]/tbody/tr[1]/td[2]"));
            return firstDate.Text;
        }
        public string GetFirstRowSite()
        {
            var firstSite = _webDriver.FindElement(By.XPath("//*[@id=\"tableFlights\"]/tbody/tr[1]/td[5]/div"));
            return firstSite.Text;
        }
        public bool ClickOnResultLinkAndCheckFlight(string firstRowDate, string firstRowSite)
        {
            IWebElement rowResultLink = null;
            try
            {
                rowResultLink = WaitForElementIsVisible(By.XPath("//*[@id=\"tableFlights\"]/tbody/tr[1]/td[8]/a"));
            }
            catch
            {
                return false;
            }

            if (rowResultLink == null)
            {
                return false;
            }
            else
            {
                rowResultLink.Click();

                var wait = new WebDriverWait(_webDriver, TimeSpan.FromSeconds(60));
                wait.Until((driver) => driver.WindowHandles.Count > 1);
                _webDriver.SwitchTo().Window(_webDriver.WindowHandles[1]);

                var date = WaitForElementExists(By.Id("start-date-picker-detail"));
                SelectElement sitesSelect = new SelectElement(WaitForElementExists(By.Id("cbSites")));
                string site = sitesSelect.SelectedOption.Text;
                string dateValue = date.GetAttribute("value");

                bool isDateOk = dateValue.IndexOf(firstRowDate, StringComparison.OrdinalIgnoreCase) >= 0;
                bool isSiteOk = site.Equals(firstRowSite);

                return isDateOk && isSiteOk;
            }
        }

        public string GetFirstFlightNumber()
        {
            if (IsDev())
            {
                var firstFlightNumber = WaitForElementIsVisible(By.Id(FIRST_FLIGHT_NUMBER_ID));
                return firstFlightNumber.Text;
            }
            else
            {
                var firstFlightNumber = WaitForElementIsVisible(By.XPath("//*[@id='tableFlights']/tbody/tr[1]/td[3]/div"));
                return firstFlightNumber.Text;
            }
        }
        public void SelectFirstFlight()
        {
            var firstCheck = WaitForElementExists(By.Id("item_IsSelected"));
            firstCheck.SetValue(ControlType.CheckBox,true);
        }
        public void Delete()
        {
            var deleteBtn = WaitForElementIsVisible(By.Id(DELETE_FLIGHT_BTN));
            deleteBtn.Click();
            var confirmBtn = WaitForElementExists(By.Id(CONFIRM_BTN));
            confirmBtn.Click();
            var okBtn = WaitForElementExists(By.XPath(OK_BTN));
            okBtn.Click();
        }

        public int NumberOfCheckboxesChekd()
        {
            IList<IWebElement> checkboxes = _webDriver.FindElements(By.XPath("//input[@type='checkbox' and @name='item.IsSelected']"));
            var nbchk = 0;

            foreach (var checkbox in checkboxes)
            {
                // Vérification si c'est coché et comptage flight coché
                bool isChecked = checkbox.Selected;
                if (isChecked)
                {
                    nbchk++;
                }
            }
            return nbchk;
        }

        public bool IsSelectedEachFlight()
        {          
            var nbpage = "//*[@id=\"list-flights-deletion\"]/nav/ul/li";
            WaitForLoad();
            var nombrPages = _webDriver.FindElements(By.XPath(nbpage)).Select(e => e.Text).ToList().Count();
            var countedFligt = 0;

            for (int i = 3; i <= nombrPages; i++)
            {
                var currentpage = "//*[@id=\"list-flights-deletion\"]/nav/ul/li[" + i + "]/a";
                WaitForLoad();
                var pagesuivante = _webDriver.FindElements(By.XPath(currentpage)).FirstOrDefault();
                var nbvol = NumberOfCheckboxesChekd();
                if (pagesuivante != null)
                {
                    pagesuivante.Click();
                }
                countedFligt = countedFligt + nbvol; 
            }
            var nombreFlight = int.Parse(GetDisplayedNumberFlight().FirstOrDefault());
            return nombreFlight == countedFligt;
        }

        public string GetLastFlightNumber()
        {
            var number = GetTotalRowsForPagination();
            var idOfLastFlightNumber = $"flights-flight-massive-delete-flight-number-{number}";
            var _flightNumber = WaitForElementIsVisible(By.Id(idOfLastFlightNumber));
            return _flightNumber.Text;
        }

        public void SelectLastFlight()
        {
            IList<IWebElement> checkboxes = _webDriver.FindElements(By.XPath("//tbody//input[@type='checkbox' and @name='item.IsSelected']"));

            // Cocher la dernière case 
            if (checkboxes.Count > 1)
            {
                IWebElement lastCheckbox = checkboxes[checkboxes.Count - 1];
                if (!lastCheckbox.Selected)
                {
                    lastCheckbox.Click();
                }
            }
        }
        public void ClickFirstUseCase()
        {
            var useCase = _webDriver.FindElement(By.XPath("//*/table[@id='tableFlights']/tbody/tr[1]/td[7]/div/span"));
            useCase.Click();
        }

        public string GetLegsNumber()
        {

            var legs = WaitForElementIsVisible(By.XPath("//*/table[contains(@class,'massiveDeleteCountDetailTable')]/tbody/tr/td[2]"));
            return legs.Text;
        }
        public string GetLegsNumberFromFlight()
        {
            var legs = _webDriver.FindElement(By.XPath("/html/body/div[2]/div/div[2]/div/div[2]/div[1]/div[1]/table/tbody/tr[2]/td[19]/div"));
            return legs.Text;
        }
        public List<string> CustomersWithoutDeplicates()
        {
            return new HashSet<string> (GetCustomersCodes()).ToList();
        }
        public void CloseMassiveDeleteModal()
        {
            _close = WaitForElementIsVisible(By.XPath(CLOSE));
            _close.Click();
            WaitPageLoading();
        }
    }



    public enum SortType
    {
        Ascending,
        Descending,
    }
    
}
