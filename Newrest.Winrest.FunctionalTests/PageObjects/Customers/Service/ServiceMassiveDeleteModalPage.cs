using DocumentFormat.OpenXml.Math;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using Newrest.Winrest.FunctionalTests.PageObjects.Menus.Datasheet;
using Newrest.Winrest.FunctionalTests.PageObjects.Shared;
using Newrest.Winrest.FunctionalTests.Utils;
using OpenQA.Selenium;
using OpenQA.Selenium.Support.PageObjects;
using OpenQA.Selenium.Support.UI;
using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;

namespace Newrest.Winrest.FunctionalTests.PageObjects.Customers.Service
{
    public class ServiceMassiveDeleteModalPage : PageBase
    {
        private const string MASSIVE_DELETE_SERVICE_NAME = "/html/body/div[3]/div/div/div[2]/div/form/div/div[1]/div/input";
        private const string MASSIVE_DELETE_SERVICE_SITE = "SelectedSiteIds_ms";
        private const string MASSIVE_DELETE_SERVICE_CUSTOMER = "SelectedCustomersForDeletion_ms";
        private const string MASSIVE_DELETE_SERVICE_FROM = "dateMenuStart";
        private const string MASSIVE_DELETE_SERVICE_TO = "dateMenuEnd";
        private const string MASSIVE_DELETE_SERVICE_STATUS = "SelectedStatus_ms";
        private const string MASSIVE_DELETE_SERVICE_SEARCH_BUTTON = "SearchServicesBtn";
        private const string MASSIVE_DELETE_PAGESIZE = "/html/body/div[3]/div/div/div[2]/div/form/div/div[8]/div/div/nav/select";
        private const string MASSIVE_DELETE_SERVICE_DELETE_BTN = "deleteServicesBtn";
        private const string MASSIVE_DELETE_SERVICE_CONFIRM_BTN = "dataConfirmOK";
        private const string MASSIVE_DELETE_SERVICE_DELETEOK_BTN = "/html/body/div[3]/div/div/div[3]/button";
        private const string MASSIVE_DELETE_PAGINATION = "/html/body/div[3]/div/div/div[2]/div/form/div/div[8]/div/div/nav/ul/li[*]/a";
        private const string MASSIVE_DELETE_SELECTALL_BUTTON = "selectAll";
        private const string MASSIVE_DELETE_CANCEL_BUTTON = "/html/body/div[3]/div/div/div[4]/button[1]";

        private const string MASSIVE_DELETE_SELECT_CHECKBOX = "/html/body/div[3]/div/div/div[2]/div/form/div/div[8]/div/div/table/tbody/tr[*]/td[1]/input[1]";
        private const string MASSIVE_DELETE_SERVICE_SELECT_COUNT = "serviceCount";
        private const string MASSIVE_DELETE_SERVICE_FROM_DATES = "/html/body/div[3]/div/div/div[2]/div/form/div/div[8]/div/div/table/tbody/tr[*]/td[5]";
        private const string MASSIVE_DELETE_SERVICE_TO_DATES = "/html/body/div[3]/div/div/div[2]/div/form/div/div[8]/div/div/table/tbody/tr[*]/td[6]";
        private const string MASSIVE_DELETE_SERVICE_SITE_LIST = "/html/body/div[3]/div/div/div[2]/div/form/div/div[8]/div/div/table/tbody/tr[*]/td[3]/div";
        private const string MASSIVE_DELETE_SERVICE_NAME_LIST = "/html/body/div[3]/div/div/div[2]/div/form/div/div[8]/div/div/table/tbody/tr[*]/td[2]/div";
        private const string MASSIVE_DELETE_SERVICE_CUSTOMER_LIST = "//*[@id=\"tableServices\"]/tbody/tr[*]/td[4]/div";
        private const string CUSTOMERHEADER_BTN = "//*[@id=\"tableServices\"]/thead/tr/th[4]/span/a";
        private const string DDL_PAGESIZE = "//*[@id=\"div-serviceResultTable\"]/descendant::select[contains(@id, 'page-size-selector')]";
        private const string CUSTOMERS_LIST = "/html/body/div[3]/div/div/div[2]/div/form/div/div[8]/div/div/table/tbody/tr[*]/td[4]/div";


        private const string CHECK_ALL_CUSTOMERS = "/html/body/div[16]/div/ul/li[1]/a";
        private const string CHECK_ALL_SITES = "/html/body/div[15]/div/ul/li[1]/a";
        private const string CHECK_ALL_STATUS = "/html/body/div[14]/div/ul/li[1]/a";

        private const string FROM_FILTER = "//*[@id=\"tableServices\"]/thead/tr/th[5]/span/a";
        private const string FROM_DATES = "//*[@id=\"tableServices\"]/tbody/tr[*]/td[5]";
        private const string TO_FILTER = "//*[@id=\"tableServices\"]/thead/tr/th[6]/span/a";
        private const string TO_DATES = "//*[@id=\"tableServices\"]/tbody/tr[*]/td[6]";


        private const string SORT_BY_SITE = "//*[@id=\"tableServices\"]/thead/tr/th[3]/span/a";
        private const string ITEM_SITE = "//*[@id=\"tableServices\"]/tbody/tr[*]/td[3]/div";
        private const string SORT_BY_SERVICE = "//*[@id=\"tableServices\"]/thead/tr/th[2]/span/a";
        private const string ITEM_SERVICE = "//*[@id=\"tableServices\"]/tbody/tr[*]/td[2]/div";
        [FindsBy(How = How.XPath, Using = MASSIVE_DELETE_CANCEL_BUTTON)]
        private IWebElement _CancelButton;


        [FindsBy(How = How.XPath, Using = SORT_BY_SITE)]
        private IWebElement _sortBySite;

        [FindsBy(How = How.XPath, Using = SORT_BY_SERVICE)]
        private IWebElement _sortByService;
        public ServiceMassiveDeleteModalPage(IWebDriver webDriver, TestContext testContext) : base(webDriver, testContext)
        {
        }
        public void Filter(ServiceMassiveDeleteFilterType filterType, object value)
        {
            WaitForLoad();

            string format = "dd/MM/yyyy";
            switch (filterType)
            {
                case ServiceMassiveDeleteFilterType.ServiceName:
                    var loadingPlanNameInput = WaitForElementIsVisible(By.XPath(MASSIVE_DELETE_SERVICE_NAME));
                    WaitPageLoading(); 
                    loadingPlanNameInput.SendKeys(value.ToString());
                    break;
                case ServiceMassiveDeleteFilterType.Site:
                    ComboBoxSelectById(new ComboBoxOptions(MASSIVE_DELETE_SERVICE_SITE, value.ToString(), false));
                    break;
                case ServiceMassiveDeleteFilterType.Customer:
                    ComboBoxSelectById(new ComboBoxOptions(MASSIVE_DELETE_SERVICE_CUSTOMER, value.ToString(), false));
                    break;
                case ServiceMassiveDeleteFilterType.From:
                    var fromInput = WaitForElementIsVisible(By.Id(MASSIVE_DELETE_SERVICE_FROM));
                    fromInput.Clear();
                    fromInput.SetValue(ControlType.DateTime, DateTime.ParseExact(value.ToString(), format, CultureInfo.InvariantCulture));
                    break;
                case ServiceMassiveDeleteFilterType.To:

                    var toInput = WaitForElementIsVisible(By.Id(MASSIVE_DELETE_SERVICE_TO));
                    toInput.Clear();
                    toInput.SetValue(ControlType.DateTime, DateTime.ParseExact(value.ToString(), format, CultureInfo.InvariantCulture));
                    break;
                case ServiceMassiveDeleteFilterType.Status:
                    ComboBoxSelectById(new ComboBoxOptions(MASSIVE_DELETE_SERVICE_STATUS, value.ToString(), false));
                    break;
            }

        }
        public void ClickSearchButton()
        {
            var searchBtn = WaitForElementIsVisible(By.Id(MASSIVE_DELETE_SERVICE_SEARCH_BUTTON));
            searchBtn.Click();
            WaitForLoad();
        }
        public new void PageSize(string size)
        {
            string pagesizetionxpath = MASSIVE_DELETE_PAGESIZE;
            if (!isElementExists(By.XPath(pagesizetionxpath)))
            {
                pagesizetionxpath = "/html/body/div[3]/div/div/div[2]/div/form/div/div[8]/div/div/div/nav/select";
            }
            var PageSizeDdl = _webDriver.FindElement(By.XPath(pagesizetionxpath));
            PageSizeDdl.SetValue(ControlType.DropDownList, size);
            WaitForLoad();
        }
        public void SelectFirst()
        {
            string checkBoxXpath = MASSIVE_DELETE_SELECT_CHECKBOX;
            if (!isElementExists(By.XPath(checkBoxXpath)))
            {
                checkBoxXpath = checkBoxXpath.Replace("div[8]/div/div/table", "div[8]/div/div/div/table");
            }
            var checkbox = _webDriver.FindElements(By.XPath(checkBoxXpath)).FirstOrDefault();
            checkbox.SetValue(ControlType.CheckBox, true);
            WaitForLoad();

        }
        public void DeleteFirstService()
        {
            SelectFirst();
            if (CheckTotalSelectCount() == 1)
            {
                var DeleteBtn = WaitForElementIsVisible(By.Id(MASSIVE_DELETE_SERVICE_DELETE_BTN));
                DeleteBtn.Click();
                var confirmBtn = WaitForElementIsVisible(By.Id(MASSIVE_DELETE_SERVICE_CONFIRM_BTN));
                confirmBtn.Click();
                var OkBtn = WaitForElementIsVisible(By.XPath(MASSIVE_DELETE_SERVICE_DELETEOK_BTN));
                OkBtn.Click();
                WaitForLoad();
            }
        }
        public void DeleteServiceForPrice()
        {
            SelectFirst();
           
                var DeleteBtn = WaitForElementIsVisible(By.Id(MASSIVE_DELETE_SERVICE_DELETE_BTN));
                DeleteBtn.Click();
                var confirmBtn = WaitForElementIsVisible(By.Id(MASSIVE_DELETE_SERVICE_CONFIRM_BTN));
                confirmBtn.Click();
                var OkBtn = WaitForElementIsVisible(By.XPath(MASSIVE_DELETE_SERVICE_DELETEOK_BTN));
                OkBtn.Click();
                WaitForLoad();    
        }
        public enum ServiceMassiveDeleteFilterType
        {
            ServiceName,
            Site,
            Customer,
            From,
            To,
            Status
        }
        public int CheckTotalPageNumber()
        {
            try
            {
                string paginationxpath = MASSIVE_DELETE_PAGINATION;
                if (!isElementExists(By.XPath(paginationxpath)))
                {
                    paginationxpath = MASSIVE_DELETE_PAGINATION.Replace("/div[8]/div/div/nav/ul/li", "/div[8]/div/div/div/nav/ul/li");
                }
                var numberPage = _webDriver.FindElements(By.XPath(paginationxpath)).Where(e => int.TryParse(e.Text, out int page)).Max(e => int.Parse(e.Text));
                return numberPage;
            }
            catch
            {
                return 0;
            }

        }
        public void ClickSelectAllButton()
        {
            var btn = WaitForElementIsVisible(By.Id(MASSIVE_DELETE_SELECTALL_BUTTON));
            btn.Click();
            WaitForLoad();
        }
        public void DeleteService()
        {
            var btn = WaitForElementIsVisible(By.Id("deleteServicesBtn"));
            btn.Click();
            WaitForLoad();
            var btnConfirm = WaitForElementIsVisible(By.Id("dataConfirmOK"));
            btnConfirm.Click();
            WaitForLoad();
            var btnOk = WaitForElementIsVisible(By.XPath("//*[@id=\"modal-1\"]/div[3]/button"));
            btnOk.Click();
            WaitForLoad();




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
        public int CheckTotalNumberByPage(int page)
        {
            GoToPage(page);
            WaitPageLoading();
            string checkBoxXpath = MASSIVE_DELETE_SELECT_CHECKBOX;
            if (!isElementExists(By.XPath(checkBoxXpath)))
            {
                checkBoxXpath = checkBoxXpath.Replace("div[8]/div/div/table", "div[8]/div/div/div/table");
            }
            var checkboxList = _webDriver.FindElements(By.XPath(checkBoxXpath));
            return checkboxList.Count;
        }
        public int CheckTotalSelectCount()
        {
            WaitPageLoading();
            var count = _webDriver.FindElement(By.Id(MASSIVE_DELETE_SERVICE_SELECT_COUNT));
            var totalnumber = count.Text.Trim();
            return Convert.ToInt32(totalnumber);
        }
        public bool VerifySelectedAllByPage(int page)
        {
            GoToPage(page);
            string chekboxXpath = MASSIVE_DELETE_SELECT_CHECKBOX;
            if (!isElementExists(By.XPath(chekboxXpath)))
            {
                chekboxXpath = chekboxXpath.Replace("/div[8]/div/div/table/tbody/", "/div[8]/div/div/div/table/tbody/");
            }
            var checkBoxlist = _webDriver.FindElements(By.XPath(chekboxXpath));
            return checkBoxlist.Where(e => e.Selected).Count() == checkBoxlist.Count();
        }
        public ServicePage Cancel()
        {
            _CancelButton = WaitForElementIsVisible(By.XPath(MASSIVE_DELETE_CANCEL_BUTTON));
            _CancelButton.Click();
            WaitForLoad();
            return new ServicePage(_webDriver, _testContext);
        }
        public bool VerifyFromTo(DateTime from, DateTime to)
        {
            string fromDateXpath = MASSIVE_DELETE_SERVICE_FROM_DATES;
            if (!isElementExists(By.XPath(fromDateXpath)))
            {
                fromDateXpath = fromDateXpath.Replace("/div[8]/div/div/table/tbody/", "/div[8]/div/div/div/table/tbody/");
            }
            string todateXpath = MASSIVE_DELETE_SERVICE_TO_DATES;
            if (!isElementExists(By.XPath(todateXpath)))
            {
                todateXpath = todateXpath.Replace("/div[8]/div/div/table/tbody/", "/div[8]/div/div/div/table/tbody/");
            }
            var listFromDates = _webDriver.FindElements(By.XPath(fromDateXpath));
            var listTODates = _webDriver.FindElements(By.XPath(todateXpath));


            for (int i = 0; i < listFromDates.Count; i++)
            {
                DateTime parsedDateFrom, parsedDateTo;
                if (DateTime.TryParseExact(listFromDates[i].Text.Substring(0, 10), "dd/MM/yyyy", CultureInfo.InvariantCulture, DateTimeStyles.None, out parsedDateFrom))
                {
                    if (DateTime.TryParseExact(listTODates[i].Text.Substring(0, 10), "dd/MM/yyyy", CultureInfo.InvariantCulture, DateTimeStyles.None, out parsedDateTo))
                    {
                        if (from > parsedDateTo && to < parsedDateFrom)
                        {
                            return false;
                        }
                    }
                }
            }
            return true;
        }
        public bool VerifySiteSearch(string site)
        {
            string sitexpath = MASSIVE_DELETE_SERVICE_SITE_LIST;
            if (!isElementExists(By.XPath(sitexpath)))
            {
                sitexpath = sitexpath.Replace("/div[8]/div/div/table/tbody/", "/div[8]/div/div/div/table/tbody/");
            }
            var listSites = _webDriver.FindElements(By.XPath(sitexpath)).Select(e => e.Text.Trim());

            return listSites.Where(s => s.Equals(site)).Count() == listSites.Count();
        }

        public void ShowInactiveCustomers()
        {
            var checkbox = WaitForElementExists(By.XPath("//*/label[@for=\"ShowInactiveCustomers\"]/../div/input"));
            checkbox.Click();
            // cascade checkbox vers combobox
            WaitForLoad();
        }

        public bool IsInactive()
        {
            WaitForLoad();
            WaitPageLoading();

            // Récupérer la liste des éléments
            var list = _webDriver.FindElements(By.XPath("//li[not(contains(@class,'ui-multiselect-excluded'))]/label"));
          
            int i = 0;
            while (i < list.Count)
            {
                var text = list[i].Text.Trim();
                if (text.StartsWith("Inactive -"))
                {
                    return true; 
                }
                i++; 
            }

            return false;
        }


        public void ClickOnCustomerFilter()
        {
            var select = WaitForElementExists(By.Id(MASSIVE_DELETE_SERVICE_CUSTOMER));
            select.Click();
        }

        public void ShowInactiveSites()
        {
            var checkbox = WaitForElementExists(By.XPath("//*/label[@for=\"ShowInactiveSites\"]/../div/input"));
            checkbox.Click();
            WaitForLoad();
        }

        public void ClickOnSiteFilter()
        {
            var select = WaitForElementExists(By.Id(MASSIVE_DELETE_SERVICE_SITE));
            select.Click();
        }

        public bool VerifyServiceNameSearch(string serviceName)
        {
            var listSites = _webDriver.FindElements(By.XPath(MASSIVE_DELETE_SERVICE_NAME_LIST));

            return listSites.Where(s => s.Text.Equals(serviceName)).Count() == listSites.Count();
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
            string xPath = "//*[@id=\"tableServices\"]/tbody/tr[*]/td[4]"; 
            var customerCodes = _webDriver.FindElements(By.XPath(xPath));

            foreach (var code in customerCodes)
            {
                ids.Add(code.Text);
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

        public string GetFirstServiceName()
        {
            IWebElement rowResult = null;
            try
            {
                rowResult = WaitForElementExists(By.XPath("//*[@id=\"tableServices\"]/tbody/tr[1]/td[2]/div"));
            }
            catch
            { 
                //silent catch : ignore if element isn't found
            }
            if (rowResult == null)
            {
                return null;
            }
            return rowResult.Text;
        }

        public bool ClickOnResultLinkAndCheckServiceName(string serviceName)
        {
            IWebElement rowResultLink = null;
            try
            {
                rowResultLink = WaitForElementIsVisible(By.XPath("//*[@id=\"tableServices\"]/tbody/tr[1]/td[8]/a")); 
            }
            catch
            { 
                //silent catch : ignore if element isn't found
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

                var title = WaitForElementExists(By.XPath("//*[@id=\"div-body\"]/div/h1"));

                return title.Text.IndexOf(serviceName, StringComparison.OrdinalIgnoreCase) >= 0;
            }
        }

        public bool CheckData()
        {
            try
            {
                var text = WaitForElementExists(By.XPath("//*[@id='tableServices']/tbody/tr/td[2]/b")).Text;

                if (!text.Equals("No data to display. Use the Search button to load data")) return true;
                return false;
            }
            catch
            {
                return true;
            }

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
            var table = _webDriver.FindElement(By.Id("tableServices"));

            var rows = table.FindElements(By.ClassName("closestFind"));

            var totalRows = rows.Count();
            return totalRows;
        }


        public bool VerifyCustomerSearch(string customerValue, string type)
        {
            var xpath = CUSTOMERS_LIST;

            if (type == "Inactive")
            {
                xpath = MASSIVE_DELETE_SERVICE_CUSTOMER_LIST;
            }
            var list = _webDriver.FindElements(By.XPath(xpath));
            foreach(var value in list)
            {
                if(value.Text.Replace(" ", "") != customerValue.Replace(" ", "")) return false;
            }
            return true;
        }


        public void CheckAllCustomer()
        {

            var customersDropDown = WaitForElementIsVisible(By.Id(MASSIVE_DELETE_SERVICE_CUSTOMER));
            customersDropDown.Click();
            var checkAllCustomers = WaitForElementIsVisible(By.XPath(CHECK_ALL_CUSTOMERS));
            checkAllCustomers.Click();
            customersDropDown.Click();
        }

        public void CheckAllSites()
        {

            var sitesDropDown = WaitForElementIsVisible(By.Id(MASSIVE_DELETE_SERVICE_SITE));
            sitesDropDown.Click();
            var checkAllSites = WaitForElementIsVisible(By.XPath(CHECK_ALL_SITES));
            checkAllSites.Click();
            sitesDropDown.Click();
        }

        public void CheckAllStatus()
        {

            var statusDropDown = WaitForElementIsVisible(By.Id(MASSIVE_DELETE_SERVICE_STATUS));
            statusDropDown.Click();
            var checkAllStatus = WaitForElementIsVisible(By.XPath(CHECK_ALL_STATUS));
            checkAllStatus.Click();
            statusDropDown.Click();
        }

      
        public void ClickFromFilter()
        {
            var fromBtn = WaitForElementIsVisible(By.XPath(FROM_FILTER));
            fromBtn.Click();
            WaitForLoad();
        }

        public void ClickToFilter()
        {
            var toBtn = WaitForElementIsVisible(By.XPath(TO_FILTER));
            toBtn.Click();
            WaitForLoad();
        }

        public bool VerifyFromDatesASC()
        {
            var datesFrom = _webDriver.FindElements(By.XPath(FROM_DATES)).Select(e=>e.Text);
            List<DateTime> dates = new List<DateTime>();
            foreach (var dateString in datesFrom)
            {
                dates.Add(DateTime.ParseExact(dateString, "dd/MM/yyyy", CultureInfo.InvariantCulture));
            }
            
            for (int i = 1; i < dates.Count; i++)
            {
                if (dates[i] < dates[i - 1])
                {
                    return false;
                }
            }
            return true;
        }

        public bool VerifyToDatesASC()
        {
            var datesTo = _webDriver.FindElements(By.XPath(TO_DATES)).Select(e => e.Text);
            List<DateTime> dates = new List<DateTime>();
            foreach (var dateString in datesTo)
            {
                dates.Add(DateTime.ParseExact(dateString, "dd/MM/yyyy", CultureInfo.InvariantCulture));
            }

            for (int i = 1; i < dates.Count; i++)
            {
                if (dates[i] < dates[i - 1])
                {
                    return false;
                }
            }
            return true;
        }



        public void SortBySite()
        {
            _sortBySite = WaitForElementIsVisible(By.XPath(SORT_BY_SITE));
            _sortBySite.Click();
            WaitForLoad();
        }

        public bool IsSortedBySite()
        {
            var elements = _webDriver.FindElements(By.XPath(ITEM_SITE));

            if (elements.Count == 0)
                return false;

            for (int i = 1; i < elements.Count; i++)
            {
                string ancienSite = elements[i - 1].GetAttribute("InnerText");
                string currentSite = elements[i].GetAttribute("InnerText");

                if (string.Compare(ancienSite, currentSite) > 0)
                    return false;
            }

            return true;
        }

        public void CheckOnlyActiveServicesStatus()
        {
            var onlyActiveServices = "Active Service";

            var statusDropDown = WaitForElementIsVisible(By.Id(MASSIVE_DELETE_SERVICE_STATUS));
            statusDropDown.Click();

            var statusDropDownInput = WaitForElementIsVisible(By.XPath("/html/body/div[14]/div/div/label/input"));
            statusDropDownInput.SetValue(ControlType.TextBox,onlyActiveServices);

            var checkonlyActiveServicetatus = WaitForElementIsVisible(By.XPath("/html/body/div[14]/ul/li[1]/label/span"));
            checkonlyActiveServicetatus.Click();
            statusDropDown.Click();
        }

        public List<string> FillAndGetDistinctServiceList()
        {
            
            ClickSearchButton();

            WaitForLoad();
            var listService = _webDriver.FindElements(By.XPath("//*[@id=\"tableServices\"]/tbody/tr[*][@class=\"closestFind\"]"));
            List<string> result = new List<string>();
            foreach (var item in listService)
            {
                result.Add(item.Text);
            }

            result = result.Distinct().ToList();

            return result;
        }
        public void SortByService()
        {
            _sortByService = WaitForElementIsVisible(By.XPath(SORT_BY_SERVICE));
            _sortByService.Click();
            WaitForLoad();
        }
        public bool IsSortedByService()
        {
            var elements = _webDriver.FindElements(By.XPath(ITEM_SERVICE));

            if (elements.Count == 0)
                return false;

            for (int i = 1; i < elements.Count; i++)
            {
                string ancienService = elements[i - 1].GetAttribute("InnerText");
                string currentService = elements[i].GetAttribute("InnerText");

                if (string.Compare(ancienService, currentService) > 0)
                    return false;
            }

            return true;
        }

    }
}
