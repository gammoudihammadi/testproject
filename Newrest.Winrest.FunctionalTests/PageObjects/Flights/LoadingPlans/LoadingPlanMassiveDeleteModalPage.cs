using Microsoft.VisualStudio.TestTools.UnitTesting;
using Newrest.Winrest.FunctionalTests.PageObjects.Shared;
using Newrest.Winrest.FunctionalTests.Utils;
using OpenQA.Selenium;
using OpenQA.Selenium.Support.UI;
using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Text.RegularExpressions;
using System.Web.UI.WebControls;
using System.Windows.Forms;

namespace Newrest.Winrest.FunctionalTests.PageObjects.Flights.LoadingPlans
{
    public class LoadingPlanMassiveDeleteModalPage : PageBase
    {
        private const string MASSIVE_DELETE_LOADING_PLAN_NAME = "SearchPattern";
        private const string MASSIVE_DELETE_LOADING_PLAN_NAME_XPATH = "/html/body/div[3]/div/div/div[2]/div/form/div/div[1]/div/input";
        private const string MASSIVE_DELETE_LOADING_PLAN_SITE = "SelectedSiteIds_ms";
        private const string MASSIVE_DELETE_LOADING_PLAN_CUSTOMER = "SelectedCustomersToDelete_ms";
        private const string MASSIVE_DELETE_LOADING_PLAN_USE_CASE = "selectedUseCase";
        private const string MASSIVE_DELETE_LOADING_PLAN_FROM = "dateloadingplanStart";
        private const string MASSIVE_DELETE_LOADING_PLAN_TO = "dateLoadingPlanEnd";
        private const string MASSIVE_DELETE_LOADING_PLAN_SEARCH_BUTTON = "SearchLoadingPlansBtn";
        private const string MASSIVE_DELETE_LOADING_PLAN_CUSTOMER_NAMES = "/html/body/div[3]/div/div/div[2]/div/form/div/div[8]/div/div/table/tbody/tr[*]/td[4]/div";
        private const string MASSIVE_DELETE_LOADING_PLAN_SHOW_INACTIVE_CUSTOMERS = "ShowInactiveCustomers";
        private const string MASSIVE_DELETE_LOADING_PLAN_SHOW_INACTIVE_SITES = "ShowInactiveSites";
        private const string MASSIVE_DELETE_LOADING_PLAN_DROPDOWN_INPUT = "SelectedCustomersToDelete_ms";
        private const string MASSIVE_DELETE_LOADING_PLAN_DROPDOWN_INPUT_SITES = "SelectedSiteIds_ms";
        private const string MASSIVE_DELETE_LOADING_PLAN_SHOW_CUSTOMERS_INACTIVE_LIST = "/html/body/div[13]/ul/li/label/span[contains(text(), 'Inactive')]";
        private const string MASSIVE_DELETE_LOADING_PLAN_SHOW_SITES_INACTIVE_LIST = "/html/body/div[13]/ul/li/label/span[contains(text(), 'Inactive')]";


        private const string MASSIVE_DELETE_LOADING_PLAN_SITE_NAMES = "/html/body/div[3]/div/div/div[2]/div/form/div/div[8]/div/div/table/tbody/tr[*]/td[3]/div";
        private const string MASSIVE_DELETE_LOADING_PLAN_LP_NAMES = "/html/body/div[3]/div/div/div[2]/div/form/div/div[8]/div/div/table/tbody/tr[*]/td[2]/div";
        private const string MASSIVE_DELETE_SORT_BY_LP_NAME = "/html/body/div[3]/div/div/div[2]/div/form/div/div[8]/div/div/table/thead/tr/th[2]/span/a";
        private const string MASSIVE_DELETE_CHECK_CUSTOMER_INACTIVE = "/html/body/div[13]/ul/li/label/span[contains(text(), 'Inactive')]/../input";
        private const string MASSIVE_DELETE_SORT_BY_SERVICENUMBER = "/html/body/div[3]/div/div/div[2]/div/form/div/div[8]/div/div/table/thead/tr/th[5]/span/a/span";
        private const string MASSIVE_DELETE_SORT_BY_LP_NAME_DESC = "/html/body/div[3]/div/div/div[2]/div/form/div/div[8]/div/div/div/table/thead/tr/th[2]/span/a";
        private const string MASSIVE_DELETE_SORT_BY_SERVICENUMBER_ASC = "/html/body/div[3]/div/div/div[2]/div/form/div/div[8]/div/div/div/table/thead/tr/th[5]/span/a";
        private const string MASSIVE_DELETE_SORT_BY_SERVICENUMBER_DESC = "/html/body/div[3]/div/div/div[2]/div/form/div/div[8]/div/div/table/thead/tr/th[5]/span/a";
        private const string MASSIVE_DELETE_SERVICE_NUMBER = "/html/body/div[3]/div/div/div[2]/div/form/div/div[8]/div/div/table/tbody/tr[*]/td[5]";
        private const string MASSIVE_DELETE_SORT_BY_SITE_ASC = "/html/body/div[3]/div/div/div[2]/div/form/div/div[8]/div/div/div/table/thead/tr/th[3]/span/a";
        private const string MASSIVE_DELETE_SORT_BY_SITE_DESC = "/html/body/div[3]/div/div/div[2]/div/form/div/div[8]/div/div/table/thead/tr/th[3]/span/a";
        private const string MASSIVE_DELETE_SITE = "/html/body/div[3]/div/div/div[2]/div/form/div/div[8]/div/div/table/tbody/tr[*]/td[3]/div";
        private const string MASSIVE_DELETE_PAGESIZE = "/html/body/div[3]/div/div/div[2]/div/form/div/div[8]/div/div/nav/select";
        private const string FIRST_SEARCH_LINK = "/html/body/div[3]/div/div/div[2]/div/form/div/div[8]/div/div/table/tbody/tr[1]/td[6]/a";
        private const string MASSIVE_DELETE_PAGINATION = "/html/body/div[3]/div/div/div[2]/div/form/div/div[8]/div/div/nav/ul/li[*]/a";


        private const string LIST_SITE_ACTIF = "/html/body/div[13]/ul";
        private const string LIST_SITE_ACTIFSS = "/html/body/div[13]/ul";
        private const string MASSIVE_DELETE_CHECK_ONLY_SITES_INACTIVE = "//*[@id=\"selectedUseCase\"]/option[3]";
        private const string MASSIVE_DELETE_CHECK_ONLY_SITES_INACTIVE_SELECT = "/html/body/div[3]/div/div/div[2]/div/form/div/div[5]/div/select/option[3]";
        private const string MASSIVE_DELETE_CHECK_ONLY_SITES_ACTIVE = "//*[@id=\"selectedUseCase\"]/option[2]";
        private const string MASSIVE_DELETE_SELECTALL_BUTTON = "selectAll";
        private const string MASSIVE_DELETE_UNSELECTALL_BUTTON = "unselectAll";
        private const string MASSIVE_DELETE_LOADING_PLAN_SELECTED = "/html/body/div[3]/div/div/div[2]/div/form/div/div[8]/div/div/table/tbody/tr[*]/td[1]/input[1]";

        private const string MASSIVE_DELETE_lOADING_PLAN_COUNTt = "loadingplanCount";
        private const string MASSIVE_DELETE_CHECK_ALL_SITES = "//*[@id=\"selectedUseCase\"]/option[1]";


        private const string MASSIVE_DELETE_SORT_BY_CUSTOMER = "/html/body/div[3]/div/div/div[2]/div/form/div/div[8]/div/div/table/thead/tr/th[4]/span";
        private const string MASSIVE_DELETE_CUSTOMERS = "/html/body/div[3]/div/div/div[2]/div/form/div/div[8]/div/div/table/tbody/tr[*]/td[4]/div";

        private const string USE_CASE_SELECT = "//*[@id=\"selectedUseCase\"]";
        private const string ONLY_INACTIVE_CUSTOMERS_USE_CASE = "//*[@id=\"selectedUseCase\"]/option[5]";
        private const string SITE_FILTER = "SelectedSiteIds_ms";
        private const string MASSIVE_DELETE_SITE_LIST = "//*[@id=\"tableLoadingPlans\"]/tbody/tr[*]/td[3]/div";

        public LoadingPlanMassiveDeleteModalPage(IWebDriver webDriver, TestContext testContext) : base(webDriver, testContext)
        {
        }
        public enum FilterType
        {
            loadingPlanName,
            Site,
            Customer,
            From,
            To,
            UseCase
        }

        public void Filter(FilterType filterType, object value)
        {
            string format = "dd/MM/yyyy";
            switch (filterType)
            {
                case FilterType.loadingPlanName:
                    var loadingPlanNameInput = WaitForElementIsVisible(By.XPath(MASSIVE_DELETE_LOADING_PLAN_NAME_XPATH));
                    loadingPlanNameInput.SendKeys(value.ToString());
                    break;
                case FilterType.Site:
                    ComboBoxSelectById(new ComboBoxOptions(MASSIVE_DELETE_LOADING_PLAN_SITE, value.ToString(),false));
                    break;
                case FilterType.Customer:
                    ComboBoxSelectById(new ComboBoxOptions(MASSIVE_DELETE_LOADING_PLAN_CUSTOMER, value.ToString(),false));
                    break;
                case FilterType.From:
                    var fromInput = WaitForElementIsVisible(By.Id(MASSIVE_DELETE_LOADING_PLAN_FROM));
                    fromInput.Clear();
                    fromInput.SetValue(ControlType.DateTime, DateTime.ParseExact((string)value, format, CultureInfo.InvariantCulture));
                    break;
                case FilterType.To:

                    var toInput = WaitForElementIsVisible(By.Id(MASSIVE_DELETE_LOADING_PLAN_TO));
                    toInput.Clear();
                    toInput.SetValue(ControlType.DateTime, DateTime.ParseExact((string)value, format, CultureInfo.InvariantCulture));
                    break;
                case FilterType.UseCase:
                    ComboBoxSelectById(new ComboBoxOptions(MASSIVE_DELETE_LOADING_PLAN_USE_CASE, value.ToString(), false));
                    break;
            }
        }
        public void ClickSearchButton()
        {
            var searchBtn = WaitForElementIsVisible(By.Id(MASSIVE_DELETE_LOADING_PLAN_SEARCH_BUTTON));
            searchBtn.Click();
            WaitForLoad();
        }
        public bool VerifyAllCustomers(string value)
        {


            var cutomerNames = _webDriver.FindElements(By.XPath(MASSIVE_DELETE_LOADING_PLAN_CUSTOMER_NAMES)).Select(e => e.Text);
            foreach (var cn in cutomerNames)
            {
                if (!RemoveSpaces(cn).Equals(RemoveSpaces(value)))
                    return false;
            }
            return true;

        }
        public void ShowInactiveCustomers(bool isChecked)
        {
            var showInactiveCustomersCheckBox = WaitForElementExists(By.Id(MASSIVE_DELETE_LOADING_PLAN_SHOW_INACTIVE_CUSTOMERS));
            showInactiveCustomersCheckBox.SetValue(ControlType.CheckBox, isChecked);
            WaitForLoad();
        }
        public bool VerifyCustomersHasInactive()
        {

            var customersInput = WaitForElementExists(By.Id(MASSIVE_DELETE_LOADING_PLAN_DROPDOWN_INPUT));
            customersInput.Click();
            var customersInactive = _webDriver.FindElements(By.XPath(MASSIVE_DELETE_LOADING_PLAN_SHOW_CUSTOMERS_INACTIVE_LIST)).Select(e => e.Text);
            if (customersInactive.Any())
                return true;
            return false;
        }

        //public void ShowInactiveSites(bool isChecked)
        //{
        //    var showInactiveSitesCheckBox = WaitForElementExists(By.Id(MASSIVE_DELETE_LOADING_PLAN_SHOW_INACTIVE_SITES));
        //    showInactiveSitesCheckBox.SetValue(ControlType.CheckBox, isChecked);
        //    WaitForLoad();
        //}

        public string RemoveSpaces(string str)
        {
            return Regex.Replace(str, @"\s+", "");
        }

        public bool VerifyAllSites(string value)
        {
            var SiteNames = getLodingPlanSiteName();
            foreach (var cn in SiteNames)
            {
                if (!RemoveSpaces(cn).Equals(RemoveSpaces(value)))
                    return false;
            }
            return true;
        }
        public void SortByLoadingPlanNameDESC()
        {
            var SortByLoadingPlanNameBtn = WaitForElementIsVisible(By.XPath(MASSIVE_DELETE_SORT_BY_LP_NAME_DESC));
            SortByLoadingPlanNameBtn.Click();
            WaitForLoad();
        }
        public IEnumerable<string> getLodingPlanNames()
        {
            return _webDriver.FindElements(By.XPath(MASSIVE_DELETE_LOADING_PLAN_LP_NAMES)).Select(e => e.Text);
        }
        public bool VerifySortByLPNameASC()
        {
            var lodingPlanNames = getLodingPlanNames().ToList();

            for (var i = 0; i < lodingPlanNames.Count - 1; i++)
            {
                if (lodingPlanNames[i].CompareTo(lodingPlanNames[i + 1]) > 0)
                {
                    return false;
                }
            }

            return true;
        }
        public bool VerifySortByLPNameDESC()
        {
            var lodingPlanNames = getLodingPlanNames().ToList();
            for (var i = 0; i < lodingPlanNames.Count - 1; i++)
            {
                if (lodingPlanNames[i].CompareTo(lodingPlanNames[i + 1]) < 0)
                {
                    return false;
                }
            }
            return true;
        }

        public IEnumerable<string> getLodingPlanServiceNumbers()
        {
            return _webDriver.FindElements(By.XPath(MASSIVE_DELETE_SERVICE_NUMBER)).Select(e => e.Text);
        }
        public void SortByServiceNumberASC()
        {
            var SortByServiceNumberBtn = WaitForElementIsVisible(By.XPath(MASSIVE_DELETE_SORT_BY_SERVICENUMBER_ASC));
            SortByServiceNumberBtn.Click();
            WaitForLoad();
        }
        public void SortByServiceNumberDESC()
        {
            var SortByServiceNumberBtn = WaitForElementIsVisible(By.XPath(MASSIVE_DELETE_SORT_BY_SERVICENUMBER_DESC));
            SortByServiceNumberBtn.Click();
            WaitForLoad();
        }
        public bool VerifySortByServiceNumberASC()
        {
            List<string> serviceNumbers = getLodingPlanServiceNumbers().ToList();
            for (int i = 0; i < serviceNumbers.Count - 1; i++)
                if (int.Parse(serviceNumbers[i]).CompareTo(int.Parse(serviceNumbers[i + 1])) > 0)
                    return false;
            return true;
        }
        public bool VerifySortByServiceNumberDESC()
        {
            List<string> serviceNumbers = getLodingPlanServiceNumbers().ToList();
            for (int i = 0; i < serviceNumbers.Count - 1; i++)
                if (int.Parse(serviceNumbers[i]).CompareTo(int.Parse(serviceNumbers[i + 1])) < 0)
                    return false;
            return true;
        }

        public void SortBySiteASC()
        {
            var SortBySiteBtn = WaitForElementIsVisible(By.XPath(MASSIVE_DELETE_SORT_BY_SITE_ASC)); //ASC
            SortBySiteBtn.Click();
            WaitForLoad();
        }
        public void SortBySiteDESC()
        {
            var SortBySiteBtn = WaitForElementIsVisible(By.XPath(MASSIVE_DELETE_SORT_BY_SITE_DESC));
            SortBySiteBtn.Click();
            WaitForLoad();
        }

        public bool VerifySortBySiteASC()
        {
            var SiteNames = getLodingPlanSiteName().ToList();
            for (var i = 0; i < SiteNames.Count - 1; i++)
            {
                if (SiteNames[i].CompareTo(SiteNames[i + 1]) > 0)
                {
                    return false;
                }
            }
            return true;
        }

        public void SelectInactiveSites()
        {
            var inactiveSelect = WaitForElementExists(By.XPath("//*[@id=\"ShowInactiveSites\"]"));
            inactiveSelect.Click();
            WaitForLoad();

            var useCaseSelect = WaitForElementExists(By.XPath("//*[@id=\"selectedUseCase\"]"));
            useCaseSelect.SetValue(ControlType.DropDownList, "Only Inactive Sites");
            WaitForLoad();
        }

        public void SelectActiveSites()
        {

            var inactiveSelect = _webDriver.FindElements(By.XPath(MASSIVE_DELETE_CHECK_ONLY_SITES_ACTIVE));
            WaitForLoad();
            inactiveSelect.FirstOrDefault().Click();
        }

        public void SelecAllSites()
        {
            var allSites = _webDriver.FindElements(By.XPath(MASSIVE_DELETE_CHECK_ALL_SITES));
            WaitForLoad();
            allSites.FirstOrDefault().Click();
        }

        public bool CheckIfAllSitesArePresent()
        {
            bool defaultval = true;
            ShowInactiveSites(true);
            var idss = "SelectedSiteIds_ms";
            var listsitss = WaitForElementExists(By.Id(idss));
            listsitss.Click();
            var xbout = "/html/body/div[13]/ul/li";
            WaitForLoad();
            var listOfAllSites = _webDriver.FindElements(By.XPath(xbout)).Select(e => e.Text).ToList();

            var listSiteResult = getcurrentloadinPlanSiteName().ToList();
            bool hasMatch = listSiteResult.Any(x => listOfAllSites.Any(y => y == x));
            defaultval = hasMatch;
            //var page2 = "//*[@id=\"list-loadingplans-deletion\"]/nav/ul/li[3]/a";
            var nbpage = "//*[@id=\"list-loadingplans-deletion\"]/nav/ul/li";
            WaitForLoad();
            var nombrPages = _webDriver.FindElements(By.XPath(nbpage)).Select(e => e.Text).ToList().Count();

            for (int i = 3; i < nombrPages; i++)
            {
                var currentpage = "//*[@id=\"list-loadingplans-deletion\"]/nav/ul/li[" + i + "]/a";
                WaitForLoad();
                var pagesuivante = _webDriver.FindElements(By.XPath(currentpage)).FirstOrDefault();
                pagesuivante.Click();

                var listSitePAgeCourante = getcurrentloadinPlanSiteName().ToList();
                bool isMatching = listSiteResult.Any(x => listOfAllSites.Any(y => y == x));
                defaultval = defaultval && isMatching;
            }

            return defaultval;
        }

        public bool CheckIfSitesAreActive()
        {

            bool defaultval = true;
            var idss = "SelectedSiteIds_ms";
            var listsitss = WaitForElementExists(By.Id(idss));
            listsitss.Click();
            var xbout = "/html/body/div[12]/ul/li";
            WaitForLoad();
            var listsitactive = _webDriver.FindElements(By.XPath(xbout)).Select(e => e.Text).ToList();


            var listSiteResult = getcurrentloadinPlanSiteName().ToList();
            bool hasMatch = listSiteResult.Any(x => listsitactive.Any(y => y == x));
            defaultval = hasMatch;
            //var page2 = "//*[@id=\"list-loadingplans-deletion\"]/nav/ul/li[3]/a";
            var nbpage = "//*[@id=\"list-loadingplans-deletion\"]/nav/ul/li";
            WaitForLoad();
            var nombrPages = _webDriver.FindElements(By.XPath(nbpage)).Select(e => e.Text).ToList().Count();

            for (int i = 3; i < nombrPages; i++)
            {
                var currentpage = "//*[@id=\"list-loadingplans-deletion\"]/nav/ul/li[" + i + "]/a";
                WaitForLoad();
                var pagesuivante = _webDriver.FindElements(By.XPath(currentpage)).FirstOrDefault();
                pagesuivante.Click();

                var listSitePAgeCourante = getcurrentloadinPlanSiteName().ToList();
                bool isMatching = listSiteResult.Any(x => listsitactive.Any(y => y == x));
                defaultval = defaultval && isMatching;
            }

            return defaultval;
        }

        public bool CheckIfSitesAreInactive()
        {
            return !CheckIfSitesAreActive();
        }

        public void SelectInactiveCustomers()
        {
            var customersInput = WaitForElementExists(By.Id(MASSIVE_DELETE_LOADING_PLAN_DROPDOWN_INPUT));
            customersInput.Click();
            var checkboxes = _webDriver.FindElements(By.XPath(MASSIVE_DELETE_CHECK_CUSTOMER_INACTIVE));
            foreach (var checkbox in checkboxes)
            {
                checkbox.Click();
                WaitForLoad();
            }
            customersInput.Click();
        }
        public bool VerifyAllCustomersAreInactive()
        {
            var customers = _webDriver.FindElements(By.XPath("/html/body/div[3]/div/div/div[2]/div/form/div/div[8]/div/div/table/tbody/tr[*]/td[4]"));
            foreach (var c in customers)
            {
                if (c.GetAttribute("title") != "Inactive")
                    return false;
            }
            return true;
        }

        public bool VerifyAllSitesAreInactive()
        {
            var sites = _webDriver.FindElements(By.XPath("/html/body/div[3]/div/div/div[2]/div/form/div/div[8]/div/div/table/tbody/tr[*]/td[3]"));
            foreach (var s in sites)
            {
                if (s.GetAttribute("title") != "Inactive")
                    return false;
            }
            return true;
        }

        public bool VerifySortBySiteDESC()
        {
            var SiteNames = getLodingPlanSiteName().ToList();
            for (var i = 0; i < SiteNames.Count - 1; i++)
            {
                if (SiteNames[i].CompareTo(SiteNames[i + 1]) < 0)
                {
                    return false;
                }
            }
            return true;
        }
        public IEnumerable<String> getLodingPlanSiteName()
        {
            return _webDriver.FindElements(By.XPath(MASSIVE_DELETE_LOADING_PLAN_SITE_NAMES)).Select(e => e.Text);
        }
        public new void PageSize(string size)
        {
            string pagesizetionxpath = MASSIVE_DELETE_PAGESIZE;
            if (!isElementExists(By.XPath(pagesizetionxpath)))
            {
                pagesizetionxpath = pagesizetionxpath.Replace("div[8]/div/div/nav/select", "div[8]/div/div/div/nav/select");
            }
            var PageSizeDdl = _webDriver.FindElement(By.XPath(pagesizetionxpath));
            PageSizeDdl.SetValue(ControlType.DropDownList, size);
            WaitForLoad();
        }

        public void ShowInactiveSites(bool isChecked)
        {
            var showInactiveSitesCheckBox = WaitForElementExists(By.Id(MASSIVE_DELETE_LOADING_PLAN_SHOW_INACTIVE_SITES));
            showInactiveSitesCheckBox.SetValue(ControlType.CheckBox, isChecked);
            WaitForLoad();
        }
        public bool VerifySitesHasInactive()
        {
            IWebElement sitesInput = WaitForElementExists(By.Id(MASSIVE_DELETE_LOADING_PLAN_DROPDOWN_INPUT_SITES));
            sitesInput.Click();
            IEnumerable<string> sitesInactive = _webDriver.FindElements(By.XPath(MASSIVE_DELETE_LOADING_PLAN_SHOW_SITES_INACTIVE_LIST)).Select(e => e.Text);
            if (sitesInactive.Any())
                return true;
            return false;
        }

        public LoadingPlansGeneralInformationsPage ClickFirstLink()
        {
            var link = WaitForElementExists(By.XPath(FIRST_SEARCH_LINK));

            link.Click();

            var wait = new WebDriverWait(_webDriver, TimeSpan.FromSeconds(60));
            wait.Until((driver) => driver.WindowHandles.Count > 1);
            _webDriver.SwitchTo().Window(_webDriver.WindowHandles[1]);

            return new LoadingPlansGeneralInformationsPage(_webDriver, _testContext);
        }

        public IEnumerable<String> getcurrentloadinPlanSiteName()
        {
            var cursite = "//*[@id=\"tableLoadingPlans\"]/tbody/tr/td[3]";
            WaitForLoad();
            return _webDriver.FindElements(By.XPath(cursite)).Select(e => e.Text);
        }
        public int getNumberPage()
        {
            string paginationxpath = MASSIVE_DELETE_PAGINATION;
            if (!isElementExists(By.XPath(paginationxpath)))
            {
                paginationxpath = MASSIVE_DELETE_PAGINATION.Replace("/div[8]/div/div/nav/ul/li", "/div[8]/div/div/div/nav/ul/li");
            }
            var numberPage = _webDriver.FindElements(By.XPath(paginationxpath)).Where(e => int.TryParse(e.Text, out int page)).Max(e => int.Parse(e.Text));
            return numberPage;
        }

        public void goToPage(int page)
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
        public new int CheckTotalNumber()
        {
            ClickSelectAllButton();
            var count = _webDriver.FindElement(By.Id(MASSIVE_DELETE_lOADING_PLAN_COUNTt));
            var totalnumber = count.Text.Trim();
            ClickUnSelectAllButton();
            return Convert.ToInt32(totalnumber);
        }
        public bool VerifySelectedAll(int selectedCount)
        {
            ClickSelectAllButton();
            var NumberPage = getNumberPage();
            var TotalNumber = 0;
            for (int i = 1; i <= NumberPage; i++)
            {
                goToPage(i);
                string chekboxXpath = MASSIVE_DELETE_LOADING_PLAN_SELECTED;
                if (!isElementExists(By.XPath(chekboxXpath)))
                {
                    chekboxXpath = chekboxXpath.Replace("/div[8]/div/div/table/tbody/", "/div[8]/div/div/div/table/tbody/");
                }
                TotalNumber += _webDriver.FindElements(By.XPath(chekboxXpath)).Where(e => e.Selected).Count();
            }

            return selectedCount == TotalNumber;

        }
        public void ClickSelectAllButton()
        {
            if (!isElementExists(By.Id(MASSIVE_DELETE_SELECTALL_BUTTON)))
            {
                ClickUnSelectAllButton();
            }
            else
            {
                var btn = WaitForElementIsVisible(By.Id(MASSIVE_DELETE_SELECTALL_BUTTON));
                btn.Click();
                WaitForLoad();

            }
        }
        public void ClickUnSelectAllButton()
        {
            if (isElementExists(By.Id(MASSIVE_DELETE_UNSELECTALL_BUTTON)))
            {
                var btn = WaitForElementIsVisible(By.Id(MASSIVE_DELETE_UNSELECTALL_BUTTON));
                btn.Click();
                WaitForLoad();
            }
        }
        public bool VerifyPagination(int size, int totalNumber)
        {
            PageSize(size.ToString());
            var NumberPage = totalNumber / size + (totalNumber % size != 0 ? 1 : 0);
            return NumberPage == getNumberPage();
        }

        public bool VerifySearchResult(string loadingPlanName)
        {
            var loadingPlanesResult = getLodingPlanNames();
            if (loadingPlanesResult.Count() == 0 || !loadingPlanesResult.Contains(loadingPlanName))
                return false;

            return true;
        }

        public void SortByCustomerDESC()
        {
            var sortByCustomer = WaitForElementExists(By.XPath(MASSIVE_DELETE_SORT_BY_CUSTOMER));
            sortByCustomer.Click();
            WaitForLoad();
        }

        public bool VerifySortByCustomerDESC()
        {
            var customers = getloadingPlanCustomers().ToList();
            for (var i = 0; i < customers.Count - 1; i++)
            {
                if (customers[i].CompareTo(customers[i + 1]) > 0)
                {
                    return false;
                }
            }
            return true;
        }


        public IEnumerable<string> getloadingPlanCustomers()
        {
            return _webDriver.FindElements(By.XPath(MASSIVE_DELETE_CUSTOMERS)).Select(e => e.Text);
        }

        public void SelectUseCaseByValue(string value)
        {
            var comboUseCase = WaitForElementIsVisible(By.XPath(USE_CASE_SELECT));
            comboUseCase.SetValue(ControlType.DropDownList, value);
        }

        public void SetSearch(string value)
        {
            var searchText = WaitForElementToBeClickable(By.XPath("/html/body/div[3]/div/div/div[2]/div/form/div/div[1]/div/input"));
            searchText.SetValue(ControlType.TextBox, value);
            WaitForLoad();
        }

        public void SetDateTimeTo(DateTime toValue)
        {
            var to = WaitForElementIsVisible(By.Id("dateLoadingPlanEnd"));
            to.SetValue(ControlType.DateTime, toValue);
            WaitForLoad();
        }

        public bool VerifySiteSearch(string site)
        {
            var listSites = _webDriver.FindElements(By.XPath(MASSIVE_DELETE_SITE_LIST));
            return listSites.Where(s => s.Text.Equals(site)).Count() == listSites.Count();
        }

        public void SelectSites(string site)
        {
            ComboBoxSelectById((new ComboBoxOptions(SITE_FILTER, site, false)));
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
    }
}
