using Microsoft.VisualStudio.TestTools.UnitTesting;
using Newrest.Winrest.FunctionalTests.PageObjects.Shared;
using Newrest.Winrest.FunctionalTests.Utils;
using OpenQA.Selenium;
using OpenQA.Selenium.Interactions;
using OpenQA.Selenium.Support.PageObjects;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text.RegularExpressions;
using System.Text;
using System.Threading;
using System;
using System.Security.Cryptography;

namespace Newrest.Winrest.FunctionalTests.PageObjects.Customers.FoodPackets
{
    public class FoodPacketPage : PageBase
    {
        public FoodPacketPage(IWebDriver webDriver, TestContext testContext) : base(webDriver, testContext)
        {
        }

        // __________________________________________ Constantes _________________________________________
        public const string SEARCH_FILTER = "SearchByPacketName";
        public const string CUSTOMER_SELECT = "/html/body/div[2]/div/div[1]/div[1]/div/form/div[5]/div/div/div[1]/button";
        public const string FILTER_CUSTOMERS = "SelectedCustomers_ms";
        public const string FILTER_SITES = "SelectedSites_ms";
        public const string NUMBER_FOOD_PACKET = "/html/body/div[2]/div/div[2]/div/div[1]/h1/span";
        public const string FIRST_FOOD_PACKET_NAME = "/html/body/div[2]/div/div[2]/div[2]/div/div/div[2]/div/div/div[2]/table/tbody/tr/td[2]";
        public const string LIST_FOOD_PACKET_CUSTOMER_GRID = "/html/body/div[2]/div/div[2]/div[2]/div/div/div[*]/div/div/div[2]/table/tbody/tr/td[3]";
        public const string LIST_FOOD_PACKET_SITE_GRID = "/html/body/div[2]/div/div[2]/div[2]/div/div/div[*]/div/div/div[2]/table/tbody/tr/td[4]";
        public const string RESET_FILTER = "/html/body/div[2]/div/div[1]/div[1]/div/form/div[1]/a";
        public const string SHOW_ONLY_SERVICES_WITH_PRICES = "/html/body/div[2]/div/div[1]/div[1]/div/form/div[7]/input";
        public const string EXTEND_MENU_EXPORT = "/html/body/div[2]/div/div[2]/div/div[1]/div/div/button";
        public const string EXPORT_BUTTON = "/html/body/div[2]/div/div[2]/div/div[1]/div/div/div/a";
        public const string FOOD_PACKET_NAME_GRID = "/html/body/div[2]/div/div[2]/div/div[2]/div/div/div/div[*]/div[1]/div/div[2]/table/tbody/tr[*]/td";
        public const string LIST_CUSTOMER_FILTER = "/html/body/div[11]/ul/li[*]/label/input";
        public const string LIST_SITE_FILTER = "/html/body/div[10]/ul/li[*]/label/input";
        public const string FIRST_FOOD_PACKET_FOLD_UNFOLD = "//*/div[@data-toggle='collapse']";
        public const string FIRST_FOOD_PACKET_ERROR_MESSAGE = "//*/span[@class='no-details-found']";
        public const string CHILDREN_FOOD_PACKETS = "children-food-packets";
        public const string HEADER_PACKETS = "//*[@id=\"tabContentItemContainer\"]/div[1]/h1";


        // tableau


        // __________________________________________ Variables __________________________________________

        [FindsBy(How = How.Id, Using = SEARCH_FILTER)]
        private IWebElement _searchFilter;
        [FindsBy(How = How.Id, Using = CHILDREN_FOOD_PACKETS)]
        private IWebElement _childrenpackets;
        [FindsBy(How = How.XPath, Using = HEADER_PACKETS)]
        private IWebElement _headerPacket;


        // __________________________________________ Méthodes ___________________________________________
        public enum FilterType
        {
            Search,
            Site,
            ShowOnlyServicesWithPrices,
            Customers,
            ChildrenFoodPackets
        }

        public void Filter(FilterType filterType, object value)

        {
            switch (filterType)
            {
                case FilterType.Search:

                    _searchFilter = WaitForElementIsVisible(By.Id(SEARCH_FILTER));
                    _searchFilter.SetValue(ControlType.TextBox, value);
                    break;

                case FilterType.Customers:
                    ComboBoxSelectById(new ComboBoxOptions(FILTER_CUSTOMERS, (string)value));
                    break;

                case FilterType.Site:
                    ComboBoxSelectById(new ComboBoxOptions(FILTER_SITES, (string)value));
                    break;

                case FilterType.ShowOnlyServicesWithPrices:
                    var showOnlyServicesWithPrices = WaitForElementIsVisible(By.XPath(SHOW_ONLY_SERVICES_WITH_PRICES));
                    showOnlyServicesWithPrices.SetValue(ControlType.CheckBox, value);
                    break;

                case FilterType.ChildrenFoodPackets:
                    WaitForElementIsVisible(By.Id(CHILDREN_FOOD_PACKETS));
                    _childrenpackets.SetValue(ControlType.CheckBox, value);
                    break;
                default:
                    break;

            }
            WaitPageLoading();
            Thread.Sleep(2000);
        }

        public string GetNumberOfFoodPacketHeader()
        {
            var number = WaitForElementIsVisible(By.XPath(NUMBER_FOOD_PACKET));
            return number.Text;
        }
        public void ResetFilters()
        {
             var resetFilter = WaitForElementIsVisible(By.XPath(RESET_FILTER));
            resetFilter.Click();

            WaitForLoad();
            // batch passage à minuit-2h UTC Renaud
            if (DateUtils.IsBeforeMidnight())
            {
                //pas de date
            }
        }
        public int GetNumberSelectedSiteFilter()
        {
            var listSitesSelectedFilters = _webDriver.FindElements(By.XPath(LIST_SITE_FILTER));
            var numberSitesSelectedSite = listSitesSelectedFilters
               .Where(p => p.Selected == true).Count();

            return numberSitesSelectedSite;
        }

        public int GetNumberSelectedCustomerFilter()
        {
            var listSelectedCustomerFilters = _webDriver.FindElements(By.XPath(LIST_CUSTOMER_FILTER));
            var numberSelectedCustomer = listSelectedCustomerFilters.Where(p => p.Selected).Count();

            return numberSelectedCustomer;
        }
        public object GetFilterValue(FilterType filterType)
        {
            switch (filterType)
            {
                case FilterType.Search:
                    _searchFilter = WaitForElementIsVisible(By.Id(SEARCH_FILTER));
                    return _searchFilter.GetAttribute("value");


                case FilterType.ShowOnlyServicesWithPrices:
                    var _showOnlyServicesWithPrices = WaitForElementIsVisible(By.XPath(SHOW_ONLY_SERVICES_WITH_PRICES));
                    return _showOnlyServicesWithPrices.Selected;
            }
            return null;
        }
        public void Export()
        {
            var extendMenu = WaitForElementIsVisible(By.XPath(EXTEND_MENU_EXPORT));
            Actions actions = new Actions(_webDriver);
            actions.MoveToElement(extendMenu).Perform();
            var exportBtn = WaitForElementIsVisible(By.XPath(EXPORT_BUTTON));
            WaitForLoad();
            exportBtn.Click();

            IsFileLoaded(By.CssSelector("[target='_blank'][class='far fa-file-excel']"));
            ClickPrintButton();
            WaitForLoad();
            Thread.Sleep(3000);
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
            Regex r = new Regex("[export?flights\\s\\d.-]", RegexOptions.IgnoreCase | RegexOptions.Singleline);
            Match m = r.Match(filePath);
            return m.Success;
        }
        public bool VerifyExcel(List<string> namesGrid, List<string> namesExcel)
        {
            Assert.AreEqual(namesGrid.Count, namesExcel.Count);
            for (int i = 0; i < namesGrid.Count(); i++)
            {
                if (namesGrid[i] != namesExcel[i])
                    return false;
            }
            return true;
        }
        public List<string> GetListFoodPacketName()
        {
            var ids = _webDriver.FindElements(By.XPath(FOOD_PACKET_NAME_GRID));
            return ids.Select(e => e.Text).ToList();
        }

        public void FoldOrUnfold()
        {
            var foldUnfold = WaitForElementIsVisible(By.XPath(FIRST_FOOD_PACKET_FOLD_UNFOLD));
            foldUnfold.Click();
            Thread.Sleep(2000);
            WaitForLoad();
        }

        public bool IsServiceError()
        {
            if (isElementVisible(By.XPath(FIRST_FOOD_PACKET_ERROR_MESSAGE)))
            {
                var elt = WaitForElementIsVisible(By.XPath(FIRST_FOOD_PACKET_ERROR_MESSAGE));
                return elt.Text == "No services could be retrieved. The packet does not have a priced service.";
            }
            return false;
        }
        public List<String> GetFilteredFoodPacketResults()
        {
            var results = new List<String>();   
            var elt = WaitForElementIsVisible(By.XPath(FIRST_FOOD_PACKET_ERROR_MESSAGE));
            return results;
        }
        public bool IsPageLoaded()
        {
            if (isElementVisible(By.XPath(HEADER_PACKETS)))
            {
                _headerPacket = _webDriver.FindElement(By.XPath(HEADER_PACKETS));
                return true;
            }
            else
            {
                return false;
            }
        }
    }
}

