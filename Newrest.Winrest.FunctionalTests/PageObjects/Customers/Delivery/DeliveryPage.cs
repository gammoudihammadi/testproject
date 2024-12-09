using iText.StyledXmlParser.Jsoup.Select;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using Newrest.Winrest.FunctionalTests.PageObjects.Shared;
using Newrest.Winrest.FunctionalTests.Utils;
using OpenQA.Selenium;
using OpenQA.Selenium.Interactions;
using OpenQA.Selenium.Support.PageObjects;
using OpenQA.Selenium.Support.UI;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Linq.Expressions;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading;
using System.Threading.Tasks;

namespace Newrest.Winrest.FunctionalTests.PageObjects.Customers.Delivery
{
    public class DeliveryPage : PageBase
    {

        public DeliveryPage(IWebDriver webDriver, TestContext testContext) : base(webDriver, testContext)
        {
        }

        //___________________________________ Constantes _______________________________________________

        // General
        private const string NEW_DELIVERY = "New delivery";
        private const string EXPORT = "btn-export-excel";
        private const string MASSIVE_DELETE = "btn-massive-delete";
        private const string FOLD_BTN = "//*[@id=\"unfoldBtn\"][@class='btn unfold-all-btn unfolded']";
        private const string UNFOLD_BTN = "//*[@id=\"unfoldBtn\"][@class='btn unfold-all-btn']";

        // Tableau
        private const string FIRST_DELIVERY_NAME = "//*[@id=\"list-item-with-action\"]/div[2]/div/div/div[2]/table/tbody/tr/td[2]";
        private const string FIRST_CUSTOMER_NAME = "//*[@id=\"list-item-with-action\"]/div[2]/div[1]/div/div[2]/table/tbody/tr/td[3]";
        private const string DETAIL_DELIVERY = "//*[starts-with(@id,\"content_\")]";
        private const string FIRST_SERVICE_NAME = "//*[starts-with(@id,\"content_\")]/div/table/tbody/tr[2]/td[1]";
        private const string SERVICE_QTY_MONDAY = "//*[starts-with(@id,\"content_\")]/div/table/tbody/tr[2]/td[3]";

        // Filtres
        private const string RESET_FILTER = "//*[@id=\"flightdelivery-filter-form\"]/div[1]/a";
        private const string SEARCH = "SearchPattern";
        private const string SORTBY = "cbSortBy";
        private const string SHOW_ALL = "//*[@id=\"ShowActive\"][@value=\"All\"]";
        private const string SHOW_ACTIVE = "//*[@id=\"ShowActive\"][@value=\"ActiveOnly\"]";
        private const string SHOW_INACTIVE = "//*[@id=\"ShowActive\"][@value=\"InactiveOnly\"]";
        private const string CUSTOMER_FILTER = "SelectedCustomers_ms";
        private const string CUSTOMER_UNSELECT_ALL = "/html/body/div[10]/div/ul/li[2]/a";
        private const string SEARCH_CUSTOMER = "/html/body/div[10]/div/div/label/input";
        private const string CUSTOMER_TYPE_FILTER = "SelectedCustomerTypes_ms";
        private const string CUSTOMER_TYPE_UNSELECT_ALL = "/html/body/div[11]/div/ul/li[2]/a";
        private const string SEARCH_CUSTOMER_TYPE = "/html/body/div[11]/div/div/label/input";
        private const string SITE_FILTER = "SelectedSites_ms";
        private const string SITES_UNSELECT_ALL = "/html/body/div[12]/div/ul/li[2]/a";
        private const string SEARCH_SITE = "/html/body/div[12]/div/div/label/input";


        private const string INACTIVE = "//*[@id=\"list-item-with-action\"]/div[{0}]/div/div/div[2]/table/tbody/tr/td[1]/img";
        private const string DELIVERY_NAME = "//*[@id=\"list-item-with-action\"]/div[*]/div/div/div[2]/table/tbody/tr/td[2]";
        private const string DELIVERY_CUSTOMER = "//*[@id=\"list-item-with-action\"]/div[*]/div[1]/div/div[2]/table/tbody/tr/td[3]";
        private const string DELIVERY_SITE = "//*[@id=\"list-item-with-action\"]/div[*]/div[1]/div/div[2]/table/tbody/tr/td[4]";

        
        //____________________________________ Variables _______________________________________________

        // General
        [FindsBy(How = How.XPath, Using = NEW_DELIVERY)]
        private IWebElement _createNewDelivery;

        [FindsBy(How = How.Id, Using = EXPORT)]
        private IWebElement _export;

        [FindsBy(How = How.Id, Using = MASSIVE_DELETE)]
        private IWebElement _massiveDelete;

        [FindsBy(How = How.XPath, Using = FOLD_BTN)]
        private IWebElement _fold;

        [FindsBy(How = How.XPath, Using = UNFOLD_BTN)]
        private IWebElement _unfold;

        // Tableau
        [FindsBy(How = How.XPath, Using = FIRST_CUSTOMER_NAME)]
        private IWebElement _firstCustomerName;

        [FindsBy(How = How.XPath, Using = FIRST_DELIVERY_NAME)]
        private IWebElement _firstDeliveryName;

        [FindsBy(How = How.XPath, Using = FIRST_SERVICE_NAME)]
        private IWebElement _firstServiceName;

        [FindsBy(How = How.XPath, Using = SERVICE_QTY_MONDAY)]
        private IWebElement _serviceQtyMonday;

        //____________________________________ Filtres __________________________________________________

        [FindsBy(How = How.XPath, Using = RESET_FILTER)]
        private IWebElement _resetFilter;

        [FindsBy(How = How.Id, Using = SEARCH)]
        private IWebElement _searchFilter;

        [FindsBy(How = How.Id, Using = SORTBY)]
        private IWebElement _sortBy;

        [FindsBy(How = How.XPath, Using = SHOW_ALL)]
        private IWebElement _showAllBtn;

        [FindsBy(How = How.XPath, Using = SHOW_ACTIVE)]
        private IWebElement _showActiveBtn;

        [FindsBy(How = How.XPath, Using = SHOW_INACTIVE)]
        private IWebElement _showInactiveBtn;

        [FindsBy(How = How.XPath, Using = CUSTOMER_FILTER)]
        private IWebElement _customerFilter;

        [FindsBy(How = How.XPath, Using = CUSTOMER_UNSELECT_ALL)]
        private IWebElement _customerUnselectAll;

        [FindsBy(How = How.XPath, Using = SEARCH_CUSTOMER)]
        private IWebElement _searchCustomer;

        [FindsBy(How = How.Id, Using = CUSTOMER_TYPE_FILTER)]
        private IWebElement _customerTypeFilter;

        [FindsBy(How = How.XPath, Using = CUSTOMER_TYPE_UNSELECT_ALL)]
        private IWebElement _customerTypeUnselectAll;

        [FindsBy(How = How.XPath, Using = SEARCH_CUSTOMER_TYPE)]
        private IWebElement _searchCustomerType;

        [FindsBy(How = How.Id, Using = SITE_FILTER)]
        private IWebElement _siteFilter;

        [FindsBy(How = How.XPath, Using = SITES_UNSELECT_ALL)]
        private IWebElement _siteUnselectAll;

        [FindsBy(How = How.XPath, Using = SEARCH_SITE)]
        private IWebElement _searchSite;

        public enum FilterType
        {
            Search,
            SortBy,
            ShowAll,
            ShowActive,
            ShowInactive,
            Customers,
            CustomerTypes,
            Sites
        }

        public void Filter(FilterType filterType, object value)
        {
            switch (filterType)
            {
                case FilterType.Search:
                    _searchFilter = WaitForElementIsVisible(By.Id(SEARCH));
                    _searchFilter.SetValue(ControlType.TextBox, value);
                    break;

                case FilterType.SortBy:
                    _sortBy = WaitForElementIsVisible(By.Id(SORTBY));
                    _sortBy.Click();
                    var element = WaitForElementIsVisible(By.XPath("//option[contains(@value,'" + value + "')]"));
                    _sortBy.SetValue(ControlType.DropDownList, element.Text);
                    _sortBy.Click();
                    break;

                case FilterType.ShowAll:
                    _showAllBtn = WaitForElementExists(By.XPath(SHOW_ALL));
                    _showAllBtn.SetValue(ControlType.RadioButton, value);
                    break;

                case FilterType.ShowActive:
                    _showActiveBtn = WaitForElementExists(By.XPath(SHOW_ACTIVE));
                    _showActiveBtn.SetValue(ControlType.RadioButton, value);
                    break;

                case FilterType.ShowInactive:
                    _showInactiveBtn = WaitForElementExists(By.XPath(SHOW_INACTIVE));
                    _showInactiveBtn.SetValue(ControlType.RadioButton, value);
                    break;

                case FilterType.Customers:
                    _customerFilter = WaitForElementIsVisible(By.Id(CUSTOMER_FILTER));
                    _customerFilter.Click();

                    _customerUnselectAll = WaitForElementIsVisible(By.XPath(CUSTOMER_UNSELECT_ALL));
                    _customerUnselectAll.Click();

                    _searchCustomer = WaitForElementIsVisible(By.XPath(SEARCH_CUSTOMER));
                    _searchCustomer.SetValue(ControlType.TextBox, value);
                    Thread.Sleep(2000);

                    var valueToCheckCustomers = WaitForElementExists(By.XPath("/html/body/div[10]/ul"));
                    valueToCheckCustomers.Click();

                    _customerFilter.Click();
                    break;

                case FilterType.CustomerTypes:
                    _customerTypeFilter = WaitForElementIsVisible(By.Id(CUSTOMER_TYPE_FILTER));
                    _customerTypeFilter.Click();

                    _customerTypeUnselectAll = WaitForElementIsVisible(By.XPath(CUSTOMER_TYPE_UNSELECT_ALL));
                    _customerTypeUnselectAll.Click();

                    _searchCustomerType = WaitForElementIsVisible(By.XPath(SEARCH_CUSTOMER_TYPE));
                    _searchCustomerType.SetValue(ControlType.TextBox, value);

                    var valueToCheckCustomersType = WaitForElementIsVisible(By.XPath("//span[text()='" + value + "']"));
                    valueToCheckCustomersType.SetValue(ControlType.CheckBox, true);

                    _customerTypeFilter.Click();
                    break;

                case FilterType.Sites:
                    _siteFilter = WaitForElementIsVisible(By.Id(SITE_FILTER));
                    _siteFilter.Click();

                    _siteUnselectAll = WaitForElementIsVisible(By.XPath(SITES_UNSELECT_ALL));
                    _siteUnselectAll.Click();

                    var siteSelected = value + " - " + value;
                    _searchSite = WaitForElementIsVisible(By.XPath(SEARCH_SITE));
                    _searchSite.SetValue(ControlType.TextBox, siteSelected);
                    WaitForLoad();

                    IWebElement siteToCheck;
                    siteToCheck = WaitForElementIsVisible(By.XPath("//span[text()='" + siteSelected + "']"));
                    siteToCheck.SetValue(ControlType.CheckBox, true);

                    _siteFilter.Click();
                    break;

                default:
                    break;
            }

            WaitPageLoading();
            Thread.Sleep(2500);
        }

        public void ResetFilter()
        {
            _resetFilter = WaitForElementIsVisible(By.XPath(RESET_FILTER));
            _resetFilter.Click();
            WaitForLoad();
            // batch passage à minuit-2h UTC Renaud
            if (DateUtils.IsBeforeMidnight())
            {
                // pas de date
            }
        }

        public bool CheckStatus(bool active)
        {
            bool isActive = false;
            int tot = CheckTotalNumber() > 100 ? 100 : CheckTotalNumber();

            if (tot == 0)
                return false;

            for (int i = 1; i <= tot; i++)
            {
                if (isElementVisible(By.XPath(String.Format(INACTIVE, i + 1))))
                {
                    _webDriver.FindElement(By.XPath(String.Format(INACTIVE, i + 1)));

                    if (active)
                        return false;
                }
                else
                {
                    isActive = true;
                    if (!active)
                        return true;
                }
            }
            return isActive;
        }

        public bool IsSortedByName()
        {
            var listeNames = _webDriver.FindElements(By.XPath(DELIVERY_NAME));

            if (listeNames.Count == 0)
                return false;

            var ancientName = listeNames[0].Text;

            foreach (var elm in listeNames)
            {
                try
                {
                    if (elm.Text.CompareTo(ancientName) < 0)
                        return false;

                    ancientName = elm.Text;
                }
                catch
                {
                    return false;
                }
            }

            return true;
        }

        public bool IsSortedByCustomer()
        {
            var listeCustomers = _webDriver.FindElements(By.XPath(DELIVERY_CUSTOMER));

            if (listeCustomers.Count == 0)
                return false;

            var ancientCustomer = listeCustomers[0].Text;

            foreach (var elm in listeCustomers)
            {
                try
                {
                    if (elm.Text.CompareTo(ancientCustomer) < 0)
                        return false;

                    ancientCustomer = elm.Text;
                }
                catch
                {
                    return false;
                }
            }

            return true;
        }

        public bool VerifyCustomer(string customer)
        {
            var listeCustomers = _webDriver.FindElements(By.XPath(DELIVERY_CUSTOMER));

            if (listeCustomers.Count == 0)
                return false;

            foreach (var elm in listeCustomers)
            {
                    if (!elm.Text.Contains(customer))
                        return false; 
            }

            return true;
        }

        public bool VerifySite(string site)
        {
            var listeSites = _webDriver.FindElements(By.XPath(DELIVERY_SITE));

            if (listeSites.Count == 0)
                return false;

            foreach (var elm in listeSites)
            {
                if (!elm.Text.Contains(site))
                    return false;
            }

            return true;
        }

        //____________________________________  Méthodes __________________________________________

        // General
        public DeliveryCreateModalPage DeliveryCreatePage()
        {
            ShowPlusMenu();
            WaitForLoad();
            Actions action = new Actions(_webDriver);
            _createNewDelivery = WaitForElementIsVisible(By.LinkText(NEW_DELIVERY));
            action.MoveToElement(_createNewDelivery).Perform();
            _createNewDelivery.Click();
            WaitForLoad();

            return new DeliveryCreateModalPage(_webDriver, _testContext);
        }

        public void ExportDelivery(bool versionPrint)
        {
            ShowExtendedMenu();

            _export = WaitForElementIsVisible(By.Id(EXPORT));
            _export.Click();
            WaitForLoad();

            if (versionPrint)
            {
                IsFileLoaded(By.CssSelector("[target='_blank'][class='far fa-file-excel']"));
                ClickPrintButton();
            }

            WaitForDownload();
            Close();
        }

        public FileInfo GetExportExcelFile(FileInfo[] taskFiles, string deliveryName = null)
        {
            List<FileInfo> correctDownloadFiles = new List<FileInfo>();

            foreach (var file in taskFiles)
            {
                if (IsExcelFileCorrect(file.Name, deliveryName))
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

        public bool IsExcelFileCorrect(string filePath, string deliveryName)
        {
            Regex reg;

            if (deliveryName != null)
            {
                reg = new Regex("^export-flightdelivery" + "-" + deliveryName + ".xlsx", RegexOptions.IgnoreCase | RegexOptions.Singleline);
            }
            else
            {
                reg = new Regex("^export-flightdelivery.xlsx", RegexOptions.IgnoreCase | RegexOptions.Singleline);
            }

            Match m = reg.Match(filePath);

            return m.Success;
        }

        public DeliveryMassiveDeletePage MassiveDelete()
        {
            ShowExtendedMenu();
           
            _massiveDelete = WaitForElementIsVisible(By.Id(MASSIVE_DELETE));
            _massiveDelete.Click();
            WaitForLoad();
            WaitPageLoading();
            return new DeliveryMassiveDeletePage(_webDriver, _testContext);
        }

        public void FoldAll()
        {
            ShowExtendedMenu();

            _fold = WaitForElementIsVisible(By.XPath(FOLD_BTN));
            _fold.Click();
            WaitForLoad();
        }

        public bool IsFoldAll()
        {
            var detailDelivery = WaitForElementExists(By.XPath(DETAIL_DELIVERY));

            // Temps nécessaire pour que l'élément change de classe
            WaitPageLoading();

            return detailDelivery.GetAttribute("class") == "panel-collapse collapse";
        }

        public void UnfoldAll()
        {
            ShowExtendedMenu();

            _unfold = WaitForElementIsVisible(By.XPath(UNFOLD_BTN));
            _unfold.Click();
            WaitForLoad();
        }

        public bool IsUnfoldAll()
        {
            var detailDelivery = WaitForElementExists(By.XPath(DETAIL_DELIVERY));

            // Temps nécessaire pour que l'élément change de classe
            WaitPageLoading();
            return detailDelivery.GetAttribute("class") == "panel-collapse collapse show";
        }

        // Tableau
        public DeliveryLoadingPage ClickOnFirstDelivery()
        {
            WaitForLoad();
            _firstDeliveryName = WaitForElementIsVisible(By.XPath(FIRST_DELIVERY_NAME));
            _firstDeliveryName.Click();
            WaitForLoad();

            return new DeliveryLoadingPage(_webDriver, _testContext);
        }

        public string GetFirstDeliveryName()
        {
            _firstDeliveryName = WaitForElementIsVisible(By.XPath(FIRST_DELIVERY_NAME));
            return _firstDeliveryName.Text;
        }


        public string GetFirstCustomerName()
        {
            _firstCustomerName = WaitForElementIsVisible(By.XPath(FIRST_CUSTOMER_NAME));
            return _firstCustomerName.GetAttribute("innerText");
        }

        public string GetServiceQuantityMonday()
        {
            _serviceQtyMonday = WaitForElementIsVisible(By.XPath(SERVICE_QTY_MONDAY));
            return _serviceQtyMonday.Text;
        }

        public string GetFirstServiceName()
        {
            _firstServiceName = WaitForElementIsVisible(By.XPath(FIRST_SERVICE_NAME));
            return _firstServiceName.Text;
        }
        public string GetSearchFilterText()
        {
            var textSearchFilter = _webDriver.FindElement(By.Id(SEARCH));
            return textSearchFilter.GetAttribute("value");
        }
        public string GetNameFilterText()
        {
            var nameFilter = _webDriver.FindElement(By.Id(SORTBY));
            return nameFilter.GetAttribute("value");
        }
        public bool GetactiveOnlyBool()
        {
            var activeBool = _webDriver.FindElement(By.XPath(SHOW_ACTIVE));
            return activeBool.Selected;
        }
        public string GetcustomerNumber()
        {
            var customerFilter = _webDriver.FindElement(By.Id(CUSTOMER_FILTER));
            return customerFilter.GetAttribute("innerText");

        }
        public string GetcustomertypeNumber()
        {
            var customerTypeFilter = _webDriver.FindElement(By.Id(CUSTOMER_TYPE_FILTER));
            return customerTypeFilter.GetAttribute("innerText");
        }
        public string GetSiteSelectedNumber()
        {
            var siteFilter = _webDriver.FindElement(By.Id(SITE_FILTER));
            return siteFilter.GetAttribute("innerText");
        }

        public bool DeliveryNameExist(string name)
        {
            var listeNames = _webDriver.FindElements(By.XPath(DELIVERY_NAME));
            if (listeNames.Count == 0)
                return false;

            foreach (var elm in listeNames)
            {
                if (elm.Text == name)
                    return true;
            }
            return false;
        }
    }
}
