using Microsoft.VisualStudio.TestTools.UnitTesting;
using Newrest.Winrest.FunctionalTests.PageObjects.Menus.Datasheet;
using Newrest.Winrest.FunctionalTests.PageObjects.Shared;
using Newrest.Winrest.FunctionalTests.Utils;
using OpenQA.Selenium;
using OpenQA.Selenium.Interactions;
using OpenQA.Selenium.Support.PageObjects;
using OpenQA.Selenium.Support.UI;
using System;
using System.Collections.ObjectModel;
using System.Data;
using System.Linq;
using System.Security.Policy;
using System.Text.RegularExpressions;
using System.Threading;

namespace Newrest.Winrest.FunctionalTests.PageObjects.Customers.Service
{
    public class ServicePricePage : PageBase
    {
        public ServicePricePage(IWebDriver _webDriver, TestContext _testContext) : base(_webDriver, _testContext)
        {
        }

        //____________________________________ Constantes______________________________________________

        // General
        private const string BACK_TO_LIST = "/html/body/div[2]/a/span[2]";
        private const string NEW_PRICE = "btn-add-item-detail";
        private const string FOLD_ALL_UNFOLD_ALL_BUTTON_STATE = "//*[@id=\"list-item-with-action\"]/div[1]/a";
        private const string FOLD_ALL = "//*[@id=\"list-item-with-action\"]/div[1]/a[@class='btn unfold-all-btn unfolded']";
        private const string UNFOLD_ALL = "//*[@id=\"list-item-with-action\"]/div[1]/a[@class='btn unfold-all-btn']";

        // Onglets       
        private const string GENERAL_INFORMATION_TAB = "hrefTabContentService";
        private const string MENU_TAB = "hrefTabContentMenus";
        private const string DELIVERY_TAB = "hrefTabContentDeliveries";
        private const string LOADINGPLAN_TAB = "hrefTabContentLoadingPlans";
        private const string FOOD_PACKETS_TAB = "hrefTabContentFoodPackets";

        // Tableau prices
        private const string FIRST_CONTENT = "//*[starts-with(@id,\"content_\")]";
        private const string LAST_CONTENT = "//*[@id=\"content_7-927\"]/div/table";
        //private const string CONTENT = "//td[contains(text(),'{0}') and contains(text(),'{1}')]/../../../../../../../div[2]/div[1]/table/tbody/tr[1]/td[*]/a[@id='btn-edit-price']";
        private const string CONTENT = "//td[contains(text(),'{0}') and contains(text(),'{1}')]/../../../../../../../div[2]//a[@id='btn-edit-price']";
        private const string PRICE = "//*[@id=\"list-item-with-action\"]/div[*]/div[1]/div/div[2]/table/tbody/tr[1]/td[contains(text(),'{0}') and contains(text(),'{1}')]";
        private const string ARROW_PRICE = "//*[@id=\"list-item-with-action\"]/div[*]/div[1]/div/div[2]/table/tbody/tr[1]/td[contains(text(),'{0}') and contains(text(),'{1}')]/../../../../../div[1]";
        private const string EDIT_PRICE = "//*[@id=\"list-item-with-action\"]/div[*]/div[1]/div/div[2]/table/tbody/tr[1]/td[contains(text(),'{0}') and contains(text(),'{1}')]/../../../../../../../div[2]/div[1]/table/tbody/tr[1]/td[*]/a[starts-with(@id, 'btn-edit-price')]";
        private const string LINE_PRICE = "/html/body/div[3]/div/div/div[2]/div/div/div[2]/div[2]/div/table/tbody/tr";

        private const string TOGGLE_FIRST_PRICE = "//*[@id=\"list-item-with-action\"]/div[2]/div[1]/div/div[1]";
        private const string TOGGLE_PRICE = "//td[contains(text(), '{0}') and contains(text(), '{1}')]/../../../../../div[contains(@class, 'col-md-1 toggle-collapse-btn-left')]";

        private const string DUPLICATE_ALL_PRICES = "btn-duplicate-price";
        private const string DESTINATION_SITE_FILTER = "SelectedDestinationSitesIds_ms";
        private const string UNCHECK_ALL_SITES = "/html/body/div[12]/div/ul/li[2]/a";
        private const string SEARCH_SITE = "/html/body/div[12]/div/div/label/input";
        private const string CUSTOMER_DESTINATION = "dropdown-customer-duplicate-service-selectized";
        private const string FIRST_CUSTOMER_DESTINATION = "//*[@id=\"modal-1\"]/div/div/div/div/form/div[2]/div[1]/div[2]/div[2]/div/div/div[2]";
        private const string DUPLICATE_BTN = "last";

        private const string PRICE_DELETE_BTN = "btn-delete-price";
        private const string LAST_PRICE_DELETE_BTN = "//*[@id=\"customers-service-deleteprice-3-1\"]";
        private const string PRICE_DELETE_CONFIRM_BTN = "dataConfirmOK";

        private const string PRICE_NAME = "//*[@id=\"list-item-with-action\"]/div[2]/div[1]/div/div[2]/table/tbody/tr/td[1]";
        private const string PRICE_DATE_FROM = "//*[starts-with(@id,\"content_\")]/div/table/tbody/tr/td[1]";
        private const string PRICE_DATE_TO = "//*[starts-with(@id,\"content_\")]/div/table/tbody/tr/td[2]";
        private const string PRICE_ITEM = "//*[starts-with(@id,\"content_\")]/div/table/tbody/tr/td[7]";
        private const string DATASHEET_ITEM = "//*[starts-with(@id,\"content_\")]/div/table/tbody/tr/td[8]";
        private const string ITEMS = "//*[starts-with(@id,\"content_\")]/div/table/tbody/tr[*]";
        private const string CUSTOMER_SRV = "//*[starts-with(@id,\"content_\")]/div/table/tbody/tr/td[3]";
        private const string DUPLICATE_BUTTON = "customers-service-duplicateprice-1";
        private const string SERVICE_PRICE = "//*[@id=\"ServiceTabNav\"]";
        private const string EDIT_FIRST_ITEM = "customers-datasheet-edit-1-1";
        private const string FIRST_ITEM = "/html/body/div[3]/div/div/div[2]/div/div/div[2]/div[1]/div/div[2]/table/tbody/tr/td[1]";
        private const string FIRST_PRICE_ITEM = "//*[@class=\"service-table\"]/tbody/tr[1]";
        private const string DATASHEET_LIST = "//*[@id=\"modal-1\"]/div/div/form/div[2]/div[3]/div/div/div[3]/div/span[1]/input[2]";
        private const string SAVE_PRICE_BTN = "//*[@value=\"Create\"]";
        private const string DATASHEET_MENU = "//*[@id=\"modal-1\"]/div/div/form/div[2]/div[3]/div/div/div[3]/div/span[1]/div";
        private const string LAST_PRICE_NAME = "//*[@id=\"list-item-with-action\"]/div[2]/div[1]/div/div[2]/table/tbody/tr/td[1]";
        private const string TOGGLE_LAST_PRICE = "//*[@id=\"list-item-with-action\"]/div[3]/div[1]/div/div[1]";
        private const string PICTO_DATASHEET = "customers-datasheet-edit-1-1";
        private const string EDIT_PRICE_LINE = "//*[contains(@id,\"content\")]/div/table/tbody/tr[*]";
        private const string EDITCYCLEMODE = "//*[@id=\"item_Method\"]/option[3]";
        private const string PRICE_CYLE = "Price";
        private const string CALENDER_BTN = "//*[@id=\"customers-service-cyclecalendar-1-1\"]/span";
        private const string PRICE_TOGGLE = "//*[@id=\"list-item-with-action\"]/div[4]/div[1]/div/div[1]";
        private const string SAVEBTN_CUST_PRICE = "last";
        private const string CALENDER_BTN_OPEN = "customers-service-cyclecalendar-1-1";
        private const string PICTO_DATASHEET_PATCH = "customers-datasheet-edit-1-1";
        private const string CALENDER_BTN_PATCH = "//*[@id='Service-CycleCalendar-1-1']/span";
        private const string EDIT_FIRST_ITEM_PATCH = "customers-datasheet-edit-1-1"; // "//*[@id=\"content_9-5061\"]/div/table/tbody/tr";
        private const string DELETE_PRICE = "//*[@id=\"list-item-with-action\"]/div[2]/div[1]/div/div[1]";
        private const string EDIT_PRICE_SERVICE = "//*[@id=\"btn-edit-price\"]/span";
        //__________________________________ Variables_________________________________________________

        // General
        [FindsBy(How = How.Id, Using = DUPLICATE_BUTTON)]
        private IWebElement _duplicateButton;
        [FindsBy(How = How.XPath, Using = SERVICE_PRICE)]
        private IWebElement _serviceprice;

        [FindsBy(How = How.XPath, Using = BACK_TO_LIST)]
        private IWebElement _backToList;

        [FindsBy(How = How.Id, Using = NEW_PRICE)]
        private IWebElement _addNewPrice;

        [FindsBy(How = How.XPath, Using = UNFOLD_ALL)]
        private IWebElement _unfoldAll;

        [FindsBy(How = How.XPath, Using = FOLD_ALL)]
        private IWebElement _foldAll;


        // Onglets
        [FindsBy(How = How.Id, Using = GENERAL_INFORMATION_TAB)]
        private IWebElement _generalInformationTab;

        [FindsBy(How = How.Id, Using = MENU_TAB)]
        private IWebElement _menusTab;

        [FindsBy(How = How.Id, Using = DELIVERY_TAB)]
        private IWebElement _deliveryTab;

        [FindsBy(How = How.Id, Using = LOADINGPLAN_TAB)]
        private IWebElement _loadingPlanTab;

        [FindsBy(How = How.Id, Using = FOOD_PACKETS_TAB)]
        private IWebElement _foodpacketstab;

        // Tableau price
        [FindsBy(How = How.XPath, Using = FIRST_CONTENT)]
        private IWebElement _content;

        [FindsBy(How = How.XPath, Using = LAST_CONTENT)]
        private IWebElement _lastcontent;

        [FindsBy(How = How.ClassName, Using = DUPLICATE_ALL_PRICES)]
        private IWebElement _duplicateAllPrices;

        [FindsBy(How = How.ClassName, Using = DESTINATION_SITE_FILTER)]
        private IWebElement _destinationSiteFilter;

        [FindsBy(How = How.ClassName, Using = UNCHECK_ALL_SITES)]
        private IWebElement _unCheckAllSites;

        [FindsBy(How = How.XPath, Using = SEARCH_SITE)]
        private IWebElement _searchSite;

        [FindsBy(How = How.Id, Using = CUSTOMER_DESTINATION)]
        private IWebElement _destinationCustomer;

        [FindsBy(How = How.XPath, Using = FIRST_CUSTOMER_DESTINATION)]
        private IWebElement _firstCustomerDestination;

        [FindsBy(How = How.Id, Using = DUPLICATE_BTN)]
        private IWebElement _duplicateBtn;

        [FindsBy(How = How.ClassName, Using = PRICE_DELETE_BTN)]
        private IWebElement _deletePriceBtn;

        [FindsBy(How = How.XPath, Using = LAST_PRICE_DELETE_BTN)]
        private IWebElement _deleteLastPriceBtn;

        [FindsBy(How = How.Id, Using = PRICE_DELETE_CONFIRM_BTN)]
        private IWebElement _deletePriceConfirmBtn;

        [FindsBy(How = How.XPath, Using = PRICE_NAME)]
        private IWebElement _priceName;

        [FindsBy(How = How.XPath, Using = DATASHEET_ITEM)]
        private IWebElement _dataSheet;

        [FindsBy(How = How.XPath, Using = PRICE_DATE_FROM)]
        private IWebElement _priceDateFrom;

        [FindsBy(How = How.XPath, Using = PRICE_DATE_TO)]
        private IWebElement _priceDateTo;

        [FindsBy(How = How.XPath, Using = CUSTOMER_SRV)]
        private IWebElement _customerSrv;

        [FindsBy(How = How.XPath, Using = TOGGLE_FIRST_PRICE)]
        private IWebElement _toggleFirstPrice;

        [FindsBy(How = How.Id, Using = EDIT_FIRST_ITEM)]
        private IWebElement _editFirstItem;

        [FindsBy(How = How.XPath, Using = FIRST_ITEM)]
        private IWebElement _firstItem;

        [FindsBy(How = How.XPath, Using = FIRST_PRICE_ITEM)]
        private IWebElement _firstPriceItem;

        [FindsBy(How = How.Id, Using = PICTO_DATASHEET)]
        private IWebElement _pictoDataSheet;

        [FindsBy(How = How.XPath, Using = DATASHEET_LIST)]
        private IWebElement _datasheetList;

        [FindsBy(How = How.XPath, Using = SAVE_PRICE_BTN)]
        private IWebElement _savePriceBtn;

        [FindsBy(How = How.XPath, Using = DATASHEET_MENU)]
        private IWebElement _datasheetMenu;

        [FindsBy(How = How.XPath, Using = LAST_PRICE_NAME)]
        private IWebElement _lastpriceName;

        [FindsBy(How = How.XPath, Using = TOGGLE_LAST_PRICE)]
        private IWebElement _toggleLastPrice;

        [FindsBy(How = How.Id, Using = PRICE_CYLE)]
        private IWebElement _priceCycle;

        [FindsBy(How = How.XPath, Using = CALENDER_BTN)]
        private IWebElement _calenderbtn;

        [FindsBy(How = How.XPath, Using = PRICE_TOGGLE)]
        private IWebElement _PriceToggle;

        [FindsBy(How = How.Id, Using = SAVEBTN_CUST_PRICE)]
        private IWebElement _create;

        [FindsBy(How = How.Id, Using = CALENDER_BTN_OPEN)]
        private IWebElement _calenderbtnOpen;

        [FindsBy(How = How.Id, Using = PICTO_DATASHEET_PATCH)]
        private IWebElement _pictoDataSheetPatch;

        [FindsBy(How = How.XPath, Using = CALENDER_BTN_PATCH)]
        private IWebElement _calenderbtnPatch;

        [FindsBy(How = How.Id, Using = EDIT_FIRST_ITEM_PATCH)]
        private IWebElement _editFirstItemPatch;

        [FindsBy(How = How.XPath, Using = EDIT_PRICE_SERVICE)]
        private IWebElement _editFirstService;
        //__________________________________ Méthodes _______________________________________

        // General
        public ServicePage BackToList()
        {
            WaitForLoad();
            WaitPageLoading();
            _backToList = WaitForElementIsVisibleNew(By.XPath(BACK_TO_LIST));
            _backToList.Click();
            WaitForLoad();

            return new ServicePage(_webDriver, _testContext);
        }

        public ServiceCreatePriceModalPage AddNewCustomerPrice()
        {
            WaitPageLoading();

            _addNewPrice = WaitForElementIsVisibleNew(By.Id(NEW_PRICE));
            _addNewPrice.Click();
            WaitPageLoading();
            WaitForLoad();
            return new ServiceCreatePriceModalPage(_webDriver, _testContext);
        }

        public void FoldAll()
        {
            _foldAll = WaitForElementIsVisible(By.XPath(FOLD_ALL));
            _foldAll.Click();
            WaitForLoad();
        }

        public Boolean IsFoldAll()
        {
            _content = WaitForElementExists(By.XPath(FIRST_CONTENT));

            // Temps nécessaire pour que l'élément change de classe
            WaitPageLoading();

            return _content.GetAttribute("class") == "panel-collapse collapse";
        }

        public void UnfoldAll()
        {
            WaitForLoad();
            var foldOrUnfold = _webDriver.FindElement(By.XPath(FOLD_ALL_UNFOLD_ALL_BUTTON_STATE));
            if (foldOrUnfold.Text != "Fold All")
            {
                _unfoldAll = WaitForElementIsVisible(By.XPath(UNFOLD_ALL));
                _unfoldAll.Click();
                WaitForLoad();
            }
        }
        public bool PriceIsVisible()
        {
            try
            {
                var linePrice = WaitForElementIsVisible(By.XPath(LINE_PRICE));

                if (linePrice != null && linePrice.Displayed)
                {
                 
                    return true;
                }
                else
                {
                    return false;
                }
            }
            catch (NoSuchElementException)
            {
                return false;
            }
        }
        public Boolean IsUnfoldAll()
        {
            _content = WaitForElementIsVisible(By.XPath(FIRST_CONTENT));

            // Temps nécessaire pour que l'élément change de classe
            WaitPageLoading();
            return _content.GetAttribute("class") == "panel-collapse collapse show";
        }

        // Onglets
        public ServiceGeneralInformationPage ClickOnGeneralInformationTab()
        {
            WaitLoading();
            _generalInformationTab = _webDriver.FindElement(By.Id(GENERAL_INFORMATION_TAB));
            _generalInformationTab.Click();
            WaitForLoad();

            return new ServiceGeneralInformationPage(_webDriver, _testContext);
        }

        public ServiceGeneralInformationPage ClickOnGeneralInformationTabOtherNavigation()
        {
            var wait = new WebDriverWait(_webDriver, TimeSpan.FromSeconds(60));
            wait.Until((driver) => driver.WindowHandles.Count > 1);
            _webDriver.SwitchTo().Window(_webDriver.WindowHandles[1]);

            _generalInformationTab = _webDriver.FindElement(By.Id(GENERAL_INFORMATION_TAB));
            _generalInformationTab.Click();
            WaitForLoad();

            return new ServiceGeneralInformationPage(_webDriver, _testContext);
        }

        public ServiceMenuPage GoToMenusTab()
        {
            _menusTab = WaitForElementToBeClickable(By.Id(MENU_TAB));
            _menusTab.Click();
            WaitForLoad();

            return new ServiceMenuPage(_webDriver, _testContext);
        }

        public ServiceDeliveryPage GoToDeliveryTab()
        {
            _deliveryTab = WaitForElementToBeClickable(By.Id(DELIVERY_TAB));
            _deliveryTab.Click();
            WaitForLoad();

            return new ServiceDeliveryPage(_webDriver, _testContext);
        }

        public ServiceLoadingPlanPage GoToLoadingPlanTab()
        {
            _loadingPlanTab = WaitForElementIsVisible(By.Id(LOADINGPLAN_TAB));
            _loadingPlanTab.Click();
            WaitPageLoading();

            return new ServiceLoadingPlanPage(_webDriver, _testContext);
        }

        public ServiceFoodPacketsPage GoToFoodPacketsTab()
        {
            _foodpacketstab = WaitForElementIsVisible(By.Id(FOOD_PACKETS_TAB));
            _foodpacketstab.Click();
            WaitPageLoading();

            return new ServiceFoodPacketsPage(_webDriver, _testContext);
        }



        // Tableau price
        public ServiceCreatePriceModalPage EditFirstPrice(string site, string customerName)
        {
            WaitForLoad();
            Actions action = new Actions(_webDriver);
            _content = WaitForElementIsVisible(By.XPath(FIRST_CONTENT));
            action.MoveToElement(_content).Perform();
            WaitForLoad();

            IWebElement edit;
            edit = WaitForElementIsVisible(By.Id("btn-edit-price"));
            edit.Click();
            WaitForLoad();

            return new ServiceCreatePriceModalPage(_webDriver, _testContext);
        }
        public ServiceCreatePriceModalPage EditPrice(string site, string customerName)
        {
            WaitForLoad();
            Actions action = new Actions(_webDriver);
            WaitForLoad();
            _content = WaitForElementExists(By.XPath(string.Format(CONTENT, site, customerName)));
            action.MoveToElement(_content).Perform();
            WaitForLoad();

            IWebElement edit;
            if (isElementVisible(By.XPath(string.Format(EDIT_PRICE, site, customerName))))
            {
                edit = WaitForElementIsVisible(By.XPath(string.Format(EDIT_PRICE, site, customerName)));
            }
            else
            {
                edit = WaitForElementIsVisible(By.Id("btn-edit-price"));
            }
            edit.Click();
            WaitForLoad();

            return new ServiceCreatePriceModalPage(_webDriver, _testContext);
        }

        // Tableau price
        public void DuplicateAllPrices(string site, string customer)
        {
            _duplicateAllPrices = WaitForElementIsVisible(By.ClassName(DUPLICATE_ALL_PRICES));
            _duplicateAllPrices.Click();
            WaitForLoad();

            // Site
            ComboBoxSelectById(new ComboBoxOptions(DESTINATION_SITE_FILTER, site, false));

            // Customer
            _destinationCustomer = WaitForElementIsVisible(By.Id(CUSTOMER_DESTINATION));
            _destinationCustomer.ClearElement();
            _destinationCustomer.SetValue(ControlType.TextBox, customer);
            WaitForLoad();

            _firstCustomerDestination = WaitForElementIsVisible(By.XPath("//*[@id=\"modal-1\"]/div/div/form/div[2]/div[1]/div[2]/div[2]/div/div/div[2]"));
            _firstCustomerDestination.Click();
            WaitForLoad();

            _duplicateBtn = WaitForElementIsVisible(By.Id(DUPLICATE_BTN));
            _duplicateBtn.Click();
            WaitForLoad();
        }

        public void DeleteFirstPrice()
        {
            Actions action = new Actions(_webDriver);
            _content = WaitForElementIsVisible(By.XPath(FIRST_CONTENT));
            action.MoveToElement(_content).Perform();

            _deletePriceBtn = WaitForElementIsVisible(By.ClassName(PRICE_DELETE_BTN));
            _deletePriceBtn.Click();
            WaitForLoad();

            _deletePriceConfirmBtn = WaitForElementIsVisible(By.Id(PRICE_DELETE_CONFIRM_BTN));
            _deletePriceConfirmBtn.Click();
            WaitForLoad();
        }
        public void DeletePriceForService()
        {
            WaitForLoad();
            Actions action = new Actions(_webDriver);
            _content = WaitForElementIsVisible(By.XPath(DELETE_PRICE));
            _content.Click();
            WaitForLoad();

            _deletePriceBtn = WaitForElementIsVisible(By.ClassName(PRICE_DELETE_BTN));
            _deletePriceBtn.Click();
            WaitForLoad();

            _deletePriceConfirmBtn = WaitForElementIsVisible(By.Id(PRICE_DELETE_CONFIRM_BTN));
            _deletePriceConfirmBtn.Click();
            WaitForLoad();
            var content = WaitForElementIsVisible(By.XPath(DELETE_PRICE));
            content.Click();
                        _deletePriceBtn = WaitForElementIsVisible(By.ClassName(PRICE_DELETE_BTN));
            _deletePriceBtn.Click();
            WaitForLoad();

            _deletePriceConfirmBtn = WaitForElementIsVisible(By.Id(PRICE_DELETE_CONFIRM_BTN));
            _deletePriceConfirmBtn.Click();
            WaitForLoad();
        }

        public void DeleteLastPrice()
        {
            Actions action = new Actions(_webDriver);
            _content = WaitForElementIsVisible(By.XPath(FIRST_CONTENT));
            action.MoveToElement(_content).Perform();
            ReadOnlyCollection<IWebElement> deleteArray;
            if (IsDev())
            {
                //FIXME scan les deux sites du service
                deleteArray = By.ClassName(PRICE_DELETE_BTN).FindElements(_webDriver);
            }
            else
            {
                deleteArray = _webDriver.FindElements(By.XPath("//*[@id=\"list-item-with-action\"]/div[2]/div[2]/div/table/tbody/tr[*]/td[10]/a[contains(@class,'btn-delete-price')]"));
            }



            if (deleteArray.Count <= 1)
            {
                // on laisse un Price sinon le Service disparait de la liste des services... et ne peut pas être recréé.
                return;
            }
           
            _deletePriceBtn = deleteArray[deleteArray.Count / 2];
            _deletePriceBtn.Click();
            WaitForLoad();

            _deletePriceConfirmBtn = WaitForElementIsVisible(By.Id(PRICE_DELETE_CONFIRM_BTN));
            _deletePriceConfirmBtn.Click();
            WaitForLoad();
        }
        public bool IsPriceVisible()
        {
            if (isElementVisible(By.XPath(FIRST_CONTENT)))
            {
                WaitForElementIsVisible(By.XPath(FIRST_CONTENT));
                return true;
            }
            else
            {
                return false;
            }
        }

        public string GetPriceName()
        {
            _priceName = WaitForElementIsVisible(By.XPath(PRICE_NAME));
            return _priceName.Text;
        }

        public bool IsPriceExist()
        {   
            return isElementVisible(By.XPath(PRICE_NAME));
        }


        public string GetPrice()
        {
            WaitPageLoading();
            _priceName = WaitForElementIsVisible(By.XPath(PRICE_ITEM));
            return _priceName.Text;
        }

        public string GetDatasheet()
        {
            _dataSheet = WaitForElementIsVisible(By.XPath(DATASHEET_ITEM));
            string datasheetText = _dataSheet.Text;
            if (datasheetText.Contains("Datasheet:"))
            {
                datasheetText = datasheetText.Replace("Datasheet:", "");
            }
            return datasheetText.Trim();
        }

        public string GetDatasheetValue()
        {
            var DatasheetValue = WaitForElementIsVisible(By.XPath("/html/body/div[4]/div/div/div/div/form/div[2]/div[3]/div/div/div[3]/div/span[1]/input[2]"));
            return DatasheetValue.GetAttribute("value");
        }



        public string GetMode()
        {

            var _dataSheet = WaitForElementIsVisible(By.XPath("//*[starts-with(@id,'content_')]/div/table/tbody/tr/td[6]"));
            string modeText = _dataSheet.Text;
            if (modeText.Contains("Mode: "))
            {
                modeText = modeText.Replace("Mode: ", "");
            }
            return modeText;
        }

        public string GetCustomerSrv()
        {
            _customerSrv = WaitForElementIsVisible(By.XPath(CUSTOMER_SRV));
            return _customerSrv.Text;
        }

        public int GetPricesItem()
        {
            var elements = _webDriver.FindElements(By.XPath(ITEMS));
            return elements.Count;
        }

        public string GetPriceDateFrom()
        {
            WaitLoading();
            _priceDateFrom = WaitForElementIsVisible(By.XPath(PRICE_DATE_FROM));
            string pattern = @"From:\s*(.*)";
            Match match = Regex.Match(_priceDateFrom.Text, pattern);
            WaitForLoad();
             return match.Success? match.Groups[1].Value.Trim():"";
        }

        public string GetPriceDateTo()
        {
            _priceDateTo = WaitForElementIsVisible(By.XPath(PRICE_DATE_TO));
            string pattern = @"To:\s*(.*)";
            Match match = Regex.Match(_priceDateTo.Text, pattern);
            WaitForLoad();
            return match.Success ? match.Groups[1].Value.Trim() : "";
        
        }

        public void SearchPriceForCustomer(string customerName, string site, DateTime fromDate, DateTime toDate, string priceAmount = "2", string datasheet = null)
        {
            Thread.Sleep(2000);
            if (isElementVisible(By.XPath(String.Format(ARROW_PRICE, site, customerName))))
            {

                var arrowPrice = WaitForElementIsVisible(By.XPath(String.Format(ARROW_PRICE, site, customerName)));
                arrowPrice.Click();
                Thread.Sleep(2000);
                WaitForLoad();

                var editPrice = WaitForElementIsVisible(By.XPath(String.Format(EDIT_PRICE, site, customerName)));
                editPrice.Click();
                WaitForLoad();

                var editPriceModalPage = new ServiceCreatePriceModalPage(_webDriver, _testContext);
                editPriceModalPage.EditPriceDates(fromDate, toDate, datasheet);

            }
            else
            {
                var priceModalPage = AddNewCustomerPrice();
                priceModalPage.FillFields_CustomerPrice(site, customerName, fromDate, toDate, priceAmount, datasheet);
            }
        }

        public bool IsPriceDuplicated(string customerName, string site)
        {
            if (isElementVisible(By.XPath(string.Format(PRICE, site, customerName))))
            {
                _webDriver.FindElements(By.XPath(string.Format(PRICE, site, customerName)));
            }
            else
            {
                return false;
            }

            return true;
        }

        public void ToggleFirstPrice()
        {
            _toggleFirstPrice = WaitForElementIsVisible(By.XPath(TOGGLE_FIRST_PRICE));
            _toggleFirstPrice.Click();
        }
        public void TogglePrice(string site, string customer)
        {
            _toggleFirstPrice = WaitForElementIsVisible(By.XPath(string.Format("//td[contains(text(), '{0}') and contains(text(), '{1}')]/../../../../../div[1]/span[@class = 'fas fa-angle-right']", site, customer)));
            _toggleFirstPrice.Click();
        }

        public bool CheckActiveService()
        {

            var wait = new WebDriverWait(_webDriver, TimeSpan.FromSeconds(60));
            wait.Until((driver) => driver.WindowHandles.Count > 1);
            _webDriver.SwitchTo().Window(_webDriver.WindowHandles[1]);

            _serviceprice = WaitForElementExists(By.XPath(SERVICE_PRICE));

            WaitForLoad();

            return _serviceprice.GetAttribute("class").Contains("active");

        }
        public void Go_To_New_Navigate()
        {
            var wait = new WebDriverWait(_webDriver, TimeSpan.FromSeconds(60));
            wait.Until((driver) => driver.WindowHandles.Count > 1);
            _webDriver.SwitchTo().Window(_webDriver.WindowHandles[1]);
            WaitPageLoading();
        }

        public DatasheetDetailsPage EditFirstItem()
        {
            if (IsDev())
            {
                _firstItem = WaitForElementIsVisible(By.XPath(FIRST_ITEM));
                _firstItem.Click();
                _editFirstItem = WaitForElementIsVisible(By.Id(EDIT_FIRST_ITEM));
                _editFirstItem.Click();

            }
            else
            {
                _firstItem = WaitForElementIsVisible(By.XPath(FIRST_ITEM));
                _firstItem.Click();
                _editFirstItemPatch = WaitForElementIsVisible(By.Id(EDIT_FIRST_ITEM_PATCH));
                _editFirstItemPatch.Click();
            }
            return new DatasheetDetailsPage(_webDriver, _testContext);
        }

        public void SetFirstPriceDataSheet(string datasheet)
        {
            var editPriceModalPage = new ServiceCreatePriceModalPage(_webDriver, _testContext);
            if (!isDatasheetExist()&&!string.IsNullOrEmpty(datasheet))
            {
                editPriceModalPage.setDatasheet(datasheet);
            }
            else
            {
                Thread.Sleep(TimeSpan.FromSeconds(2));
                editPriceModalPage.Close();
            }
        }
        public bool isDatasheetExist()
        {
            _firstPriceItem = WaitForElementIsVisible(By.XPath(FIRST_PRICE_ITEM));
            _firstPriceItem.Click();
            WaitForLoad();

            var _datasheetName = WaitForElementIsVisible(By.XPath("//*[@id=\"modal-1\"]/div/div/form/div[2]/div[3]/div/div/div[3]/div/span[1]/input[2]"));
            if (_datasheetName.GetAttribute("value") != null && _datasheetName.GetAttribute("value") != "")
            {
                return true;
            }
            return false;
        }
        public string GetCustomerSrvFromInput()
        {
            WaitForLoad();
            WaitPageLoading();
            var customerSrvElement = WaitForElementIsVisible(By.Id("item_CustomerSrv"));

            if (customerSrvElement.TagName == "input" || customerSrvElement.TagName == "textarea")
            {
                return customerSrvElement.GetAttribute("value");
            }
            else
            {
                return customerSrvElement.Text;
            }
        }
        public void ClickFirstItem()
        {
            var x = WaitForElementIsVisible(By.XPath(FIRST_PRICE_ITEM));
            x.Click();
            WaitForLoad();
        }
        public void ClickPictoDataSheet()
        {

            if (IsDev())
            {
                _pictoDataSheet = WaitForElementIsVisible(By.Id(PICTO_DATASHEET));
                _pictoDataSheet.Click();
                WaitForLoad();

            }
            else
            {
                _pictoDataSheetPatch = WaitForElementIsVisible(By.Id(PICTO_DATASHEET_PATCH));
                _pictoDataSheetPatch.Click();
                WaitForLoad();
            }


            
        }
        public string GetLastPriceName()
        {
            _priceName = WaitForElementIsVisible(By.XPath(LAST_PRICE_NAME));
            return _priceName.Text;
        }
        public void ToggleLastPrice()
        {
            _toggleFirstPrice = WaitForElementIsVisible(By.XPath("//*[@id=\"list-item-with-action\"]/div[2]/div[1]/div/div[1]"));
            _toggleFirstPrice.Click();
        }
        public void DeletePrice()
        {
            Actions action = new Actions(_webDriver);
            _lastcontent = WaitForElementIsVisible(By.XPath(LAST_CONTENT));
            action.MoveToElement(_lastcontent).Perform();
            _deleteLastPriceBtn = WaitForElementIsVisible(By.XPath(LAST_PRICE_DELETE_BTN));
            _deleteLastPriceBtn.Click();
            WaitForLoad();

            _deletePriceConfirmBtn = WaitForElementIsVisible(By.Id(PRICE_DELETE_CONFIRM_BTN));
            _deletePriceConfirmBtn.Click();
            WaitForLoad();
        }

        public bool VerifyIfNewWindowIsOpened()
        {

            var wait = new WebDriverWait(_webDriver, TimeSpan.FromSeconds(60));
            if (_webDriver.WindowHandles.Count == 1)
            {
                return false;
            }
            return true;
        }
        public ServiceCreatePriceModalPage ServiceEitPriceModal(int LinePriceToEdit)
        {
            var lines = _webDriver.FindElements(By.XPath(EDIT_PRICE_LINE));
            WaitPageLoading();
            lines[LinePriceToEdit - 1].Click();
            WaitForLoad();
            return new ServiceCreatePriceModalPage(_webDriver, _testContext);
        }
        public void ClickEditModeCycle()
        {
            var _editcyclemode = WaitForElementExists(By.XPath(EDITCYCLEMODE));
            _editcyclemode.Click();
            WaitForLoad();
        }

        public void SetPrice(string price)
        {
            if (isElementVisible(By.Id(PRICE_CYLE)))
            {
                _priceCycle = WaitForElementIsVisible(By.Id(PRICE_CYLE));
                _priceCycle.SetValue(PageBase.ControlType.TextBox, price);
                _priceCycle.SendKeys(Keys.Enter);
                
            }
            _savePriceBtn = WaitForElementIsVisible(By.XPath(SAVE_PRICE_BTN));
            _savePriceBtn.Click();
            WaitForLoad();
        }

        public void OpenModalCalender()
        {
            var firstContent = WaitForElementIsVisible(By.XPath(FIRST_CONTENT));
            new Actions(_webDriver).MoveToElement(firstContent).Perform();
            var _calenderbtnOpen = WaitForElementIsVisible(By.XPath(CALENDER_BTN));
            WaitForLoad();
            _calenderbtnOpen.Click();
        }

        public bool VerifyIfModalIsOpened()
        {
            var modalTitle = WaitForElementIsVisible(By.XPath("//*[@id=\"modal-1\"]/div[1]/h4"));
            if (modalTitle != null)
            {
                return true;
            }
            return false;
        }
        public void DuplicatePrice()
        {
            if (IsDev())
            {
                _duplicateButton = WaitForElementIsVisible(By.Id(DUPLICATE_BUTTON));
            }
            else
            {
                _duplicateButton = WaitForElementIsVisible(By.Id(DUPLICATE_BUTTON));
            }
            _duplicateButton.Click();
            WaitForLoad();
        }
        public void ConfirmDuplicatePrice()
        {
            _duplicateBtn = WaitForElementIsVisible(By.Id(DUPLICATE_BTN));
            _duplicateBtn.Click();
            WaitForLoad();
        }
        public void UpdateSiteWhenDuplicatingPrice(string site)
        {
            ComboBoxSelectById(new ComboBoxOptions(DESTINATION_SITE_FILTER, site, false));
            WaitForLoad();
        }
        public bool CheckErrorMessage()
        {
            return isElementVisible(By.XPath("//*[@id=\"modal-1\"]/div/div/div/div/form/div[2]/div[2]/span"));
        }
        public void TogglePrice()
        {
            _PriceToggle = WaitForElementIsVisible(By.XPath(PRICE_TOGGLE));
            _PriceToggle.Click();
        }

        public int GetPricesRow()
        {
            var elements = _webDriver.FindElements(By.XPath("//*[@id=\"list-item-with-action\"]/div[*]/div[1]/div/div[2]"));
            return elements.Count;
        }
        public ServiceCreatePriceModalPage EditFirstPriceService()
        {
            Actions a = new Actions(_webDriver);
            _editFirstService = WaitForElementExists(By.XPath(EDIT_PRICE_SERVICE));
            a.MoveToElement(_editFirstService).Perform();
            WaitForLoad();
            _editFirstService.Click();
            WaitForLoad();
            return new ServiceCreatePriceModalPage(_webDriver, _testContext);
        }
             
    }
}
