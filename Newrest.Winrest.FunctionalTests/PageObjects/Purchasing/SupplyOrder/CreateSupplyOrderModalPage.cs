using Microsoft.VisualStudio.TestTools.UnitTesting;
using Newrest.Winrest.FunctionalTests.PageObjects.Shared;
using Newrest.Winrest.FunctionalTests.Utils;
using OpenQA.Selenium;
using OpenQA.Selenium.Support.PageObjects;
using OpenQA.Selenium.Support.UI;
using System;
using System.Collections.Generic;

namespace Newrest.Winrest.FunctionalTests.PageObjects.Purchasing
{
    public class CreateSupplyOrderModalPage : PageBase
    {

        public CreateSupplyOrderModalPage(IWebDriver webDriver, TestContext testContext) : base(webDriver, testContext)
        {
        }
        // _______________________________________ Constantes _____________________________________________

        private const string SELECTED_SITE_ID = "SelectedSiteId";
        private const string DATE_FROM = "datapicker-new-supply-order-start";
        private const string DATE_TO = "datapicker-new-supply-order-end";
        private const string DELIVERY_DATE = "datapicker-new-supply-order-delivery";
        private const string DELIVERY_LOCATION = "SelectedSitePlaceId";
        private const string ACTIVATED = "SupplyOrder_IsActive";
        private const string SUBMIT = "btn-submit-form-create-supply-order";
        private const string ROUND_QUANTITIES = "checkBoxRoundPackagingQtys";
        private const string BACK_TO_LIST = "/html/body/div[2]/a/span[2]";

        private const string SUPPLYORDER_NUMBER = "tb-new-supplyorder-number";
        private const string SELECTED_SO_CHECKBOX = "item_IsSelected";
        private const string FROM_NEEDS_CHECKBOX = "checkBoxPrefillNeedRawMaterials";
        private const string PQ_FROM_MENU_PLANIFICATION = "checkBoxPrefillMenuPlanif";
        private const string PQ_FROM_MENU_CUSTOMERORDERS = "checkBoxCopyFromCustomerOrders";
        private const string PERFILL_QUANTITIES_FROM_PRODUCTION_MANAGEMENT = "//*[@id=\"checkBoxPrefillRawMaterials\"]";

        private const string COPY_ITEMS_ANOTHER_SUPPLY_ORDER_CHECKBOX = "//*[@id=\"form-create-supply-order\"]/div/div[7]/div[1]/div/div";
        private const string FIRST_SUPPLY_ORDER_ROW = "//*[@id=\"div-copy-from-items\"]/table/tbody/tr[2]";
        private const string FIST_SO = "//*[@id=\"div-copy-from-items\"]/table/tbody/tr[2]/td[6]/div";
        private const string FIRST_SUPPLY = "//*[@id=\"list-item-with-action\"]/table/tbody/tr";
        private const string ID_SUPPLY = "/html/body/div[3]/div/div/div/div[2]/div/form/div/div[1]/div[1]/div/div/input";



        // _______________________________________ Variables ______________________________________________


        [FindsBy(How = How.XPath, Using = PERFILL_QUANTITIES_FROM_PRODUCTION_MANAGEMENT)]
        private IWebElement _perfillmanagement;

        [FindsBy(How = How.Id, Using = SELECTED_SITE_ID)]
        private IWebElement _site;

        [FindsBy(How = How.XPath, Using = BACK_TO_LIST)]
        private IWebElement _backToList;

        [FindsBy(How = How.Id, Using = DATE_FROM)]
        private IWebElement _dateFrom;

        [FindsBy(How = How.Id, Using = DATE_TO)]
        private IWebElement _dateTo;

        [FindsBy(How = How.Id, Using = DELIVERY_DATE)]
        private IWebElement _deliveryDate;

        [FindsBy(How = How.Id, Using = DELIVERY_LOCATION)]
        private IWebElement _deliveryLocation;

        [FindsBy(How = How.Id, Using = ACTIVATED)]
        private IWebElement _activated;

        [FindsBy(How = How.Id, Using = SUBMIT)]
        private IWebElement _submit;

        [FindsBy(How = How.Id, Using = ROUND_QUANTITIES)]
        private IWebElement _roundQuantities;

        [FindsBy(How = How.Id, Using = SUPPLYORDER_NUMBER)]
        private IWebElement _supplyOrderNumber;

        [FindsBy(How = How.Id, Using = SELECTED_SO_CHECKBOX)]
        private IWebElement _checkBoxSelectSO;

        [FindsBy(How = How.Id, Using = PQ_FROM_MENU_PLANIFICATION)]
        private IWebElement _checkFromMenuPlanification;

        [FindsBy(How = How.Id, Using = PQ_FROM_MENU_CUSTOMERORDERS)]
        private IWebElement _checkFromMenuCustomerorders;

        [FindsBy(How = How.Id, Using = FROM_NEEDS_CHECKBOX)]
        private IWebElement _fromneedscheckbox;

        [FindsBy(How = How.XPath, Using = FIST_SO)]
        private IWebElement _fistsupplyorder;

        [FindsBy(How = How.XPath, Using = FIRST_SUPPLY)]
        private IWebElement _firstsupply;

        



        [FindsBy(How = How.XPath, Using = COPY_ITEMS_ANOTHER_SUPPLY_ORDER_CHECKBOX)]
        private IWebElement _copyItems;
        [FindsBy(How = How.XPath, Using = FIRST_SUPPLY_ORDER_ROW)]
        private IWebElement _firstSupplyOrderRow;

        // ___________________________________________ Méthodes _______________________________________________



        public void FillPrincipalField_CreationNewSupplyOrder(string site, DateTime from, DateTime to, DateTime deliveryDate, string deliveryLocation, bool isActive = true, bool roundQuantities = false)
        {
            //Attente que la popup modale soit complètement chargée
            _site = WaitForElementIsVisible(By.Id(SELECTED_SITE_ID));
            _site.SetValue(ControlType.DropDownList, site);
            WaitForLoad();

            _dateFrom = WaitForElementIsVisible(By.Id(DATE_FROM));
            _dateFrom.SetValue(ControlType.DateTime, from);
            _dateFrom.SendKeys(Keys.Tab);

            _dateTo = WaitForElementIsVisible(By.Id(DATE_TO));
            _dateTo.SetValue(ControlType.DateTime, to);
            _dateTo.SendKeys(Keys.Tab);

            if (deliveryDate.Date <= from.Date)
            {
                _deliveryDate = WaitForElementIsVisible(By.Id(DELIVERY_DATE));
                _deliveryDate.SetValue(ControlType.DateTime, deliveryDate);
                _deliveryDate.SendKeys(Keys.Tab);
                // WaitPageLoading();
            }

            _deliveryLocation = WaitForElementIsVisible(By.Id(DELIVERY_LOCATION));
            _deliveryLocation.SetValue(ControlType.DropDownList, deliveryLocation);

            _activated = WaitForElementExists(By.Id(ACTIVATED));
            _activated.SetValue(ControlType.CheckBox, isActive);

            _roundQuantities = WaitForElementExists(By.Id(ROUND_QUANTITIES));
            _roundQuantities.SetValue(ControlType.CheckBox, roundQuantities);
        }
        public void FillPrincipalField_CreationNewSupplyOrderActiveFalse(string site, DateTime from, DateTime to, DateTime deliveryDate, string deliveryLocation, bool isActive = false, bool roundQuantities = false)
        {
            //Attente que la popup modale soit complètement chargée
            _site = WaitForElementIsVisible(By.Id(SELECTED_SITE_ID));
            _site.SetValue(ControlType.DropDownList, site);
            WaitForLoad();

            _dateFrom = WaitForElementIsVisible(By.Id(DATE_FROM));
            _dateFrom.SetValue(ControlType.DateTime, from);
            _dateFrom.SendKeys(Keys.Tab);

            _dateTo = WaitForElementIsVisible(By.Id(DATE_TO));
            _dateTo.SetValue(ControlType.DateTime, to);
            _dateTo.SendKeys(Keys.Tab);

            if (deliveryDate.Date <= from.Date)
            {
                _deliveryDate = WaitForElementIsVisible(By.Id(DELIVERY_DATE));
                _deliveryDate.SetValue(ControlType.DateTime, deliveryDate);
                _deliveryDate.SendKeys(Keys.Tab);
                // WaitPageLoading();
            }

            _deliveryLocation = WaitForElementIsVisible(By.Id(DELIVERY_LOCATION));
            _deliveryLocation.SetValue(ControlType.DropDownList, deliveryLocation);

            _activated = WaitForElementExists(By.Id(ACTIVATED));
            _activated.SetValue(ControlType.CheckBox, isActive);

            _roundQuantities = WaitForElementExists(By.Id(ROUND_QUANTITIES));
            _roundQuantities.SetValue(ControlType.CheckBox, roundQuantities);
        }


        public SupplyOrderItem Submit()
        {
            _submit = WaitForElementToBeClickable(By.Id(SUBMIT));
            _submit.Click();
            WaitForLoad();
            CheckErrorModalExist();
            return new SupplyOrderItem(_webDriver, _testContext);
        }
        public SupplyOrderItem CreateButton()
        {
            _submit = WaitForElementToBeClickable(By.Id(SUBMIT));
            _submit.Click();
            WaitForLoad();
            return new SupplyOrderItem(_webDriver, _testContext);
        }
        public SupplyOrderPage BackToList()
        {
            WaitPageLoading();
            _backToList = WaitForElementIsVisible(By.XPath(BACK_TO_LIST));
            _backToList.Click();
            WaitForLoad();

            return new SupplyOrderPage(_webDriver, _testContext);
        }

        public void CheckErrorModalExist()
        {
            WaitForLoad();
            var xpath = "//*[@id=\"modal-1\"]//div[@class=\"errors-panel\"]";
            var ErrorModalIsVisible = isElementVisible(By.XPath(xpath));
            
            if (ErrorModalIsVisible)
            {
                var errorMsg = WaitForElementIsVisible(By.XPath(xpath)).Text;
                Assert.Fail($"A modal error appeared with the following message : '{errorMsg}'");
            }
        }

        public string GetSupplyOrderNumber()
        {
            _supplyOrderNumber = WaitForElementIsVisible(By.Id(SUPPLYORDER_NUMBER));
            return _supplyOrderNumber.GetAttribute("value");
        }

        public void FillFromAnotherSupplyOrder(string site, string supplyOrderNumber)
        {
            var checkBoxCopyFrom = WaitForElementExists(By.Id("checkBoxCopyFrom"));
            checkBoxCopyFrom.SetValue(PageBase.ControlType.CheckBox, true);
            var selectSite = new SelectElement(WaitForElementIsVisible(By.Id("form-copy-site-dropdown")));
            selectSite.SelectByText(site);
            WaitForLoad();

            var soNumber = WaitForElementIsVisible(By.Id("form-copy-from-tbSearchName"));
            soNumber.SendKeys(supplyOrderNumber);
            WaitForLoad();
            WaitPageLoading();

            _checkBoxSelectSO = WaitForElementExists(By.Id(SELECTED_SO_CHECKBOX));
            _checkBoxSelectSO.SetValue(PageBase.ControlType.CheckBox, true);
            WaitForLoad();
        }
        public void CheckedPrefillQtesFromMenuPlanification()
        {
            _checkFromMenuPlanification = WaitForElementExists(By.Id(PQ_FROM_MENU_PLANIFICATION));
            _checkFromMenuPlanification.SetValue(ControlType.CheckBox, true);
            WaitForLoad();
        }

        public void CheckedPrefillQtesFromMenuCustomerorders()
        {
            _checkFromMenuCustomerorders = WaitForElementExists(By.Id(PQ_FROM_MENU_CUSTOMERORDERS));
            _checkFromMenuCustomerorders.SetValue(ControlType.CheckBox, true);
            WaitForLoad();
        }
        public void CheckCustomer()
        {
            var checkbox = WaitForElementExists(By.XPath("/html/body/div[3]/div/div/div/div[4]/div/div/div/form/div/table/tbody/tr[2]/td[5]/div/input[1]"));
            checkbox.SetValue(ControlType.CheckBox, true);
            WaitForLoad();
        }

        public void IsErrorReport()
        {
            if (isElementExists(By.XPath("//*[@id=\"modal-1\"]/div[2]/p[5]/a")))
            {
                var link = WaitForElementExists(By.XPath("//*[@id=\"modal-1\"]/div[2]/p[5]/a"));
                link.Click();
                WaitForLoad();
            }

        }

        public void SetDateFromCustomer(DateTime from)
        {
            var dateFrom = WaitForElementIsVisible(By.XPath("/html/body/div[3]/div/div/div/div[4]/div/div/div/form/table/tbody/tr/td[3]/input"));
            dateFrom.SetValue(ControlType.DateTime, from);
        }

        public void CheckFromNeeds()
        {
            _fromneedscheckbox = WaitForElementExists(By.Id(FROM_NEEDS_CHECKBOX));
            _fromneedscheckbox.SetValue(PageBase.ControlType.CheckBox, true);
            WaitForLoad();
        }
        public void CheckQuantitiesFromProdCOs()
        {
            var quantitiesFromProdCOs = WaitForElementExists(By.XPath("//*[@id=\"checkBoxCopyFromProdCO\"]"));
            quantitiesFromProdCOs.SetValue(PageBase.ControlType.CheckBox, true);
            WaitForLoad();
        }
        public void Check_Perfill_Quantities_From_Production_Management()
        {
            _perfillmanagement = WaitForElementExists(By.XPath(PERFILL_QUANTITIES_FROM_PRODUCTION_MANAGEMENT));
            _perfillmanagement.SetValue(PageBase.ControlType.CheckBox, true);
            WaitForLoad();
        }

        public List<string> GetListOfNumbersFromList()
        {
            var listOfNumbersElements = _webDriver.FindElements(By.XPath("//*[@id=\"div-copy-from-items\"]/table/tbody/tr[*]/td[2]"));
            var listOfNumbers = new List<string>();
            foreach (var element in listOfNumbersElements)
            {
                listOfNumbers.Add(element.Text);
            }
            return listOfNumbers;
        }

        public void FillFromAnotherSupplyOrderFucntion(string site)
        {
            var checkBoxCopyFrom = WaitForElementExists(By.Id("checkBoxCopyFrom"));
            checkBoxCopyFrom.SetValue(PageBase.ControlType.CheckBox, true);
            var selectSite = new SelectElement(WaitForElementIsVisible(By.Id("form-copy-site-dropdown")));
            selectSite.SelectByText(site);
            WaitForLoad();

            var listOfNumbersElements = _webDriver.FindElements(By.XPath("//*[@id=\"div-copy-from-items\"]/table/tbody/tr[*]/td[2]"));
            var listOfNumbers = new List<string>();
            foreach (var element in listOfNumbersElements)
            {
                listOfNumbers.Add(element.Text);
            }
            Random random = new Random();
            int randomIndex = random.Next(listOfNumbers.Count);
            string randomItem = listOfNumbers[randomIndex];
            var soNumber = WaitForElementIsVisible(By.Id("form-copy-from-tbSearchName"));
            soNumber.SendKeys(randomItem);
            WaitForLoad();
            WaitPageLoading();

            _checkBoxSelectSO = WaitForElementExists(By.Id(SELECTED_SO_CHECKBOX));
            _checkBoxSelectSO.SetValue(PageBase.ControlType.CheckBox, true);
            WaitForLoad();
        }

        public void CopyItemsFromAnotherSupplyOrder()
        {
            _copyItems = WaitForElementIsVisible(By.XPath(COPY_ITEMS_ANOTHER_SUPPLY_ORDER_CHECKBOX));
            _copyItems.Click();
            WaitForLoad();
        }
        public void SelectFirstSO()
        {
            _fistsupplyorder = WaitForElementIsVisible(By.XPath(FIST_SO));
            _fistsupplyorder.Click();
            WaitForLoad();
        }
        public void SelectFirstSUPPLY()
        {
            _firstsupply = WaitForElementIsVisible(By.XPath(FIRST_SUPPLY));
            _firstsupply.Click();
            WaitForLoad();
        }
        public bool IsDataAvailable()
        {
            return isElementExists(By.XPath(FIRST_SUPPLY_ORDER_ROW));
        }
        public void filterNumberInPageCreateSO(string number)
        {
            var _FilterNumberInPageCreateSO = WaitForElementIsVisible(By.XPath("//*[@id=\"form-copy-from-tbSearchName\"]"));
            _FilterNumberInPageCreateSO.SetValue(PageBase.ControlType.TextBox, number);
            WaitForLoad();
            WaitPageLoading();
        }
        public string GetFirstItemId()
        {
            var element = WaitForElementIsVisible(By.XPath(ID_SUPPLY));
            return element.GetAttribute("value");
        }
    }
}
