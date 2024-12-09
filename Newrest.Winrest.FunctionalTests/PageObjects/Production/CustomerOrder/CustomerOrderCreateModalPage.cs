using Microsoft.VisualStudio.TestTools.UnitTesting;
using Newrest.Winrest.FunctionalTests.PageObjects.Shared;
using Newrest.Winrest.FunctionalTests.Utils;
using OpenQA.Selenium;
using OpenQA.Selenium.Support.PageObjects;
using System;

namespace Newrest.Winrest.FunctionalTests.PageObjects.Production.CustomerOrder
{
    public class CustomerOrderCreateModalPage : PageBase
    {
        public CustomerOrderCreateModalPage(IWebDriver webDriver, TestContext testContext) : base(webDriver, testContext)
        {
        }

        //__________________________________ Constantes _____________________________________

        private const string SELECTED_SITE_ID = "dropdown-site";
        private const string CUSTOMER = "//*[@id=\"createFormOrder\"]/div/div[1]/div[2]/div/span[1]/span/input[2]";
        private const string DELIVERY = "DeliveryName";
        private const string SELECTED_CUSTOMER = "//*[@id=\"createFormOrder\"]/div/div[1]/div[2]/div/span[1]/span/div/div/div";
        private const string SELECTED_DELIVERY = "//*[@id=\"createFormOrder\"]/div/div[1]/div[4]/div/span/span/div/div/div";
        private const string FLIGHT_NUMBER = "FlightNumber";
        private const string AIRCRAFT = "Aircraftt";
        private const string DATE = "OrderDate";
        private const string SUBMIT = "btn-submit-form-create-order";
        private const string ORIGIN = "AirportOrig";
        private const string FLIGHTDATE = "flightDate";

        private const string SOURCE_SITE_DROPDOWN = "//*[@id=\"SelectedSiteId\"]";
        private const string SOURCE_CUSTOMER_TEXTBOX = "//*[@id=\"SelectedCustomerId\"]";
        private const string DESTINATION_SITE_DROPDOWN = "//*[@id=\"SelectedDestinationSiteId\"]";
        private const string DESTINATION_CUSTOMER_TEXTBOX = "//*[@id=\"SelectedDestinationCustomerId\"]";
        private const string COPY_ITEMS_BUTTON = "//*[@id=\"checkBoxCopyFrom\"]";
        private const string FIRST_ORDER_CHECKBOX = "//*[@id=\"item_IsSelected\"]"; ////*[@id="table-copy-from-orders"]/tbody/tr[2]/td[5]/div
        private const string NEW_ITEM = "addOrderDetailBtn";
        private const string ITEM_NAME = "[placeholder = 'Item']";
        private const string ITEM_SELECTED = "//*[@id=\"dispatchTable\"]/tbody/tr[2]/td[2]/span/span/div/div/div[1]";
        private const string ITEM_QUANTITY = "[placeholder='Quantity']";
        private const string ITEM_CATEGORY = "/html/body/div[3]/div/div[3]/div[2]/div/div[1]/div[2]/div/div/table/tbody/tr[2]/td[3]/select";
        private const string ICON_SAVED = "//*[@id=\"dispatchTable\"]/tbody/tr[2]/td[15]/span[@class='fas fa-save']";



        private const string FIRSTMESSAGE = "//*[@id=\"flight-detail\"]/div/div[1]/div[2]/div/div/span";
        private const string SECONDMESSAGE = "//*[@id=\"flight-detail\"]/div/div[1]/div[4]/div/span";
        private const string APPLYVAT = "//*[@id=\"ApplyVAT\"]";
        private const string ORDER_TYPE = "dropdown-order-type";
        private const string FIRST_ORDER_TYPE = "//*[@id=\"dropdown-order-type\"]/option[1]";
        private const string HANDLER = "dropdown-handler-type";
        private const string ORDER_NUMBER = "//*[@id=\"div-body\"]/div/div[1]/h1";


        //__________________________________ Variables ______________________________________

        [FindsBy(How = How.XPath, Using = ORDER_NUMBER)]
        private IWebElement _orderNumber;

        [FindsBy(How = How.Id, Using = SELECTED_SITE_ID)]
        private IWebElement _site;

        [FindsBy(How = How.Id, Using = ORIGIN)]
        private IWebElement _orgin;

        [FindsBy(How = How.Id, Using = FLIGHTDATE)]
        private IWebElement _flightDate;

        [FindsBy(How = How.XPath, Using = CUSTOMER)]
        private IWebElement _customer;

        [FindsBy(How = How.Id, Using = DELIVERY)]
        private IWebElement _delivery;

        [FindsBy(How = How.XPath, Using = SELECTED_CUSTOMER)]
        private IWebElement _selectedCustomer;

        [FindsBy(How = How.XPath, Using = SELECTED_DELIVERY)]
        private IWebElement _selectedDelivery;

        [FindsBy(How = How.Id, Using = FLIGHT_NUMBER)]
        private IWebElement _flightNumber;

        private int _flightNumberInt;

        [FindsBy(How = How.Id, Using = AIRCRAFT)]
        private IWebElement _aircraft;

        [FindsBy(How = How.Id, Using = DATE)]
        private IWebElement _date;

        [FindsBy(How = How.Id, Using = SUBMIT)]
        private IWebElement _submit;

        [FindsBy(How = How.Id, Using = ORDER_TYPE)]
        private IWebElement _orderType;

        [FindsBy(How = How.Id, Using = HANDLER)]
        private IWebElement _handler;

        [FindsBy(How = How.Id, Using = NEW_ITEM)]
        private IWebElement _createNewItem;

        [FindsBy(How = How.CssSelector, Using = ITEM_NAME)]
        private IWebElement _itemName;

        [FindsBy(How = How.XPath, Using = ITEM_SELECTED)]
        private IWebElement _itemSelected;

        [FindsBy(How = How.CssSelector, Using = ITEM_QUANTITY)]
        private IWebElement _itemQuantity;

        [FindsBy(How = How.CssSelector, Using = ITEM_CATEGORY)]
        private IWebElement _itemCategory;
        //__________________________________ Méthodes __________________________________________

        public void FillField_CreatNewCustomerOrder(string site, string customer, string aircraft, DateTime? date = null , string orderType = null)
        {
            Random rnd = new Random();

            _site = WaitForElementExists(By.Id(SELECTED_SITE_ID));
            _site.SetValue(ControlType.DropDownList, site);
            WaitForLoad();

            _customer = WaitForElementExists(By.XPath(CUSTOMER));
            _customer.SetValue(ControlType.TextBox, customer);

            _selectedCustomer = WaitForElementExists(By.XPath(SELECTED_CUSTOMER));
            _selectedCustomer.Click();
            WaitForLoad();

            _aircraft = WaitForElementExists(By.Id(AIRCRAFT));
            _aircraft.Click();
            _aircraft.SetValue(ControlType.DropDownList, aircraft);

            _flightNumber = WaitForElementExists(By.Id(FLIGHT_NUMBER));
            _flightNumberInt = rnd.Next(1000, 6000);
            _flightNumber.SetValue(ControlType.TextBox, _flightNumberInt.ToString());

            if (date != null)
            {
                _date = WaitForElementIsVisible(By.Id(DATE));
                _date.SetValue(ControlType.DateTime, date);
                _date.SendKeys(Keys.Tab);
            }
            else
            {
                _date = WaitForElementIsVisible(By.Id(DATE));
                _date.SetValue(ControlType.DateTime, DateUtils.Now);
                _date.SendKeys(Keys.Tab);
            }

            if (orderType != null)
            {
                _orderType = WaitForElementExists(By.Id(ORDER_TYPE));
                _orderType.SetValue(ControlType.DropDownList, orderType);
                WaitForLoad();
            }
            else
            {
                _orderType = WaitForElementIsVisible(By.Id(ORDER_TYPE));
                _orderType.Click();
                var firstOrderType = WaitForElementIsVisible(By.XPath(FIRST_ORDER_TYPE));
                firstOrderType.Click();
            }         
        }

        public void FillField_CreatNewCustomerOrderWOUTflight(string site, string customer, string delivery, DateTime? date = null)
        {
            _site = WaitForElementIsVisible(By.Id(SELECTED_SITE_ID));
            _site.SetValue(ControlType.DropDownList, site);
            WaitForLoad();

            _customer = WaitForElementIsVisible(By.XPath(CUSTOMER));
            _customer.SetValue(ControlType.TextBox, customer);
            WaitForLoad();

            _selectedCustomer = WaitForElementIsVisible(By.XPath(SELECTED_CUSTOMER));
            _selectedCustomer.Click();
            WaitForLoad();
            
            if (date != null)
            {
                _date = WaitForElementIsVisible(By.Id(DATE));
                _date.SetValue(ControlType.DateTime, date);
                _date.SendKeys(Keys.Tab);
            }
            else
            {
                _date = WaitForElementIsVisible(By.Id(DATE));
                _date.SetValue(ControlType.DateTime, DateUtils.Now);
                _date.SendKeys(Keys.Tab);
            }
            WaitForLoad();

            _delivery = WaitForElementIsVisible(By.Id(DELIVERY));
            _delivery.SetValue(ControlType.TextBox, delivery);
            WaitForLoad();

            if (isElementVisible(By.XPath(SELECTED_DELIVERY)))
            { 
                _selectedDelivery = WaitForElementIsVisible(By.XPath(SELECTED_DELIVERY));
                _selectedDelivery.Click();
                WaitForLoad();
            }
        }

        public CustomerOrderItem Submit()
        {
            _submit = WaitForElementIsVisible(By.Id(SUBMIT));
            _submit.Click();
            WaitForLoad();

            return new CustomerOrderItem(_webDriver, _testContext);
        }

        public string GetFlightNumber()
        {
            return _flightNumberInt.ToString();
        }

        public void FillField_CreatFromCustomerOrder(string CustomerOrderSite, string CustomerOrderCustomer)
        {
            // Sélection du site de source
            var sourceSiteDropDown = WaitForElementExists(By.XPath(SOURCE_SITE_DROPDOWN));
            sourceSiteDropDown.SetValue(ControlType.DropDownList, CustomerOrderSite);
            WaitPageLoading();
            // Sélection du client source
            var sourceCustomerTextBox = WaitForElementExists(By.XPath(SOURCE_CUSTOMER_TEXTBOX));
            sourceCustomerTextBox.SetValue(ControlType.TextBox, CustomerOrderCustomer);
            WaitPageLoading();
            // Sélection du site de destination
            var destinationSiteDropDown = WaitForElementExists(By.XPath(DESTINATION_SITE_DROPDOWN));
            destinationSiteDropDown.SetValue(ControlType.DropDownList, CustomerOrderSite);
            WaitPageLoading();
            // Sélection du client destination
            var destinationCustomerTextBox = WaitForElementExists(By.XPath(DESTINATION_CUSTOMER_TEXTBOX));
            destinationCustomerTextBox.SetValue(ControlType.TextBox, CustomerOrderCustomer);
            WaitForLoad();
        }

        public void CopyItemsFromAnotherOrder()
        {
            // Cliquer sur le bouton "Copy items from another order"
            var copyItemsButton = WaitForElementExists(By.XPath(COPY_ITEMS_BUTTON));
            copyItemsButton.Click();
            WaitForLoad();
        }
        public void SelectFirstOrderToDuplicate()
        {
            var firstOrderCheckbox = WaitForElementExists(By.XPath(FIRST_ORDER_CHECKBOX));
            firstOrderCheckbox.Click();
            WaitForLoad();
        }

        public bool verifierApplyVAT()
        {
            var elements1 = _webDriver.FindElements(By.XPath(APPLYVAT));

            bool premierMessageExiste = elements1.Count > 0;
            return premierMessageExiste;
        }
        public bool VerifierLesMessages()
        {
            var elements1 = _webDriver.FindElements(By.XPath(FIRSTMESSAGE));
            var elements2 = _webDriver.FindElements(By.XPath(SECONDMESSAGE));
            bool premierMessageExiste = elements1.Count > 0;
            bool secondMessageExiste = elements2.Count > 0;
            return premierMessageExiste && secondMessageExiste;
        }
        public void FillField_CreatNewCustomerOrder_WithoutFlight(string site, string customer, DateTime? date = null)
        {
            _site = WaitForElementExists(By.Id(SELECTED_SITE_ID));
            _site.SetValue(ControlType.DropDownList, site);
            WaitForLoad();

            _customer = WaitForElementExists(By.XPath(CUSTOMER));
            _customer.SetValue(ControlType.TextBox, customer);

            _selectedCustomer = WaitForElementExists(By.XPath(SELECTED_CUSTOMER));
            _selectedCustomer.Click();

            if (date != null)
            {
                _date = WaitForElementIsVisible(By.Id(DATE));
                _date.SetValue(ControlType.DateTime, date);
                _date.SendKeys(Keys.Tab);
            }
            else
            {
                _date = WaitForElementIsVisible(By.Id(DATE));
                _date.SetValue(ControlType.DateTime, DateUtils.Now);
                _date.SendKeys(Keys.Tab);
            }
        }

        public void FillField_CreatNewCustomerOrderWithHandler(string site, string customer, string aircraft, string handler)
        {
            Random rnd = new Random();

            _site = WaitForElementExists(By.Id(SELECTED_SITE_ID));
            _site.SetValue(ControlType.DropDownList, site);
            WaitForLoad();

            _customer = WaitForElementExists(By.XPath(CUSTOMER));
            _customer.SetValue(ControlType.TextBox, customer);

            _selectedCustomer = WaitForElementExists(By.XPath(SELECTED_CUSTOMER));
            _selectedCustomer.Click();
            WaitForLoad();

            _aircraft = WaitForElementExists(By.Id(AIRCRAFT));
            _aircraft.Click();
            _aircraft.SetValue(ControlType.DropDownList, aircraft);

            _handler = WaitForElementExists(By.Id(HANDLER));
            _handler.Click();
            _handler.SetValue(ControlType.DropDownList, handler);
            WaitForLoad();

            _flightNumber = WaitForElementExists(By.Id(FLIGHT_NUMBER));
            _flightNumberInt = rnd.Next(1000, 6000);
            _flightNumber.SetValue(ControlType.TextBox, _flightNumberInt.ToString());
        }
        public void FillField_CreatNewCustomerOrder2(string site, string customer, string aircraft, string flightNumber,  DateTime? date = null, string orderType = null, DateTime? dateFlight = null)
        {
            Random rnd = new Random();

            _site = WaitForElementExists(By.Id(SELECTED_SITE_ID));
            _site.SetValue(ControlType.DropDownList, site);
            WaitForLoad();

            _customer = WaitForElementExists(By.XPath(CUSTOMER));
            _customer.SetValue(ControlType.TextBox, customer);

            _selectedCustomer = WaitForElementExists(By.XPath(SELECTED_CUSTOMER));
            _selectedCustomer.Click();
            WaitForLoad();

            _aircraft = WaitForElementExists(By.Id(AIRCRAFT));
            _aircraft.Click();
            _aircraft.SetValue(ControlType.DropDownList, aircraft);

            _flightNumber = WaitForElementExists(By.Id(FLIGHT_NUMBER));
            _flightNumber.SetValue(ControlType.TextBox, flightNumber);

            _orgin = WaitForElementExists(By.Id(ORIGIN));
            _orgin.SetValue(ControlType.DropDownList, site);

            if (date != null)
            {
                _date = WaitForElementIsVisible(By.Id(DATE));
                _date.SetValue(ControlType.DateTime, date);
                _date.SendKeys(Keys.Tab);
            }
            else
            {
                _date = WaitForElementIsVisible(By.Id(DATE));
                _date.SetValue(ControlType.DateTime, DateUtils.Now);
                _date.SendKeys(Keys.Tab);
            }

            if (orderType != null)
            {
                _orderType = WaitForElementExists(By.Id(ORDER_TYPE));
                _orderType.SetValue(ControlType.DropDownList, orderType);
                WaitForLoad();
            }
            else
            {
                _orderType = WaitForElementIsVisible(By.Id(ORDER_TYPE));
                _orderType.Click();
                var firstOrderType = WaitForElementIsVisible(By.XPath(FIRST_ORDER_TYPE));
                firstOrderType.Click();
            }
            if(dateFlight != null)
            {
                _flightDate = WaitForElementIsVisible(By.Id(FLIGHTDATE));
                _flightDate.SetValue(ControlType.DateTime, dateFlight);
            }
        }
        public bool AddNewItemWithCategory(string itemName, string quantity, string category)
        {
            try
            {
                // Click sur le bouton +
                _createNewItem = WaitForElementIsVisible(By.Id(NEW_ITEM));
                _createNewItem.Click();
                WaitForLoad();

                _itemName = WaitForElementIsVisible(By.CssSelector(ITEM_NAME));
                _itemName.SetValue(ControlType.TextBox, itemName);
                WaitForLoad();

                // Selection du premier élément de la liste
                _itemSelected = WaitForElementIsVisible(By.XPath(ITEM_SELECTED));
                _itemSelected.Click();
                WaitForLoad();

                // Temps d'attente obligatoire pour les informations du service créé se chargent
                _itemQuantity = WaitForElementIsVisible(By.CssSelector(ITEM_QUANTITY));
                _itemQuantity.SetValue(ControlType.TextBox, quantity);

                _itemCategory = WaitForElementIsVisible(By.XPath(ITEM_CATEGORY));
                _itemCategory.SetValue(ControlType.DropDownList, category);

                // Temps d'attente obligatoire pour la prise en compte de l'item
                WaitForElementIsVisible(By.XPath(ICON_SAVED));
            }
            catch
            {
                return false;
            }

            return true;
        }
        public string GetOrderNumber()
        {
            // Localiser l'élément contenant le titre complet
            IWebElement titleElement = WaitForElementIsVisible(By.XPath("//*[@id=\"div-body\"]/div/div[1]/h1"));

            // Récupérer le texte complet
            string fullTitle = titleElement.Text;

            // Utiliser une expression régulière pour extraire le numéro de commande (partie numérique uniquement)
            var match = System.Text.RegularExpressions.Regex.Match(fullTitle, @"\d+");
            if (match.Success)
            {
                return match.Value; // Retourner uniquement le nombre
            }
            throw new Exception($"Aucun numéro trouvé dans le titre : {fullTitle}");
        }


    }
}
