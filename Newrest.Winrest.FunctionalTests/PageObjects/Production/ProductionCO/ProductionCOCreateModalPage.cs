using DocumentFormat.OpenXml.Drawing;
using DocumentFormat.OpenXml.Drawing.Charts;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using Newrest.Winrest.FunctionalTests.PageObjects.Menus.Datasheet;
using Newrest.Winrest.FunctionalTests.PageObjects.Menus.Recipes;
using Newrest.Winrest.FunctionalTests.PageObjects.Shared;
using Newrest.Winrest.FunctionalTests.Utils;
using OfficeOpenXml.FormulaParsing.Excel.Functions.DateTime;
using OfficeOpenXml.FormulaParsing.Excel.Functions.Text;
using OpenQA.Selenium;
using OpenQA.Selenium.Support.PageObjects;
using OpenQA.Selenium.Support.UI;
using System;
using System.Collections.Generic;
using System.Data;
using System.Globalization;
using System.Linq;
using System.Runtime.InteropServices;
using System.Security.Policy;
using System.Threading;
using System.Windows;

namespace Newrest.Winrest.FunctionalTests.PageObjects.Production.CustomerOrder
{
    public class ProductionCOCreateModalPage : PageBase
    {
        public ProductionCOCreateModalPage(IWebDriver webDriver, TestContext testContext) : base(webDriver, testContext)
        {
        }

        //__________________________________ Constantes _____________________________________

        private const string SELECTED_SITE_ID = "drop-down-site";
        private const string DATE_FROM = "date-from-picker";
        private const string DATE_TO = "date-to-picker";
        private const string SEARCH_BUTTON = "searchServiceButton";
        private const string CHECK_BUTTON = "Services_Items_0__IsChecked";
        private const string NEXT_BUTTON = "btn-submit-services";
        private const string NEW_DATE_BUTTON = "btn-create-new-row";
        private const string DATE = "date-picker-0";
        private const string QUANTITY = "input-quantity-0";
        private const string OK_BUTTON = "btn-create-co";
        private const string ALL_SERVICES_NAME = "//*[@id=\"service-details-form\"]/table/tbody/tr[*]/td[1]";
        private const string CASE_HIDDEN = "//*[@id=\"Services_Items_0__IsChecked\"]";
        private const string Date_EXPEDITION_DEV = "//*[@id=\"orders-service-93580\"]/tbody/tr/td[1]";
        private const string Date_EXPEDITION_PATCH = "//*[starts-with(@id,'orders-service-')]/tbody/tr/td[1]";
        private const string PLANNED_QTY = "//*[@id=\"services-list\"]/div[1]/div[2]/span[1]";
        private const string DIFF_QTY = "//*[@id=\"services-list\"]/div[1]/div[4]/span";
       private const string DUPLICATE = "//*[@id=\"services-list\"]/div[2]/div[2]/div/div/div[3]/a[1]";
       private const string CORBEILLE = "//*[@id=\"services-list\"]/div[2]/div[2]/div[1]/div[2]/div[3]/a[2]";
        private const string CHECK_ALL_BUTTON = "check-all";
        private const string SEARCH_SERVICE_NAME = "//*[@id=\"form-copy-from-tbSearchName\"]";

        private const string LIGNEDUPLICATE = "//*[@id=\"services-list\"]/div[2]/div[2]/div[1]/div[2]";
        private const string STYLE = "//*[@id=\"form-create-search\"]/div[1]/div[4]";
        private const string MODAL = "//*[@id=\"modal-1\"]/div[1]";
        private const string SERVICE_NAME = "//*[@id=\"service-details-form\"]/table/tbody/tr[*]/td[1][contains(text(), '{0}')]";
        private const string DATA_SHEET = "datasheet-span";
        private const string CLOSE = "//*[@id=\"service-details-form\"]/div/button[1]";
        private const string SERVICE_NAME_LIST = "//*[@id=\"service-details-form\"]/table/tbody/tr[*]/td[1]";
        //__________________________________ Variables ______________________________________

        [FindsBy(How = How.XPath, Using = LIGNEDUPLICATE)]
        private IWebElement _ligneduplicate;

        [FindsBy(How = How.XPath, Using = CORBEILLE)]
        private IWebElement _corbeille;

        [FindsBy(How = How.XPath, Using = DUPLICATE)]
        private IWebElement _duplicate;

        [FindsBy(How = How.XPath, Using = DIFF_QTY)]
        private IWebElement _diffqty;

        [FindsBy(How = How.XPath, Using = PLANNED_QTY)]
        private IWebElement _plannedqty;

        [FindsBy(How = How.Id, Using = SELECTED_SITE_ID)]
        private IWebElement _site;

        [FindsBy(How = How.Id, Using = DATE_FROM)]
        private IWebElement _dateFrom;

        [FindsBy(How = How.Id, Using = DATE_TO)]
        private IWebElement _dateTo;

        [FindsBy(How = How.Id, Using = SEARCH_BUTTON)]
        private IWebElement _search;

        [FindsBy(How = How.Id, Using = CHECK_BUTTON)]
        private IWebElement _check_button;

        [FindsBy(How = How.Id, Using = NEXT_BUTTON)]
        private IWebElement _next;

        [FindsBy(How = How.Id, Using = NEW_DATE_BUTTON)]
        private IWebElement _new_date_button;

        [FindsBy(How = How.Id, Using = DATE)]
        private IWebElement _date;

        [FindsBy(How = How.Id, Using = QUANTITY)]
        private IWebElement _quantity;

        [FindsBy(How = How.Id, Using = OK_BUTTON)]
        private IWebElement _ok_button;

        [FindsBy(How = How.XPath, Using = Date_EXPEDITION_DEV)]
        private IWebElement _dateExpeditionDev;

        [FindsBy(How = How.XPath, Using = Date_EXPEDITION_PATCH)]
        private IWebElement _dateExpeditionPatch;
        [FindsBy(How = How.Id, Using = CHECK_ALL_BUTTON)]
        private IWebElement _checkAllButton;

        [FindsBy(How = How.XPath, Using = STYLE)]
        private IWebElement _style;

        [FindsBy(How = How.XPath, Using = MODAL)]
        private IWebElement _modal;

        [FindsBy(How = How.XPath, Using = SEARCH_SERVICE_NAME)]
        private IWebElement _searchServiceName;

        [FindsBy(How = How.Id, Using = DATA_SHEET)]
        private IWebElement _datasheet;

        //__________________________________ Méthodes __________________________________________

        public void FillField_NewProductionCOSearch(string site, DateTime? dateFrom, DateTime? dateTo)
        {
            _site = WaitForElementIsVisible(By.Id(SELECTED_SITE_ID));
            _site.SetValue(ControlType.DropDownList, site);
            WaitForLoad();

            _dateFrom = WaitForElementIsVisible(By.Id(DATE_FROM));
            _dateFrom.SetValue(ControlType.DateTime, dateFrom);
            _dateFrom.SendKeys(Keys.Tab);

            _dateTo = WaitForElementIsVisible(By.Id(DATE_TO));
            _dateTo.SetValue(ControlType.DateTime, dateTo);
            _dateTo.SendKeys(Keys.Tab);

            _search = WaitForElementIsVisible(By.Id(SEARCH_BUTTON));
            _search.Click();
            WaitForLoad();
        }

        public bool CheckServiceToProduce(string targetService)
        {
            bool serviceFound = isElementVisible(By.XPath(String.Format(SERVICE_NAME, targetService)));
            return serviceFound;
        }

        public void FillField_NewDate(DateTime? date, bool isActivate = true)
        {
             Random rnd = new Random();

            //_check_button = WaitForElementExists(By.Id(CHECK_BUTTON));
            //_check_button.SetValue(ControlType.CheckBox, isActivate);

            _checkAllButton= WaitForElementExists(By.Id(CHECK_ALL_BUTTON));
            _checkAllButton.Click();

            _next = WaitForElementIsVisible(By.Id(NEXT_BUTTON));
            _next.Click();

            _new_date_button = WaitForElementIsVisible(By.Id(NEW_DATE_BUTTON));
            _new_date_button.Click();

            _date = WaitForElementIsVisible(By.Id(DATE));
            _date.SetValue(ControlType.DateTime, date);
            _date.SendKeys(Keys.Tab);

            _quantity = WaitForElementIsVisible(By.Id(QUANTITY));
            _quantity.SetValue(ControlType.TextBox, rnd.Next(1, 20).ToString());

            _ok_button = WaitForElementIsVisible(By.Id(OK_BUTTON));
            _ok_button.Click();

            WaitForLoad();
        }

        public void Submit()
        {
            _ok_button = WaitForElementIsVisible(By.Id(OK_BUTTON));
            _ok_button.Click();

            WaitForLoad();
        }

        public enum FilterType
        {
            Search,
            DateFrom,
            DateTo,
            Site
        }

        public void Filter(FilterType filterType, object value)
        {
            switch (filterType)
            {
                case FilterType.Site:
                    _site = WaitForElementIsVisible(By.Id(SELECTED_SITE_ID));
                    _site.Click();
                    // var element = WaitForElementIsVisible(By.XPath("//option[contains(@value,'" + value + "')]"));
                    // _site.SetValue(ControlType.DropDownList, element.Text);
                    _site.SetValue(ControlType.DropDownList, value);
                    WaitForLoad();
                    _site.Click();

                    break;
                case FilterType.Search:
                    _search = WaitForElementIsVisible(By.Id(SEARCH_BUTTON));
                    _search.SetValue(ControlType.TextBox, value);
                    break;
                case FilterType.DateFrom:

                    _dateFrom = WaitForElementIsVisible(By.Id(DATE_FROM));
                    _dateFrom.SetValue(ControlType.DateTime, value);
                    _dateFrom.SendKeys(Keys.Tab);
                    break;
                case FilterType.DateTo:
                    _dateTo = WaitForElementIsVisible(By.Id(DATE_TO));
                    _dateTo.SetValue(ControlType.DateTime, value);
                    _dateTo.SendKeys(Keys.Tab);
                    break;
                default:
                    break;
            }
        }
        public ProductionCOCreateModalPage Search()
        {
            WaitForLoad();
            var price = WaitForElementIsVisible(By.Id(SEARCH_BUTTON));
            price.Click();
            WaitPageLoading();
            return new ProductionCOCreateModalPage(_webDriver, _testContext);
        }
        public bool isServicehidden()
        {
            var caseservice = WaitForElementIsVisible(By.XPath(CASE_HIDDEN));
            var ishidden = caseservice.GetAttribute("disabled");
            return bool.Parse(ishidden);
        }
        public string GetDateExpedition()
        {
            if (IsDev())
            {
                var dateElement = WaitForElementExists(By.XPath(Date_EXPEDITION_DEV));
                return dateElement.Text;
            }
            else
            {
                var dateElement = WaitForElementExists(By.XPath(Date_EXPEDITION_PATCH));
                return dateElement.Text;
            }
        }

        public void CheckBox_Service(bool isActivate = true)
        {
            _check_button = WaitForElementExists(By.Id(CHECK_BUTTON));
            _check_button.SetValue(ControlType.CheckBox, isActivate);

            _next = WaitForElementIsVisible(By.Id(NEXT_BUTTON));
            _next.Click();
        }
        public void checkFirstService()
        {
            var services = _webDriver.FindElements(By.XPath("//*[contains(@id,\"Services_Items\")]"));
            IWebElement fisrtServiceWithDataSheet = services.FirstOrDefault(serv=>serv.Enabled==true);
            fisrtServiceWithDataSheet.Click();
            WaitForLoad();
            _next = WaitForElementIsVisible(By.Id(NEXT_BUTTON));
            _next.Click();
            WaitForLoad();
        } 
        public void createNewDate(DateTime? date)
        {
            Random rnd = new Random();
            WaitForLoad();
            _new_date_button = WaitForElementIsVisible(By.Id(NEW_DATE_BUTTON));
            _new_date_button.Click();
            _date = WaitForElementIsVisible(By.Id(DATE));
            _date.SetValue(ControlType.DateTime, date);
            _date.SendKeys(Keys.Tab);
            _quantity = WaitForElementIsVisible(By.Id(QUANTITY));
            _quantity.SetValue(ControlType.TextBox, rnd.Next(1, 20).ToString());
            WaitForLoad();
        }
        public bool VerifyDateExpedition(string expeditionDateStr, DateTime dateFrom, DateTime dateTo)
        {
            DateTime expeditionDate = DateTime.ParseExact(expeditionDateStr, "dd/MM/yyyy", CultureInfo.InvariantCulture);

            if (expeditionDate >= dateFrom && expeditionDate <= dateTo)
            {
                return true;
            }
            return false;
        }

        public bool VerifyRedirectionToDatasheetPage(string ongletTitle)
        {
            int initialWindowCount = _webDriver.WindowHandles.Count;
            var datasheet = WaitForElementIsVisible(By.Id(DATA_SHEET));
            datasheet.Click();
            WaitForLoad();            
            if (_webDriver.WindowHandles.Count > initialWindowCount )
            {
                // switch driver to the opened tab
                var wait = new WebDriverWait(_webDriver, TimeSpan.FromSeconds(60));
                wait.Until((driver) => driver.WindowHandles.Count > 1);
                _webDriver.SwitchTo().Window(_webDriver.WindowHandles[1]);
                WaitForLoad();
                string newTabTitle = _webDriver.Title;
                if (newTabTitle.Equals(ongletTitle)) 
                {
                    return true;
                }                
            }
            return false;
        }
        public void AddNewPRCOWithQty(DateTime? date, bool isActivate = true, string serviceName = null)
        {
            if(serviceName!=null)
            {
                 _searchServiceName = WaitForElementIsVisible(By.XPath(SEARCH_SERVICE_NAME));
                _searchServiceName.SendKeys(serviceName);
                WaitForLoad();
               

            }
           
            _check_button = WaitForElementIsVisible(By.Id(CHECK_BUTTON));
            _check_button.SetValue(ControlType.CheckBox, isActivate);
            WaitForLoad();
            Thread.Sleep(4000);
            _next = WaitForElementIsVisible(By.Id(NEXT_BUTTON));
            bool isDisplayed = _next.Enabled&&isActivate;
            _next.Click();
            WaitForLoad();
            WaitPageLoading();

            _new_date_button = WaitForElementIsVisible(By.Id(NEW_DATE_BUTTON));
            _new_date_button.Click();
              WaitForLoad();

            _date = WaitForElementIsVisible(By.Id(DATE));
            _date.SetValue(ControlType.DateTime, date);
            _date.SendKeys(Keys.Tab);
             WaitForLoad();

        }
        public int GetPlannedQty()
        {
            _plannedqty = WaitForElementIsVisible(By.XPath(PLANNED_QTY));     
            return (int.Parse(_plannedqty.Text));
        }
        public int GetDiffQty()
        {
            _diffqty = WaitForElementIsVisible(By.XPath(DIFF_QTY));
            return (int.Parse(_diffqty.Text));
        }
        public void SetNewQtyForNewDate(string qty)
        {            
            var setquantity = WaitForElementIsVisible(By.XPath("//*[@id=\"services-list\"]/div[2]/div[2]/div/div/div[2]/input"));
            setquantity.SetValue(ControlType.TextBox, qty);
        }
        public void ClickDuplicateNewDate()
        {
            _duplicate = WaitForElementIsVisible(By.XPath(DUPLICATE));
            _duplicate.Click();
            WaitForLoad();
        }
        public void ClickSupprimerNewDate()
        {
            _corbeille = WaitForElementIsVisible(By.XPath(CORBEILLE));
            _corbeille.Click();
            WaitForLoad();
        }
        public bool VerifyDuplicateInputDate()
        {
            var elements_input_date = _webDriver.FindElements(By.XPath("//*[@id=\"services-list\"]/div[2]/div[2]/div/div/div[1]/div/div[1]/div/input"));

            if (elements_input_date.Count == 0)
                return false;

            if (elements_input_date.Count == 2)
            {
                foreach (var elm in elements_input_date)
                {
                    if (!elm.GetAttribute("value").Equals(DateTime.Now.ToString("dd/MM/yyyy")))
                    {
                        return false;
                    }
                }
            }
            return true;
        }
        public bool VerifyDuplicateInputQuantity(string qty)
        {
            var elements_input_qty = _webDriver.FindElements(By.XPath("//*[@id=\"services-list\"]/div[2]/div[2]/div/div/div[2]/input"));
            if (elements_input_qty.Count == 0)
                return false;
            if (elements_input_qty.Count == 2)
            {
                foreach (var elm in elements_input_qty)
                {
                    if (!elm.GetAttribute("value").Equals(qty))
                    {
                        return false;
                    }
                }
            }

            return true;
        }
        public bool VerifySupprimerDuplicate()
        {
            var elements_input_qty = _webDriver.FindElements(By.XPath("//*[@id=\"services-list\"]/div[2]/div[2]/div[1]/div[*]"));
            if (elements_input_qty.Count == 2 && isElementVisible(By.XPath("//*[@id=\"services-list\"]/div[2]/div[2]/div[1]/div[2]")))
                return false; 
            return true;
        }
        public void OpenProductionCOPopUp(DateTime? date, bool isActivate = true)
        {
            Random rnd = new Random();        

            _checkAllButton = WaitForElementExists(By.Id(CHECK_ALL_BUTTON));
            _checkAllButton.Click();

            _next = WaitForElementIsVisible(By.Id(NEXT_BUTTON));
            _next.Click();

            _new_date_button = WaitForElementIsVisible(By.Id(NEW_DATE_BUTTON));
            _new_date_button.Click();
        }
        public bool IsStyleAffiche()
        {

            _style = WaitForElementIsVisible(By.XPath(STYLE));
            _modal = WaitForElementIsVisible(By.XPath(MODAL));

            if (_style != null || _modal != null) 
            {
                return true;
            }

            else return false;
        }
        public List<string> GetAllServiceToProduce()
        {
            var allServices = _webDriver.FindElements(By.XPath(ALL_SERVICES_NAME));   
            List<string> services = new List<string>(); 

            foreach (var service in allServices)
            {
                services.Add(service.Text);
            }
            return services;
        }
        public bool IsCustomerOrderPresent(string name, string category, int quantity)
        {
            // Locate the rows in the search result table
            var rows = _webDriver.FindElements(By.XPath("//*[@id=\"service-details-form\"]/table/tbody/tr"));

            foreach (var row in rows)
            {
                var nameText = row.FindElement(By.XPath(".//td[1]")).Text;
                var categoryText = row.FindElement(By.XPath(".//td[2]")).Text;
                var quantityText = row.FindElement(By.XPath(".//td[3]")).Text;

                // Compare the details to ensure the order is not present
                if (nameText.Equals(name, StringComparison.OrdinalIgnoreCase) &&
                    categoryText.Equals(category, StringComparison.OrdinalIgnoreCase) &&
                    quantityText == quantity.ToString())
                {
                    return true;  // Order found, test will fail
                }
            }
            return false;  
        }
        public bool VerifierPresenceService(string nomServiceCherche)
        {
            var services = _webDriver.FindElements(By.XPath(SERVICE_NAME_LIST));

            foreach (var service in services)
            {
                if (service.Text.Equals(nomServiceCherche))
                {
                    return true;
                }
            }

            return false;
        }

        public void CancelBtn()
        {
           var _btnClose = WaitForElementIsVisible(By.XPath(CLOSE));
            _btnClose.Click();
            WaitPageLoading();
        }
    }
}
