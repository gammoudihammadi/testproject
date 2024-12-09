using Microsoft.VisualStudio.TestTools.UnitTesting;
using Newrest.Winrest.FunctionalTests.PageObjects.Shared;
using Newrest.Winrest.FunctionalTests.Utils;
using OpenQA.Selenium;
using OpenQA.Selenium.Support.PageObjects;
using System;
using System.Linq;

namespace Newrest.Winrest.FunctionalTests.PageObjects.Flights.LpCart
{
    public class LpCartCreateModalPage : PageBase
    {


        // ____________________________________ Constantes __________________________________________

        private const string CODE = "tb-new-lpcart-number";
        private const string NAME = "tb-new-lpcart-name";
        private const string SITE = "selectSite";
        private const string CUSTOMER = "selectCustomer";
        private const string AIRCRAFT = "drop-down-suppliers";
        private const string DATE_FROM = "datapicker-new-lpcart-from";
        private const string DATE_TO = "datapicker-new-lpcart-to";
        private const string COMMENT = "Comment";
        private const string CREATE = "last";

        private const string ROUTES = "/html/body/div[3]/div/div/div/div/form/div[2]/div[5]/div/div/div/span/input[2]";
        private const string ADD_ROUTE = "/html/body/div[3]/div/div/div/div/form/div[2]/div[5]/div/div/div/a";
        private const string MODAL_FORM = "//*[@id=\"tabContentItemContainer\"]/div/h1";

        // ____________________________________ Variables ___________________________________________

        [FindsBy(How = How.Id, Using = CODE)]
        private IWebElement _code;

        [FindsBy(How = How.Id, Using = NAME)]
        private IWebElement _name;

        [FindsBy(How = How.Id, Using = SITE)]
        private IWebElement _site;
        
        [FindsBy(How = How.Id, Using = CUSTOMER)]
        private IWebElement _customer;

        [FindsBy(How = How.Id, Using = AIRCRAFT)]
        private IWebElement _aircraft;

        [FindsBy(How = How.Id, Using = DATE_FROM)]
        private IWebElement _dateFrom;

        [FindsBy(How = How.Id, Using = DATE_TO)]
        private IWebElement _dateTo;

        [FindsBy(How = How.Id, Using = COMMENT)]
        private IWebElement _comment;

        [FindsBy(How = How.Id, Using = CREATE)]
        private IWebElement _createBtn;

        //___________________________________ Méthodes ___________________________________________
        public LpCartCreateModalPage(IWebDriver webDriver, TestContext testContext) : base(webDriver, testContext)
        {
        }

        public LpCartCartDetailPage FillField_CreateNewLpCart(string lpCartCode, string lpCartName, string site, string customer, string aircraft, DateTime from, DateTime to, string comment)
        {
            // Définition du nom du lpCode
            _code = WaitForElementIsVisible(By.Id(CODE));
            _code.SetValue(ControlType.TextBox, lpCartCode);
            WaitForLoad();
            Assert.IsTrue(_code.GetAttribute("value") == lpCartCode , "code not changed");
            // Définition du nom du lpCartName
            _name = WaitForElementIsVisible(By.Id(NAME));
            _name.SetValue(ControlType.TextBox, lpCartName);
            Assert.IsTrue(_name.GetAttribute("value") == lpCartName, "name not changed");
            // Définition du site
            _site = WaitForElementIsVisible(By.Id(SITE));
            _site.SetValue(ControlType.DropDownList, site);
            WaitForLoad();
            var selectedSite = _webDriver.FindElements(By.XPath("//*[@id=\"selectSite\"]/option")).Where(e=>e.Selected).FirstOrDefault();
            Assert.IsTrue(selectedSite.Text == site, "site not changed");
            // Définition du Customer
            _customer = WaitForElementIsVisible(By.Id(CUSTOMER));
            _customer.SetValue(ControlType.DropDownList, customer);
            WaitForLoad();
            var selectedCustomer = _webDriver.FindElements(By.XPath("//*[@id=\"selectCustomer\"]/option")).Where(e => e.Selected).FirstOrDefault();
            Assert.IsTrue(selectedCustomer.Text == customer, "customer not changed");
            // Définition du aircraft
            _aircraft = WaitForElementIsVisible(By.XPath("//*/input[contains(@class,'new-aircraft') and contains(@class,'tt-input')]"));
            _aircraft.SetValue(ControlType.TextBox, aircraft);
            var choice = WaitForElementIsVisible(By.XPath("//*/div[@class='tt-suggestion tt-selectable']"));
            choice.Click();
            WaitForLoad();

            // Dating
            _dateFrom = WaitForElementIsVisible(By.Id(DATE_FROM));
            _dateFrom.SetValue(ControlType.DateTime, from);
            _dateFrom.SendKeys(Keys.Tab);
            WaitForLoad();

            _dateTo = WaitForElementIsVisible(By.Id(DATE_TO));
            _dateTo.SetValue(ControlType.DateTime, to);
            _dateTo.SendKeys(Keys.Tab);
            WaitForLoad();

            _comment = WaitForElementIsVisible(By.Id(COMMENT));
            _comment.SetValue(ControlType.TextBox, comment);
            _comment.SendKeys(Keys.Tab);
            WaitForLoad();

            _createBtn = WaitForElementToBeClickable(By.Id(CREATE));
            _createBtn.Click();
            WaitLoading();
            WaitPageLoading();
            var modalIsClosed = isElementVisible(By.Id("modal-1"));
            Assert.IsFalse(modalIsClosed, "Modal is still showing even after creation");
            return new LpCartCartDetailPage(_webDriver, _testContext);

        }


        public LpCartCartDetailPage FillField_CreateNewLpCartWithRoutes(string lpCartCode, string lpCartName, string site, string customer, string aircraft, DateTime from, DateTime to, string comment,string route)
        {
            // Définition du nom du lpCode
            _code = WaitForElementIsVisible(By.Id(CODE));
            _code.SetValue(ControlType.TextBox, lpCartCode);

            // Définition du nom du lpCartName
            _name = WaitForElementIsVisible(By.Id(NAME));
            _name.SetValue(ControlType.TextBox, lpCartName);

            // Définition du site
            _site = WaitForElementIsVisible(By.Id(SITE));
            _site.SetValue(ControlType.DropDownList, site);
            WaitForLoad();

            // Définition du Customer
            _customer = WaitForElementIsVisible(By.Id(CUSTOMER));
            _customer.SetValue(ControlType.DropDownList, customer);
            WaitForLoad();

            // Définition du aircraft
            _aircraft = WaitForElementIsVisible(By.XPath("//*/input[contains(@class,'new-aircraft') and contains(@class,'tt-input')]"));
            _aircraft.SetValue(ControlType.TextBox, aircraft);
            var choice = WaitForElementIsVisible(By.XPath("//*/div[@class='tt-suggestion tt-selectable']"));
            choice.Click();
            WaitForLoad();

            // Dating
            _dateFrom = WaitForElementIsVisible(By.Id(DATE_FROM));
            _dateFrom.SetValue(ControlType.DateTime, from);
            _dateFrom.SendKeys(Keys.Tab);
            WaitForLoad();

            _dateTo = WaitForElementIsVisible(By.Id(DATE_TO));
            _dateTo.SetValue(ControlType.DateTime, to);
            _dateTo.SendKeys(Keys.Tab);
            WaitForLoad();

            //comment
            _comment = WaitForElementIsVisible(By.Id(COMMENT));
            _comment.SetValue(ControlType.TextBox, comment);
            _comment.SendKeys(Keys.Tab);
            WaitForLoad();

            //afffect routes
            var routes = WaitForElementIsVisible(By.XPath(ROUTES));
            routes.SendKeys(route);
            var addRoute = WaitForElementIsVisible(By.XPath(ADD_ROUTE));
            addRoute.Click();
            WaitForLoad();
            //save
            _createBtn = WaitForElementToBeClickable(By.Id(CREATE));
            _createBtn.Click();
            WaitForLoad();

            return new LpCartCartDetailPage(_webDriver, _testContext);

        }

        public bool IsFormPopupClosed()
        {
            var _modalForm = _webDriver.FindElement(By.XPath(MODAL_FORM));
            return _modalForm.Displayed;
        }

    }
}
