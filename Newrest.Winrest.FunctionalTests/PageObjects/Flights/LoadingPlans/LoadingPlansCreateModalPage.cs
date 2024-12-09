using Microsoft.VisualStudio.TestTools.UnitTesting;
using Newrest.Winrest.FunctionalTests.PageObjects.Shared;
using Newrest.Winrest.FunctionalTests.Utils;
using OpenQA.Selenium;
using OpenQA.Selenium.Support.PageObjects;
using System;

namespace Newrest.Winrest.FunctionalTests.PageObjects.Flights.LoadingPlans
{
    public class LoadingPlansCreateModalPage : PageBase
    {
        public LoadingPlansCreateModalPage(IWebDriver webDriver, TestContext testContext) : base(webDriver, testContext)
        {
        }
        // ____________________________________ Constantes __________________________________________

        private const string LP_NAME = "Name";
        private const string SITE = "dropdown-site";
        private const string CUSTOMER = "dropdown-customer-selectized";
        private const string FIRST_CUSTOMER = "//*[@id=\"createFormVipOrder\"]/div[1]/div[1]/table/tbody/tr[2]/td[2]/div/div[2]/div/div[1]";
        private const string FROM_HEURES = "//*[@id=\"createFormVipOrder\"]/div[1]/div[1]/table/tbody/tr[4]/td[2]/input[2]";
        private const string FROM_MINUTES = "//*[@id=\"createFormVipOrder\"]/div[1]/div[1]/table/tbody/tr[4]/td[2]/input[3]";
        private const string TO_HEURES = "//*[@id=\"createFormVipOrder\"]/div[1]/div[1]/table/tbody/tr[4]/td[4]/input[2]";
        private const string TO_MINUTES = "//*[@id=\"createFormVipOrder\"]/div[1]/div[1]/table/tbody/tr[4]/td[4]/input[3]";
        private const string ROUTE = "//*[@id=\"div-routes\"]/div[1]/div/span/input[2]";
        private const string FIRST_ROUTE = "//*[@id=\"div-routes\"]/div[1]/div/span/div/div/div";
        private const string REGISTRATION = "//*[@id=\"div-registration-types\"]/div[1]/div/span/input[2]";
        private const string FIRST_REGISTRATION = "//*[@id=\"div-registration-types\"]/div[1]/div/span/div/div/div";
        private const string CREATE = "//*[@id=\"createFormVipOrder\"]/div[2]/div/button";
        private const string END_DATE = "popup-end-date-picker";
        private const string START_DATE = "popup-start-date-picker";

        // ____________________________________ Variables ___________________________________________

        [FindsBy(How = How.Id, Using = LP_NAME)]
        private IWebElement _loadingPlanName;

        [FindsBy(How = How.Id, Using = SITE)]
        private IWebElement _site;

        [FindsBy(How = How.Id, Using = CUSTOMER)]
        private IWebElement _customer;

        [FindsBy(How = How.XPath, Using = FIRST_CUSTOMER)]
        private IWebElement _firstCustomer;

        [FindsBy(How = How.XPath, Using = FROM_HEURES)]
        private IWebElement _edtFromHours;

        [FindsBy(How = How.XPath, Using = FROM_MINUTES)]
        private IWebElement _edtFromMinutes;

        [FindsBy(How = How.XPath, Using = TO_HEURES)]
        private IWebElement _edtToHours;

        [FindsBy(How = How.XPath, Using = TO_MINUTES)]
        private IWebElement _edtToMinutes;

        [FindsBy(How = How.XPath, Using = ROUTE)]
        private IWebElement _route;

        [FindsBy(How = How.XPath, Using = FIRST_ROUTE)]
        private IWebElement _firstRoute;

        [FindsBy(How = How.XPath, Using = REGISTRATION)]
        private IWebElement _aircraftRegistration;

        [FindsBy(How = How.XPath, Using = FIRST_REGISTRATION)]
        private IWebElement _firstAircraft;

        [FindsBy(How = How.Id, Using = END_DATE)]
        private IWebElement _endDate;

        [FindsBy(How = How.Id, Using = START_DATE)]
        private IWebElement _startDate;

        [FindsBy(How = How.XPath, Using = CREATE)]
        private IWebElement _createBtn;

        // ________________________________Pages______________________________________________________


        public void FillField_CreateNewLoadingPlan(string loadingPlanName, string customer, string route, string aircraftRegistration, string site, string type = null, string fromHeures = "00", string fromMinutes = "00", string toHeures = "23", string toMinutes = "59")
        {
            // Définition du nom du loading plan
            _loadingPlanName = WaitForElementIsVisible(By.Id(LP_NAME));
            _loadingPlanName.SetValue(ControlType.TextBox, loadingPlanName);
            WaitForLoad();

            // Définition du site
            _site = WaitForElementIsVisible(By.Id(SITE));
            _site.SetValue(ControlType.DropDownList, site);
            WaitForLoad();

            // Définition du Customer
            _customer = WaitForElementIsVisible(By.Id(CUSTOMER));
            _customer.SetValue(ControlType.TextBox, customer);
            _customer.SendKeys(Keys.Tab);
            WaitForLoad();

            // Définition du type
            if (type != null)
            {
                var typeOflp = WaitForElementIsVisible(By.XPath("/html/body/div[3]/div/div/div[2]/div/div/form/div[1]/div[1]/div[1]/table/tbody/tr[2]/td[4]/select"));
                typeOflp.SetValue(ControlType.DropDownList, type);
                WaitForLoad();
            }


            // Définition du ETD From
            _edtFromHours = WaitForElementIsVisible(By.XPath(FROM_HEURES));
            _edtFromHours.SendKeys(Keys.ArrowRight);
            _edtFromHours.SendKeys(Keys.ArrowRight);
            _edtFromHours.SendKeys(Keys.Backspace);
            _edtFromHours.SendKeys(Keys.Backspace);
            _edtFromHours.SendKeys(fromHeures);
            WaitForLoad();

            _edtFromMinutes = WaitForElementIsVisible(By.XPath(FROM_MINUTES));
            _edtFromMinutes.SendKeys(Keys.ArrowRight);
            _edtFromMinutes.SendKeys(Keys.ArrowRight);
            _edtFromMinutes.SendKeys(Keys.Backspace);
            _edtFromMinutes.SendKeys(Keys.Backspace);
            _edtFromMinutes.SendKeys(fromMinutes);
            WaitForLoad();

            // Définition du ETD To
            _edtToHours = WaitForElementIsVisible(By.XPath(TO_HEURES));
            _edtToHours.SendKeys(Keys.ArrowRight);
            _edtToHours.SendKeys(Keys.ArrowRight);
            _edtToHours.SendKeys(Keys.Backspace);
            _edtToHours.SendKeys(Keys.Backspace);
            _edtToHours.SendKeys(toHeures);
            WaitForLoad();

            _edtToMinutes = WaitForElementIsVisible(By.XPath(TO_MINUTES));
            _edtToMinutes.SendKeys(Keys.ArrowRight);
            _edtToMinutes.SendKeys(Keys.ArrowRight);
            _edtToMinutes.SendKeys(Keys.Backspace);
            _edtToMinutes.SendKeys(Keys.Backspace);
            _edtToMinutes.SendKeys(toMinutes);
            WaitForLoad();

            // Définition de l'attribut "Routes"
            // Définition du type
            if (route != null)
            {
                _route = WaitForElementIsVisible(By.XPath(ROUTE));
                _route.SetValue(ControlType.TextBox, route);
                WaitForLoad();

                try
                {
                    _firstRoute = WaitForElementIsVisible(By.XPath(FIRST_ROUTE));
                    _firstRoute.Click();
                    WaitForLoad();
                }
                catch (Exception)
                {
                    //try to create it if it does not exist
                    var addRouteBtn = WaitForElementExists(By.XPath("//*[@id=\"div-routes\"]/div[1]/div/a"));
                    addRouteBtn.Click();
                    WaitForLoad();
                }
            }
            // Définition de l'attribut "Aircraft & Registration"
            _aircraftRegistration = WaitForElementIsVisible(By.XPath(REGISTRATION));
            _aircraftRegistration.SetValue(ControlType.TextBox, aircraftRegistration);
            WaitForLoad();

            _firstAircraft = WaitForElementIsVisible(By.XPath(FIRST_REGISTRATION));
            _firstAircraft.Click();
            WaitForLoad();

        }

        public LoadingPlansDetailsPage Create()
        {
            _createBtn = WaitForElementToBeClickable(By.XPath(CREATE));
            _createBtn.Click();
            WaitForLoad();

            return new LoadingPlansDetailsPage(_webDriver, _testContext);
        }

        public void WaiForLoad()
        {
            base.WaitForLoad();
        }

        public void FillFieldLoadingPlanInformations(DateTime endDate, DateTime? startDate = null)
        {
            if (startDate != null)
            {
                _startDate = WaitForElementIsVisible(By.Id(START_DATE));
                _startDate.SetValue(ControlType.DateTime, startDate);
                _startDate.SendKeys(Keys.Enter);
            }
            // On modifie le champ End Date
            _endDate = WaitForElementIsVisible(By.Id(END_DATE));
            _endDate.SetValue(ControlType.DateTime, endDate);
            _endDate.SendKeys(Keys.Enter);

            WaitForLoad();
        }
    }
}
