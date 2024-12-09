using Microsoft.VisualStudio.TestTools.UnitTesting;
using Newrest.Winrest.FunctionalTests.PageObjects.Shared;
using Newrest.Winrest.FunctionalTests.Utils;
using OpenQA.Selenium;
using OpenQA.Selenium.Support.PageObjects;


namespace Newrest.Winrest.FunctionalTests.PageObjects.Parameters.Flights
{
    public class ParametersFlightAircraft : PageBase
    {
        private const string NEW_BUTTON = "//*[@id=\"tabContentParameters\"]/div[1]/a";
        private const string AIRCRAFT_NAME_INPUT = "//*[@id=\"first\"]";
        private const string EQUIPEMENTTYPE_INPUT = "EquipementType";
        private const string SAVE_BUTTON = "last";
        private const string DELETE_BUTTON = "//*[@id=\"tabContentItemContainer\"]/div/table/tbody/tr[2]/td[4]/a[2]";
        private const string CONFIRM_DELETE = "first";
        private const string SEARCH_INPUT = "tbSearchPattern";



        [FindsBy(How = How.XPath, Using = NEW_BUTTON)]
        private IWebElement _newButton;

        [FindsBy(How = How.XPath, Using = AIRCRAFT_NAME_INPUT)]
        private IWebElement _aircraftNameInput;

        [FindsBy(How = How.Id, Using = EQUIPEMENTTYPE_INPUT)]
        private IWebElement _equipementTypeInput;

        [FindsBy(How = How.Id, Using = SAVE_BUTTON)]
        private IWebElement _saveButton;

        [FindsBy(How = How.XPath, Using = DELETE_BUTTON)]
        private IWebElement _deleteButton;

        [FindsBy(How = How.Id, Using = CONFIRM_DELETE)]
        private IWebElement _confirmDelete;

        [FindsBy(How = How.Id, Using = SEARCH_INPUT)]
        private IWebElement _searchInput;

        public ParametersFlightAircraft(IWebDriver webDriver, TestContext testContext) : base(webDriver, testContext)
        {
        }

        public void CreateAircraftModelPage()
        {
            IWebElement newAircraftButton = WaitForElementIsVisible(By.XPath(NEW_BUTTON));
            newAircraftButton.Click();
            WaitPageLoading();
        }

        public void FillFieldsAircraftModelPage(string aircraftName , string equipementype)
        {
            _aircraftNameInput = WaitForElementIsVisible(By.XPath(AIRCRAFT_NAME_INPUT));
            _aircraftNameInput.SetValue(ControlType.TextBox, aircraftName);
            WaitForLoad();

            _equipementTypeInput = WaitForElementIsVisible(By.Id(EQUIPEMENTTYPE_INPUT));
            _equipementTypeInput.SetValue(ControlType.TextBox , equipementype);
            WaitForLoad();
        }

        public void  Save()
        {
            _saveButton = WaitForElementIsVisible(By.Id(SAVE_BUTTON));
            _saveButton.Click();
            WaitForLoad();
           
        }

        public void DeleteFirstAircraft()
        {
            _deleteButton = WaitForElementIsVisible(By.XPath(DELETE_BUTTON));
            _deleteButton.Click();
            WaitForLoad();

            _confirmDelete = WaitForElementIsVisible(By.Id(CONFIRM_DELETE));
            _confirmDelete.Click();
            WaitForLoad();
        }

        public void FilterBySearch(string aircraftName)
        {
            _searchInput = WaitForElementIsVisible(By.Id(SEARCH_INPUT));
            _searchInput.Clear();
            _searchInput.SetValue(ControlType.TextBox, aircraftName);
            WaitForLoad();
        }
    }
}
