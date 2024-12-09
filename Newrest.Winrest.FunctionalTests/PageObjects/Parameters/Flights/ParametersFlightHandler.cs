using Microsoft.VisualStudio.TestTools.UnitTesting;
using Newrest.Winrest.FunctionalTests.PageObjects.Shared;
using Newrest.Winrest.FunctionalTests.Utils;
using OpenQA.Selenium;
using OpenQA.Selenium.Support.PageObjects;

namespace Newrest.Winrest.FunctionalTests.PageObjects.Parameters.Flights
{
    public class ParametersFlightHandler : PageBase
    {
        private const string ADD_NEW_BUTTON = "parameters-handler-new";
        private const string HANDLER_NAME_INPUT = "first";
        private const string SAVE_BUTTON = "last";
        private const string SAVE_ADD_NEW_BUTTON = "SaveAndNew";
        private const string DELETE_BUTTON = "//*[@id=\"tabContentParameters\"]/table/tbody/tr[2]/td[2]/a[2]";
        private const string CONFIRM_DELETE = "first";



        [FindsBy(How = How.XPath, Using = ADD_NEW_BUTTON)]
        private IWebElement _addNewButton;

        [FindsBy(How = How.XPath, Using = SAVE_ADD_NEW_BUTTON)]
        private IWebElement _saveAddNewButton;

        [FindsBy(How = How.XPath, Using = HANDLER_NAME_INPUT)]
        private IWebElement _handlerNameInput;

        [FindsBy(How = How.Id, Using = SAVE_BUTTON)]
        private IWebElement _saveButton;

        [FindsBy(How = How.XPath, Using = DELETE_BUTTON)]
        private IWebElement _deleteButton;

        [FindsBy(How = How.Id, Using = CONFIRM_DELETE)]
        private IWebElement _confirmDelete;

        public ParametersFlightHandler(IWebDriver webDriver, TestContext testContext) : base(webDriver, testContext)
        {
        }



        public void CreateHandlerModelPage()
        {
            _addNewButton = WaitForElementIsVisible(By.Id(ADD_NEW_BUTTON));
            _addNewButton.Click();
            WaitPageLoading();
        }

        public void CreateHandlerModelPageFromPopUp()
        {
            _saveAddNewButton = WaitForElementIsVisible(By.Id(SAVE_ADD_NEW_BUTTON));
            _saveAddNewButton.Click();
            WaitPageLoading();
        }

        public void FillFieldsHandlerModelPage(string handlertName)
        {
            _handlerNameInput = WaitForElementIsVisible(By.Id(HANDLER_NAME_INPUT));
            _handlerNameInput.SetValue(ControlType.TextBox, handlertName);
            WaitForLoad();

        }

        public void Save()
        {
            _saveButton = WaitForElementIsVisible(By.Id(SAVE_BUTTON));
            _saveButton.Click();
            WaitForLoad();

        }

        public void DeleteAircraftByHandlerName(string handlerName)
        {
            string deleteXPath = $"//td[normalize-space(text())='{handlerName}']/following-sibling::td//a[contains(@href, 'DeleteAjax')]";

            var deleteButton = WaitForElementIsVisible(By.XPath(deleteXPath));

            deleteButton.Click();
            WaitForLoad();

            var confirmDelete = WaitForElementIsVisible(By.Id(CONFIRM_DELETE));
            confirmDelete.Click();
            WaitForLoad();
        }

    }
}