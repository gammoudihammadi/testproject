using Microsoft.VisualStudio.TestTools.UnitTesting;
using Newrest.Winrest.FunctionalTests.PageObjects.Shared;
using Newrest.Winrest.FunctionalTests.Utils;
using OpenQA.Selenium;
using OpenQA.Selenium.Support.PageObjects;

namespace Newrest.Winrest.FunctionalTests.PageObjects.Customers.Customer
{
    public class CustomerAgreementPage : PageBase
    {

        public CustomerAgreementPage(IWebDriver webDriver, TestContext testContext) : base(webDriver, testContext)
        {
        }

        //_______________________________________________________Constantes____________________________________________________________

        // General
        private const string NEW_AGREEMENT = "//*[@id=\"tabContentDetails\"]/div/div[1]/div/a[2]";
        private const string NUMBER = "AgreementNumber_Number";

        private const string CREATE = "//*[@id=\"modal-1\"]//button[2]";
        private const string CANCEL = "//*[@id=\"modal-1\"]//button[1]";
        private const string ERROR_MESSAGE = "//*[@id=\"modal-1\"]//form/div[1]/div[2]/div/span";

        // Tableau
        private const string EDIT_AGREEMENT = "//*[@id=\"list-item-with-action\"]/div/div[1]/div/div[3]/a[1]";
        private const string DELETE_AGREEMENT = "//*[@id=\"list-item-with-action\"]/div/div[1]/div/div[3]/a[2]";
        private const string CONFIRM_DELETE = "/html/body/div[11]/div/div/div[3]/a[1]";

        private const string FIRST_AGREEMENT = "//*[@id=\"list-item-with-action\"]/div/div[1]/div/div[2]";
        private const string AGREEMENT = "//*[@id=\"list-item-with-action\"]/div/div[1]/div/div[1]";

        //_______________________________________________________Variables_____________________________________________________________

        // General
        [FindsBy(How = How.XPath, Using = NEW_AGREEMENT)]
        private IWebElement _newAgreement;

        [FindsBy(How = How.Id, Using = NUMBER)]
        private IWebElement _agreementNumber;

        [FindsBy(How = How.XPath, Using = CREATE)]
        private IWebElement _createBtn;

        [FindsBy(How = How.XPath, Using = CANCEL)]
        private IWebElement _cancelBtn;


        // Tableau
        [FindsBy(How = How.XPath, Using = EDIT_AGREEMENT)]
        private IWebElement _editAgreement;

        [FindsBy(How = How.XPath, Using = DELETE_AGREEMENT)]
        private IWebElement _deleteAgreement;

        [FindsBy(How = How.XPath, Using = CONFIRM_DELETE)]
        private IWebElement _confirmDelete;

        [FindsBy(How = How.XPath, Using = FIRST_AGREEMENT)]
        private IWebElement _firstAgreement;

        //_______________________________________ Méthodes______________________________________

        // General
        public void CreateAgreement(string agreementNumber)
        {
            _newAgreement = WaitForElementToBeClickable(By.XPath(NEW_AGREEMENT));
            _newAgreement.Click();
            WaitForLoad();

            _agreementNumber = WaitForElementIsVisible(By.Id(NUMBER));
            _agreementNumber.SetValue(ControlType.TextBox, agreementNumber);
            WaitForLoad();

            _createBtn = WaitForElementIsVisible(By.XPath(CREATE));
            _createBtn.Click();
            WaitForLoad();
        }

        public void Cancel()
        {
            _cancelBtn = WaitForElementIsVisible(By.XPath(CANCEL));
            _cancelBtn.Click();
            WaitForLoad();
        }

        public bool IsErrorMessageDisplayed()
        {
            if (isElementVisible(By.XPath(ERROR_MESSAGE)))
            {
                _webDriver.FindElement(By.XPath(ERROR_MESSAGE));
                return true;
            }
            else
            {
                return false;
            }

        }

        // Tableau
        public void EditAgreement(string agreementNumber)
        {
            _editAgreement = WaitForElementIsVisible(By.XPath(EDIT_AGREEMENT));
            _editAgreement.Click();
            WaitForLoad();

            _agreementNumber = WaitForElementIsVisible(By.Id(NUMBER));
            _agreementNumber.SetValue(ControlType.TextBox, agreementNumber);
            WaitForLoad();

            _createBtn = WaitForElementIsVisible(By.XPath(CREATE));
            _createBtn.Click();
            WaitForLoad();
        }

        public void ClickFirstAgreement()
        {
            _firstAgreement = WaitForElementIsVisible(By.XPath(FIRST_AGREEMENT));
            _firstAgreement.Click();
            WaitPageLoading();
        }

        public string GetFirstAgreementNumber()
        {
            _firstAgreement = WaitForElementIsVisible(By.XPath(FIRST_AGREEMENT));

            return _firstAgreement.Text;
        }

        public void DeleteAgreement()
        {
            _deleteAgreement = WaitForElementIsVisible(By.XPath(DELETE_AGREEMENT));
            _deleteAgreement.Click();
            WaitForLoad();

            _confirmDelete = WaitForElementToBeClickable(By.XPath(CONFIRM_DELETE));
            _confirmDelete.Click();

            WaitForLoad();
        }

        public bool IsAgreementDisplayed()
        {
            bool valueBool = true;

            if(isElementVisible(By.XPath(AGREEMENT)))
            {
                _webDriver.FindElement(By.XPath(AGREEMENT));
            }
            else
            {
                valueBool = false;
            }

            return valueBool;
        }
    }
}
