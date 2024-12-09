using Microsoft.VisualStudio.TestTools.UnitTesting;
using Newrest.Winrest.FunctionalTests.PageObjects.Shared;
using Newrest.Winrest.FunctionalTests.Utils;
using OfficeOpenXml.FormulaParsing.Excel.Functions.Text;
using OpenQA.Selenium;
using OpenQA.Selenium.Support.PageObjects;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using static System.Resources.ResXFileRef;

namespace Newrest.Winrest.FunctionalTests.PageObjects.Jobs.Settings
{
    public class FileFlowJobNotificationCreateModalPage : PageBase
    {
        private const string EMAIL = "Email";
        private const string CUSTOMER = "drop-down-customers";
        private const string SCHEDULE_JOB = "SelectedFileFlowScheduledJobIds_ms";
        private const string CREATE_BUTTON = "last";
        private const string EMAIL_VALIDATOR_MESSAGE = "//*[@id=\"item-edit-form\"]/div[2]/div[1]/div/span"; 
        private const string CUSTOMER_VALIDATOR_MESSAGE = "//*[@id=\"SelectedCustomer-Group\"]/div/span";
        private const string FIRST_ITEM_DROPDOWN_LIST = "//*[@id=\"ui-multiselect-3-SelectedFileFlowScheduledJobIds-option-0\"]";

        
        private const string SEARCH = "/html/body/div[13]/div/div/label/input";
        private const string UNSELECT_ALL = "/html/body/div[13]/div/ul/li[2]/a/span[2]";
        public FileFlowJobNotificationCreateModalPage(IWebDriver webDriver, TestContext testContext) : base(webDriver, testContext)
        {
        }
        [FindsBy(How = How.Id, Using = EMAIL)]
        private IWebElement _email;
        [FindsBy(How = How.Id, Using = CUSTOMER)]
        private IWebElement _cutomer;
        [FindsBy(How = How.Id, Using = SCHEDULE_JOB)]
        private IWebElement _scheduleJob;
        [FindsBy(How = How.Id, Using = CREATE_BUTTON)]
        private IWebElement _createButton; 
        //[FindsBy(How = How.XPath, Using = EMAIL_VALIDATOR_MESSAGE)]
        //private IWebElement _emailValidatorMessage;
        //[FindsBy(How = How.XPath, Using = CUSTOMER_VALIDATOR_MESSAGE)]
        //private IWebElement _customerValidatorMessage;
        public JobNotifications FillFieldAddNewFileFlowJobNotification(string email, string customer, string ScheduledJobs = null)
        {
            _email = WaitForElementIsVisible(By.Id(EMAIL));
            _email.SetValue(ControlType.TextBox, email);
            WaitForLoad();
            if (customer != null) 
            {
                _cutomer = WaitForElementExists(By.Id(CUSTOMER));
                _cutomer.SetValue(ControlType.DropDownList, customer);
                WaitForLoad();
            }
           
            if (ScheduledJobs != null)
            {
                _scheduleJob = WaitForElementExists(By.Id(SCHEDULE_JOB));
                _scheduleJob.Click();
                WaitForLoad();

                var unselectAll = WaitForElementIsVisible(By.XPath(UNSELECT_ALL));
                unselectAll.Click();
                WaitForLoad();

                // Localiser et cliquer sur la liste déroulante
                var dropDownList = WaitForElementIsVisible(By.XPath(SEARCH));
                dropDownList.Click();
                WaitForLoad();

                // Sélectionner et cliquer sur le premier élément dans la liste déroulante
                var firstItemInDropDown = WaitForElementIsVisible(By.XPath(FIRST_ITEM_DROPDOWN_LIST));
                firstItemInDropDown.Click();
                WaitForLoad();

            }
            SubmitForm();
            WaitForLoad();
            return new JobNotifications(_webDriver, _testContext);
        }
        public void SubmitForm()
        {
            _createButton = WaitForElementIsVisible(By.Id(CREATE_BUTTON));
            _createButton.Click();
        }
        public bool CheckValidator()
        {
            return isElementExists(By.XPath(CUSTOMER_VALIDATOR_MESSAGE)) && isElementExists(By.XPath(CUSTOMER_VALIDATOR_MESSAGE));
        }
    }
}
