using Microsoft.VisualStudio.TestTools.UnitTesting;
using Newrest.Winrest.FunctionalTests.PageObjects.Shared;
using Newrest.Winrest.FunctionalTests.Utils;
using OpenQA.Selenium;
using OpenQA.Selenium.Support.PageObjects;
using OpenQA.Selenium.Support.UI;
using System;

namespace Newrest.Winrest.FunctionalTests.PageObjects.Customers.DeliveryRound
{
    public class DeliveryRoundCreateModalpage : PageBase
    {
        public DeliveryRoundCreateModalpage(IWebDriver webDriver, TestContext testContext) : base(webDriver, testContext)
        {
        }

        //__________________________________Constantes_______________________________________

        private const string DELIVERY_ROUND_NAME = "first";
        private const string SITE = "SiteId";
        private const string START_DATE = "StartDate";
        private const string END_DATE = "EndDate";


        private const string MONDAY_BTN = "//label[@for='Monday']";
        private const string TUESDAY_BTN = "//label[@for='Tuesday']";
        private const string WEDNESDAY_BTN = "//label[@for='Wednesday']";
        private const string THURSDAY_BTN = "//label[@for='Thursday']";
        private const string FRIDAY_BTN = "//label[@for='Friday']";
        private const string SATURDAY_BTN = "//label[@for='Saturday']";
        private const string SUNDAY_BTN = "//label[@for='Sunday']";

        private const string SAVE_BTN = "last";

        private const string ERROR_MESSAGE_NAME_REQUIRED = "//*[@id=\"deliveryround-filter-form\"]/div/div[1]/div[1]/div/div[1]/div/span";

        //__________________________________Variables_______________________________________

        [FindsBy(How = How.Id, Using = DELIVERY_ROUND_NAME)]
        private IWebElement _deliveryRoundName;
        
        [FindsBy(How = How.XPath, Using = SITE)]
        private IWebElement _site;

        [FindsBy(How = How.XPath, Using = START_DATE)]
        private IWebElement _startDate;

        [FindsBy(How = How.XPath, Using = END_DATE)]
        private IWebElement _endDate;

        [FindsBy(How = How.XPath, Using = MONDAY_BTN)]
        private IWebElement _mondayBtn;
        
        [FindsBy(How = How.XPath, Using = TUESDAY_BTN)]
        private IWebElement _tuesdayBtn;
        
        [FindsBy(How = How.XPath, Using = WEDNESDAY_BTN)]
        private IWebElement _wednesdayBtn;
        
        [FindsBy(How = How.XPath, Using = THURSDAY_BTN)]
        private IWebElement _thursdayBtn;
        
        [FindsBy(How = How.XPath, Using = FRIDAY_BTN)]
        private IWebElement _fridayBtn;

        [FindsBy(How = How.XPath, Using = SATURDAY_BTN)]
        private IWebElement _saturdayBtn;
        
        [FindsBy(How = How.XPath, Using = SUNDAY_BTN)]
        private IWebElement _sundayBtn;

        [FindsBy(How = How.Id, Using = SAVE_BTN)]
        private IWebElement _saveBtn;


        [FindsBy(How = How.XPath, Using = ERROR_MESSAGE_NAME_REQUIRED)]
        private IWebElement _errorMessage;

        
        // _______________________________________________ Méthodes _____________________________________________________________

        public DeliveryRoundGeneralInformationPage FillField_CreateNewDeliveryRound(string name, string site, DateTime startDateTime, DateTime endDateTime)
        {
            // Définition du nom
            _deliveryRoundName = WaitForElementIsVisible(By.Id(DELIVERY_ROUND_NAME));
            _deliveryRoundName.SetValue(ControlType.TextBox, name);

            // Définition du site
            _site = WaitForElementIsVisible(By.Id(SITE));
            //_site.Click();
            SelectElement select = new SelectElement(_site);
            select.SelectByText((string)site, true);
            //_site.SetValue(ControlType.DropDownList, site);
            //_site.Click();

            // Définition de Start Date
            _startDate = WaitForElementIsVisible(By.Id(START_DATE));
            _startDate.SetValue(ControlType.DateTime, startDateTime);
            _startDate.SendKeys(Keys.Tab);

            // Définition de End Date
            _endDate = WaitForElementIsVisible(By.Id(END_DATE));
            _endDate.SetValue(ControlType.DateTime, endDateTime);
            _endDate.SendKeys(Keys.Tab);

            // Définition des jours de livraison
            _tuesdayBtn = WaitForElementIsVisible(By.XPath(TUESDAY_BTN));
            _tuesdayBtn.Click();

            _saturdayBtn = WaitForElementIsVisible(By.XPath(SATURDAY_BTN));
            _saturdayBtn.Click();

            // Click sur le bouton "Create"
            Save();

            return new DeliveryRoundGeneralInformationPage(_webDriver, _testContext);
        }

        public DeliveryRoundGeneralInformationPage FillField_CreateNewDeliveryRoundForAllDays(string name, string site, DateTime startDateTime, DateTime endDateTime)
        {
            // Définition du nom
            _deliveryRoundName = WaitForElementIsVisible(By.Id(DELIVERY_ROUND_NAME));
            _deliveryRoundName.SetValue(ControlType.TextBox, name);

            // Définition du site
            _site = WaitForElementIsVisible(By.Id(SITE));
            _site.SetValue(ControlType.DropDownList, site);
            _site.Click();

            // Définition de Start Date
            _startDate = WaitForElementIsVisible(By.Id(START_DATE));
            _startDate.SetValue(ControlType.DateTime, startDateTime);
            _startDate.SendKeys(Keys.Tab);

            // Définition de End Date
            _endDate = WaitForElementIsVisible(By.Id(END_DATE));
            _endDate.SetValue(ControlType.DateTime, endDateTime);
            _endDate.SendKeys(Keys.Tab);

            // Définition des jours de livraison
            _mondayBtn = WaitForElementIsVisible(By.XPath(MONDAY_BTN));
            _mondayBtn.Click();

            _tuesdayBtn = WaitForElementIsVisible(By.XPath(TUESDAY_BTN));
            _tuesdayBtn.Click();
            
            _wednesdayBtn = WaitForElementIsVisible(By.XPath(WEDNESDAY_BTN));
            _wednesdayBtn.Click();
            
            _thursdayBtn = WaitForElementIsVisible(By.XPath(THURSDAY_BTN));
            _thursdayBtn.Click();
            
            _fridayBtn = WaitForElementIsVisible(By.XPath(FRIDAY_BTN));
            _fridayBtn.Click();

            _saturdayBtn = WaitForElementIsVisible(By.XPath(SATURDAY_BTN));
            _saturdayBtn.Click();
            
            _sundayBtn = WaitForElementIsVisible(By.XPath(SUNDAY_BTN));
            _sundayBtn.Click();

            // Click sur le bouton "Create"
            Save();

            return new DeliveryRoundGeneralInformationPage(_webDriver, _testContext);
        }

        public void Save()
        {
            _saveBtn = WaitForElementIsVisible(By.Id(SAVE_BTN));
            _saveBtn.Click();
            WaitPageLoading();
            WaitForLoad();
        }

        public bool ErrorMessageNameRequired()
        {
            _errorMessage = WaitForElementIsVisible(By.XPath(ERROR_MESSAGE_NAME_REQUIRED));

            if (_errorMessage.Text == "Name is required")
            {
                return true;
            }

            return false;
        }

        public bool ErrorMessageNameAlreadyExists()
        {
           _errorMessage = WaitForElementIsVisible(By.XPath(ERROR_MESSAGE_NAME_REQUIRED));

            if (_errorMessage.Text == "A Delivery round with the same name already exists.")
            {
                return true;
            }

            return false;
        }
    }
}